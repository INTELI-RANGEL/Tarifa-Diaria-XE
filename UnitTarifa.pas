unit UnitTarifa;

interface

uses frxClass,Forms, frxDBSet,ZConnection, DB, ZAbstractRODataset, Dialogs,DBCtrls,ExtCtrls,
  ZAbstractDataset, ZDataset,math,strUtils,DateUtils,SysUtils, RxMemDS, controls,
  classes,JvMemoryDataset,DBClient,AdvProgr,cxProgressBar,ComObj,OleCtrls,
  UFunctionsGHH, UnitMetodos,Variants,dxmdaset;

type
  // Dummy class to hold event handlers:
  EventHandlers = class
    // Event handlers:
   // class procedure OnBeforeOpen(Sender: TObject);
    //class procedure OnAfterOpen(Sender: TObject);
     class procedure ActaPdf_AfterScroll(DataSet: TDataSet);
     class procedure ActaPdf_DetalleAfterScroll(DataSet: TDataSet);

  end;

Type
  FtTipo=(FtTierra,FtAbordo);
  RtRecurso=(RtPersonal=1,RtEquipo=2,RtPernocta=3,RtBarco=4,RtMaterial=5,RtExtraordinaria=6);
  //  TCCOpciones = (CCOpenAntes, CCOpenDespues, CCPostAntes, CCPostDespues, CCListaErrores, CCPrimerError);
  TSetRtRecurso = Set of RtRecurso;
  //Secciones reórte diario
  FtSeccion=(FtsCaratula,FtsActividades,FtsRMaterial,FtsBGeneral,FtsTPersonal,FtsNGenerales,FtsLPersonal,FtsCCombustible,FtsCPersonal,FtsCEquipos,
             FtsCPernoctas,FtsRRecursos, FtsRRecursos_detalle,FtsRRecursos_detalleV2,FtsFotografico,FtsAvancesCortes, FtsAll, FtsNone);
  //Tipos de generadores
  FtGenerador=(FtGBarco=1,FtGPersonal=2,FtGEquipo=3,FtGMaterial=4,FtGPernocta=5,FtGAnexoPU=6,FtTiempoExtra=7);
  //Tipos de estimaicones
  FtEstimacion=(FtEAll=0,FtEBarco=1,FtEPersonal=2,FtEEquipo=3,FtEMaterial=4,FtEPernocta=5,FtEAnexoPU=6);
  //Opciones de reportes
  FtSeccionBitacora=(FtsTareas,FtsAvancesParciales);

  Procedure EncabezadoPDF_Horizontal(Reportediario :TzREadOnlyQuery;        Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure FirmasPDF_Generales(Reportediario :TzREadOnlyQuery;             Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure ReportePDF_MovimientosLogisticos2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_ActividadesPorFolio(ReporteDiario:TzREadOnlyQuery;   Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion); overload;

  {N/A}     Procedure ReportePDF_AvancesCortes(ReporteDiario:TzREadOnlyQuery;         Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  {N/A}     Procedure ReportePDF_ResumenMaterial2(ReporteDiario:TzREadOnlyQuery;      Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);

  Procedure ReportePDF_BalanceGeneral(ReporteDiario:TzREadOnlyQuery;        Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_ActividadesPorFolio(sFiltroFolio: string; ReporteDiario:TzREadOnlyQuery;   Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion); overload;
  Procedure ReportePDF_NotasGenerales(ReporteDiario:TzREadOnlyQuery;        Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);

  {N/A}     Procedure ReportePDF_ConsumosDiesel(ReporteDiario:TzREadOnlyQuery;        Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_ConcentradoDePersonalTierra(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_ConcentradoDePersonal2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_ConcentradoDePernoctas2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_DistribucionDeEquipos2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_ConcentradoDeEquipos2(ReporteDiario:TzREadOnlyQuery; Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_PartidasAnexoC_detalleV2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
 // Procedure ReportePDF_ConcentradoDePernoctas2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_TotalDePersonal2(ReporteDiario:TzREadOnlyQuery;      Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  Procedure ReportePDF_DistribucionDeEquiposTierra(ReporteDiario:TzREadOnlyQuery; Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
 
  {N/A}     Procedure ReportePDF_Listadepersonal(ReporteDiario:TzREadOnlyQuery;       Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  {N/A}     procedure ReportePDF_ResumenDeRecursos(ReporteDiario:TzREadOnlyQuery;     Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  {N/A}     procedure ReportePDF_PartidasAnexoC(ReporteDiario:TzREadOnlyQuery;        Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  {N/A}     Procedure ReportePDF_PartidasAnexoC_detalle(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);

   procedure ReportePDF_Reportefotografico(ReporteDiario:TzREadOnlyQuery;    Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
  procedure ReportePDF_ClearDataset(Var Reporte: TfrxReport);

  {Generadores}
  Procedure TdConfiguracion(ParamContrato,ParamFolio:String;Var Reporte: TfrxReport);
  Procedure TdConfiguracionGenerador(ParamContrato,ParamFolio:String;Var Reporte: TfrxReport); overload;
  Procedure TdConfiguracionGenerador(ParamContrato,ParamFolio:String;Var Reporte: TfrxReport;FechaI,FechaT:TDateTime)  ; overload;
  Procedure NotaCampo(ParamContrato,ParamFolio:String;Var Reporte: TfrxReport);
  Procedure TdGenerador(ParamContrato:String;ParamFechaI,ParamFechaT:TDateTime;Var Reporte: TfrxReport;IsGeneral:Boolean=True); Overload;
  Procedure TdGenerador(ParamContrato:String;ParamFechaI,ParamFechaT:TDateTime;Var MryRecursos:TjvMemoryData;IsGeneral:Boolean=True;Progress:TAdvProgress=Nil); Overload;
  Procedure TdGenerador(ParamContrato:String;ParamFechaI,ParamFechaT:TDateTime;Var MryRecursos:TClientDataSet;IsGeneral:Boolean=True); Overload;
  Procedure ExportOTProject(ParamContrato,ParamConvenio,ParamFolio:String;PgbAvance: TcxProgressBar=nil);
  function  VolumenActividades(var MsProject:Variant;ListaParams:TstringList;var Ren:Integer;const ItotalReg:Integer;PgbAvance: TcxProgressBar=nil):Boolean;
  Function  TdProrrateoFolio(const ParamContrato:String;Const ParamFecha:TDatetime;Const IdDiario:Integer):Boolean;
  Function  TdAjustaFolio(const ParamContrato:String;Const ParamFecha:TDatetime):Boolean;
  Procedure TdProrrateaPernocta(const ParamContrato:String;Const ParamFecha:TDatetime);
  Procedure TdGenerador(ListaParams:TstringList;TipoGenerador:FtGenerador;Var Reporte: TfrxReport;IsMonto:Boolean=False);Overload;
  Procedure TdEstimacionCamsa(ListaParams:TstringList;TipoEstimacion:FtEstimacion;Var Reporte: TfrxReport;IsMonto:Boolean=False);
  {Actividades}
  Procedure BitacoraPDF_ActividadesTareas(ReporteDiario:TzREadOnlyQuery;   Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccionBitacora; sParamEstatus, sParamNivelImprime : string);
  Procedure BitacoraPDF_AvancesCortes(ReporteDiario:TzREadOnlyQuery;   Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccionBitacora);
  Procedure BitacoraPDF_ActividadesDetalle(ReporteDiario:TzREadOnlyQuery;   Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccionBitacora);
  {Bitacora}
  Procedure ReportePDF_ActividadesBitacora(ReporteDiario:TzREadOnlyQuery;   Var Reporte: TfrxReport);
  Procedure procReporteGenerador(sParamContrato, sParamtipo, sParamCategoria :string; lParamContinuo :boolean; dParamFechaI, dParamFechaF: tDate; tOrigen: TComponent; tProcedure: tfrxGetValueEvent; FormatosExp: string = ''; PermisosExp: string = '');
  
  {Procedimientos basicos}
  {busqueda de maximo id ->}      function Td_maximoId(sParamContrato, sParamTabla, sParamCampo :string) : Integer;
  {Eliminar o actualizar }
  {Contrato, Folio, Wbs o act ->} procedure BuscaElimina_datos(sParamTabla, sLlevaContrato, sLlevaFolio, sLlevaWbs, sLLevaAct, sParamContrato, sParamFolio, sParamWbs, sParamAct, sParamNuevoContrato, sParamNuevoFolio, sParamNuevaWbs, sParamNuevaAct : string; accion :string; lMensaje :boolean);
  {Avances acumulados ->}         function AvanceFolio(sParamContrato, sParamFolio : string; dParamFecha : tDateTime) : double;

  Procedure ActaPdf_Informacion(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport);
  Procedure ActaPdf_Actividades(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport);
  Procedure ActaPdf_CostoActividad(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure ActaEx_CostoActividad(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure ActaPdf_NotaCampo(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure ActaEx_NotaCampo(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure ActaPdf_DesgloseCostos(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure ActaEx_DesgloseCostos(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
  Procedure ActaPdf_CostoInterferencia(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);

//  RtRecurso=(RtPersonal=1,RtEquipo=2,RtPernocta=3,RtBarco=4,RtExtraordinaria=5);
//  TSetRtRecurso = Set of RtRecurso;
  Procedure Load_NotaCampo(DatosActa:TzREadOnlyQuery;Opciones:TSetRtRecurso=[RtPersonal,RtEquipo,RtPernocta,RtBarco,RtExtraordinaria];Clean:Boolean=true);
  Procedure Load_DesgloseCostos(DatosActa:TzREadOnlyQuery;Opciones:TSetRtRecurso=[RtPersonal,RtEquipo,RtPernocta,RtBarco,RtExtraordinaria];Clean:Boolean=true);
  Procedure Load_Ajuste(DatosActa:TzREadOnlyQuery;Opciones:TSetRtRecurso=[RtPersonal,RtEquipo,RtPernocta,RtBarco,RtExtraordinaria]);
 // procedure ActaPdf_AfterScroll(DataSet: TDataSet);
  const
    TotalCol=10;
    TotalColPer=11;//9;
    Mantisa=2;

	{abordo
	
	TotalCol=10;
    TotalColPer=11;
	}

  var
    MemoryTmp,
    MemoryDetalleTmp:TJvMemoryData;
    Montos:array[1..4,1..2] of Double;    // 1 personal, 2 equipo , 3 barco, 4 pernocta

implementation

uses frm_connection, global,Utilerias, UnitExcepciones;

{     class procedure ActaPdf_AfterScroll(DataSet: TDataSet);
  end;

class procedure TXMLEventHandlers.OnBeforeOpen(Sender: TObject);
begin
  MessageBox(0, PChar(ClassName + '.OnBeforeOpen'), nil, 0)
end;

class procedure TXMLEventHandlers.OnAfterOpen(Sender: TObject);
begin
  MessageBox(0, PChar(ClassName + '.OnAfterOpen'), nil, 0)
end;}
class procedure EventHandlers.ActaPdf_DetalleAfterScroll(DataSet: TDataSet);
begin
  if Assigned(MemoryDetalleTmp) then
  begin
    MemoryDetalleTmp.Filtered:=False;
    MemoryDetalleTmp.Filter:='sNumeroActividad=' + QuotedStr(DataSet.FieldByName('sNumeroActividad').AsString);
    MemoryDetalleTmp.Filtered:=True;
  end;

end;

class procedure EventHandlers.ActaPdf_AfterScroll(DataSet: TDataSet);
begin
  if Assigned(MemoryTmp) then
  begin
    MemoryTmp.Filtered:=False;
    MemoryTmp.Filter:='sNumeroActividad=' + QuotedStr(DataSet.FieldByName('sNumeroActividad').AsString);
    MemoryTmp.Filtered:=True;
  end;
end;


Procedure ActaPdf_CostoInterferencia(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
const
  SQlRef: array[1..2,1..4] of string=(('bitacoradepersonal','personal','sIdPersonal','PERSONAL'),('bitacoradeequipos','equipos','sIdEquipo','EQUIPO'));
var
  QrInterferencias,QrRecursos:TZReadOnlyQuery;
  Td_Movimientos: TfrxDBDataset;
  MDInterferencias:TdxMemData;
  i:Integer;
  dImporteMn,
  dImporteDll :Double;
begin
  QrInterferencias:=TZReadOnlyQuery.Create(nil);
  QrRecursos :=TZReadOnlyQuery.Create(nil);
  try
    Td_Movimientos:= TfrxDBDataset.Create(nil);
    Td_Movimientos.UserName:='Td_Movimientos';
    MDInterferencias:=TdxMemData.Create(nil);
    QrRecursos.Connection:=connection.zConnection;
    QrInterferencias.Connection:=connection.zConnection;
    QrInterferencias.SQL.Text:= 'select b.*,CONVERT(TIME_FORMAT(SEC_TO_TIME(Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b.didfecha,b.sHorainicio), TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ,"%H:%i") using utf8)  AS diferencia, '  +
                        'tm.sIdTipoMovimiento,tm.sDescripcion,tm.dVentaMn,tm.dVentaDll, '  +
                        '( EXTRACT( HOUR FROM sec_to_time(Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b.didfecha,b.sHorainicio), TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ) * 3600 + ' + #13#10 +
                          'EXTRACT( MINUTE FROM sec_to_time(Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b.didfecha,b.sHorainicio), TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ) * 60' + #13#10 +
                        ')  as Hora ' +
                        'from bitacoradeactividades b ' + #13#10 +
                        'inner join tiposdemovimiento tm '  +
                        'on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=b.sIdClasificacion)' +
                        'where b.sContrato=:Contrato and b.snumeroorden=:Orden ' + #13#10 +
                        'and b.sIdTipoMovimiento=:Tipo and FIND_IN_SET(tm.sIdTipoMovimiento,:Movimientos)' + #13#10 +
                        'group by tm.sIdtipoMovimiento '   +
                        'order by b.snumeroactividad,b.didfecha,time(b.sHoraInicio)' ;
    QrInterferencias.ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;

    QrInterferencias.ParamByName('contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
    QrInterferencias.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
    //QrAvance.ParamByName('Actividad').AsString:=QrActividad.FieldByName('sNumeroActividad').AsString;
    QrInterferencias.ParamByName('tipo').AsString:='ED';
    QrInterferencias.ParamByName('Movimientos').AsString:='APO,ATA,CIA,CNT,DESMOV,EAA,FAA,FP,FT,HR,IB,MOT,MT,ILG,PO,RO,RRE,RX,SI';
    QrInterferencias.Open;

    MDInterferencias.LoadFromDataSet(QrInterferencias);

    QrInterferencias.first;
    while not QrInterferencias.Eof do
    begin
      dImporteMn:=0;
      dImporteDll :=0;
      for I := 1 to Length(SQlRef) do
      begin
        QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select br.' +SQlRef[i,3] + ' as sIdRecurso,br.sDescripcion, r.sMedida, ' +
                              '(IFNULL(sum(cast(br.dCanthh, as char)), 0) + ' +
                              ' IFNULL((SELECT ' +
                              '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                              '         FROM ' +
                              '           bitacoradeajustes ba ' +
                              '         WHERE ' +
                              '           ba.sContrato = br.sContrato AND ba.sNumeroOrden = br.sNumeroOrden AND ' +
                              '           ba.sNumeroActividad = br.sNumeroActividad AND ' +
                              '           ba.sIdPartidaAnexo = br.' + SQlRef[i,3] + ' AND ba.dFecha = br.dIdFecha), 0)) as dCanthh, ' +
                              'r.dVentaMn,r.dVentaDll, ' +
                              'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) as Ajuste,' +
                              'br.didfecha ' +
                              'from '+ SQlRef[i,1]  + ' br ' +
                              'inner join bitacoradeactividades ba ' +
                              'on (ba.sContrato=br.sContrato and ba.dIdFecha=br.dIdFecha and ba.iIdDiario=br.iIdDiario and ba.sNumeroOrden=br.sNumeroOrden) ' +
                              'left join ' + SQlRef[i,2] + ' r ' +
                              'on(r.sContrato=:Contrato and br.'+SQlRef[i,3]+'=r.'+ SQlRef[i,3] +') ' +
                              'where br.sContrato=:Orden and br.sNumeroOrden=:Folio and ba.sIdClasificacion=:Clasificacion ' +
                              'group by sIdRecurso,br.didfecha order by r.iitemorden' ;
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QrRecursos.ParamByName('Clasificacion').AsString:=QrInterferencias.FieldByName('sIdClasificacion').AsString;
        QrRecursos.Open;

        while not QrRecursos.Eof do
        begin
          dImporteMn:= dImporteMn + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa) * QrRecursos.FieldByName('dVentaMn').AsFloat);
          dImporteDll:= dImporteDll + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat);
          QrRecursos.Next;
        end;
      end;
      if MDInterferencias.Locate('sIdClasificacion',QrInterferencias.FieldByName('sIdClasificacion').AsString,[]) then
      begin
        MDInterferencias.Edit;
        MDInterferencias.FieldByName('dVentaMn').AsFloat:=xRound(dImporteMn,Mantisa);
        MDInterferencias.FieldByName('dVentaDll').AsFloat:=xRound(dImporteDll,Mantisa);
        MDInterferencias.Post;
      end;

      QrInterferencias.Next;
    end;
    MDInterferencias.First;
    Td_Movimientos.DataSet:= MDInterferencias;
    Td_Movimientos.FieldAliases.Clear;
    Reporte.DataSets.Add(Td_Movimientos);



  finally
    QrInterferencias.Destroy;
    QrRecursos.Destroy;
  end;

end;


Procedure ActaPdf_Informacion(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport);
var
  Td_MovFolios: TfrxDBDataset;
begin
  Td_MovFolios:= TfrxDBDataset.Create(nil);
  Td_MovFolios.UserName:='Td_MovFolios';
  Td_MovFolios.DataSet:= DatosActa;
  Td_MovFolios.FieldAliases.Clear;
  Reporte.DataSets.Add(Td_MovFolios);
end;
Procedure ActaPdf_Actividades(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport);
var
  QrActividades,QImportes:TzReadOnlyQuery;
  Td_partidas,Acta_Costos: TfrxDBDataset;

begin
  QrActividades:=TzReadOnlyQuery.create(nil);
  QImportes :=TzReadOnlyQuery.create(nil);
  Td_partidas:= TfrxDBDataset.Create(nil);
  Acta_Costos:= TfrxDBDataset.Create(nil);
  try
    Td_partidas.UserName:='td_partidas';
    Acta_Costos.UserName:='Acta_Costos';

    QrActividades.connection:=Connection.zConnection;
    QImportes.Connection:=connection.zConnection;

    QImportes.SQL.Text:='select sum(ac.dImporteMn) as dImporteMn,sum(ac.dImporteDll) as dImportedll, '+
                        'ac.*, ac.mDescripcion as sDescripcion,ac.dImporteMn as sImporte, '+
                        'if(ac.eTipo="BARCO","UTILIZANDO POSICIONAMIENTO DINÁMICO",'+
                        'if(ac.eTipo="PERNOCTA","SERVICIOS DE HOTELERIA",if(ac.eTipo="MATERIAL","",ac.eTipo))) as sDescripcionAnexo ' +
                        'from acta_campo ac inner join actividadesxorden ao ' +
                        'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                        'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                        'where iIdActa=:Acta and ao.sTipoActividad=:Tipo and '   +
                        'ac.sIdRecurso like "$IMPORTE%" and ac.eTipo<>"ACTIVIDAD" ' +
                        'group by ac.etipo order by ao.iItemOrden,iOrdenTipo,ac.sAnexo';

    QImportes.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
    QImportes.ParamByName('Tipo').AsString:='Paquete';
    QImportes.Open;

    QrActividades.SQL.Text:=  'select ot.*,pl.sDescripcion as plataforma,ot.subicacion as ubicacion,' + #13#10 +
                              'round(avancesanteriores(:Fecha,ot.sContrato,ot.sNumeroOrden),2) as avance, '  +
                              '(select min(ba.dIdFecha) from bitacoradeactividades ba where ba.sContrato=ot.sContrato and ba.sNumeroOrden=ot.sNumeroOrden) as FInicio,' +
                              '(select max(ba2.dIdFecha) from bitacoradeactividades ba2 where ba2.sContrato=ot.sContrato and ba2.sNumeroOrden= ot.sNumeroOrden) as FTermino ' +
                              'from ordenesdetrabajo ot' + #13#10 +
                              'inner join plataformas pl' + #13#10 +
                              'on(pl.sIdPlataforma=ot.sIdPLataforma)' + #13#10 +
                              'inner join pernoctan pn' + #13#10 +
                              'on(pn.sIdPernocta=ot.sIdPernocta)' + #13#10 +
                              'where ot.sContrato=:Contrato and ot.sNumeroOrden=:Orden';
    QrActividades.ParamByName('Fecha').AsDate:=IncDay(DatosActa.FieldByName('dFecha').AsDateTime);
    QrActividades.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
    QrActividades.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
    try
      QrActividades.Open;
    except
      on e : exception do
        UnitExcepciones.manejarExcep(E.Message, E.ClassName,'Acta de Entrega', 'Al leer las actividades', 0);
    end;



    Td_Partidas.DataSet:= QrActividades;
    Td_Partidas.FieldAliases.Clear;

    Acta_Costos.DataSet:=QImportes;
    Acta_Costos.FieldAliases.Clear;


    Reporte.DataSets.Add(Td_partidas);
    Reporte.DataSets.Add(Acta_Costos);
  finally

  end;
end;

Procedure ActaEx_CostoActividad(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
var
  QrActividades,
  QImportes:TZReadOnlyQuery;
  Acta_CostoPartidas,
  Acta_CostoRecursosDetalle: TfrxDBDataset;
  dsUnion:TDataSource;
begin
  QrActividades:=TZReadOnlyQuery.Create(nil);
  QImportes:=TZReadOnlyQuery.Create(nil);
  dsUnion:=TDataSource.Create(nil);
  try

    Acta_CostoPartidas:= TfrxDBDataset.Create(nil);
    Acta_CostoRecursosDetalle:= TfrxDBDataset.Create(nil);

    QrActividades.Connection:=connection.zConnection;
    QImportes.Connection:=connection.zConnection;

    Acta_CostoPartidas.UserName:='Acta_CostoPartidas';
    Acta_CostoRecursosDetalle.UserName:='Acta_CostoRecursosDetalle';

    QrActividades.SQL.Text:='select ao.* from actividadesxorden ao inner join acta_campo ac '+
                            'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                            'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                            'where ac.iIdActa=:Acta and ao.sTipoActividad=:Tipo ' +
                            'group by ao.swbs order by ao.iItemOrden';

    QImportes.SQL.Text:='select ac.*, ac.mDescripcion as sDescripcion,ac.dImporteMn as sImporte from acta_campo ac inner join actividadesxorden ao ' +
                        'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                        'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                        'where iIdActa=:Acta and ao.sTipoActividad=:Tipo and '   +
                        'ac.sIdRecurso like "$IMPORTE%" and ac.eTipo<>"ACTIVIDAD" order by ao.iItemOrden,iOrdenTipo';

    QrActividades.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
    QrActividades.ParamByName('Tipo').AsString:='Actividad';
    QrActividades.Open;

    QImportes.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
    QImportes.ParamByName('Tipo').AsString:='Actividad';
    QImportes.Open;


    dsUnion.DataSet:=QrActividades;
    QImportes.MasterSource:=dsUnion;
    QImportes.MasterFields:='swbs';
    QImportes.LinkedFields:='swbs';

    Acta_CostoPartidas.DataSet:= QrActividades;
    Acta_CostoRecursosDetalle.DataSet:= QImportes;

    Acta_CostoPartidas.FieldAliases.Clear;
    Acta_CostoRecursosDetalle.FieldAliases.Clear;

    Reporte.DataSets.Add(Acta_CostoPartidas);
    Reporte.DataSets.Add(Acta_CostoRecursosDetalle);

  finally

  end;
end;

Procedure ActaPdf_CostoActividad(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
const
  SQlRef: array[1..2,1..4] of string=(('bitacoradepersonal','personal','sIdPersonal','PERSONAL'),('bitacoradeequipos','equipos','sIdEquipo','EQUIPO'));
var
  QrActividades,
  QrRecursos,
  QrPernoctas,
  QrAdicional,
  QrTotalHH:TZReadOnlyQuery;
  mryRecursos:TjvMemoryData;
  i:Integer;
  dImporteMn,dImporteDll,
  dImporteMnTm,dImporteDllTm:Double;
  Acta_CostoPartidas,
  Acta_CostoRecursosDetalle: TfrxDBDataset;
  sPdasOmitidas:String;
  dTotalImporteMn,dTotalImporteDll:Double;
  dTotalHH:Double;
  tmpRedondeo:String;
begin

  QrRecursos:=TZReadOnlyQuery.Create(nil);
  QrPernoctas:=TZReadOnlyQuery.Create(nil);
  QrAdicional:=TZReadOnlyQuery.Create(nil);
  QrTotalHH:=TZReadOnlyQuery.Create(nil);
  try
    QrActividades:=TZReadOnlyQuery.Create(nil);
    MryRecursos:=TjvMemoryData.Create(nil);
    Acta_CostoPartidas:= TfrxDBDataset.Create(nil);
    Acta_CostoRecursosDetalle:= TfrxDBDataset.Create(nil);

    QrActividades.Connection:=connection.zConnection;
    QrRecursos.Connection:=connection.zConnection;
    QrPernoctas.Connection:=connection.zConnection;
    QrAdicional.Connection:=connection.zConnection;
    QrTotalHH.Connection:=connection.zConnection;
    Acta_CostoPartidas.UserName:='Acta_CostoPartidas';
    Acta_CostoRecursosDetalle.UserName:='Acta_CostoRecursosDetalle';

    with MryRecursos do
    begin
      Active:=False;
      FieldDefs.Add('sContrato', ftString, 100, False);
      FieldDefs.Add('sFolio', ftString, 100, False);
      FieldDefs.Add('sNumeroActividad', ftString, 100, False);
      FieldDefs.Add('sdescripcion', ftString, 500, False);
      FieldDefs.Add('sImporte', ftString, 50, False);
      FieldDefs.Add('dImporteMn', FtCurrency, 0, False);
      FieldDefs.Add('dImporteDll', FtCurrency, 0, False);
      Active:=true;
    end;

    QrActividades.SQL.Text:='select * from actividadesxorden where scontrato=:Contrato and sNumeroOrden=:Orden and sTipoActividad="Actividad" order by iItemOrden';
    QrActividades.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
    QrActividades.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
    QrActividades.Open;
   // QrActividades.AfterScroll:=EventHandlers.ActaPdf_AfterScroll;

    while not QrActividades.Eof do
    begin
      dTotalImporteMn:=0;
      dTotalImporteDll:=0;



      if Tipo=FtAbordo then
      begin
        dImporteMn:=0;
        dImporteDll :=0;
        QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select me.*,sum(IF(mf.sFactor=NULL,me.sFactor,mf.sFactor)) as Total,tm.dVentaMn,tm.dVentaDll ' + #13#10 +
                              'from movimientosdeembarcacion me' + #13#10 +
                              'inner join tiposdemovimiento tm' + #13#10 +
                              'on(me.sContrato=tm.sContrato and me.sClasificacion=tm.sIdTipoMovimiento)' + #13#10 +
                              'left join movimientosxfolios mf' + #13#10 +
                              'on(me.sContrato=mf.sContrato and me.iIdDiario = mf.iIdDiario and mf.sFolio=:Folio)' + #13#10 +
                              'where me.sContrato=:Contrato and me.sOrden=:Orden and me.sNumeroActividad=:Actividad' + #13#10 +
                              'and tm.sClasificacion="Movimiento de Barco" and tm.lAplicaNotaCampo="Si"' + #13#10 +
                              'group by tm.sIdTipoMovimiento' ;
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QrRecursos.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QrRecursos.Open;

        while not QrRecursos.Eof do
        begin
          dImporteMn:= dImporteMn + xRound((QrRecursos.FieldByName('Total').AsFloat *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
          dImporteDll:= dImporteDll + xRound((QrRecursos.FieldByName('Total').AsFloat * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
          QrRecursos.Next;
        end;

        if (dImporteMn>0) or (dImporteDll>0) then
        Begin
          with MryRecursos do
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            FieldByName('sdescripcion').AsString:='MOVIMIENTO DE EMBARCACIÓN';
            FieldByName('sImporte').AsString:='IMPORTE BARCO' ;
            FieldByName('dImporteMn').AsFloat:=dImporteMn;
            FieldByName('dImportedLL').AsFloat:=dImporteDll;
            post;
          end;
        End
        else
        begin
          QrRecursos.Active:=False;
          QrRecursos.SQL.Text:= 'select * from tiposdemovimiento ' + #13#10 +
                                'where sContrato=:Contrato and sClasificacion="Movimiento de Barco" and lAplicaNotaCampo="Si"' + #13#10 +
                                'and lImprimeCeros="Si" group by sIdTipoMovimiento' ;
          QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
          QrRecursos.Open;
          if QrRecursos.recordcount>0 then
          begin
              with MryRecursos do
              begin
                Append;
                FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                FieldByName('sdescripcion').AsString:='MOVIMIENTO DE EMBARCACIÓN';
                FieldByName('sImporte').AsString:='IMPORTE BARCO' ;
                FieldByName('dImporteMn').AsFloat:=dImporteMn;
                FieldByName('dImportedLL').AsFloat:=dImporteDll;
                post;
              end;
          end;



        end;

        dTotalImporteMn:=dTotalImporteMn + dImporteMn ;
        dTotalImporteDll:=dTotalImporteDll + dImporteDll;

      end;



      for I := 1 to Length(SQlRef) do
      begin
        dImporteMn:=0;
        dImporteDll :=0;
        dImporteMnTm :=0;
        dImporteDllTm :=0;
        QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select br.' +SQlRef[i,3] + ' as sIdRecurso,br.sDescripcion,r.sMedida, ' +
                              '(IFNULL(SUM(CAST(br.dCanthh AS CHAR)), 0) + ' +
                              ' IFNULL((SELECT ' +
                              '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                              '         FROM ' +
                              '           bitacoradeajustes ba ' +
                              '         WHERE ' +
                              '           ba.sContrato = br.sContrato AND ba.sNumeroOrden = br.sNumeroOrden AND ' +
                              '           ba.sNumeroActividad = br.sNumeroActividad AND ' +
                              '           ba.sIdPartidaAnexo = br.' + SQlRef[i,3] + ' AND ba.dFecha = br.dIdFecha), 0)) as dCanthh, ' +
                              'r.dVentaMn,r.dVentaDll, ' +
                              'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) as dAjuste ' +
                              'from '+ SQlRef[i,1]  + ' br ' +
                              'left join ' + SQlRef[i,2] + ' r ' +
                              'on(r.sContrato=:Contrato and br.'+SQlRef[i,3]+'=r.'+ SQlRef[i,3] +') ' +
                              'where br.sContrato=:Orden and br.sNumeroOrden=:Folio and br.sNumeroActividad=:Actividad '  +
                              'group by sIdRecurso,br.didfecha ';
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QrRecursos.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QrRecursos.Open;
        tmpRedondeo:='';
        while not QrRecursos.Eof do
        begin
          if (tmpRedondeo<>QrRecursos.fieldByNAme('sIdRecurso').asstring) then
          begin
            if tmpRedondeo<>'' then
            begin
              dImporteMn:= dImporteMn + xRound(dImporteMnTm,2) ;
              dImporteDll:= dImporteDll +  xRound(dImporteDllTm,2);
            end;

            dImporteMnTm :=0;
            dImporteDllTm :=0;
            tmpRedondeo:=QrRecursos.fieldByNAme('sIdRecurso').asstring;
          end;



          dImporteMnTm :=dImporteMnTm+ ((xRound(QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('dAjuste').AsFloat,Mantisa) ) *  QrRecursos.FieldByName('dVentaMn').AsFloat);
          dImporteDllTm :=dImporteDllTm+((xRound(QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('dAjuste').AsFloat,Mantisa) ) * QrRecursos.FieldByName('dVentaDll').AsFloat);

          QrRecursos.Next;
        end;

        if (tmpRedondeo<>'') then
        begin
          dImporteMn:= dImporteMn + xRound(dImporteMnTm,2) ;
          dImporteDll:= dImporteDll +  xRound(dImporteDllTm,2);
        end;


        with MryRecursos do
        begin
          Append;
          FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          FieldByName('sdescripcion').AsString:=SQlRef[i,4];
          FieldByName('sImporte').AsString:='IMPORTE ' + SQlRef[i,4];
          FieldByName('dImporteMn').AsFloat:=xRound(dImporteMn,2);
          FieldByName('dImportedLL').AsFloat:=xRound(dImporteDll,2);
          post;
        end;

        dTotalImporteMn:=dTotalImporteMn + dImporteMn ;
        dTotalImporteDll:=dTotalImporteDll + dImporteDll;
      end;

      if DatosActa.FieldByName('lPernocta').AsString='Si' then
      begin
        dImporteMn:=0;
        dImporteDll :=0;
        dTotalHH:=0;
        with QrRecursos do
        begin
          active := false;

          // Modificación por E. Rangel - 21/08/2016
          sql.text := 'select (sum(xround(bp.dCantHH,2)) + ' +
                      'IFNULL((SELECT ' +
                      '          SUM(XROUND(CAST(ba.dAjuste AS CHAR), 2)) ' +
                      '        FROM ' +
                      '          bitacoradeajustes ba ' +
                      '        WHERE ' +
                      '          ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                      '          ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                      '          ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, ' +
                      'ifnull(SUM(CAST(bp.dAjuste AS CHAR)),0) AS Ajuste ' +
                      'from bitacoradepersonal bp ' +
                      'where bp.sContrato =:Contrato and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.sNumeroActividad=:Actividad and bp.lImprime = "Si" ' +
                      'group by bp.sTipoPernocta,bp.didfecha';

          {sql.text := 'select xround(sum( bp.dCantHH),2) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                      'from bitacoradepersonal bp '+
                      'where bp.sContrato =:Contrato '+
                      'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.sNumeroActividad=:Actividad and bp.lImprime = "Si" '+
                      'group by bp.sTipoPernocta,bp.didfecha ';}
          parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
          parambyname('Folio').asstring      := DatosActa.FieldByName('sNumeroOrden').AsString;
          ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        end;

        with QrTotalHH do
        begin
          active := false;

          // Modificacion por E. Rangel - 21/08/2016
          sql.text := 'select (sum(bp.dCantHH) + ' +
                      ' IFNULL((SELECT ' +
                      '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                      '          FROM ' +
                      '            bitacoradeajustes ba ' +
                      '          WHERE ' +
                      '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                      '            ba.sNumeroActividad = ba.sNumeroActividad AND ' +
                      '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, ' +
                      'ifnull(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste ' +
                      'from bitacoradepersonal bp ' +
                      'where bp.sContrato =:Contrato and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" ' +
                      'group by bp.sTipoPernocta, bp.didfecha';

          {sql.text := 'select sum(bp.dCantHH) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                      'from bitacoradepersonal bp '+
                      'where bp.sContrato =:Contrato '+
                      'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" '+
                      'group by bp.sTipoPernocta, bp.didfecha ';}

          parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
          parambyname('Folio').asstring      := DatosActa.FieldByName('sNumeroOrden').AsString;
        end;


        QrAdicional.Active:=false;
        {QrAdicional.SQL.Text:='select ifnull(sum(dCantidad),0) as dCantidad from bitacoradepernocta where ' +
                              'sContrato=:Contrato and sNumeroOrden=:Folio and '+
                              'sIdCuenta=:Tipo group by sContrato ';}

        QrAdicional.SQL.Text := 'select ' +
                                '  (ifnull(sum(bpr.dCantidad),0) + ' +
                                '   IFNULL((SELECT ' +
                                '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '           FROM ' +
                                '             bitacoradeajustes ba ' +
                                '           WHERE ' +
                                '             ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                                '             ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0)) as dCantidad ' +
                                'from ' +
                                '  bitacoradepernocta bpr ' +
                                'where ' +
                                '  bpr.sContrato=:Contrato and ' +
                                '  bpr.sNumeroOrden=:Folio and ' +
                                '  bpr.sIdCuenta=:Tipo ' +
                                'group by ' +
                                '  bpr.sContrato';

        QrAdicional.parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
        QrAdicional.parambyname('Folio').asstring      := DatosActa.FieldByName('sNumeroOrden').AsString;



        QrPernoctas.Active:=False;
        QrPernoctas.SQL.Text:='select * from cuentas';
        QrPernoctas.Open;
        while not QrPernoctas.Eof do
        begin
          with QrRecursos do
          begin
            Active:=False;
            ParamByName('tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
            Open;
            while not Eof do
            begin
              dImporteMn:= dImporteMn + xRound(((xRound(FieldByName('dCantidad').AsFloat + FieldByName('ajuste').AsFloat,Mantisa) ) *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
              dImporteDll:= dImporteDll + xRound(((FieldByName('dCantidad').AsFloat + FieldByName('ajuste').AsFloat ) * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);
              Next;
            end;
          end;

          with QrTotalHH do
          begin
            Active:=False;
            ParamByName('tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
            Open;
            if RecordCount=1 then
              dTotalHH:=xRound(FieldByName('dCantidad').AsFloat+ FieldByName('aJuste').AsFloat,Mantisa);
          end;



          with QrAdicional do
          begin
            Active:=False;
            ParamByName('tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
            Open;
            while not Eof do
            begin
              dTotalHH:=(xRound(FieldByName('dCantidad').AsFloat,Mantisa) * (xRound(QrRecursos.FieldByName('dCantidad').AsFloat+ QrRecursos.FieldByName('aJuste').AsFloat,Mantisa))) / dTotalHH;
              dImporteMn:= dImporteMn + xRound((dTotalHH *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
              dImporteDll:= dImporteDll + xRound((dTotalHH * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);




              //dImporteMn:= dImporteMn + (FieldByName('dCantidad').AsFloat *  QrPernoctas.FieldByName('dVentaMn').AsFloat);
              //dImporteDll:= dImporteDll + (FieldByName('dCantidad').AsFloat * QrPernoctas.FieldByName('dVentaDll').AsFloat);
              Next;
            end;
          end;

          QrPernoctas.Next;
        end;

        with MryRecursos do
        begin
          Append;
          FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          FieldByName('sdescripcion').AsString:='PERNOCTA';
          FieldByName('sImporte').AsString:='IMPORTE PERNOCTAS';
          FieldByName('dImporteMn').AsFloat:=dImporteMn;
          FieldByName('dImportedLL').AsFloat:=dImporteDll;
          post;

        end;

        dTotalImporteMn:=dTotalImporteMn + dImporteMn ;
        dTotalImporteDll:=dTotalImporteDll + dImporteDll;


      end;

      {if Tipo=FtAbordo then
      begin
        with MryRecursos do
        begin
          Append;
          FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          FieldByName('sdescripcion').AsString:='PERNOCTA';
          FieldByName('sImporte').AsString:='IMPORTE PERNOCTAS';
          FieldByName('dImporteMn').AsFloat:=0;
          FieldByName('dImportedLL').AsFloat:=0;
          post;

          Append;
          FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          FieldByName('sdescripcion').AsString:='MATERIAL';
          FieldByName('sImporte').AsString:='IMPORTE DE MATERIAL';
          FieldByName('dImporteMn').AsFloat:=0;
          FieldByName('dImportedLL').AsFloat:=0;
          post;
        end;

      end; }

      if (dTotalImporteMn=0) and (dTotalImporteDll=0) then
        if sPdasOmitidas='' then
          sPdasOmitidas:= QrActividades.fieldByName('sNumeroActividad').asstring
        else
          sPdasOmitidas:= sPdasOmitidas  + ',' + QrActividades.fieldByName('sNumeroActividad').asstring ;



      QrActividades.Next;
    end;
    if sPdasOmitidas<>'' then
    begin
      QrActividades.Active:=False;
      QrActividades.SQL.Text:='select * from actividadesxorden where scontrato=:Contrato and sNumeroOrden=:Orden and sTipoActividad="Actividad" ' +
                              'and not FIND_IN_SET(snumeroactividad,:PdasOmitidas) order by iItemOrden';
      QrActividades.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
      QrActividades.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
      QrActividades.ParamByName('PdasOmitidas').AsString:=sPdasOmitidas;
      QrActividades.Open;
    end;
    QrActividades.First;
    QrActividades.AfterScroll:=EventHandlers.ActaPdf_AfterScroll;
    MryRecursos.First;
    MemoryTmp:=MryRecursos;

    {TdAvancesPartidas.UserName:='TdAvancesPartidas';
    TdDistribucion_detalle.UserName:='TdDistribucion_detalle';   }

    

    Acta_CostoPartidas.DataSet:= QrActividades;
    Acta_CostoRecursosDetalle.DataSet:= MryRecursos;

    Acta_CostoPartidas.FieldAliases.Clear;
    Acta_CostoRecursosDetalle.FieldAliases.Clear;

    Reporte.DataSets.Add(Acta_CostoPartidas);
    Reporte.DataSets.Add(Acta_CostoRecursosDetalle);

  finally
    QrRecursos.Destroy;
    QrPernoctas.destroy;
    QrAdicional.destroy;
    QrTotalHH.Destroy;
  end;
end;

Procedure Load_NotaCampo(DatosActa:TzREadOnlyQuery;Opciones:TSetRtRecurso=[RtPersonal,RtEquipo,RtPernocta,RtBarco,RtExtraordinaria];Clean:Boolean=true);
const
  SQlRef: array[1..2,1..4] of string=(('bitacoradepersonal','personal','sIdPersonal','PERSONAL'),('bitacoradeequipos','equipos','sIdEquipo','EQUIPO'));
var
  QrActividades,
  QrRecursos,
  QrPernoctas,
  QrAdicional,
  QrMovimientos,
  QrTotalHH:TZReadOnlyQuery;
  i:Integer;
  dImporteMn,dImporteDll,
  tmpCantidad,tmpCantidadReal,
  dImpRecMn,dImpRecDll,
  dImpActMn,dImpActDll:Double;

  sumHrsTotales,SumHrsPda,dTotalHH:Double;

  tmpRecursos:string;
  SumDiaBarco:Variant;
  QActaCampo,QElimina:TZQuery;
  TmpDiaPernocta:TDate;
  PosReg:TBookmark;
begin
  QrRecursos:=TZReadOnlyQuery.Create(nil);
  QrPernoctas:=TZReadOnlyQuery.Create(nil);
  QrAdicional:=TZReadOnlyQuery.Create(nil);
  QrTotalHH:=TZReadOnlyQuery.Create(nil);
  QrMovimientos:=TZReadOnlyQuery.Create(nil);
  QActaCampo:=TZQuery.Create(nil);
  QElimina:=TZQuery.Create(nil);
  QrActividades:=TZReadOnlyQuery.Create(nil);
  try
    QrActividades.Connection:=connection.zConnection;
    QrRecursos.Connection:=connection.zConnection;
    QrPernoctas.Connection:=connection.zConnection;
    QrAdicional.Connection:=connection.zConnection;
    QrTotalHH.Connection:=connection.zConnection;
    QrMovimientos.Connection:=connection.zConnection;

    QActaCampo.Connection:=connection.zConnection;
    QElimina.Connection:=connection.zConnection;

    QActaCampo.SQL.Text:= 'select * from acta_campo where sContrato=:Contrato and iIdActa=:Acta and '+
                          'sNumeroOrden=:Orden and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                          'eTipo=:Tipo ';

    QElimina.SQL.Text:='delete from acta_campo where sContrato=:Contrato and iIdActa=:Acta and ' +
                       'sNumeroOrden=:Orden and swbs=:wbs and sNumeroActividad=:Actividad and ' +
                       'eTipo=:Tipo ';


    QrActividades.SQL.Text:='select * from actividadesxorden where scontrato=:Contrato and sNumeroOrden=:Orden and sTipoActividad="Actividad" order by iItemOrden ';
    QrActividades.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
    QrActividades.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
    QrActividades.Open;


    if DatosActa.FieldByName('lPernocta').AsString='Si' then
    begin
      with QrPernoctas do
      begin
        Active:=False;
        SQL.Text:='select * from cuentas order by iOrden';
        Open;
      end;

      QrAdicional.Active:=false;
      {QrAdicional.SQL.Text:='select ifnull(sum(dCantidad),0) as dCantidad from bitacoradepernocta where ' +
                            'sContrato=:Contrato and sNumeroOrden=:Folio and '+
                            'sIdCuenta=:Tipo group by sContrato ';}

      QrAdicional.SQL.Text := 'select ' +
                              '  (ifnull(sum(bpr.dCantidad),0) + ' +
                              '   IFNULL((SELECT ' +
                              '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                              '           FROM ' +
                              '             bitacoradeajustes ba ' +
                              '           WHERE ' +
                              '             ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                              '             ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0)) as dCantidad ' +
                              'from ' +
                              '  bitacoradepernocta bpr ' +
                              'where ' +
                              '  bpr.sContrato=:Contrato and ' +
                              '  bpr.sNumeroOrden=:Folio and ' +
                              '  bpr.sIdCuenta=:Tipo ' +
                              'group by ' +
                              '  bpr.sContrato';

      QrAdicional.parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
      QrAdicional.parambyname('Folio').asstring      := DatosActa.FieldByName('sNumeroOrden').AsString;

      with QrTotalHH do
      begin
        active := false;

        // Cambios por E. Rangel - 21/08/2016
        sql.text := 'select (sum(bp.dCantHH) + ' +
                    ' IFNULL((SELECT ' +
                    '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                    '          FROM ' +
                    '            bitacoradeajustes ba ' +
                    '          WHERE ' +
                    '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                    '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                    '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantidad, ' +
                    'bp.sNumeroOrden, bp.sIdPlataforma, ' +
                    'ifnull(SUM(CAST(bp.dAjuste AS CHAR)),0) AS Ajuste ' +
                    'from bitacoradepersonal bp ' +
                    'where bp.sContrato =:Contrato and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" ' +
                    'group by bp.sTipoPernocta';

        {sql.text := 'select sum(bp.dCantHH) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                    'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                    'from bitacoradepersonal bp '+
                    'where bp.sContrato =:Contrato '+
                    'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" '+
                    'group by bp.sTipoPernocta ';}

        parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
        parambyname('Folio').asstring      := DatosActa.FieldByName('sNumeroOrden').AsString;
      end;
    end;



    while not QrActividades.Eof do
    begin

      dImpActMn:=0;
      dImpActDll:=0;

      {$REGION 'BARCO'}
      if (RtBarco in Opciones) then
      begin
        dImpRecMn:=0;
        dImpRecDll:=0;
        SumHrsPda:=0;
        sumHrsTotales:=0;
        SumDiaBarco:=0;

        if Clean then
        begin
          QElimina.Active:=False;
          QElimina.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QElimina.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          QElimina.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QElimina.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QElimina.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QElimina.ParamByName('Tipo').AsString:='BARCO';
          QElimina.ExecSQL;
        end;

        QActaCampo.Active:=False;
        QActaCampo.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QActaCampo.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
        QActaCampo.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QActaCampo.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
        QActaCampo.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QActaCampo.ParamByName('Tipo').AsString:='BARCO';
        QActaCampo.Open;

        //QrMovimientos
        QrMovimientos.Active:=False;
        QrMovimientos.SQL.Text:= 'select * from tiposdemovimiento ' + #13#10 +
                              'where sContrato=:Contrato and sClasificacion="Movimiento de Barco" and lAplicaNotaCampo="Si"' + #13#10 +
                              'group by sIdTipoMovimiento order by iOrden' ;
        QrMovimientos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrMovimientos.Open;
        while not QrMovimientos.Eof do
        Begin
          QrRecursos.Active:=False;
          QrRecursos.SQL.Text:= 'select b.*,CONVERT(TIME_FORMAT(SEC_TO_TIME(Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b.didfecha,b.sHorainicio),' + #13#10 +
                                'TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ,"%H:%i") using utf8)  AS diferencia,tm.sIdTipoMovimiento,' + #13#10 +
                                'tm.sDescripcion,tm.dVentaMn,tm.dVentaDll, ( EXTRACT( HOUR FROM sec_to_time(Sum( (TIMESTAMPDIFF(MINUTE ,' + #13#10 +
                                'TIMEstamp(b.didfecha,b.sHorainicio),TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ) * 3600 +' + #13#10 +
                                'EXTRACT( MINUTE FROM sec_to_time(Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b.didfecha,b.sHorainicio),' + #13#10 +
                                 'TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ) * 60)  as Hora,xRound(( (CAST((xRound( Cast((' + #13#10 +
                                 'cast(( EXTRACT( HOUR FROM sec_to_time(Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b.didfecha,b.sHorainicio),' + #13#10 +
                                 'TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ) * 3600 + EXTRACT( MINUTE FROM sec_to_time(' + #13#10 +
                                 'Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b.didfecha,b.sHorainicio), TIMEstamp(b.didfecha,b.sHorafinal) )) ) *' + #13#10 +
                                 '60) ) * 60)  as decimal(50,30)) /86400   ) as decimal(50,30)),16) / (CAST((' + #13#10 +
                                 'select ( EXTRACT( HOUR FROM sec_to_time(Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b2.didfecha,b2.sHorainicio),' + #13#10 +
                                 'TIMEstamp(b2.didfecha,b2.sHorafinal) )) ) * 60) ) * 3600 + EXTRACT( MINUTE FROM sec_to_time(' + #13#10 +
                                 'Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b2.didfecha,b2.sHorainicio), TIMEstamp(b2.didfecha,b2.sHorafinal) )) )' + #13#10 +
                                 '* 60) ) * 60)  as Hora' + #13#10 +
                                 'from bitacoradeactividades b2' + #13#10 +
                                 'inner join tiposdemovimiento tm2 on(tm2.sContrato=:ContratoBarco and tm2.sIdTipoMovimiento=b2.sIdClasificacion)' + #13#10 +
                                 'where b2.sContrato=:Contrato and b2.snumeroorden=:Orden and b2.sIdTipoMovimiento=:Tipo and b2.didfecha=b.didfecha' + #13#10 +
                                 'group by b2.didfecha ) as decimal(50,30)) /86400) ) as decimal(50,30))) *' + #13#10 +
                                 'xRound((select ifnull(SUM(mf.sFactor),0) from movimientosxfolios mf' + #13#10 +
                                 'inner join movimientosdeembarcacion me' + #13#10 +
                                 'on(me.sContrato=mf.sContrato and me.iIdDiario = mf.iIdDiario)' + #13#10 +
                                 'inner join tiposdemovimiento tm' + #13#10 +
                                 'on(me.sContrato=tm.sContrato and me.sClasificacion=tm.sIdTipoMovimiento)' + #13#10 +
                                 'where mf.sContrato=:ContratoBarco and mf.dIdFecha=b.dIdFecha and mf.sFolio=b.sNumeroOrden' + #13#10 +
                                 'and tm.sClasificacion="Movimiento de Barco" and tm.lAplicaNotaCampo="Si" and tm.sIdTipoMovimiento=:Movimiento ' + #13#10 +
                                 'group by mf.sFolio),6) ),17) as diapor' + #13#10 +
                                 'from bitacoradeactividades b' + #13#10 +
                                 'inner join tiposdemovimiento tm on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=b.sIdClasificacion)' + #13#10 +
                                 'where b.sContrato=:Contrato and b.snumeroorden=:Orden and b.sIdTipoMovimiento=:Tipo' + #13#10 +
                                 'group by b.swbs,b.didfecha order by b.didfecha' ;
          QrRecursos.ParamByName('contratoBarco').AsString:=global_Contrato_Barco;
          QrRecursos.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QrRecursos.ParamByName('Movimiento').AsString:=QrMovimientos.FieldByName('sIdTipoMovimiento').AsString;
          QrRecursos.ParamByName('Tipo').AsString:='ED';
          try
            QrRecursos.Open;
          except

          end;

          if QrRecursos.Active then
            while not QrRecursos.Eof do
            begin
              if QrRecursos.FieldByName('swbs').AsString=QrActividades.FieldByName('swbs').AsString then
                SumHrsPda:=SumHrsPda +  QrRecursos.FieldByName('hora').AsFloat;

              if not QrRecursos.FieldByName('diapor').IsNull then
                SumDiaBarco:= SumDiaBarco +  QrRecursos.FieldByName('diapor').AsVariant;

              sumHrsTotales := sumHrsTotales + QrRecursos.FieldByName('hora').AsFloat;
              QrRecursos.Next;
            end;

          if SumDiaBarco>0 then
          begin
            with QActaCampo do
            begin
              //86400
              if not QActaCampo.Locate('sIdRecurso',QrMovimientos.FieldByName('sIdTipoMovimiento').AsString,[]) then
              begin
                Append;
                FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                FieldByName('sIdRecurso').AsString:=QrMovimientos.FieldByName('sIdTipoMovimiento').AsString;
                FieldByName('mdescripcion').AsString:=QrMovimientos.FieldByName('sDescripcion').AsString;
                FieldByName('sMedida').AsString:=QrMovimientos.FieldByName('sTipo').AsString;
                FieldByName('eTipo').AsString:='BARCO';
              end
              else
                Edit;

              FieldByName('dCantidad').AsFloat:= xRound((SumDiaBarco * (SumHrsPda / 86400))/ (sumHrsTotales/ 86400),6);
              FieldByName('dCantidadReal').AsFloat:=(SumDiaBarco * (SumHrsPda / 86400))/ (sumHrsTotales/ 86400) ;
              FieldByName('dCostoMn').AsFloat:=QrMovimientos.FieldByName('dVentaMn').AsFloat;
              FieldByName('dCostoDll').AsFloat:=QrMovimientos.FieldByName('dVentaDll').AsFloat;
              FieldByName('dImporteMn').AsFloat:=xRound((xRound((SumDiaBarco * (SumHrsPda / 86400))/ (sumHrsTotales/ 86400),6) *  QrMovimientos.FieldByName('dVentaMn').AsFloat),2);
              FieldByName('dImportedLL').AsFloat:=xRound((xRound((SumDiaBarco * (SumHrsPda / 86400))/ (sumHrsTotales/ 86400),6)* QrMovimientos.FieldByName('dVentaDll').AsFloat),2);
              FieldByName('iOrdenTipo').AsInteger:=0;
              FieldByName('iOrdenRecurso').AsInteger:=QrMovimientos.FieldByName('iOrden').AsInteger;
              FieldByName('sFormulaMn').AsString:='##';
              FieldByName('sFormulaDll').AsString:='##';
              FieldByName('xCol').AsInteger:=-1;
              FieldByName('xRow').AsInteger:=-1;
              FieldByName('sAnexo').AsString:='';
              post;
            end;

            dImpRecMn:=dImpRecMn + xRound((xRound((SumDiaBarco * (SumHrsPda / 86400))/ (sumHrsTotales/ 86400),6) *  QrMovimientos.FieldByName('dVentaMn').AsFloat),2);
            dImpRecDll:= dImpRecDll + xRound((xRound((SumDiaBarco * (SumHrsPda / 86400))/ (sumHrsTotales/ 86400),6)* QrMovimientos.FieldByName('dVentaDll').AsFloat),2);


          end
          else
          begin
            if not QActaCampo.Locate('sIdRecurso',QrMovimientos.FieldByName('sIdTipoMovimiento').AsString,[]) then
            begin
              if QrMovimientos.FieldByName('lImprimeCeros').AsString='Si' then
                with QActaCampo do
              begin
                Append;
                FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                FieldByName('sIdRecurso').AsString:=QrMovimientos.FieldByName('sIdTipoMovimiento').AsString;
                FieldByName('mdescripcion').AsString:=QrMovimientos.FieldByName('sDescripcion').AsString;
                FieldByName('sMedida').AsString:=QrMovimientos.FieldByName('sTipo').AsString;
                FieldByName('eTipo').AsString:='BARCO';
                FieldByName('dCantidad').AsFloat:=0 ;
                FieldByName('dCantidadReal').AsFloat:=0 ;
                FieldByName('dCostoMn').AsFloat:=QrMovimientos.FieldByName('dVentaMn').AsFloat;
                FieldByName('dCostoDll').AsFloat:=QrMovimientos.FieldByName('dVentaDll').AsFloat;
                FieldByName('dImporteMn').AsFloat:=0;
                FieldByName('dImportedLL').AsFloat:=0;
                FieldByName('iOrdenTipo').AsInteger:=0;
                FieldByName('iOrdenRecurso').AsInteger:=QrMovimientos.FieldByName('iOrden').AsInteger;
                FieldByName('sFormulaMn').AsString:='##';
                FieldByName('sFormulaDll').AsString:='##';
                FieldByName('xCol').AsInteger:=-1;
                FieldByName('xRow').AsInteger:=-1;
                FieldByName('sAnexo').AsString:='';
                post;
              end;
            end
            else
              if QrMovimientos.FieldByName('lImprimeCeros').AsString='No' then
                QActaCampo.Delete;
          end;

          QrMovimientos.Next;
        End;

        with QActaCampo do
        begin
          if REcordCoUnt>0 then
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
            FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            FieldByName('sIdRecurso').AsString:='$IMPORTEBARCO';
            FieldByName('mdescripcion').AsString:='IMPORTE BARCO:';
            FieldByName('sMedida').AsString:='*';
            FieldByName('eTipo').AsString:='BARCO';
            FieldByName('dCantidad').AsFloat:=0;
            FieldByName('dCantidadReal').AsFloat:=0 ;
            FieldByName('dCostoMn').AsFloat:=0;
            FieldByName('dCostoDll').AsFloat:=0;
            FieldByName('dImporteMn').AsFloat:=dImpRecMn;
            FieldByName('dImportedLL').AsFloat:=dImpRecDll;
            FieldByName('iOrdenTipo').AsInteger:=0;
            FieldByName('iOrdenRecurso').AsInteger:=0;
            FieldByName('sFormulaMn').AsString:='##';
            FieldByName('sFormulaDll').AsString:='##';
            FieldByName('xCol').AsInteger:=-1;
            FieldByName('xRow').AsInteger:=-1;
            FieldByName('sAnexo').AsString:='';
            post;
          end;
        end;
        dImpActMn:=dImpActMn +  dImpRecMn;
        dImpActDll:=dImpActDll +  dImpRecDll;
      end;
      {$ENDREGION}

      {$REGION 'PERSONAL Y EQUIPO'}
      for I := 1 to Length(SQlRef) do
      begin
        tmpRecursos:='';
        tmpCantidad:=0;
        dImporteMn:=0;
        dImporteDll :=0;
        tmpCantidadReal:=0;
        dImpRecMn:=0;
        dImpRecDll:=0;
        if ((RtPersonal in Opciones) and (i=1)) or ((RtEquipo in Opciones) and (i=2)) then
        begin
          if Clean then
          begin
            QElimina.Active:=False;
            QElimina.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            QElimina.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
            QElimina.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            QElimina.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            QElimina.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            QElimina.ParamByName('Tipo').AsString:=SQlRef[i,4];
            QElimina.ExecSQL;
          end;

          QActaCampo.Active:=False;
          QActaCampo.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QActaCampo.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          QActaCampo.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QActaCampo.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QActaCampo.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QActaCampo.ParamByName('Tipo').AsString:=SQlRef[i,4];
          QActaCampo.Open;

          QrRecursos.Active:=False;
          QrRecursos.SQL.Text:= 'select br.' +SQlRef[i,3] + ' as sIdRecurso,br.sDescripcion,r.sMedida, ' +
                                '(sum(br.dCanthh) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = br.sContrato AND ba.sNumeroOrden = br.sNumeroOrden AND ' +
                                '           ba.sNumeroActividad = :Actividad AND ' +
                                '           ba.sIdPartidaAnexo = br.' + SQLRef[i, 3] + ' AND ba.dFecha = bp.dIdFecha), 0)) as dCanthh, ' +
                                'Ifnull(sum(CAST(br.dAjuste AS CHAR)), 0) as Ajuste, ' +
                                'r.dVentaMn,r.dVentaDll,r.iitemorden ' +
                                'from '+ SQlRef[i,1]  + ' br ' +
                                'left join ' + SQlRef[i,2] + ' r ' +
                                'on(r.sContrato=:Contrato and br.'+SQlRef[i,3]+'=r.'+ SQlRef[i,3] +') ' +
                                'where br.sContrato=:Orden and br.sNumeroOrden=:Folio and br.sNumeroActividad=:Actividad ' +
                                'group by sIdRecurso,br.dIdFecha order by r.iitemorden' ;
          QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
          QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QrRecursos.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QrRecursos.Open;

          while not QrRecursos.Eof do
          begin
            if tmpRecursos<> QrRecursos.FieldByName('sIdRecurso').AsString then
            begin
              with QActaCampo do
              begin
                if tmpRecursos<>'' then
                begin
                  if QActaCampo.Locate('sIdRecurso',tmpRecursos,[]) then
                  begin
                    QActaCampo.Edit;
                    QActaCampo.FieldByName('dCantidad').AsFloat:=tmpCantidad;
                    QActaCampo.FieldByName('dCantidadReal').AsFloat:=tmpCantidadReal;
                    QActaCampo.FieldByName('dImporteMn').AsFloat:=xRound(dImporteMn,Mantisa);
                    QActaCampo.FieldByName('dImportedLL').AsFloat:=xRound(dImporteDll,Mantisa);
                    QActaCampo.Post;
                    dImpRecMn:=dImpRecMn + xRound(dImporteMn,Mantisa);
                    dImpRecDll:=dImpRecDll + xRound(dImporteDll,Mantisa);



                  end;
                  tmpCantidad:=0;
                  tmpCantidadReal:=0;
                  dImporteMn:=0;
                  dImporteDll :=0;

                end;

                if not Locate('sIdRecurso',QrRecursos.FieldByName('sIdRecurso').AsString,[]) then
                begin
                  Append;
                  FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                  FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                  FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                  FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                  FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                  FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sIdRecurso').AsString;
                  FieldByName('mdescripcion').AsString:=QrRecursos.FieldByName('sDescripcion').AsString;
                  FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sMedida').AsString;
                  FieldByName('eTipo').AsString:=SQlRef[i,4];
                  FieldByName('dCantidad').AsFloat:=xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa);
                  FieldByName('dCantidadReal').AsFloat:=(QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat );
                  FieldByName('dCostoMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
                  FieldByName('dCostodLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
                  FieldByName('dImporteMn').AsFloat:=xRound((xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
                  FieldByName('dImportedLL').AsFloat:=xRound((xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
                  FieldByName('iOrdenTipo').AsInteger:=i;
                  FieldByName('iOrdenRecurso').AsInteger:=QrRecursos.FieldByName('iitemorden').AsInteger;
                  FieldByName('sFormulaMn').AsString:='##';
                  FieldByName('sFormulaDll').AsString:='##';
                  FieldByName('xCol').AsInteger:=-1;
                  FieldByName('xRow').AsInteger:=-1;
                  FieldByName('sAnexo').AsString:='';
                  post;
                end;
                tmpRecursos:=QrRecursos.FieldByName('sIdRecurso').AsString;

              end;
            end;

            tmpCantidad:=tmpCantidad + xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa);
            tmpCantidadReal:=tmpCantidadReal + (QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat );
            dImporteMn:= dImporteMn + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) *  QrRecursos.FieldByName('dVentaMn').AsFloat);
            dImporteDll:= dImporteDll + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat);

            QrRecursos.Next;
          end;

          with QActaCampo do
          begin
            if tmpRecursos<>'' then
            begin
              if QActaCampo.Locate('sIdRecurso',tmpRecursos,[]) then
              begin
                QActaCampo.Edit;
                QActaCampo.FieldByName('dCantidad').AsFloat:=tmpCantidad;
                QActaCampo.FieldByName('dCantidadReal').AsFloat:=tmpCantidadReal;
                QActaCampo.FieldByName('dImporteMn').AsFloat:=xRound(dImporteMn,Mantisa);
                QActaCampo.FieldByName('dImportedLL').AsFloat:=xRound(dImporteDll,Mantisa);
                QActaCampo.Post;

                dImpRecMn:=dImpRecMn + xRound(dImporteMn,Mantisa);
                dImpRecDll:=dImpRecDll + xRound(dImporteDll,Mantisa);
              end;





            end;
          end;

          with QActaCampo do
          begin
            if REcordCoUnt>0 then
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sIdRecurso').AsString:='$IMPORTE'+SQlRef[i,4];
              FieldByName('mdescripcion').AsString:='IMPORTE '+SQlRef[i,4]+':';
              FieldByName('sMedida').AsString:='*';
              FieldByName('eTipo').AsString:=SQlRef[i,4];
              FieldByName('dCantidad').AsFloat:=0;
              FieldByName('dCantidadReal').AsFloat:=0 ;
              FieldByName('dCostoMn').AsFloat:=0;
              FieldByName('dCostoDll').AsFloat:=0;
              FieldByName('dImporteMn').AsFloat:=dImpRecMn;
              FieldByName('dImportedLL').AsFloat:=dImpRecDll;
              FieldByName('iOrdenTipo').AsInteger:=i;
              FieldByName('iOrdenRecurso').AsInteger:=0;
              FieldByName('sFormulaMn').AsString:='##';
              FieldByName('sFormulaDll').AsString:='##';
              FieldByName('xCol').AsInteger:=-1;
              FieldByName('xRow').AsInteger:=-1;
              FieldByName('sAnexo').AsString:='';
              post;
            end;
          end;


          dImpActMn:=dImpActMn +  dImpRecMn;
          dImpActDll:=dImpActDll + dImpRecDll;
        end;

      end;
      {$ENDREGION}

      {$REGION 'PERNOCTA'}
      if DatosActa.FieldByName('lPernocta').AsString='Si' then
        if (RtPernocta in Opciones) then
        begin
          dImpRecMn:=0;
          dImpRecDll:=0;
          TmpDiaPernocta:=0;
          if Clean then
          begin
            QElimina.Active:=False;
            QElimina.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            QElimina.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
            QElimina.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            QElimina.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            QElimina.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            QElimina.ParamByName('Tipo').AsString:='PERNOCTA';
            QElimina.ExecSQL;
          end;

          QActaCampo.Active:=False;
          QActaCampo.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QActaCampo.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          QActaCampo.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QActaCampo.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QActaCampo.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QActaCampo.ParamByName('Tipo').AsString:='PERNOCTA';
          QActaCampo.Open;

          with QrRecursos do
          begin
            active := false;
            {sql.text := 'select sum(bp.dCantHH) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                        'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                        'from bitacoradepersonal bp '+
                        'where bp.sContrato =:Contrato '+
                        'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" '+
                        'and bp.sNumeroActividad=:Actividad ' +
                        'group by bp.sTipoPernocta ';
                        }
// Rangel Aquí hay que revisar esta madre
            sql.text := 'select b.*,CONVERT(TIME_FORMAT(SEC_TO_TIME(Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b.didfecha,b.sHorainicio), ' + #13#10 +
                        'TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ,"%H:%i") using utf8)  AS diferencia,tm.sIdTipoMovimiento,' + #13#10 +
                        'tm.sDescripcion,tm.dVentaMn,tm.dVentaDll, ( EXTRACT( HOUR FROM sec_to_time(Sum( (TIMESTAMPDIFF(MINUTE ,' + #13#10 +
                        'TIMEstamp(b.didfecha,b.sHorainicio),TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ) * 3600 +' + #13#10 +
                        'EXTRACT( MINUTE FROM sec_to_time(Sum( (TIMESTAMPDIFF(MINUTE ,TIMEstamp(b.didfecha,b.sHorainicio),' + #13#10 +
                        'TIMEstamp(b.didfecha,b.sHorafinal) )) ) * 60) ) * 60)  as Hora,' + #13#10 +
                        '(( select sum(b3.dCanthh ) + ' +
                        ' IFNULL((SELECT ' +
                        '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                        '         FROM ' +
                        '           bitacoradeajustes ba ' +
                        '         WHERE ' +
                        '           ba.sContrato = b3.sContrato AND ba.sNumeroOrden = b3.sNumeroOrden AND ' +
                        '           ba.sNumeroActividad = b3.sNumeroActividad AND ' +
                        '           ba.sIdPartidaAnexo = b3.sIdPersonal AND ba.dFecha = b3.dIdFecha), 0)) from (select sContrato,sNumeroOrden,didfecha,sidpersonal,' + #13#10 +
                        '(xRound(sum(b2.dCanthh + b2.dAjuste ), 2) + ' +
                        ' IFNULL((SELECT ' +
                        '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                        '         FROM ' +
                        '           bitacoradeajustes ba ' +
                        '         WHERE ' +
                        '           ba.sContrato = b2.sContrato AND ba.sNumeroOrden = b2.sNumeroOrden AND ' +
                        '           ba.sNumeroActividad = b2.sNumeroActividad AND ' +
                        '           ba.sIdPartidaAnexo = b2.sIdPersonal AND ba.dFecha = b2.dIdFecha), 0)) as dCanthh  from bitacoradepersonal b2' + #13#10 +
                        'where b2.sContrato=:Contrato and b2.sNumeroOrden=:Orden and b2.sTipoPernocta=:Pernocta and b2.lAplicaPernocta="Si" group by b2.didfecha,b2.sIdPersonal) b3' + #13#10 +
                        'where b3.sContrato=b.sContrato and b3.sNumeroOrden=b.sNumeroOrden and b3.didfecha=b.didfecha ' + #13#10 +
                        'group by b3.didfecha ) as Dia,( select ifnull(sum(bp.dCantidad),0)' + #13#10 +
                        'from bitacoradepernocta bp where (bp.sContrato=b.sContrato and bp.sNumeroOrden=b.sNumeroOrden and ' + #13#10 +
                        'bp.sIdCuenta=:Pernocta and bp.didfecha=b.didfecha) group by bp.sIdCuenta ) as Adicional' + #13#10 +
                        'from bitacoradeactividades b ' + #13#10 +
                        'inner join tiposdemovimiento tm on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=b.sIdClasificacion)' + #13#10 +
                        'where b.sContrato=:Contrato and b.snumeroorden=:Orden and b.sIdTipoMovimiento=:Tipo ' + #13#10 +
                        'group by b.swbs,b.didfecha order by b.didfecha';

            parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
            parambyname('Orden').asstring           := DatosActa.FieldByName('sNumeroOrden').AsString;
            ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;
            parambyname('Tipo').asstring           := 'ED';
          end;
          dImpRecMn:=0;
          dImpRecDll:=0;
          QrPernoctas.First;
          while not QrPernoctas.Eof do
          begin

            dTotalHH:=0;
            QrRecursos.Active:=False;
            QrRecursos.ParamByName('Pernocta').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
            QrRecursos.Open;

            {QrAdicional.Active:=False;
            QrAdicional.ParamByName('Tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
            QrAdicional.Open;

            if QrAdicional.RecordCount>0 then
            begin
              QrTotalHH.Active:=False;
              QrTotalHH.ParamByName('Tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
              QrTotalHH.Open;
              if QrTotalHH.RecordCount>0 then
                dTotalHH:=xRound(QrTotalHH.FieldByName('dCantidad').asfloat +  QrTotalHH.FieldByName('Ajuste').asfloat,Mantisa) ;

              dTotalHH:=(xRound((QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa)  * QrAdicional.FieldByName('dCantidad').AsFloat)  / dTotalHH ;


            end; }
            dTotalHH:=0;
            while not QrRecursos.Eof do
            begin
              if QrRecursos.FieldByName('swbs').AsString=QrActividades.FieldByName('swbs').AsString {TmpDiaPernocta<>QrRecursos.FieldByName('dIdFecha').AsDateTime} then
              begin
                TmpDiaPernocta:=QrRecursos.FieldByName('dIdFecha').AsDateTime;
                sumHrsTotales:=0;
                SumHrsPda:=0;
                PosReg:=QrRecursos.GetBookmark;
                QrRecursos.First;
                while not QrRecursos.Eof do
                begin
                  if QrRecursos.FieldByName('didfecha').AsDateTime=TmpDiaPernocta then
                  begin
                    sumHrsTotales:=sumHrsTotales + QrRecursos.FieldByName('Hora').AsFloat;
                    if QrRecursos.FieldByName('swbs').AsString=QrActividades.FieldByName('swbs').AsString then
                      SumHrsPda:=SumHrsPda + + QrRecursos.FieldByName('Hora').AsFloat;
                  end;
                  QrRecursos.Next;
                end;
                QrRecursos.GotoBookmark(PosReg);

                if sumHrsTotales>0 then
                  dTotalHH:=dTotalHH + (((QrRecursos.FieldByName('dia').AsFloat + QrRecursos.FieldByName('adicional').AsFloat) * (SumHrsPda/86400)) /
                          (sumHrsTotales/86400));
              end;
             // (SumHrsPda/86400)
             // (sumHrsTotales/86400)

              QrRecursos.Next;
            end;

            if dTotalHH>0 then
            begin
              with QActaCampo do
              begin
                if not Locate('sIdRecurso',QrPernoctas.FieldByName('sIdPernocta').AsString,[]) then
                begin
                  Append;
                  FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                  FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                  FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                  FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                  FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                  FieldByName('sIdRecurso').AsString:=QrPernoctas.FieldByName('sIdPernocta').AsString;
                  FieldByName('mdescripcion').AsString:=QrPernoctas.FieldByName('sDescripcion').AsString;
                  FieldByName('sMedida').AsString:=QrPernoctas.FieldByName('sMedida').AsString;
                  FieldByName('eTipo').AsString:='PERNOCTA';
                end
                else
                  edit ;

                FieldByName('dCantidad').AsFloat:=xRound(dTotalHH,2);
                FieldByName('dCantidadReal').AsFloat:=dTotalHH ;
                FieldByName('dCostoMn').AsFloat:=QrPernoctas.FieldByName('dVentaMn').AsFloat;
                FieldByName('dCostodLL').AsFloat:=QrPernoctas.FieldByName('dVentaDll').AsFloat;
                FieldByName('dImporteMn').AsFloat:=xRound((xRound(dTotalHH,2)  *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
                FieldByName('dImportedLL').AsFloat:=xRound((xRound(dTotalHH,2)  * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);
                FieldByName('iOrdenTipo').AsInteger:=3;
                FieldByName('iOrdenRecurso').AsInteger:=QrPernoctas.FieldByName('iorden').AsInteger;
                FieldByName('sFormulaMn').AsString:='##';
                FieldByName('sFormulaDll').AsString:='##';
                FieldByName('xCol').AsInteger:=-1;
                FieldByName('xRow').AsInteger:=-1;
                FieldByName('sAnexo').AsString:='';
                post;
              end;

                  dImpRecMn:=dImpRecMn + xRound((xRound(dTotalHH,2)  *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
                  dImpRecDll:=dImpRecDll + xRound((xRound(dTotalHH,2)  * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);

            end
            else
            begin
              with QActaCampo do
                if not Locate('sIdRecurso',QrPernoctas.FieldByName('sIdPernocta').AsString,[]) then
                begin
                  Append;
                  FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                  FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                  FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                  FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                  FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                  FieldByName('sIdRecurso').AsString:=QrPernoctas.FieldByName('sIdPernocta').AsString;
                  FieldByName('mdescripcion').AsString:=QrPernoctas.FieldByName('sDescripcion').AsString;
                  FieldByName('sMedida').AsString:=QrPernoctas.FieldByName('sMedida').AsString;
                  FieldByName('eTipo').AsString:='PERNOCTA';
                  FieldByName('dCantidad').AsFloat:=xRound(dTotalHH,2);
                  FieldByName('dCantidadReal').AsFloat:=dTotalHH;
                  FieldByName('dCostoMn').AsFloat:=QrPernoctas.FieldByName('dVentaMn').AsFloat;
                  FieldByName('dCostodLL').AsFloat:=QrPernoctas.FieldByName('dVentaDll').AsFloat;
                  FieldByName('dImporteMn').AsFloat:=xRound((xRound(dTotalHH,2) *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
                  FieldByName('dImportedLL').AsFloat:=xRound((xRound(dTotalHH,2)   * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);
                  FieldByName('iOrdenTipo').AsInteger:=3;
                  FieldByName('iOrdenRecurso').AsInteger:=QrPernoctas.FieldByName('iorden').AsInteger;
                  FieldByName('sFormulaMn').AsString:='##';
                  FieldByName('sFormulaDll').AsString:='##';
                  FieldByName('xCol').AsInteger:=-1;
                  FieldByName('xRow').AsInteger:=-1;
                  FieldByName('sAnexo').AsString:='';
                  post;
                  dImpRecMn:=dImpRecMn+ xRound((xRound(dTotalHH,2) *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
                  dImpRecDll:=dImpRecDll +  xRound((xRound(dTotalHH,2)  * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);
                end;
            end;
            QrPernoctas.Next;
          end;

          with QActaCampo do
          begin
            if REcordCoUnt>0 then
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sIdRecurso').AsString:='$IMPORTEPERNOCTA';
              FieldByName('mdescripcion').AsString:='IMPORTE PERNOCTA :';
              FieldByName('sMedida').AsString:='*';
              FieldByName('eTipo').AsString:='PERNOCTA';
              FieldByName('dCantidad').AsFloat:=0;
              FieldByName('dCantidadReal').AsFloat:=0 ;
              FieldByName('dCostoMn').AsFloat:=0;
              FieldByName('dCostoDll').AsFloat:=0;
              FieldByName('dImporteMn').AsFloat:=dImpRecMn;
              FieldByName('dImportedLL').AsFloat:=dImpRecDll;
              FieldByName('iOrdenTipo').AsInteger:=3;
              FieldByName('iOrdenRecurso').AsInteger:=0;
              FieldByName('sFormulaMn').AsString:='##';
              FieldByName('sFormulaDll').AsString:='##';
              FieldByName('xCol').AsInteger:=-1;
              FieldByName('xRow').AsInteger:=-1;
              FieldByName('sAnexo').AsString:='';
              post;
            end;
          end;


          dImpActMn:=dImpActMn +  dImpRecMn;
          dImpActDll:=dImpActDll + dImpRecDll;


        end;
      {$ENDREGION}

      {$REGION 'MATERIAL'}
      ////////////////////Material////////////////////////////
      if (RtMaterial in Opciones) then
      begin
        if Clean then
        begin
          QElimina.Active:=False;
          QElimina.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QElimina.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          QElimina.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QElimina.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QElimina.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QElimina.ParamByName('Tipo').AsString:='MATERIAL';
          QElimina.ExecSQL;
        end;

        QActaCampo.Active:=False;
        QActaCampo.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QActaCampo.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
        QActaCampo.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QActaCampo.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
        QActaCampo.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QActaCampo.ParamByName('Tipo').AsString:='MATERIAL';
        QActaCampo.Open;

        with QrRecursos do
        begin
          active := false;
          {sql.text := 'select sum(bp.dCantHH) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                      'from bitacoradepersonal bp '+
                      'where bp.sContrato =:Contrato '+
                      'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" '+
                      'and bp.sNumeroActividad=:Actividad ' +
                      'group by bp.sTipoPernocta ';
                      }

          sql.text := 'SELECT aa.mDescripcion as sDescripcion,aa.dVentaMn,aa.dVentaDll,aa.sMedida,aa.iItemOrden, ' +
                      'bp.*,sum(bp.dCantidad) as Total,0 AS Ajuste,bp.sIdMaterial as sidRecurso ' + #13#10 +
                      'FROM bitacorademateriales bp' + #13#10 +
                      'inner join actividadesxorden ao '+
                      'on(ao.sContrato=bp.sContrato and ao.sNumeroOrden=bp.sNumeroOrden and ao.swbs=bp.swbs) ' +
                     'inner join actividadesxanexo aa '+
                     'on(aa.sContrato=:ContratoBarco and aa.sNumeroActividad=bp.sIdMaterial) ' +
                     'where bp.sContrato=:Contrato and bp.sNumeroOrden=:Folio and bp.swbs=:wbs '+
                     'group by bp.sIdMaterial order by aa.iItemOrden';
          parambyname('contratobarco').asstring      :=Global_Contrato_Barco;
          parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
          parambyname('Folio').asstring           := DatosActa.FieldByName('sNumeroOrden').AsString;
          ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          Open;
        end;
        tmpRecursos:='';
        QrRecursos.First;
        dImpRecMn:=0;
        dImpRecDll:=0;

        while not QrRecursos.Eof do
        begin
          with QActaCampo do
          begin
            if (tmpRecursos<>'') and (tmpRecursos<>QrRecursos.FieldByName('sAnexo').AsString) then
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sIdRecurso').AsString:='$IMPORTEMATERIAL';
              FieldByName('mdescripcion').AsString:='IMPORTE MATERIAL:';
              FieldByName('sMedida').AsString:='*';
              FieldByName('eTipo').AsString:='MATERIAL';
              FieldByName('dCantidad').AsFloat:=0;
              FieldByName('dCantidadReal').AsFloat:=0 ;
              FieldByName('dCostoMn').AsFloat:=0;
              FieldByName('dCostoDll').AsFloat:=0;
              FieldByName('dImporteMn').AsFloat:=dImpRecMn;
              FieldByName('dImportedLL').AsFloat:=dImpRecDll;
              FieldByName('iOrdenTipo').AsInteger:=3;
              FieldByName('iOrdenRecurso').AsInteger:=0;
              FieldByName('sFormulaMn').AsString:='##';
              FieldByName('sFormulaDll').AsString:='##';
              FieldByName('xCol').AsInteger:=-1;
              FieldByName('xRow').AsInteger:=-1;
              FieldByName('sAnexo').AsString:=tmpRecursos;
              post;

              dImpActMn:=dImpActMn +  dImpRecMn;
              dImpActDll:=dImpActDll + dImpRecDll;
              dImpRecMn:=0;
              dImpRecDll:=0;
              tmpRecursos:=QrRecursos.FieldByName('sAnexo').AsString;
            end;

            if not Locate('sIdRecurso',QrRecursos.FieldByName('sidRecurso').AsString,[]) then
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sidRecurso').AsString;
              FieldByName('mdescripcion').AsString:=QrRecursos.FieldByName('sDescripcion').AsString;
              FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sMedida').AsString;
              FieldByName('eTipo').AsString:='MATERIAL';
            end
            else
              Edit;

            FieldByName('dCantidad').AsFloat:= xRound((QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa) ;
            FieldByName('dCantidadReal').AsFloat:=(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat);
            FieldByName('dCostoMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
            FieldByName('dCostodLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
            FieldByName('dImporteMn').AsFloat:=xRound(((xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
            FieldByName('dImportedLL').AsFloat:=xRound((( xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
            FieldByName('iOrdenTipo').AsInteger:=4;
            FieldByName('iOrdenRecurso').AsInteger:=QrRecursos.RecNo;
            FieldByName('sFormulaMn').AsString:='##';
            FieldByName('sFormulaDll').AsString:='##';
            FieldByName('xCol').AsInteger:=-1;
            FieldByName('xRow').AsInteger:=-1;
            FieldByName('sAnexo').AsString:=QrRecursos.FieldByName('sAnexo').AsString;
            post;

            dImpRecMn:=dImpRecMn + xRound(((xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
            dImpRecDll:=dImpRecDll + xRound((( xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
          end;

          QrRecursos.Next;
        end;

        if (dImpRecMn>0) or (dImpRecDll>0) then
          with QActaCampo do
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
            FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            FieldByName('sIdRecurso').AsString:='$IMPORTEMATERIAL';
            FieldByName('mdescripcion').AsString:='IMPORTE MATERIAL:';
            FieldByName('sMedida').AsString:='*';
            FieldByName('eTipo').AsString:='MATERIAL';
            FieldByName('dCantidad').AsFloat:=0;
            FieldByName('dCantidadReal').AsFloat:=0 ;
            FieldByName('dCostoMn').AsFloat:=0;
            FieldByName('dCostoDll').AsFloat:=0;
            FieldByName('dImporteMn').AsFloat:=dImpRecMn;
            FieldByName('dImportedLL').AsFloat:=dImpRecDll;
            FieldByName('iOrdenTipo').AsInteger:=3;
            FieldByName('iOrdenRecurso').AsInteger:=0;
            FieldByName('sFormulaMn').AsString:='##';
            FieldByName('sFormulaDll').AsString:='##';
            FieldByName('xCol').AsInteger:=-1;
            FieldByName('xRow').AsInteger:=-1;
            FieldByName('sAnexo').AsString:=QrRecursos.FieldByName('sAnexo').AsString;
            post;

            dImpActMn:=dImpActMn +  dImpRecMn;
            dImpActDll:=dImpActDll + dImpRecDll;
          end;
      end;
      {$ENDREGION}


      with QActaCampo do
      begin
        Active:=False;
        ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
        ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
        ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
        ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        ParamByName('Tipo').AsString:='ACTIVIDAD';
        Open;
        if Recordcount=0 then
        begin
          Append;
          FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          FieldByName('sIdRecurso').AsString:='$IMPORTEACTIVIDAD';
          FieldByName('mdescripcion').AsString:='IMPORTE ACTIVIDAD :';
          FieldByName('sMedida').AsString:='*';
          FieldByName('eTipo').AsString:='ACTIVIDAD';
          FieldByName('dCantidad').AsFloat:=0;
          FieldByName('dCantidadReal').AsFloat:=0 ;
          FieldByName('dCostoMn').AsFloat:=0;
          FieldByName('dCostoDll').AsFloat:=0;
          FieldByName('dImporteMn').AsFloat:=dImpActMn;
          FieldByName('dImportedLL').AsFloat:=dImpActDll;
          FieldByName('iOrdenTipo').AsInteger:=5;
          FieldByName('iOrdenRecurso').AsInteger:=0;
          FieldByName('sFormulaMn').AsString:='##';
          FieldByName('sFormulaDll').AsString:='##';
          FieldByName('xCol').AsInteger:=-1;
          FieldByName('xRow').AsInteger:=-1;
          FieldByName('sAnexo').AsString:='';
          post;
        end
        else
        begin
          Edit;
          FieldByName('sFormulaMn').AsString:='##';
          FieldByName('sFormulaDll').AsString:='##';
          FieldByName('dImporteMn').AsFloat:=dImpActMn;
          FieldByName('dImportedLL').AsFloat:=dImpActDll;
          Post;
        end;
      end;


      QrActividades.Next;
    end;
  finally
    QrRecursos.Destroy;
    QrPernoctas.Destroy;
    QrAdicional.destroy;
    QrTotalHH.Destroy;
    QActaCampo.Destroy;
    QElimina.Destroy;
    QrActividades.Destroy;
  end;
end;

Procedure Load_DesgloseCostos(DatosActa:TzREadOnlyQuery;Opciones:TSetRtRecurso=[RtPersonal,RtEquipo,RtPernocta,RtBarco,RtExtraordinaria];Clean:Boolean=true);
const
  SQlRef: array[1..2,1..4] of string=(('bitacoradepersonal','personal','sIdPersonal','PERSONAL'),('bitacoradeequipos','equipos','sIdEquipo','EQUIPO'));
var
  QrActividades,
  QrRecursos,
  QrPernoctas,
  QrAdicional:TZReadOnlyQuery;
  i:Integer;
  dImporteMn,dImporteDll,
  tmpCantidad,tmpCantidadReal,
  dImpRecMn,dImpRecDll,
  dImpActMn,dImpActDll,
  dImpRecTeMn,dImpRecTeDll:Double;

  dTotalHH:Double;
  tmpRecursos:string;
  QrAnexoTE:TZReadOnlyQuery;
  sAnexosTE,sCampoAnexo:string;
  QActaCampo,QElimina:TZQuery;

begin
  QrRecursos:=TZReadOnlyQuery.Create(nil);
  QrPernoctas:=TZReadOnlyQuery.Create(nil);
  QrAdicional:=TZReadOnlyQuery.Create(nil);
  QActaCampo:=TZQuery.Create(nil);
  QElimina:=TZQuery.Create(nil);
  QrActividades:=TZReadOnlyQuery.Create(nil);
  QrAnexoTE:=TZReadOnlyQuery.Create(nil);
  try
    QrActividades.Connection:=connection.zConnection;
    QrRecursos.Connection:=connection.zConnection;
    QrPernoctas.Connection:=connection.zConnection;
    QrAdicional.Connection:=connection.zConnection;
    QrAnexoTE.Connection:=connection.zConnection;
    QActaCampo.Connection:=connection.zConnection;
    QElimina.Connection:=connection.zConnection;

    QActaCampo.SQL.Text:= 'select * from acta_campo where sContrato=:Contrato and iIdActa=:Acta and '+
                          'sNumeroOrden=:Orden and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                          'eTipo=:Tipo ';

    QElimina.SQL.Text:='delete from acta_campo where sContrato=:Contrato and iIdActa=:Acta and ' +
                       'sNumeroOrden=:Orden and swbs=:wbs and sNumeroActividad=:Actividad and ' +
                       'eTipo=:Tipo ';


    QrActividades.SQL.Text:='select * from actividadesxorden where scontrato=:Contrato and sNumeroOrden=:Orden and iNivel=0 order by iItemOrden limit 1';
    QrActividades.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
    QrActividades.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
    QrActividades.Open;
    

    if DatosActa.FieldByName('lPernocta').AsString='Si' then
    begin
      with QrPernoctas do
      begin
        Active:=False;
        SQL.Text:='select * from cuentas order by iOrden';
        Open;
      end;

      QrAdicional.Active:=false;
      {QrAdicional.SQL.Text:='select ifnull(sum(dCantidad),0) as dCantidad from bitacoradepernocta where ' +
                            'sContrato=:Contrato and sNumeroOrden=:Folio and '+
                            'sIdCuenta=:Tipo group by sContrato ';}

      QrAdicional.SQL.Text:='select ' +
                            '  (ifnull(sum(bpr.dCantidad),0) + ' +
                            '   IFNULL((SELECT ' +
                            '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                            '           FROM ' +
                            '             bitacoradeajustes ba ' +
                            '           WHERE ' +
                            '             ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                            '             ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0))as dCantidad ' +
                            'from ' +
                            '  bitacoradepernocta bpr ' +
                            'where ' +
                            '  bpr.sContrato=:Contrato and ' +
                            '  bpr.sNumeroOrden=:Folio and ' +
                            '  bpr.sIdCuenta=:Tipo ' +
                            'group by ' +
                            '  bpr.sContrato';

      QrAdicional.parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
      QrAdicional.parambyname('Folio').asstring      := DatosActa.FieldByName('sNumeroOrden').AsString;

    end;


    while not QrActividades.Eof do
    begin
      dImpActMn:=0;
      dImpActDll:=0;
      //if CCListaErrores in Opciones then
      if RtBarco in Opciones  then
      begin
        dImpRecMn:=0;
        dImpRecDll:=0;
        dImpRecTeMn :=0;
        dImpRecTeDll:=0;

        if Clean then
        begin
          QElimina.Active:=False;
          QElimina.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QElimina.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          QElimina.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QElimina.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QElimina.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QElimina.ParamByName('Tipo').AsString:='BARCO';
          QElimina.ExecSQL;
        end;

        QActaCampo.Active:=False;
        QActaCampo.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QActaCampo.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
        QActaCampo.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QActaCampo.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
        QActaCampo.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QActaCampo.ParamByName('Tipo').AsString:='BARCO';
        QActaCampo.Open;

        QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select me.*,sum(IF(mf.sFactor=NULL,me.sFactor,mf.sFactor)) as Total,tm.dVentaMn,tm.dVentaDll, ' + #13#10 +
                              'tm.sDescripcion as Descripcion, tm.iOrden '   +
                              'from movimientosdeembarcacion me' + #13#10 +
                              'inner join tiposdemovimiento tm' + #13#10 +
                              'on(me.sContrato=tm.sContrato and me.sClasificacion=tm.sIdTipoMovimiento)' + #13#10 +
                              'left join movimientosxfolios mf' + #13#10 +
                              'on(me.sContrato=mf.sContrato and me.iIdDiario = mf.iIdDiario and mf.sFolio=:Folio)' + #13#10 +
                              'where me.sContrato=:Contrato and me.sOrden=:Orden ' + #13#10 +
                              'and tm.sClasificacion="Movimiento de Barco" and tm.lAplicaNotaCampo="Si"' + #13#10 +
                              'group by tm.sIdTipoMovimiento order by tm.iOrden' ;
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QrRecursos.Open;

        while not QrRecursos.Eof do
        begin
          if QrRecursos.FieldByName('Total').AsFloat>0 then
          begin

            if not QActaCampo.Locate('sIdRecurso',QrRecursos.FieldByName('sClasificacion').AsString,[]) then
            begin
              with QActaCampo do
              begin
                Append;
                FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sClasificacion').AsString;
                FieldByName('mdescripcion').AsString:=QrRecursos.FieldByName('Descripcion').AsString;
                FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sTipo').AsString;
                FieldByName('eTipo').AsString:='BARCO';
                FieldByName('dCantidad').AsFloat:=xRound(QrRecursos.FieldByName('Total').AsFloat,6) ;
                FieldByName('dCantidadReal').AsFloat:=QrRecursos.FieldByName('Total').AsFloat ;
                FieldByName('dCostoMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
                FieldByName('dCostoDll').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
                FieldByName('dImporteMn').AsFloat:=xRound((xRound(QrRecursos.FieldByName('Total').AsFloat,6) *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
                FieldByName('dImportedLL').AsFloat:=xRound((xRound(QrRecursos.FieldByName('Total').AsFloat,6) * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
                FieldByName('iOrdenTipo').AsInteger:=0;
                FieldByName('iOrdenRecurso').AsInteger:=QrRecursos.FieldByName('iOrden').AsInteger;
                FieldByName('sFormulaMn').AsString:='##';
                FieldByName('sFormulaDll').AsString:='##';
                FieldByName('xCol').AsInteger:=-1;
                FieldByName('xRow').AsInteger:=-1;
                FieldByName('sAnexo').AsString:='';
                post;
              end;
            end
            else
            begin
              with QActaCampo do
              begin
                Edit;
                FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sTipo').AsString;
                FieldByName('dCantidad').AsFloat:=QrRecursos.FieldByName('Total').AsFloat ;
                FieldByName('dCantidadReal').AsFloat:=QrRecursos.FieldByName('Total').AsFloat ;
                FieldByName('dCostoMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
                FieldByName('dCostoDll').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
                FieldByName('dImporteMn').AsFloat:=xRound((QrRecursos.FieldByName('Total').AsFloat *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
                FieldByName('dImportedLL').AsFloat:=xRound((QrRecursos.FieldByName('Total').AsFloat* QrRecursos.FieldByName('dVentaDll').AsFloat),2);
                FieldByName('iOrdenRecurso').AsInteger:=QrRecursos.FieldByName('iOrden').AsInteger;
                post;
              end;

            end;
            dImpRecMn:=dImpRecMn + xRound((xRound(QrRecursos.FieldByName('Total').AsFloat,6) *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
            dImpRecDll:= dImpRecDll + xRound((xRound(QrRecursos.FieldByName('Total').AsFloat,6)* QrRecursos.FieldByName('dVentaDll').AsFloat),2);
          end;
          QrRecursos.Next;
        end;

        QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select * from tiposdemovimiento ' + #13#10 +
                              'where sContrato=:Contrato and sClasificacion="Movimiento de Barco" and lAplicaNotaCampo="Si"' + #13#10 +
                              'and lImprimeCeros="Si" group by sIdTipoMovimiento order by iOrden' ;
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.Open;
        while not QrRecursos.Eof do
        begin
          if not QActaCampo.Locate('sIdRecurso',QrRecursos.FieldByName('sIdTipoMovimiento').AsString,[]) then
            with QActaCampo do
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sIdTipoMovimiento').AsString;
              FieldByName('mdescripcion').AsString:=QrRecursos.FieldByName('Descripcion').AsString;
              FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sTipo').AsString;
              FieldByName('eTipo').AsString:='BARCO';
              FieldByName('dCantidad').AsFloat:=0;
              FieldByName('dCantidadReal').AsFloat:=0 ;
              FieldByName('dCostoMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
              FieldByName('dCostoDll').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
              FieldByName('dImporteMn').AsFloat:=0;
              FieldByName('dImportedLL').AsFloat:=0;
              FieldByName('iOrdenTipo').AsInteger:=0;
              FieldByName('iOrdenRecurso').AsInteger:=QrRecursos.FieldByName('iOrden').AsInteger;
              FieldByName('sFormulaMn').AsString:='##';
              FieldByName('sFormulaDll').AsString:='##';
              FieldByName('xCol').AsInteger:=-1;
              FieldByName('xRow').AsInteger:=-1;
              FieldByName('sAnexo').AsString:='';
              post;
            end;
          QrRecursos.Next;
        end;

        with QActaCampo do
        begin
          if REcordCoUnt>0 then
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
            FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            FieldByName('sIdRecurso').AsString:='$IMPORTEBARCO';
            FieldByName('mdescripcion').AsString:='IMPORTE BARCO:';
            FieldByName('sMedida').AsString:='*';
            FieldByName('eTipo').AsString:='BARCO';
            FieldByName('dCantidad').AsFloat:=0;
            FieldByName('dCantidadReal').AsFloat:=0 ;
            FieldByName('dCostoMn').AsFloat:=0;
            FieldByName('dCostoDll').AsFloat:=0;
            FieldByName('dImporteMn').AsFloat:=dImpRecMn;
            FieldByName('dImportedLL').AsFloat:=dImpRecDll;
            FieldByName('iOrdenTipo').AsInteger:=0;
            FieldByName('iOrdenRecurso').AsInteger:=0;
            FieldByName('sFormulaMn').AsString:='##';
            FieldByName('sFormulaDll').AsString:='##';
            FieldByName('xCol').AsInteger:=-1;
            FieldByName('xRow').AsInteger:=-1;
            FieldByName('sLeyendaAnexo').AsString:='ANEXO C 1.1';
            FieldByName('sAnexo').AsString:='';
            post;
          end;
        end;
        dImpActMn:=dImpActMn +  dImpRecMn;
        dImpActDll:=dImpActDll +  dImpRecDll;
      end;


      for I := 1 to Length(SQlRef) do
      begin
        tmpRecursos:='';
        tmpCantidad:=0;
        dImporteMn:=0;
        dImporteDll :=0;
        tmpCantidadReal:=0;
        dImpRecMn:=0;
        dImpRecDll:=0;
        dImpRecTeMn:=0;
        dImpRecTeDll:=0;

        sAnexosTE:='';
        sCampoAnexo:='';
        if I=1 then
        begin
          QrAnexoTE.Active:=False;
          QrAnexoTE.SQL.Text:='select GROUP_CONCAT(sanexo) as sAnexo from anexos where stipo="TIEMPO_EXTRA" ';
          QrAnexoTE.Open;
          if QrAnexoTE.RecordCount=1 then
            sAnexosTE:=QrAnexoTE.FieldByName('sAnexo').AsString;

          sCampoAnexo:= ', r.sAnexo ';
        end;







        if ((RtPersonal in Opciones) and (i=1)) or ((RtEquipo in Opciones) and (i=2)) then
        begin
          if Clean then
          begin
            QElimina.Active:=False;
            QElimina.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            QElimina.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
            QElimina.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            QElimina.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            QElimina.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            QElimina.ParamByName('Tipo').AsString:=SQlRef[i,4];
            QElimina.ExecSQL;

            {if sAnexosTE<>'' then
            begin
              QElimina.Active:=False;
              QElimina.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              QElimina.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              QElimina.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              QElimina.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              QElimina.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              QElimina.ParamByName('Tipo').AsString:='TIEMPO EXTRA';
              QElimina.ExecSQL;
            end;   }

          end;

          QActaCampo.Active:=False;
          QActaCampo.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QActaCampo.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          QActaCampo.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QActaCampo.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QActaCampo.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QActaCampo.ParamByName('Tipo').AsString:=SQlRef[i,4];
          QActaCampo.Open;

          QrRecursos.Active:=False;
          QrRecursos.SQL.Text:= 'select br.' +SQlRef[i,3] + ' as sIdRecurso,br.sDescripcion,r.sMedida, (sum(br.dCanthh) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                '           ba.sIdPartidaAnexo = '+ SQlRef[i,3] + ' AND ba.dFecha = bp.dIdFecha), 0)) as dCanthh, ' +
                                'Ifnull(sum(CAST(br.dAjuste AS CHAR)),0) as Ajuste, ' +
                                'r.dVentaMn,r.dVentaDll,r.iitemorden '+ sCampoAnexo  +
                                'from '+ SQlRef[i,1]  + ' br ' +
                                'left join ' + SQlRef[i,2] + ' r ' +
                                'on(r.sContrato=:Contrato and br.'+SQlRef[i,3]+'=r.'+ SQlRef[i,3] +') ' +
                                'where br.sContrato=:Orden and br.sNumeroOrden=:Folio  ' +
                                'group by sIdRecurso,br.dIdFecha order by r.iitemorden' ;
          QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
          QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QrRecursos.Open;

          while not QrRecursos.Eof do
          begin
            if tmpRecursos<> QrRecursos.FieldByName('sIdRecurso').AsString then
            begin
              with QActaCampo do
              begin
                if tmpRecursos<>'' then
                begin
                  if Locate('sIdRecurso',tmpRecursos,[]) then
                  begin
                    Edit;
                    FieldByName('dCantidad').AsFloat:=tmpCantidad;
                    FieldByName('dCantidadReal').AsFloat:=tmpCantidadReal;
                    FieldByName('dImporteMn').AsFloat:=xRound(dImporteMn,2);
                    FieldByName('dImportedLL').AsFloat:=xRound(dImporteDll,2);
                    Post;

                    if FieldByName('sAnexo').AsString<>'' then
                    begin
                      dImpRecTeMn:=dImpRecTeMn + xRound(dImporteMn,2);
                      dImpRecTeDll:=dImpRecTeDll + xRound(dImporteDll,2);
                    end
                    else
                    begin
                      dImpRecMn:=dImpRecMn + xRound(dImporteMn,2);
                      dImpRecDll:=dImpRecDll + xRound(dImporteDll,2);

                    end;




                  end;
                  tmpCantidad:=0;
                  tmpCantidadReal:=0;
                  dImporteMn:=0;
                  dImporteDll :=0;

                end;

                if not Locate('sIdRecurso',QrRecursos.FieldByName('sIdRecurso').AsString,[]) then
                begin
                  Append;
                  FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                  FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                  FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                  FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                  FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                  FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sIdRecurso').AsString;
                  FieldByName('mdescripcion').AsString:=QrRecursos.FieldByName('sDescripcion').AsString;
                  FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sMedida').AsString;

                  if i=1 then
                  begin
                    if sAnexosTE<>'' then
                    begin
                      if AnsiMatchText(sAnexosTE,[QrRecursos.FieldByName('sAnexo').AsString]) then
                        FieldByName('sAnexo').AsString:='TIEMPO EXTRA'
                      else
                        FieldByName('sAnexo').AsString:='';
                    end
                    else
                      FieldByName('sAnexo').AsString:='';

                  end
                  else
                    FieldByName('sAnexo').AsString:='';

                  FieldByName('eTipo').AsString:=SQlRef[i,4];




                  //FieldByName('eTipo').AsString:=SQlRef[i,4];
                  FieldByName('dCantidad').AsFloat:=xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa);
                  FieldByName('dCantidadReal').AsFloat:=(QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat );
                  FieldByName('dCostoMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
                  FieldByName('dCostodLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
                  FieldByName('dImporteMn').AsFloat:=xRound((xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
                  FieldByName('dImportedLL').AsFloat:=xRound((xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
                  FieldByName('iOrdenTipo').AsInteger:=i;
                  FieldByName('iOrdenRecurso').AsInteger:=QrRecursos.FieldByName('iitemorden').AsInteger;
                  FieldByName('sFormulaMn').AsString:='##';
                  FieldByName('sFormulaDll').AsString:='##';
                  FieldByName('xCol').AsInteger:=-1;
                  FieldByName('xRow').AsInteger:=-1;
                  post;
                end;
                tmpRecursos:=QrRecursos.FieldByName('sIdRecurso').AsString;

              end;
            end;

            tmpCantidad:=tmpCantidad + xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa);
            tmpCantidadReal:=tmpCantidadReal + (QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat );
            dImporteMn:= dImporteMn + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) *  QrRecursos.FieldByName('dVentaMn').AsFloat);
            dImporteDll:= dImporteDll + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat);

            QrRecursos.Next;
          end;

          with QActaCampo do
          begin
            if tmpRecursos<>'' then
            begin
              if Locate('sIdRecurso',tmpRecursos,[]) then
              begin
                Edit;
                FieldByName('dCantidad').AsFloat:=tmpCantidad;
                FieldByName('dCantidadReal').AsFloat:=tmpCantidadReal;
                FieldByName('dImporteMn').AsFloat:=xRound(dImporteMn,2);
                FieldByName('dImportedLL').AsFloat:=xRound(dImporteDll,2);
                Post;

                {dImpRecMn:=dImpRecMn + xRound(dImporteMn,2);
                dImpRecDll:=dImpRecDll + xRound(dImporteDll,2); }
                if FieldByName('sAnexo').AsString<>'' then
                begin
                  dImpRecTeMn:=dImpRecTeMn + xRound(dImporteMn,2);
                  dImpRecTeDll:=dImpRecTeDll + xRound(dImporteDll,2);
                end
                else
                begin
                  dImpRecMn:=dImpRecMn + xRound(dImporteMn,2);
                  dImpRecDll:=dImpRecDll + xRound(dImporteDll,2);

                end;
              end;





            end;
          end;

          with QActaCampo do
          begin
            if REcordCoUnt>0 then
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sIdRecurso').AsString:='$IMPORTE'+SQlRef[i,4];
              FieldByName('mdescripcion').AsString:='IMPORTE '+SQlRef[i,4]+':';
              FieldByName('sMedida').AsString:='*';
              FieldByName('eTipo').AsString:=SQlRef[i,4];
              FieldByName('dCantidad').AsFloat:=0;
              FieldByName('dCantidadReal').AsFloat:=0 ;
              FieldByName('dCostoMn').AsFloat:=0;
              FieldByName('dCostoDll').AsFloat:=0;
              FieldByName('dImporteMn').AsFloat:=dImpRecMn;
              FieldByName('dImportedLL').AsFloat:=dImpRecDll;
              FieldByName('iOrdenTipo').AsInteger:=i;
              FieldByName('iOrdenRecurso').AsInteger:=0;
              FieldByName('sFormulaMn').AsString:='##';
              FieldByName('sFormulaDll').AsString:='##';
              FieldByName('xCol').AsInteger:=-1;
              FieldByName('xRow').AsInteger:=-1;
              if DatosActa.FieldByName('eLugarOT').AsString='Tierra' then
              begin
                if SQlRef[i,4]='PERSONAL' then
                  FieldByName('sLeyendaAnexo').AsString:='ANEXO C-5 "PERSONAL OPT."';

                if SQlRef[i,4]='EQUIPO' then
                  FieldByName('sLeyendaAnexo').AsString:='ANEXO C-5 "EQUIPO OPT."';
              end
              else
              begin
                if SQlRef[i,4]='PERSONAL' then
                  FieldByName('sLeyendaAnexo').AsString:='ANEXO C-2 "PERSONAL OPT."';

                if SQlRef[i,4]='EQUIPO' then
                  FieldByName('sLeyendaAnexo').AsString:='ANEXO C-3 "EQUIPO OPT."';
              end;
              FieldByName('sAnexo').AsString:='';
              post;

              if (dImpRecTeMn>0) or (dImpRecTeDll>0) then
              begin
                Append;
                FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                FieldByName('sIdRecurso').AsString:='$IMPORTETIEMPOEXTRA';
                FieldByName('mdescripcion').AsString:='IMPORTE '+SQlRef[i,4]+':';
                FieldByName('sMedida').AsString:='*';
                FieldByName('eTipo').AsString:=SQlRef[i,4];
                FieldByName('dCantidad').AsFloat:=0;
                FieldByName('dCantidadReal').AsFloat:=0 ;
                FieldByName('dCostoMn').AsFloat:=0;
                FieldByName('dCostoDll').AsFloat:=0;
                FieldByName('dImporteMn').AsFloat:=dImpRecTeMn;
                FieldByName('dImportedLL').AsFloat:=dImpRecTeDll;
                FieldByName('iOrdenTipo').AsInteger:=i;
                FieldByName('iOrdenRecurso').AsInteger:=0;
                FieldByName('sFormulaMn').AsString:='##';
                FieldByName('sFormulaDll').AsString:='##';
                FieldByName('xCol').AsInteger:=-1;
                FieldByName('xRow').AsInteger:=-1;
                if DatosActa.FieldByName('eLugarOT').AsString='Tierra' then
                begin
                  if SQlRef[i,4]='PERSONAL' then
                    FieldByName('sLeyendaAnexo').AsString:='ANEXO C-5 "PERSONAL OPT."';

                  if SQlRef[i,4]='EQUIPO' then
                    FieldByName('sLeyendaAnexo').AsString:='ANEXO C-5 "EQUIPO OPT."';
                end
                else
                begin
                  if SQlRef[i,4]='PERSONAL' then
                    FieldByName('sLeyendaAnexo').AsString:='ANEXO C-2 "PERSONAL OPT."';

                  if SQlRef[i,4]='EQUIPO' then
                    FieldByName('sLeyendaAnexo').AsString:='ANEXO C-3 "EQUIPO OPT."';
                end;
                FieldByName('sAnexo').AsString:='TIEMPO EXTRA';
                post;



              end;

            end;
          end;


          dImpActMn:=dImpActMn +  dImpRecMn;
          dImpActDll:=dImpActDll + dImpRecDll;
        end;

      end;

      if DatosActa.FieldByName('lPernocta').AsString='Si' then
        if (RtPernocta in Opciones) then
        begin
          dImpRecMn:=0;
          dImpRecDll:=0;
          if Clean then
          begin
            QElimina.Active:=False;
            QElimina.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            QElimina.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
            QElimina.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            QElimina.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            QElimina.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            QElimina.ParamByName('Tipo').AsString:='PERNOCTA';
            QElimina.ExecSQL;
          end;

          QActaCampo.Active:=False;
          QActaCampo.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QActaCampo.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          QActaCampo.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QActaCampo.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QActaCampo.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QActaCampo.ParamByName('Tipo').AsString:='PERNOCTA';
          QActaCampo.Open;

          with QrRecursos do
          begin
            active := false;
            {sql.text := 'select sum(bp.dCantHH) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                        'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                        'from bitacoradepersonal bp '+
                        'where bp.sContrato =:Contrato '+
                        'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" '+
                        'and bp.sNumeroActividad=:Actividad ' +
                        'group by bp.sTipoPernocta ';
                        }
// Rangel
            sql.text := 'SELECT sum(bp.dCantidad) as Total, ' +
                        'IFNULL(sum(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste ' + #13#10 +
                        'FROM' + #13#10 +
                        '( SELECT bp3.scontrato,bp3.didfecha, bp3.iIdDiario,bp3.sNumeroOrden,bp3.stipopernocta, ' + #13#10 +
                        '(sum(xround(bp3.dCanthh,2)) + ' +
                        ' IFNULL((SELECT ' +
                        '           SUM(ROUND(CAST(ba.dAjuste AS CHAR), 2)) ' +
                        '         FROM ' +
                        '           bitacoradeajustes ba ' +
                        '         WHERE ' +
                        '           ba.sContrato = bp3.sContrato AND ba.sNumeroOrden = bp3.sNumeroOrden AND ' +
                        '           ba.sNumeroActividad = bp3.sNumeroActividad AND ' +
                        '           ba.sIdPartidaAnexo = bp3.sIdPersonal AND ba.dFecha = bp3.dIdFecha), 0)) as dCantidad, ' +
                        '(IFNULL(sum(CAST(bp3.dAjuste AS CHAR)), 0) + ' +
                        ' IFNULL((SELECT ' +
                        '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                        '         FROM ' +
                        '           bitacoradeajustes ba ' +
                        '         WHERE ' +
                        '           ba.sContrato = bp3.sContrato AND ba.sNumeroOrden = bp3.sNumeroOrden AND ' +
                        '           ba.sNumeroActividad = bp3.sNumeroActividad AND ' +
                        '           ba.sIdPartidaAnexo = bp3.sIdPersonal AND ba.dFecha = bp3.dIdFecha), 0)) AS Ajuste ' + #13#10 +
                        'FROM bitacoradepersonal bp3 ' + #13#10 +
                        'WHERE bp3.scontrato =:Contrato ' + #13#10 +
                        'and bp3.stipopernocta=:Tipo and bp3.sNumeroOrden=:Folio and bp3.lAplicaPernocta="Si"' + #13#10 +
                        'group by bp3.dIdFecha, bp3.sNumeroOrden, bp3.stipopernocta,bp3.sIdPersonal' + #13#10 +
                        ') bp ' + #13#10 +
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden )' + #13#10 +
                        'inner join cuentas c on (c.sIdCuenta = bp.stipopernocta) ' + #13#10 +
                        'group by bp.sNumeroOrden, bp.stipopernocta order by c.sIdCuenta, o.iOrden, bp.dIdFecha';

            parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
            parambyname('Folio').asstring           := DatosActa.FieldByName('sNumeroOrden').AsString;

          end;
          dImpRecMn:=0;
          dImpRecDll:=0;

          QrPernoctas.First;
          while not QrPernoctas.Eof do
          begin
            dTotalHH:=0;
            QrRecursos.Active:=False;
            QrRecursos.ParamByName('Tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
            QrRecursos.Open;

            QrAdicional.Active:=False;
            QrAdicional.ParamByName('Tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
            QrAdicional.Open;

            if QrAdicional.RecordCount>0 then
              dTotalHH:=QrAdicional.FieldByName('dCantidad').AsFloat;



            while not QrRecursos.Eof do
            begin
              with QActaCampo do
                if not Locate('sIdRecurso',QrPernoctas.FieldByName('sIdPernocta').AsString,[]) then
                begin
                  Append;
                  FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                  FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                  FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                  FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                  FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                  FieldByName('sIdRecurso').AsString:=QrPernoctas.FieldByName('sIdPernocta').AsString;
                  FieldByName('mdescripcion').AsString:=QrPernoctas.FieldByName('sDescripcion').AsString;
                  FieldByName('sMedida').AsString:=QrPernoctas.FieldByName('sMedida').AsString;
                  FieldByName('eTipo').AsString:='PERNOCTA';
                  FieldByName('dCantidad').AsFloat:=dTotalHH + xRound((QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa) ;
                  FieldByName('dCantidadReal').AsFloat:=dTotalHH + (QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat);
                  FieldByName('dCostoMn').AsFloat:=QrPernoctas.FieldByName('dVentaMn').AsFloat;
                  FieldByName('dCostodLL').AsFloat:=QrPernoctas.FieldByName('dVentaDll').AsFloat;
                  FieldByName('dImporteMn').AsFloat:=xRound(((dTotalHH + xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
                  FieldByName('dImportedLL').AsFloat:=xRound(((dTotalHH + xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);
                  FieldByName('iOrdenTipo').AsInteger:=3;
                  FieldByName('iOrdenRecurso').AsInteger:=QrPernoctas.FieldByName('iorden').AsInteger;
                  FieldByName('sFormulaMn').AsString:='##';
                  FieldByName('sFormulaDll').AsString:='##';
                  FieldByName('xCol').AsInteger:=-1;
                  FieldByName('xRow').AsInteger:=-1;
                  FieldByName('sAnexo').AsString:='';
                  post;

                  dImpRecMn:=dImpRecMn + xRound(((dTotalHH + xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
                  dImpRecDll:=dImpRecDll + xRound(((dTotalHH + xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);
                end;
              QrRecursos.Next;
            end;

            if QrRecursos.RecordCount=0 then
            begin
              with QActaCampo do
                if not Locate('sIdRecurso',QrPernoctas.FieldByName('sIdPernocta').AsString,[]) then
                begin
                  Append;
                  FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
                  FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
                  FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
                  FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                  FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                  FieldByName('sIdRecurso').AsString:=QrPernoctas.FieldByName('sIdPernocta').AsString;
                  FieldByName('mdescripcion').AsString:=QrPernoctas.FieldByName('sDescripcion').AsString;
                  FieldByName('sMedida').AsString:=QrPernoctas.FieldByName('sMedida').AsString;
                  FieldByName('eTipo').AsString:='PERNOCTA';
                  FieldByName('dCantidad').AsFloat:=dTotalHH;
                  FieldByName('dCantidadReal').AsFloat:=dTotalHH;
                  FieldByName('dCostoMn').AsFloat:=QrPernoctas.FieldByName('dVentaMn').AsFloat;
                  FieldByName('dCostodLL').AsFloat:=QrPernoctas.FieldByName('dVentaDll').AsFloat;
                  FieldByName('dImporteMn').AsFloat:=xRound((dTotalHH *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
                  FieldByName('dImportedLL').AsFloat:=xRound((dTotalHH   * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);
                  FieldByName('iOrdenTipo').AsInteger:=3;
                  FieldByName('iOrdenRecurso').AsInteger:=QrPernoctas.FieldByName('iorden').AsInteger;
                  FieldByName('sFormulaMn').AsString:='##';
                  FieldByName('sFormulaDll').AsString:='##';
                  FieldByName('xCol').AsInteger:=-1;
                  FieldByName('xRow').AsInteger:=-1;
                  FieldByName('sAnexo').AsString:='';
                  post;
                  dImpRecMn:=dImpRecMn+ xRound((dTotalHH *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
                  dImpRecDll:=dImpRecDll +  xRound((dTotalHH   * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);
                end;
            end;
            QrPernoctas.Next;
          end;

          with QActaCampo do
          begin
            if REcordCoUnt>0 then
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sIdRecurso').AsString:='$IMPORTEPERNOCTA';
              FieldByName('mdescripcion').AsString:='IMPORTE PERNOCTA :';
              FieldByName('sMedida').AsString:='*';
              FieldByName('eTipo').AsString:='PERNOCTA';
              FieldByName('dCantidad').AsFloat:=0;
              FieldByName('dCantidadReal').AsFloat:=0 ;
              FieldByName('dCostoMn').AsFloat:=0;
              FieldByName('dCostoDll').AsFloat:=0;
              FieldByName('dImporteMn').AsFloat:=dImpRecMn;
              FieldByName('dImportedLL').AsFloat:=dImpRecDll;
              FieldByName('iOrdenTipo').AsInteger:=3;
              FieldByName('iOrdenRecurso').AsInteger:=0;
              FieldByName('sFormulaMn').AsString:='##';
              FieldByName('sFormulaDll').AsString:='##';
              FieldByName('xCol').AsInteger:=-1;
              FieldByName('xRow').AsInteger:=-1;
              FieldByName('sLeyendaAnexo').AsString:='ANEXO C-4 "SERV. DE HOTEL."';
              FieldByName('sAnexo').AsString:='';
              post;
            end;
          end;


          dImpActMn:=dImpActMn +  dImpRecMn;
          dImpActDll:=dImpActDll + dImpRecDll;


        end;

      if (RtMaterial in Opciones) then
      begin
        if Clean then
        begin
          QElimina.Active:=False;
          QElimina.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          QElimina.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          QElimina.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          QElimina.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QElimina.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QElimina.ParamByName('Tipo').AsString:='MATERIAL';
          QElimina.ExecSQL;
        end;

        QActaCampo.Active:=False;
        QActaCampo.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QActaCampo.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
        QActaCampo.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QActaCampo.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
        QActaCampo.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QActaCampo.ParamByName('Tipo').AsString:='MATERIAL';
        QActaCampo.Open;

        with QrRecursos do
        begin
          active := false;
          {sql.text := 'select sum(bp.dCantHH) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                      'from bitacoradepersonal bp '+
                      'where bp.sContrato =:Contrato '+
                      'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" '+
                      'and bp.sNumeroActividad=:Actividad ' +
                      'group by bp.sTipoPernocta ';
                      }

          sql.text := 'SELECT aa.mDescripcion as sDescripcion,aa.dVentaMn,aa.dVentaDll,aa.sMedida,aa.iItemOrden, ' +
                      'bp.*,sum(bp.dCantidad) as Total,0 AS Ajuste,bp.sIdMaterial as sidRecurso ' + #13#10 +
                      'FROM bitacorademateriales bp' + #13#10 +
                      'inner join actividadesxorden ao '+
                      'on(ao.sContrato=bp.sContrato and ao.sNumeroOrden=bp.sNumeroOrden and ao.swbs=bp.swbs) ' +
                     'inner join actividadesxanexo aa '+
                     'on(aa.sContrato=:ContratoBarco and aa.sNumeroActividad=bp.sIdMaterial) ' +
                     'where bp.sContrato=:Contrato and bp.sNumeroOrden=:Folio '+
                     'group by bp.sIdMaterial order by aa.iItemOrden';
          parambyname('contratobarco').asstring      :=Global_Contrato_Barco;
          parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
          parambyname('Folio').asstring           := DatosActa.FieldByName('sNumeroOrden').AsString;
          //ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          Open;
        end;
        tmpRecursos:='';
        QrRecursos.First;
        dImpRecMn:=0;
        dImpRecDll:=0;

        while not QrRecursos.Eof do
        begin
          with QActaCampo do
          begin
            if (tmpRecursos<>'') and (tmpRecursos<>QrRecursos.FieldByName('sAnexo').AsString) then
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sIdRecurso').AsString:='$IMPORTEMATERIAL';
              FieldByName('mdescripcion').AsString:='IMPORTE MATERIAL:';
              FieldByName('sMedida').AsString:='*';
              FieldByName('eTipo').AsString:='MATERIAL';
              FieldByName('dCantidad').AsFloat:=0;
              FieldByName('dCantidadReal').AsFloat:=0 ;
              FieldByName('dCostoMn').AsFloat:=0;
              FieldByName('dCostoDll').AsFloat:=0;
              FieldByName('dImporteMn').AsFloat:=dImpRecMn;
              FieldByName('dImportedLL').AsFloat:=dImpRecDll;
              FieldByName('iOrdenTipo').AsInteger:=3;
              FieldByName('iOrdenRecurso').AsInteger:=0;
              FieldByName('sFormulaMn').AsString:='##';
              FieldByName('sFormulaDll').AsString:='##';
              FieldByName('xCol').AsInteger:=-1;
              FieldByName('xRow').AsInteger:=-1;
              FieldByName('sAnexo').AsString:=tmpRecursos;
              FieldByName('sLeyendaAnexo').AsString:='ANEXO ' + tmpRecursos;
              post;

              dImpActMn:=dImpActMn +  dImpRecMn;
              dImpActDll:=dImpActDll + dImpRecDll;
              dImpRecMn:=0;
              dImpRecDll:=0;
              tmpRecursos:=QrRecursos.FieldByName('sAnexo').AsString;
            end;

            if not Locate('sIdRecurso',QrRecursos.FieldByName('sidRecurso').AsString,[]) then
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
              FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sidRecurso').AsString;
              FieldByName('mdescripcion').AsString:=QrRecursos.FieldByName('sDescripcion').AsString;
              FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sMedida').AsString;
              FieldByName('eTipo').AsString:='MATERIAL';
            end
            else
              Edit;

            FieldByName('dCantidad').AsFloat:= xRound((QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa) ;
            FieldByName('dCantidadReal').AsFloat:=(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat);
            FieldByName('dCostoMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
            FieldByName('dCostodLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
            FieldByName('dImporteMn').AsFloat:=xRound(((xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
            FieldByName('dImportedLL').AsFloat:=xRound((( xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
            FieldByName('iOrdenTipo').AsInteger:=4;
            FieldByName('iOrdenRecurso').AsInteger:=QrRecursos.RecNo;
            FieldByName('sFormulaMn').AsString:='##';
            FieldByName('sFormulaDll').AsString:='##';
            FieldByName('xCol').AsInteger:=-1;
            FieldByName('xRow').AsInteger:=-1;
            FieldByName('sAnexo').AsString:=QrRecursos.FieldByName('sAnexo').AsString;
            post;

            dImpRecMn:=dImpRecMn + xRound(((xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
            dImpRecDll:=dImpRecDll + xRound((( xRound(QrRecursos.FieldByName('total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
          end;

          QrRecursos.Next;
        end;

        if (dImpRecMn>0) or (dImpRecDll>0) then
          with QActaCampo do
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
            FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            FieldByName('sIdRecurso').AsString:='$IMPORTEMATERIAL';
            FieldByName('mdescripcion').AsString:='IMPORTE MATERIAL:';
            FieldByName('sMedida').AsString:='*';
            FieldByName('eTipo').AsString:='MATERIAL';
            FieldByName('dCantidad').AsFloat:=0;
            FieldByName('dCantidadReal').AsFloat:=0 ;
            FieldByName('dCostoMn').AsFloat:=0;
            FieldByName('dCostoDll').AsFloat:=0;
            FieldByName('dImporteMn').AsFloat:=dImpRecMn;
            FieldByName('dImportedLL').AsFloat:=dImpRecDll;
            FieldByName('iOrdenTipo').AsInteger:=3;
            FieldByName('iOrdenRecurso').AsInteger:=0;
            FieldByName('sFormulaMn').AsString:='##';
            FieldByName('sFormulaDll').AsString:='##';
            FieldByName('xCol').AsInteger:=-1;
            FieldByName('xRow').AsInteger:=-1;
            FieldByName('sAnexo').AsString:=QrRecursos.FieldByName('sAnexo').AsString;
            FieldByName('sLeyendaAnexo').AsString:='ANEXO ' + QrRecursos.FieldByName('sAnexo').AsString;
            post;

            dImpActMn:=dImpActMn +  dImpRecMn;
            dImpActDll:=dImpActDll + dImpRecDll;
          end;
      end;

      ///  //////////////////////////////////////////


      with QActaCampo do
      begin
        Active:=False;
        ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
        ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
        ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
        ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        ParamByName('Tipo').AsString:='ACTIVIDAD';
        Open;
        if Recordcount=0 then
        begin
          Append;
          FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          FieldByName('iIdActa').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
          FieldByName('sNumeroOrden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          FieldByName('swbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          FieldByName('sIdRecurso').AsString:='$IMPORTEACTIVIDAD';
          FieldByName('mdescripcion').AsString:='IMPORTE ACTIVIDAD :';
          FieldByName('sMedida').AsString:='*';
          FieldByName('eTipo').AsString:='ACTIVIDAD';
          FieldByName('dCantidad').AsFloat:=0;
          FieldByName('dCantidadReal').AsFloat:=0 ;
          FieldByName('dCostoMn').AsFloat:=0;
          FieldByName('dCostoDll').AsFloat:=0;
          FieldByName('dImporteMn').AsFloat:=dImpActMn;
          FieldByName('dImportedLL').AsFloat:=dImpActDll;
          FieldByName('iOrdenTipo').AsInteger:=5;
          FieldByName('iOrdenRecurso').AsInteger:=0;
          FieldByName('sFormulaMn').AsString:='##';
          FieldByName('sFormulaDll').AsString:='##';
          FieldByName('xCol').AsInteger:=-1;
          FieldByName('xRow').AsInteger:=-1;
          FieldByName('sAnexo').AsString:='';
          post;
        end
        else
        begin
          Edit;
          FieldByName('sFormulaMn').AsString:='##';
          FieldByName('sFormulaDll').AsString:='##';
          FieldByName('dImporteMn').AsFloat:=dImpActMn;
          FieldByName('dImportedLL').AsFloat:=dImpActDll;
          Post;
        end;
      end;


      QrActividades.Next;
    end;
  finally
    QrRecursos.Destroy;
    QrPernoctas.Destroy;                                              
    QrAdicional.destroy;
    QActaCampo.Destroy;
    QElimina.Destroy;
    QrActividades.Destroy;
    QrAnexoTE.Destroy;
  end;
end;

Procedure Load_Ajuste(DatosActa:TzREadOnlyQuery;Opciones:TSetRtRecurso=[RtPersonal,RtEquipo,RtPernocta,RtBarco,RtExtraordinaria]);
var
  QrDesgCostos:TZReadOnlyQuery;
  QNotaCampo:TZQuery;
  DifCant,DifMontoMn,DifMontoDll:Double;
  swbs:string;
begin
  QrDesgCostos:=TZReadOnlyQuery.Create(nil);
  QNotaCampo:=TZQuery.Create(nil);
  try
    QrDesgCostos.Connection:=connection.zConnection;
    QNotaCampo.Connection:=connection.zConnection;

    QrDesgCostos.SQL.Text:= 'select ac.* from acta_campo ac  ' +
                            'inner join actividadesxorden ao ' +
                            'on(ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ac.swbs=ao.swbs) ' +
                            'where ac.sContrato=:Contrato and ac.iIdActa=:Acta and '+
                            'ac.sNumeroOrden=:Orden and ao.stipoactividad="Paquete" and ac.sIdRecurso not like "$IMPORTE%" '   +
                            'order by ac.iOrdenTipo,ac.iOrdenRecurso' ;
    QrDesgCostos.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
    QrDesgCostos.ParamByName('Acta').AsInteger:= DatosActa.FieldByName('iIdActa').AsInteger;
    QrDesgCostos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
    QrDesgCostos.Open;

    while not QrDesgCostos.eof do
    begin
      DifCant:=0;
      QNotaCampo.Active:=False;
      QNotaCampo.SQL.Text:='select sum(ac.dCantidad) as dCantidad from acta_campo ac  ' +
                            'inner join actividadesxorden ao ' +
                            'on(ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ac.swbs=ao.swbs) ' +
                            'where ac.sContrato=:Contrato and ac.iIdActa=:Acta and '+
                            'ac.sNumeroOrden=:Orden and ao.stipoactividad="Actividad" and '   +
                            'ac.sIdRecurso=:Recurso and eTipo=:Tipo '+
                            'group by ac.sIdRecurso ' ;
      QNotaCampo.ParamByName('Contrato').AsString:=QrDesgCostos.FieldByName('sContrato').AsString;
      QNotaCampo.ParamByName('Acta').AsInteger:= QrDesgCostos.FieldByName('iIdActa').AsInteger;
      QNotaCampo.ParamByName('Orden').AsString:=QrDesgCostos.FieldByName('sNumeroOrden').AsString;
      QNotaCampo.ParamByName('Recurso').AsString:=QrDesgCostos.FieldByName('sIdRecurso').AsString;
      QNotaCampo.ParamByName('Tipo').AsString:=QrDesgCostos.FieldByName('eTipo').AsString;
      QNotaCampo.Open;
      if QNotaCampo.RecordCount=1 then
        DifCant:=QrDesgCostos.FieldByName('dCantidad').AsFloat - QNotaCampo.FieldByName('dCantidad').AsFloat;

      if DifCant<>0 then
      begin
        QNotaCampo.Active:=False;
        QNotaCampo.SQL.Text:= 'select * from acta_campo' + #13#10 +
                              'where iidacta=:Acta and sidrecurso=:recurso and swbs<>:wbs' + #13#10 +
                              'order by dCantidad desc limit 1';
        QNotaCampo.ParamByName('Acta').AsInteger:= QrDesgCostos.FieldByName('iIdActa').AsInteger;
        QNotaCampo.ParamByName('Recurso').AsString:=QrDesgCostos.FieldByName('sIdRecurso').AsString;
        QNotaCampo.ParamByName('wbs').AsString:=QrDesgCostos.FieldByName('swbs').AsString;
        QNotaCampo.Open;
        DifCant:=QNotaCampo.FieldByName('dCantidad').AsFloat +  DifCant;
        DifMontoMn:=xRound((DifCant * QNotaCampo.FieldByName('dCostoMn').AsFloat),Mantisa) - QNotaCampo.FieldByName('dImporteMn').AsFloat;
        DifMontoDll:=xRound((DifCant * QNotaCampo.FieldByName('dCostoDll').AsFloat),Mantisa) - QNotaCampo.FieldByName('dImporteDll').AsFloat ;
        swbs:=QNotaCampo.FieldByName('swbs').AsString;

        QNotaCampo.Edit;
        QNotaCampo.FieldByName('dCantidad').AsFloat:=DifCant;
        QNotaCampo.FieldByName('dImporteMn').AsFloat:=xRound((DifCant * QNotaCampo.FieldByName('dCostoMn').AsFloat),Mantisa);
        QNotaCampo.FieldByName('dImporteDll').AsFloat:=xRound((DifCant * QNotaCampo.FieldByName('dCostoDll').AsFloat),Mantisa);
        QNotaCampo.Post;

        QNotaCampo.Active:=False;
        QNotaCampo.SQL.Text:= 'select * from acta_campo' + #13#10 +
                              'where iidacta=:Acta and Find_In_set(eTipo,:Tipos) and swbs=:wbs' + #13#10 +
                              'and sIdRecurso like "$IMPORTE%"';
        QNotaCampo.ParamByName('Acta').AsInteger:= QrDesgCostos.FieldByName('iIdActa').AsInteger;
        QNotaCampo.ParamByName('wbs').AsString:=swbs;
        QNotaCampo.ParamByName('tipos').AsString:='ACTIVIDAD,' + QrDesgCostos.FieldByName('eTipo').AsString;
        QNotaCampo.Open;

        while not QNotaCampo.Eof do
        begin
          QNotaCampo.Edit;
          QNotaCampo.FieldByName('dImporteMn').AsFloat:=QNotaCampo.FieldByName('dImporteMn').AsFloat +  DifMontoMn;
          QNotaCampo.FieldByName('dImportedll').AsFloat:=QNotaCampo.FieldByName('dImportedll').AsFloat + DifMontoDll;
          QNotaCampo.Post;
          QNotaCampo.Next;
        end;
      end;

      QrDesgCostos.Next;
    end;
  finally
    QrDesgCostos.Destroy;
    QNotaCampo.Destroy;
  end;
end;

Procedure ActaEx_NotaCampo(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
var
  QrActividades,
  QrAvance,
 { QImportes,  }
  QRecursos  :TZReadOnlyQuery;

  Td_partidasAnexo,
  Td_Distribucion_detalle2,
  Td_resumenMaterial: TfrxDBDataset;
  dsUnion:TDataSource;

begin
  QRecursos:=TZReadOnlyQuery.Create(nil);
  QrActividades:=TZReadOnlyQuery.Create(nil);
  QrAvance:=TZReadOnlyQuery.Create(nil);
  try

    Td_partidasAnexo:= TfrxDBDataset.Create(nil);
    Td_Distribucion_detalle2:= TfrxDBDataset.Create(nil);
    Td_resumenMaterial:= TfrxDBDataset.Create(nil);
    dsUnion:=TDataSource.Create(nil);

    QrActividades.Connection:=connection.zConnection;
    QRecursos.Connection:=connection.zConnection;

    Td_partidasAnexo.UserName:='Td_partidasAnexo';
    Td_Distribucion_detalle2.UserName:='Td_Distribucion_detalle2';
    Td_resumenMaterial.UserName:='Td_resumenMaterial';

    QrActividades.SQL.Text:='select ao.* from actividadesxorden ao inner join acta_campo ac '+
                        'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                        'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                        'where ac.iIdActa=:Acta and ao.sTipoActividad=:Tipo ' +
                        'group by ao.swbs order by ao.iItemOrden';

    dsUnion.DataSet:=QrActividades;

    QrAvance.Connection:=connection.zConnection;
    QrAvance.SQL.Text:= 'select b.*, ' +
                        '( SELECT (ifnull(sum(ba.dAvance), 0)) ' +
                                          '		FROM ' +
                                          '			bitacoradeactividades AS ba ' +
                                          '		WHERE ' +
                                          '			ba.sContrato = b.sContrato ' +
                                          '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                          '		AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento ' +
                                          '		AND ba.swbs = b.swbs ' +
                                          '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                          '		AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
                                          '   < cast(b.sHoraInicio AS Time))  )	) AS AvanceAnterior ' +
                        ' from bitacoradeactividades b' + #13#10 +
                        'where b.sContrato=:Contrato and b.snumeroorden=:Orden ' + #13#10 +
                        'and b.sIdTipoMovimiento=:Tipo' + #13#10 +
                        'order by b.snumeroactividad,b.didfecha,time(b.sHoraInicio)' ;
    QrAvance.ParamByName('contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
    QrAvance.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
    QrAvance.ParamByName('tipo').AsString:='ED';
    QrAvance.Open;

    {QImportes.SQL.Text:='select ac.* from acta_campo ac inner join actividadesxorden ao ' +
                    'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                    'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                    'where iIdActa=:Acta and ao.sTipoActividad=:Tipo and '   +
                    'ac.sIdRecurso like "$IMPORTE%" and ac.eTipo<>"ACTIVIDAD" order by ao.iItemOrden,iOrdenTipo';

    QRecursos.SQL.Text:='select * from acta_campo where iIdActa=:Acta and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                        'eTipo=:Tipo and sIdRecurso not like "$IMPORTE%" order by iOrdenRecurso';   }

    {QRecursos.SQL.Text:='select ac.sContrato,ac.sNumeroActividad,etipo as sTipo,ac.mDescripcion as sDescripcion, ac.dCostoMn as dVentaMn,'+
                    'ac.dCostoDll as dVentaDll from acta_campo ac inner join actividadesxorden ao ' +
                    'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                    'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                    'where iIdActa=:Acta and ao.sTipoActividad=:Tipo and '   +
                    'ac.sIdRecurso not like "$IMPORTE%" order by ao.iItemOrden,ac.iOrdenTipo,ac.iOrdenRecurso'; }
    if DatosActa.FieldByName('lMaterial').AsString='Si' then
      QRecursos.SQL.Text:='select sContrato,swbs,sNumeroActividad,sIdRecurso,sTipo,sDescripcion,dCantidad,sMedida,dVentaMn,dVentaDll,dImporteMn,'+
                          'dImporteDll,ImporteMn,ImporteDll,ActImporteMn,ActImporteDll,iItemOrden,iOrdenTipo,iOrdenRecurso' + #13#10 +
                          'FROM(' + #13#10 + 
                          '' + #13#10 + 
                          '(' + #13#10 +
                          'select ac.sContrato,ac.swbs,ac.sNumeroActividad,ac.sIdRecurso,ac.etipo as sTipo,ac.mDescripcion as sDescripcion,' + #13#10 +
                          'ac.dCantidad,ac.sMedida,ac.dCostoMn as dVentaMn,ac.dCostoDll as dVentaDll,ac.dImporteMn,ac.dImporteDll,' + #13#10 +
                          'ac2.dImporteMn as ImporteMn,ac2.dImportedll as ImporteDll,ac3.dImporteMn as ActImporteMn,ac3.dImportedll as ActImporteDll, ' +
                          'ao.iItemOrden,ac.iOrdenTipo,ac.iOrdenRecurso from acta_campo ac inner join actividadesxorden ao' + #13#10 +
                          'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and' + #13#10 +
                          'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad)' + #13#10 +
                          'inner join acta_campo ac2' + #13#10 +
                          'on(ac2.iIdActa=ac.iIdActa and ac2.sWbs=ac.swbs and ac2.eTipo=ac.eTipo and ac2.sIdRecurso like "$IMPORTE%") '  +
                          'inner join acta_campo ac3' + #13#10 +
                          'on(ac3.iIdActa=ac.iIdActa and ac3.sWbs=ac.swbs and ac3.eTipo="ACTIVIDAD") ' +
                          'where ac.iIdActa=:Acta and ao.sTipoActividad=:Tipo and' + #13#10 +
                          'ac.sIdRecurso not like "$IMPORTE%" order by ao.iItemOrden,ac.iOrdenTipo,ac.iOrdenRecurso)' + #13#10 +
                          'UNION' + #13#10 + 
                          '(select bp.sContrato,ao2.swbs,bp.sNumeroActividad,bp.sTrazabilidad as sIdRecurso, "MATERIALES" as sTipo,i.mDescripcion as sDescripcion,' + #13#10 +
                          'xround(sum( bp.dCantidad),2) as dCantidad,bp.sMedida,0.00 as dVentaMn,0.00 as dVentaDll,0.00 as dImporteMN,0.00 as dImporteDll,' + #13#10 +
                          '0.00 as ImporteMn,0.00 as ImporteDll,ac3.dImporteMn as ActImporteMn,ac3.dImportedll as ActImporteDll, '+
                          'ao2.iItemOrden,4 as iOrdenTipo,1 as iOrdenRecurso' + #13#10 +
                           'from acta_material bp' + #13#10 +
                           'inner join insumos i' + #13#10 + 
                           'on(i.sContrato=bp.sContrato and i.sIdInsumo=bp.sIdInsumo)' + #13#10 +
                           'inner join acta_entrega ae' + #13#10 +
                           'on (ae.sContrato=bp.sContrato and ae.iIdActa=bp.iIdActa)' + #13#10 +
                           'inner join actividadesxorden ao2' + #13#10 +
                           'on(ao2.sContrato=bp.sContrato and ae.sNumeroorden=ao2.sNumeroOrden and ao2.sNumeroActividad=bp.sNumeroActividad)' + #13#10 +
                           'inner join acta_campo ac3' + #13#10 + 
                           'on(ac3.iIdActa=bp.iIdActa and ac3.sWbs=ao2.swbs and ac3.eTipo="ACTIVIDAD") ' +

                           'where  bp.iIdActa=:Acta' + #13#10 +
                           'group by bp.sIdInsumo order by bp.sIdInsumo' + #13#10 +
                           '' + #13#10 +
                           '' + #13#10 +
                          ')' + #13#10 +
                          '' + #13#10 + 
                          ') t' + #13#10 +
                          'order by iItemOrden,iOrdenTipo,iOrdenRecurso'
    else
      QRecursos.SQL.Text:='select ac.*,ac.etipo as sTipo,ac.mDescripcion as sDescripcion, ac.dCostoMn as dVentaMn,'+
                    'ac.dCostoDll as dVentaDll,ac2.dImporteMn as ImporteMn,ac2.dImportedll as ImporteDll '+ 
                    ',ac3.dImporteMn as ActImporteMn,ac3.dImportedll as ActImporteDll ' +
                    'from acta_campo ac inner join actividadesxorden ao ' +
                    'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                    'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                    'inner join acta_campo ac2' + #13#10 +
                    'on(ac2.iIdActa=ac.iIdActa and ac2.sWbs=ac.swbs and ac2.eTipo=ac.eTipo and ac2.sIdRecurso like "$IMPORTE%") '  +
                    'inner join acta_campo ac3' + #13#10 +
                    'on(ac3.iIdActa=ac.iIdActa and ac3.sWbs=ac.swbs and ac3.eTipo="ACTIVIDAD") ' +
                    'where ac.iIdActa=:Acta and ao.sTipoActividad=:Tipo and '   +
                    'ac.sIdRecurso not like "$IMPORTE%" order by ao.iItemOrden,ac.iOrdenTipo,ac.iOrdenRecurso';


    QrActividades.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
    QrActividades.ParamByName('Tipo').AsString:='Actividad';
    QrActividades.Open;

    QRecursos.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
    QRecursos.ParamByName('Tipo').AsString:='Actividad';
    QRecursos.Open;


    QrAvance.MasterSource:=dsUnion;
    QrAvance.MasterFields:='swbs';
    QrAvance.LinkedFields:='swbs';

    QRecursos.MasterSource:=dsUnion;
    QRecursos.MasterFields:='swbs';
    QRecursos.LinkedFields:='swbs';

    Td_partidasAnexo.DataSet:= QrActividades;
    Td_Distribucion_detalle2.DataSet:= QrAvance;
    Td_resumenMaterial.DataSet:= QRecursos;

    Td_partidasAnexo.FieldAliases.Clear;
    Td_Distribucion_detalle2.FieldAliases.Clear;
    Td_resumenMaterial.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_partidasAnexo);
    Reporte.DataSets.Add(Td_Distribucion_detalle2);
    Reporte.DataSets.Add(Td_resumenMaterial);


  finally

  end;
end;

Procedure ActaPdf_NotaCampo(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
const
  SQlRef: array[1..2,1..4] of string=(('bitacoradepersonal','personal','sIdPersonal','PERSONAL'),('bitacoradeequipos','equipos','sIdEquipo','EQUIPO'));
var
  QrActividades,
  QrRecursos,QrAvance,
  QrPernoctas,
  QrAdicional,
  QrTotalHH:TZReadOnlyQuery;
  mryRecursos:TjvMemoryData;
  i:Integer;
  dImporteMn,dImporteDll,tmpCantidad:Double;
  Td_partidasAnexo,
  Td_Distribucion_detalle2,
  Td_resumenMaterial: TfrxDBDataset;
  dsUnion:TDataSource;
  dTotalHH:Double;
  tmpRecursos:string;
begin
  QrRecursos:=TZReadOnlyQuery.Create(nil);
  QrPernoctas:=TZReadOnlyQuery.Create(nil);
  QrAdicional:=TZReadOnlyQuery.Create(nil);
  QrTotalHH:=TZReadOnlyQuery.Create(nil);
  try
    QrActividades:=TZReadOnlyQuery.Create(nil);
    QrAvance:=TZReadOnlyQuery.Create(nil);
    MryRecursos:=TjvMemoryData.Create(nil);

    Td_partidasAnexo:= TfrxDBDataset.Create(nil);
    Td_Distribucion_detalle2:= TfrxDBDataset.Create(nil);
    Td_resumenMaterial:= TfrxDBDataset.Create(nil);
    dsUnion:=TDataSource.Create(nil);

    QrActividades.Connection:=connection.zConnection;
    QrRecursos.Connection:=connection.zConnection;
    QrPernoctas.Connection:=connection.zConnection;
    QrAdicional.Connection:=connection.zConnection;
    QrTotalHH.Connection:=connection.zConnection;
    Td_partidasAnexo.UserName:='Td_partidasAnexo';
    Td_Distribucion_detalle2.UserName:='Td_Distribucion_detalle2';
    Td_resumenMaterial.UserName:='Td_resumenMaterial';

    with MryRecursos do
    begin
      Active:=False;
      FieldDefs.Add('sContrato', ftString, 100, False);
      FieldDefs.Add('sFolio', ftString, 100, False);
      FieldDefs.Add('sNumeroActividad', ftString, 100, False);
      FieldDefs.Add('sTipo', ftString, 50, False);
      FieldDefs.Add('sIdRecurso', ftString, 100, False);
      FieldDefs.Add('sdescripcion', ftString, 500, False);
      FieldDefs.Add('sMedida', ftString, 50, False);
      FieldDefs.Add('dCantidad', ftFloat, 0, False);
      FieldDefs.Add('dVentaMn', FtCurrency, 0, False);
      FieldDefs.Add('dVentaDll', FtCurrency, 0, False);
      FieldDefs.Add('dImporteMn', FtCurrency, 0, False);
      FieldDefs.Add('dImporteDll', FtCurrency, 0, True);
      Active:=true;
    end;




    QrActividades.SQL.Text:='select * from actividadesxorden where scontrato=:Contrato and sNumeroOrden=:Orden and sTipoActividad="Actividad" order by iItemOrden ';
    QrActividades.ParamByName('Contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
    QrActividades.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
    QrActividades.Open;
    dsUnion.DataSet:=QrActividades;

    QrAvance.Connection:=connection.zConnection;
    QrAvance.SQL.Text:= 'select b.*, ' +
                        '( SELECT (ifnull(sum(ba.dAvance), 0)) ' +
                                          '		FROM ' +
                                          '			bitacoradeactividades AS ba ' +
                                          '		WHERE ' +
                                          '			ba.sContrato = b.sContrato ' +
                                          '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                          '		AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento ' +
                                          '		AND ba.swbs = b.swbs ' +
                                          '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                          '		AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
                                          '   < cast(b.sHoraInicio AS Time))  )	) AS AvanceAnterior ' +
                        ' from bitacoradeactividades b' + #13#10 +
                        'where b.sContrato=:Contrato and b.snumeroorden=:Orden ' + #13#10 +
                        'and b.sIdTipoMovimiento=:Tipo' + #13#10 +
                        'order by b.snumeroactividad,b.didfecha,time(b.sHoraInicio)' ;
    QrAvance.ParamByName('contrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
    QrAvance.ParamByName('Orden').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
    //QrAvance.ParamByName('Actividad').AsString:=QrActividad.FieldByName('sNumeroActividad').AsString;
    QrAvance.ParamByName('tipo').AsString:='ED';
    QrAvance.Open;

    if DatosActa.FieldByName('lPernocta').AsString='Si' then
    begin
      with QrPernoctas do
      begin
        Active:=False;
        SQL.Text:='select * from cuentas order by iOrden';
        Open;
      end;

      QrAdicional.Active:=false;
      {QrAdicional.SQL.Text:='select ifnull(sum(dCantidad),0) as dCantidad from bitacoradepernocta where ' +
                            'sContrato=:Contrato and sNumeroOrden=:Folio and '+
                            'sIdCuenta=:Tipo group by sContrato ';}

      QrAdicional.SQL.Text:='select ' +
                            '  (ifnull(sum(bpr.dCantidad),0) + ' +
                            '   IFNULL((SELECT ' +
                            '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                            '           FROM ' +
                            '             bitacoradeajustes ba ' +
                            '           WHERE ' +
                            '             ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                            '             ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0))as dCantidad ' +
                            'from ' +
                            '  bitacoradepernocta bpr ' +
                            'where ' +
                            '  bpr.sContrato=:Contrato and ' +
                            '  bpr.sNumeroOrden=:Folio and ' +
                            '  bpr.sIdCuenta=:Tipo ' +
                            'group by ' +
                            '  bpr.sContrato';

      QrAdicional.parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
      QrAdicional.parambyname('Folio').asstring      := DatosActa.FieldByName('sNumeroOrden').AsString;

      with QrTotalHH do
      begin
        active := false;

        // mODIFICACION por E. Rangel - 21/08/2016
        sql.text := 'select (sum(bp.dCantHH) + ' +
                    ' IFNULL((SELECT ' +
                    '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                    '         FROM ' +
                    '           bitacoradeajustes ba ' +
                    '         WHERE ' +
                    '           ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                    '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                    '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantidad, ' +
                    'bp.sNumeroOrden, bp.sIdPlataforma, ' +
                    'ifnull(SUM(CAST(bp.dAjuste AS CHAR)),0) AS Ajuste ' +
                    'from bitacoradepersonal bp ' +
                    'where bp.sContrato =:Contrato and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" ' +
                    'group by bp.sTipoPernocta';

        {sql.text := 'select sum(bp.dCantHH) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                    'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                    'from bitacoradepersonal bp '+
                    'where bp.sContrato =:Contrato '+
                    'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" '+
                    'group by bp.sTipoPernocta ';}

        parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
        parambyname('Folio').asstring      := DatosActa.FieldByName('sNumeroOrden').AsString;
      end;
    end;


    QrAvance.MasterSource:=dsUnion;
    QrAvance.MasterFields:='sNumeroActividad';
    QrAvance.LinkedFields:='sNumeroActividad';

    while not QrActividades.Eof do
    begin
      if Tipo=FtAbordo then
      begin
        QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select me.*,sum(IF(mf.sFactor=NULL,me.sFactor,mf.sFactor)) as Total,tm.dVentaMn,tm.dVentaDll, ' + #13#10 +
                              'tm.sDescripcion as Descripcion '   +
                              'from movimientosdeembarcacion me' + #13#10 +
                              'inner join tiposdemovimiento tm' + #13#10 +
                              'on(me.sContrato=tm.sContrato and me.sClasificacion=tm.sIdTipoMovimiento)' + #13#10 +
                              'left join movimientosxfolios mf' + #13#10 +
                              'on(me.sContrato=mf.sContrato and me.iIdDiario = mf.iIdDiario and mf.sFolio=:Folio)' + #13#10 +
                              'where me.sContrato=:Contrato and me.sOrden=:Orden and me.sNumeroActividad=:Actividad' + #13#10 +
                              'and tm.sClasificacion="Movimiento de Barco" and tm.lAplicaNotaCampo="Si"' + #13#10 +
                              'group by tm.sIdTipoMovimiento' ;
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QrRecursos.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QrRecursos.Open;

        while not QrRecursos.Eof do
        begin
          with MryRecursos do
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            FieldByName('sTipo').AsString:='BARCO';
            FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sClasificacion').AsString;
            FieldByName('sdescripcion').AsString:=QrRecursos.FieldByName('Descripcion').AsString;
            FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sTipo').AsString;
            FieldByName('dCantidad').AsFloat:=xRound(QrRecursos.FieldByName('Total').AsFloat,Mantisa) ;
            FieldByName('dVentaMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
            FieldByName('dVentadLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
            FieldByName('dImporteMn').AsFloat:=xRound((xRound(QrRecursos.FieldByName('Total').AsFloat,Mantisa) *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
            FieldByName('dImportedLL').AsFloat:=xRound((xRound(QrRecursos.FieldByName('Total').AsFloat,Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
            post;
          end;
          QrRecursos.Next;
        end;

        QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select * from tiposdemovimiento ' + #13#10 +
                              'where sContrato=:Contrato and sClasificacion="Movimiento de Barco" and lAplicaNotaCampo="Si"' + #13#10 +
                              'and lImprimeCeros="Si" group by sIdTipoMovimiento' ;
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.Open;
        while not QrRecursos.Eof do
        begin
          if not MryRecursos.Locate('sIdRecurso;sNumeroActividad',VarArrayOf([QrRecursos.FieldByName('sIdTipoMovimiento').AsString,QrActividades.FieldByName('sNumeroActividad').AsString]),[]) then
            with MryRecursos do
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sTipo').AsString:='BARCO';
              FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sIdTipoMovimiento').AsString;
              FieldByName('sdescripcion').AsString:=QrRecursos.FieldByName('sDescripcion').AsString;
              FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sTipo').AsString;
              FieldByName('dCantidad').AsFloat:=0 ;
              FieldByName('dVentaMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
              FieldByName('dVentadLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
              FieldByName('dImporteMn').AsFloat:=(0 *  QrRecursos.FieldByName('dVentaMn').AsFloat);
              FieldByName('dImportedLL').AsFloat:=(0 * QrRecursos.FieldByName('dVentaDll').AsFloat);
              post;
            end;
          QrRecursos.Next;
        end;



      end;


      for I := 1 to Length(SQlRef) do
      begin
        tmpRecursos:='';
        tmpCantidad:=0;
        dImporteMn:=0;
        dImporteDll :=0;
        QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select br.' +SQlRef[i,3] + ' as sIdRecurso,br.sDescripcion,r.sMedida, ' +
                              '(sum(br.dCanthh) + ' +
                              ' IFNULL((SELECT ' +
                              '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                              '         FROM ' +
                              '           bitacoradeajustes ba ' +
                              '         WHERE ' +
                              '           ba.sContrato = :Orden AND ba.sNumeroOrden = :Folio AND ' +
                              '           ba.sNumeroActividad = :Actividad AND ' +
                              '           ba.sIdPartidaAnexo = br. ' + SQlRef[i,3] + ' AND ba.dFecha = bp.dIdFecha), 0)) as dCanthh, ' +
                              'Ifnull(sum(CAST(br.dAjuste AS CHAR)),0) as Ajuste, ' +
                              'r.dVentaMn,r.dVentaDll ' +
                              'from '+ SQlRef[i,1]  + ' br ' +
                              'left join ' + SQlRef[i,2] + ' r ' +
                              'on(r.sContrato=:Contrato and br.'+SQlRef[i,3]+'=r.'+ SQlRef[i,3] +') ' +
                              'where br.sContrato=:Orden and br.sNumeroOrden=:Folio and br.sNumeroActividad=:Actividad ' +
                              'group by sIdRecurso,br.dIdFecha order by r.iitemorden' ;
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
        QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
        QrRecursos.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QrRecursos.Open;

        while not QrRecursos.Eof do
        begin
          if tmpRecursos<> QrRecursos.FieldByName('sIdRecurso').AsString then
          begin
            with MryRecursos do
            begin
              if tmpRecursos<>'' then
              begin
                Edit;
                FieldByName('dCantidad').AsFloat:=tmpCantidad;
                FieldByName('dImporteMn').AsFloat:=xRound(dImporteMn,2);
                FieldByName('dImportedLL').AsFloat:=xRound(dImporteDll,2);
                Post;
                tmpCantidad:=0;
                dImporteMn:=0;
                dImporteDll :=0;
              end;



              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sTipo').AsString:=SQlRef[i,4];
              FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sIdRecurso').AsString;
              FieldByName('sdescripcion').AsString:=QrRecursos.FieldByName('sDescripcion').AsString;
              FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sMedida').AsString;
              FieldByName('dCantidad').AsFloat:=xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa);
              FieldByName('dVentaMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
              FieldByName('dVentadLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
              FieldByName('dImporteMn').AsFloat:=xRound((xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
              FieldByName('dImportedLL').AsFloat:=xRound((xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat),2);
              post;
              tmpRecursos:=QrRecursos.FieldByName('sIdRecurso').AsString;

            end;
          end;

          tmpCantidad:=tmpCantidad + xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa);
          dImporteMn:= dImporteMn + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) *  QrRecursos.FieldByName('dVentaMn').AsFloat);
          dImporteDll:= dImporteDll + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat);

          QrRecursos.Next;
        end;
        with MryRecursos do
        begin
          if tmpRecursos<>'' then
          begin
            Edit;
            FieldByName('dCantidad').AsFloat:=tmpCantidad;
            FieldByName('dImporteMn').AsFloat:=xRound(dImporteMn,2);
            FieldByName('dImportedLL').AsFloat:=xRound(dImporteDll,2);
            Post;
          end;
        end;

      end;

      if DatosActa.FieldByName('lPernocta').AsString='Si' then
      begin

        with QrRecursos do
        begin
          active := false;

          sql.text := 'select (sum(bp.dCantHH) + ' +
                      ' IFNULL((SELECT ' +
                      '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                      '         FROM ' +
                      '           bitacoradeajustes ba ' +
                      '         WHERE ' +
                      '           ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                      '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                      '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantidad, ' +
                      'bp.sNumeroOrden, bp.sIdPlataforma, ' +
                      'ifnull(SUM(CAST(bp.dAjuste AS CHAR)),0) AS Ajuste ' +
                      'from bitacoradepersonal bp ' +
                      'where bp.sContrato =:Contrato and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" and bp.sNumeroActividad=:Actividad ' +
                      'group by bp.sTipoPernocta';

{          sql.text := 'select sum(bp.dCantHH) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                      'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                      'from bitacoradepersonal bp '+
                      'where bp.sContrato =:Contrato '+
                      'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" '+
                      'and bp.sNumeroActividad=:Actividad ' +
                      'group by bp.sTipoPernocta ';}
          parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
          parambyname('Folio').asstring           := DatosActa.FieldByName('sNumeroOrden').AsString;
          ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        end;

        QrPernoctas.First;
        while not QrPernoctas.Eof do
        begin
          dTotalHH:=0;
          QrRecursos.Active:=False;
          QrRecursos.ParamByName('Tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
          QrRecursos.Open;

          QrAdicional.Active:=False;
          QrAdicional.ParamByName('Tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
          QrAdicional.Open;

          if QrAdicional.RecordCount>0 then
          begin
            QrTotalHH.Active:=False;
            QrTotalHH.ParamByName('Tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
            QrTotalHH.Open;
            if QrTotalHH.RecordCount>0 then
              dTotalHH:=xRound(QrTotalHH.FieldByName('dCantidad').asfloat +  QrTotalHH.FieldByName('Ajuste').asfloat,Mantisa) ;

            dTotalHH:=(xRound((QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa)  * QrAdicional.FieldByName('dCantidad').AsFloat)  / dTotalHH ;


          end;

          while not QrRecursos.Eof do
          begin

            with MryRecursos do
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sTipo').AsString:='PERNOCTA';
              FieldByName('sIdRecurso').AsString:=QrPernoctas.FieldByName('sIdPernocta').AsString;
              FieldByName('sdescripcion').AsString:=QrPernoctas.FieldByName('sDescripcion').AsString;
              FieldByName('sMedida').AsString:=QrPernoctas.FieldByName('sMedida').AsString;
              FieldByName('dCantidad').AsFloat:=dTotalHH + xRound((QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa) ;
              FieldByName('dVentaMn').AsFloat:=QrPernoctas.FieldByName('dVentaMn').AsFloat;
              FieldByName('dVentadLL').AsFloat:=QrPernoctas.FieldByName('dVentaDll').AsFloat;
              FieldByName('dImporteMn').AsFloat:=xRound(((dTotalHH + xRound(QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
              FieldByName('dImportedLL').AsFloat:=xRound(((dTotalHH + xRound(QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat,Mantisa))  * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);
              post;
            end;
            QrRecursos.Next;
          end;

          if QrRecursos.RecordCount=0 then
          begin
            with MryRecursos do
            begin
              Append;
              FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
              FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
              FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              FieldByName('sTipo').AsString:='PERNOCTA';
              FieldByName('sIdRecurso').AsString:=QrPernoctas.FieldByName('sIdPernocta').AsString;
              FieldByName('sdescripcion').AsString:=QrPernoctas.FieldByName('sDescripcion').AsString;
              FieldByName('sMedida').AsString:=QrPernoctas.FieldByName('sMedida').AsString;
              FieldByName('dCantidad').AsFloat:=0.00;//(QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat) ;
              FieldByName('dVentaMn').AsFloat:=QrPernoctas.FieldByName('dVentaMn').AsFloat;
              FieldByName('dVentadLL').AsFloat:=QrPernoctas.FieldByName('dVentaDll').AsFloat;
              FieldByName('dImporteMn').AsFloat:=0.00;//((QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat)  *  QrRecursos.FieldByName('dVentaMn').AsFloat);
              FieldByName('dImportedLL').AsFloat:=0.00;//((QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat)  * QrRecursos.FieldByName('dVentaDll').AsFloat);
              post;
            end;
          end;
          QrPernoctas.Next;
        end;
      
      end;

      if DatosActa.FieldByName('lMaterial').AsString='Si' then
      begin
        with QrRecursos do
        begin
          active := false;
          sql.text := 'select bp.sIdInsumo as sIdMaterial,bp.sTrazabilidad,round(sum( bp.dCantidad),2) as dCantidad, '+
                      'i.mDescripcion as sDescripcion,bp.sMedida '    +
                      'from acta_material bp '+
                      'inner join insumos i '+
                      'on(i.sContrato=bp.sContrato and i.sIdInsumo=bp.sIdInsumo) ' +
                      'where bp.sContrato =:Contrato '+
                      'and bp.iIdActa=:Acta '+
                      'and bp.sNumeroActividad=:Actividad ' +
                      'group by bp.sIdInsumo ';
          parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
          parambyname('Acta').AsInteger           := DatosActa.FieldByName('iIdActa').AsInteger;
          ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          Open;
        end;

        while not QrRecursos.Eof do
        begin



          with MryRecursos do
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('sNumeroActividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            FieldByName('sTipo').AsString:='MATERIAL';
            FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sTrazabilidad').AsString;
            FieldByName('sdescripcion').AsString:=QrRecursos.FieldByName('sDescripcion').AsString;
            FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sMedida').AsString;
            FieldByName('dCantidad').AsFloat:=(QrRecursos.FieldByName('dCantidad').AsFloat) ;
            FieldByName('dVentaMn').AsFloat:=0.00;
            FieldByName('dVentadLL').AsFloat:=0.00;//QrPernoctas.FieldByName('dVentaDll').AsFloat;
            FieldByName('dImporteMn').AsFloat:=0.00;//((QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat)  *  QrPernoctas.FieldByName('dVentaMn').AsFloat);
            FieldByName('dImportedLL').AsFloat:=0.00;//(//(QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat)  * QrPernoctas.FieldByName('dVentaDll').AsFloat);
            post;
          end;



          QrRecursos.Next;
        end;
      
      end;

      QrActividades.Next;
    end;
    QrActividades.First;
    MryRecursos.First;
    QrActividades.AfterScroll:=EventHandlers.ActaPdf_DetalleAfterScroll;
    MemoryDetalleTmp:=MryRecursos;

      

    Td_partidasAnexo.DataSet:= QrActividades;
    Td_Distribucion_detalle2.DataSet:= QrAvance;
    Td_resumenMaterial.DataSet:= MryRecursos;

    Td_partidasAnexo.FieldAliases.Clear;
    Td_Distribucion_detalle2.FieldAliases.Clear;
    Td_resumenMaterial.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_partidasAnexo);
    Reporte.DataSets.Add(Td_Distribucion_detalle2);
    Reporte.DataSets.Add(Td_resumenMaterial);


  finally
    QrRecursos.Destroy;
    QrPernoctas.Destroy;
    QrAdicional.destroy;
  end;
end;
Procedure ActaEx_DesgloseCostos(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
var
  QRecursos  :TZReadOnlyQuery;
  Td_ResumenPersonal: TfrxDBDataset;
begin
  QRecursos:=TZReadOnlyQuery.Create(nil);
  try

    QRecursos.Connection:=connection.zConnection;
    Td_ResumenPersonal:= TfrxDBDataset.Create(nil);
    Td_ResumenPersonal.UserName:='Td_ResumenPersonal';


    QRecursos.SQL.Text:='select ac.*,ac.etipo as sTipo,ao.sNumeroOrden as sFolio,ac.dCostoMn as dVentaMn,'+
                'ac.dCostoDll as dVentaDll,ac.mDescripcion as sDescripcion, '+
                'ac2.dImporteMn as ImporteMn,ac2.dImportedll as ImporteDll ' +
                ',ac3.dImporteMn as ActImporteMn,ac3.dImportedll as ActImporteDll '  +
                'from acta_campo ac inner join actividadesxorden ao ' +
                'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                'inner join acta_campo ac2' + #13#10 +
                'on(ac2.iIdActa=ac.iIdActa and ac2.sWbs=ac.swbs and ac2.eTipo=ac.eTipo and ac2.sIdRecurso like "$IMPORTE%" and ac2.sAnexo=ac.sAnexo)' +
                'inner join acta_campo ac3' + #13#10 +
                'on(ac3.iIdActa=ac.iIdActa and ac3.sWbs=ac.swbs and ac3.eTipo="ACTIVIDAD") '   +
                'where ac.iIdActa=:Acta and ao.sTipoActividad=:Tipo and '   +
                'ac.sIdRecurso not like "$IMPORTE%" order by ao.iItemOrden,ac.iOrdenTipo,ac.sAnexo,ac.iOrdenRecurso';

    QRecursos.ParamByName('Acta').AsInteger:=DatosActa.FieldByName('iIdActa').AsInteger;
    QRecursos.ParamByName('Tipo').AsString:='Paquete';
    QRecursos.Open;


    Td_ResumenPersonal.DataSet:= QRecursos;
    Td_ResumenPersonal.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_ResumenPersonal);


  finally

  end;
end;

Procedure ActaPdf_DesgloseCostos(DatosActa:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
const
  SQlRef: array[1..2,1..4] of string=(('bitacoradepersonal','personal','sIdPersonal','PERSONAL'),('bitacoradeequipos','equipos','sIdEquipo','EQUIPO'));
var
  {QrActividades, }
  QrRecursos{,QrAvance},QrPernoctas,
  QrAdicional:TZReadOnlyQuery;
  mryRecursos:TjvMemoryData;
  i:Integer;
  dImporteMn,dImporteDll,dAdicionalHH,tmpCantidad:Double;
  Td_ResumenPersonal: TfrxDBDataset;
  QrAnexoTE:TZReadOnlyQuery;
  sAnexosTE,sCampoAnexo:string;
  tmpRecursos:string;
begin
  Montos[1,1]:=0;
  Montos[1,2]:=0;
  Montos[2,1]:=0;
  Montos[2,2]:=0;
  Montos[3,1]:=0;
  Montos[3,2]:=0;
  Montos[4,1]:=0;
  Montos[4,2]:=0;
  QrRecursos:=TZReadOnlyQuery.Create(nil);
  QrAnexoTE:=TZReadOnlyQuery.Create(nil);
  QrPernoctas:=TZReadOnlyQuery.Create(nil);
  QrAdicional:=TZReadOnlyQuery.Create(nil);
  try

    MryRecursos:=TjvMemoryData.Create(nil);
    Td_ResumenPersonal:= TfrxDBDataset.Create(nil);

    QrAnexoTE.Connection:=connection.zConnection;
    QrRecursos.Connection:=connection.zConnection;
    QrPernoctas.Connection:=connection.zConnection;
    QrAdicional.Connection:=connection.zConnection;
    Td_ResumenPersonal.UserName:='Td_ResumenPersonal';

    with MryRecursos do
    begin
      Active:=False;
      FieldDefs.Add('sContrato', ftString, 100, False);
      FieldDefs.Add('sFolio', ftString, 100, False);
      FieldDefs.Add('sNumeroActividad', ftString, 100, False);
      FieldDefs.Add('sTipo', ftString, 50, False);
      FieldDefs.Add('sAnexo', ftString, 50, False);
      FieldDefs.Add('sIdRecurso', ftString, 100, False);
      FieldDefs.Add('sdescripcion', ftString, 500, False);
      FieldDefs.Add('sMedida', ftString, 50, false);
      FieldDefs.Add('dCantidad', ftFloat, 0, False);
      FieldDefs.Add('dVentaMn', FtCurrency, 0, False);
      FieldDefs.Add('dVentaDll', FtCurrency, 0, False);
      FieldDefs.Add('dImporteMn', FtCurrency, 0, False);
      FieldDefs.Add('dImporteDll', FtCurrency, 0, False);
      FieldDefs.Add('iItemOrden', ftInteger, 0, False);
      FieldDefs.Add('iPostipo', ftInteger, 0, False);
      Active:=true;
    end;

    if Tipo=FtAbordo then
    begin
      dImporteMn:=0;
      dImporteDll :=0;
      QrRecursos.Active:=False;
      QrRecursos.SQL.Text:= 'select me.*,sum(IF(mf.sFactor=NULL,me.sFactor,mf.sFactor)) as dCAntidad,tm.dVentaMn,tm.dVentaDll, ' + #13#10 +
                            'tm.sDescripcion as Descripcion,tm.iOrden '   +
                            'from movimientosdeembarcacion me' + #13#10 +
                            'inner join tiposdemovimiento tm' + #13#10 +
                            'on(me.sContrato=tm.sContrato and me.sClasificacion=tm.sIdTipoMovimiento)' + #13#10 +
                            'left join movimientosxfolios mf' + #13#10 +
                            'on(me.sContrato=mf.sContrato and me.iIdDiario = mf.iIdDiario and mf.sFolio=:Folio)' + #13#10 +
                            'where me.sContrato=:Contrato and me.sOrden=:Orden ' + #13#10 +
                            'and tm.sClasificacion="Movimiento de Barco" and tm.lAplicaNotaCampo="Si"' + #13#10 +
                            'group by tm.sIdTipoMovimiento' ;
      QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
      QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
      QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
      QrRecursos.Open;

      while not QrRecursos.Eof do
      begin
        dImporteMn:= {dImporteMn + }xRound((xRound((QrRecursos.FieldByName('dCantidad').AsFloat ),Mantisa) *  QrRecursos.FieldByName('dVentaMn').AsFloat),2);
        dImporteDll:= {dImporteDll +} xRound((xRound((QrRecursos.FieldByName('dCantidad').AsFloat ),Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat),2);

        with MryRecursos do
        begin
          Append;
          FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
          FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
          FieldByName('sNumeroActividad').AsString:='*';//QrRecursos.FieldByName('sNumeroActividad').AsString;
          FieldByName('sTipo').AsString:='BARCO';
          FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sClasificacion').AsString;;//QrPernoctas.FieldByName('sIdPernocta').AsString;
          FieldByName('sAnexo').AsString:='';
          FieldByName('sdescripcion').AsString:=QrRecursos.FieldByName('Descripcion').AsString;
          FieldByName('dCantidad').AsFloat:=xRound((QrRecursos.FieldByName('dCantidad').AsFloat),Mantisa);
          FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sTipo').AsString;
          FieldByName('dVentaMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
          FieldByName('dVentadLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
          FieldByName('dImporteMn').AsFloat:=dImporteMn;
          FieldByName('dImportedLL').AsFloat:=dImporteDll;
          FieldByName('iPosTipo').AsInteger:=0;
          FieldByName('iitemorden').AsInteger:=QrRecursos.FieldByName('iOrden').AsInteger;
          post;
        end;
        Montos[3,1]:=Montos[3,1] + xRound(dImporteMn,2);
        Montos[3,2]:=Montos[3,2] + xRound(dImporteDll,2);
        QrRecursos.Next;
      end;

      if QrRecursos.RecordCount=0 then
      begin
        QrRecursos.Active:=False;
        QrRecursos.SQL.Text:= 'select * from tiposdemovimiento ' + #13#10 +
                              'where sContrato=:Contrato and sClasificacion="Movimiento de Barco" and lAplicaNotaCampo="Si"' + #13#10 +
                              'and lImprimeCeros="Si" group by sIdTipoMovimiento' ;
        QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
        QrRecursos.Open;
        while not QrRecursos.Eof do
        begin
          with MryRecursos do
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('sNumeroActividad').AsString:='*';//QrRecursos.FieldByName('sNumeroActividad').AsString;
            FieldByName('sTipo').AsString:='BARCO';
            FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sIdTipoMovimiento').AsString;//QrPernoctas.FieldByName('sIdPernocta').AsString;
            FieldByName('sAnexo').AsString:='';
            FieldByName('sdescripcion').AsString:=QrRecursos.FieldByName('sDescripcion').AsString;
            FieldByName('dCantidad').AsFloat:=0.00;
            FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sTipo').AsString;
            FieldByName('dVentaMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
            FieldByName('dVentadLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
            FieldByName('dImporteMn').AsFloat:=0.00;
            FieldByName('dImportedLL').AsFloat:=0.00;
            FieldByName('iPosTipo').AsInteger:=0;
            FieldByName('iitemorden').AsInteger:=QrRecursos.FieldByName('iOrden').AsInteger;
            post;
          end;
          QrRecursos.Next;
        end;
      end;



    end;

    for I := 1 to Length(SQlRef) do
    begin
      tmpRecursos:='';
      tmpCantidad:=0;
      sAnexosTE:='';
      sCampoAnexo:='';
      if I=1 then
      begin
        QrAnexoTE.Active:=False;
        QrAnexoTE.SQL.Text:='select GROUP_CONCAT(sanexo) as sAnexo from anexos where stipo="TIEMPO_EXTRA" ';
        QrAnexoTE.Open;
        if QrAnexoTE.RecordCount=1 then
          sAnexosTE:=QrAnexoTE.FieldByName('sAnexo').AsString;

        sCampoAnexo:= ', r.sAnexo ';
      end;
      dImporteMn:=0;
      dImporteDll :=0;

      QrRecursos.Active:=False;
      QrRecursos.SQL.Text:= 'select br.' +SQlRef[i,3] + ' as sIdRecurso,br.sDescripcion,r.sMedida, ' +
                            '(sum(br.dCanthh) + ' +
                            ' IFNULL((SELECT ' +
                            '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                            '         FROM ' +
                            '           bitacoradeajustes ba ' +
                            '         WHERE ' +
                            '           ba.sContrato = br.sContrato AND ba.sNumeroOrden = br.sNumeroOrden AND ' +
                            '           ba.sNumeroActividad = br.sNumeroActividad AND ' +
                            '           ba.sIdPartidaAnexo = br.' + SQlRef[i,3] + ' AND ba.dFecha = br.dIdFecha), 0)) as dCanthh, ' +
                            'ifnull(sum(CAST(br.dAjuste AS CHAR)),0) as Ajuste, ' +
                            'r.dVentaMn,r.dVentaDll,r.iitemorden ' + sCampoAnexo +
                            'from '+ SQlRef[i,1]  + ' br ' +
                            'left join ' + SQlRef[i,2] + ' r ' +
                            'on(r.sContrato=:Contrato and br.'+SQlRef[i,3]+'=r.'+ SQlRef[i,3] +') ' +
                            'where br.sContrato=:Orden and br.sNumeroOrden=:Folio ' +
                            'group by sIdRecurso,br.didfecha order by r.iitemorden' ;
      QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
      QrRecursos.ParamByName('Orden').AsString:=DatosActa.FieldByName('sContrato').AsString;
      QrRecursos.ParamByName('Folio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
      //QrRecursos.ParamByName('Actividad').AsString:=DatosActa.FieldByName('sNumeroActividad').AsString;
      QrRecursos.Open;

      while not QrRecursos.Eof do
      begin
        if tmpRecursos<> QrRecursos.FieldByName('sIdRecurso').AsString then
        begin

          with MryRecursos do
          begin
            if tmpRecursos<>'' then
            begin
              Edit;
              FieldByName('dCantidad').AsFloat:=tmpCantidad;
              FieldByName('dImporteMn').AsFloat:=xRound(dImporteMn,2);
              FieldByName('dImportedLL').AsFloat:=xRound(dImporteDll,2);
              Post;
              Montos[i,1]:=Montos[i,1] + xRound(dImporteMn,2);
              Montos[i,2]:=Montos[i,2] + xRound(dImporteDll,2);
              tmpCantidad:=0;
              dImporteMn:=0;
              dImporteDll :=0;
            end;

            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('sNumeroActividad').AsString:='*';//QrRecursos.FieldByName('sNumeroActividad').AsString;
            FieldByName('sTipo').AsString:=SQlRef[i,4];
            FieldByName('sIdRecurso').AsString:=QrRecursos.FieldByName('sIdRecurso').AsString;

            if i=1 then
            begin
              if sAnexosTE<>'' then
              begin
                if AnsiMatchText(sAnexosTE,[QrRecursos.FieldByName('sAnexo').AsString]) then
                  FieldByName('sAnexo').AsString:='TIEMPO EXTRA'
                else
                  FieldByName('sAnexo').AsString:='';
              end
              else
                FieldByName('sAnexo').AsString:='';

            end
            else
              FieldByName('sAnexo').AsString:='';


            FieldByName('sdescripcion').AsString:=QrRecursos.FieldByName('sDescripcion').AsString;
            FieldByName('dCantidad').AsFloat:=xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa);
            FieldByName('sMedida').AsString:=QrRecursos.FieldByName('sMedida').AsString;
            FieldByName('dVentaMn').AsFloat:=QrRecursos.FieldByName('dVentaMn').AsFloat;
            FieldByName('dVentadLL').AsFloat:=QrRecursos.FieldByName('dVentaDll').AsFloat;
            FieldByName('dImporteMn').AsFloat:=0;
            FieldByName('dImportedLL').AsFloat:=0;
            FieldByName('iPosTipo').AsInteger:=i ;
            FieldByName('iitemorden').AsInteger:=QrRecursos.FieldByName('iItemOrden').AsInteger;

            post;


            tmpRecursos:=QrRecursos.FieldByName('sIdRecurso').AsString;

          end;

        
        end;

        tmpCantidad:=tmpCantidad + xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat ),Mantisa);
        dImporteMn:= dImporteMn + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa) *  QrRecursos.FieldByName('dVentaMn').AsFloat);
        dImporteDll:= dImporteDll + (xRound((QrRecursos.FieldByName('dCanthh').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat),Mantisa) * QrRecursos.FieldByName('dVentaDll').AsFloat);



        QrRecursos.Next;
      end;
      with MryRecursos do
      begin
        if tmpRecursos<>'' then
        begin
          Edit;
          FieldByName('dCantidad').AsFloat:=tmpCantidad;
          FieldByName('dImporteMn').AsFloat:=xRound(dImporteMn,2);
          FieldByName('dImportedLL').AsFloat:=xRound(dImporteDll,2);
          Post;
          Montos[i,1]:=Montos[i,1] + xRound(dImporteMn,2);
          Montos[i,2]:=Montos[i,2] + xRound(dImporteDll,2);
        end;
      end;

    end;

    if DatosActa.FieldByName('lPernocta').AsString='Si' then
    begin
      with QrRecursos do
      begin
        active := false;
        sql.text := 'select (sum(bp.dCantHH) + ' +
                    ' IFNULL((SELECT ' +
                    '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                    '         FROM ' +
                    '           bitacoradeajustes ba ' +
                    '         WHERE ' +
                    '           ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                    '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                    '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantidad, ' +
                    'bp.sNumeroOrden, bp.sIdPlataforma, '+
                    'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste ' +
                    'from bitacoradepersonal bp '+
                    'where bp.sContrato =:Contrato '+
                    'and bp.sNumeroOrden=:Folio and bp.sTipoPernocta =:Tipo and bp.lImprime = "Si" '+
                 //   'and bp.sNumeroActividad=:Actividad ' +
                    'group by bp.sTipoPernocta ';
        parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
        parambyname('Folio').asstring           := DatosActa.FieldByName('sNumeroOrden').AsString;
        //ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
      end;

      QrAdicional.Active:=false;
      {QrAdicional.SQL.Text:='select ifnull(sum(dCantidad),0) as dCantidad from bitacoradepernocta where ' +
                            'sContrato=:Contrato and sNumeroOrden=:Folio and '+
                            'sIdCuenta=:Tipo group by sContrato ';}

      QrAdicional.SQL.Text:='select ' +
                            '  (ifnull(sum(dCantidad),0) + ' +
                            '   IFNULL((SELECT ' +
                            '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                            '           FROM ' +
                            '             bitacoradeajustes ba ' +
                            '           WHERE ' +
                            '             ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                            '             ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0))as dCantidad ' +
                            'from ' +
                            '  bitacoradepernocta bpr ' +
                            'where ' +
                            '  bpr.sContrato=:Contrato and ' +
                            '  bpr.sNumeroOrden=:Folio and ' +
                            '  bpr.sIdCuenta=:Tipo ' +
                            'group by ' +
                            '  bpr.sContrato';

      QrAdicional.parambyname('contrato').asstring      := DatosActa.FieldByName('sContrato').AsString;
      QrAdicional.parambyname('Folio').asstring      := DatosActa.FieldByName('sNumeroOrden').AsString;




      with QrPernoctas do
      begin
        Active:=False;
        SQL.Text:='select * from cuentas order by iOrden';
        Open;
      end;

      while not QrPernoctas.Eof do
      begin
        dAdicionalHH:=0;
        QrRecursos.Active:=False;
        QrRecursos.ParamByName('Tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
        QrRecursos.Open;

        QrAdicional.Active:=False;
        QrAdicional.ParamByName('Tipo').AsString:=QrPernoctas.FieldByName('sIdCuenta').AsString;
        QrAdicional.Open;
        if QrAdicional.RecordCount=1 then
          dAdicionalHH:= xRound(QrAdicional.FieldByName('dCantidad').AsFloat,Mantisa);


        while not QrRecursos.Eof do
        begin
          dImporteMn:= {dImporteMn + }xRound(((dAdicionalHH + xRound(QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('ajuste').AsFloat,Mantisa)) *  QrPernoctas.FieldByName('dVentaMn').AsFloat),2);
          dImporteDll:= {dImporteDll +} xRound(((dAdicionalHH + xRound(QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('ajuste').AsFloat,Mantisa)) * QrPernoctas.FieldByName('dVentaDll').AsFloat),2);

          with MryRecursos do
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('sNumeroActividad').AsString:='*';//QrRecursos.FieldByName('sNumeroActividad').AsString;
            FieldByName('sTipo').AsString:='PERNOCTA';
            FieldByName('sIdRecurso').AsString:=QrPernoctas.FieldByName('sIdPernocta').AsString;
            FieldByName('sAnexo').AsString:='';
            FieldByName('sdescripcion').AsString:=QrPernoctas.FieldByName('sDescripcion').AsString;
            FieldByName('dCantidad').AsFloat:=(dAdicionalHH + xRound(QrRecursos.FieldByName('dCantidad').AsFloat + QrRecursos.FieldByName('ajuste').AsFloat,Mantisa));
            FieldByName('sMedida').AsString:=QrPernoctas.FieldByName('sMedida').AsString;
            FieldByName('dVentaMn').AsFloat:=QrPernoctas.FieldByName('dVentaMn').AsFloat;
            FieldByName('dVentadLL').AsFloat:=QrPernoctas.FieldByName('dVentaDll').AsFloat;
            FieldByName('dImporteMn').AsFloat:=dImporteMn;
            FieldByName('dImportedLL').AsFloat:=dImporteDll;
            FieldByName('iPosTipo').AsInteger:=i+1;
            FieldByName('iitemorden').AsInteger:=QrPernoctas.FieldByName('iOrden').AsInteger;
            post;
          end;
          Montos[4,1]:=Montos[4,1] + xRound(dImporteMn,2);
          Montos[4,2]:=Montos[4,2] + xRound(dImporteDll,2);
          QrRecursos.Next;
        end;

        if QrRecursos.RecordCount=0 then
        begin
          with MryRecursos do
          begin
            Append;
            FieldByName('sContrato').AsString:=DatosActa.FieldByName('sContrato').AsString;
            FieldByName('sFolio').AsString:=DatosActa.FieldByName('sNumeroOrden').AsString;
            FieldByName('sNumeroActividad').AsString:='*';//QrRecursos.FieldByName('sNumeroActividad').AsString;
            FieldByName('sTipo').AsString:='PERNOCTA';
            FieldByName('sIdRecurso').AsString:=QrPernoctas.FieldByName('sIdPernocta').AsString;
            FieldByName('sAnexo').AsString:='';
            FieldByName('sdescripcion').AsString:=QrPernoctas.FieldByName('sDescripcion').AsString;
            FieldByName('dCantidad').AsFloat:=0.00;//QrRecursos.FieldByName('dCantidad').AsFloat;
            FieldByName('sMedida').AsString:=QrPernoctas.FieldByName('sMedida').AsString;
            FieldByName('dVentaMn').AsFloat:=QrPernoctas.FieldByName('dVentaMn').AsFloat;
            FieldByName('dVentadLL').AsFloat:=QrPernoctas.FieldByName('dVentaDll').AsFloat;
            FieldByName('dImporteMn').AsFloat:=0.00;//dImporteMn;
            FieldByName('dImportedLL').AsFloat:=0.00;//dImporteDll;
            FieldByName('iPosTipo').AsInteger:=i+1;
            FieldByName('iitemorden').AsInteger:=QrPernoctas.FieldByName('iOrden').AsInteger;
            post;
          end;
        end;
        QrPernoctas.Next;
      end;

    end;



    MryRecursos.SortOnFields('iPostipo;sAnexo;iitemorden');
    MryRecursos.First;

    Td_ResumenPersonal.DataSet:= MryRecursos;
    Td_ResumenPersonal.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_ResumenPersonal);
  finally
    QrAnexoTE.Destroy;
    QrRecursos.Destroy;
    QrPernoctas.Destroy;
    QrAdicional.Destroy;
  end;
end;

Procedure TdEstimacionCamsa(ListaParams:TstringList;TipoEstimacion:FtEstimacion;Var Reporte: TfrxReport;IsMonto:Boolean=False);
const
{$Region 'SQL'}
  Asql:Array[1..6,1..3] of String =( ('SELECT me.sOrden as Contrato,mf.sFolio as Folio,tm.sIdTipoMovimiento as sIdRecurso,tm.sDescripcion as Descripcion,' + #10 +
                                'tm.sTipo,tm.smedida as medida,me.didfecha,mf.sFolio,ifnull(SUM( mf.sFactor ),0) AS Total,' + #10 +
                                'ifnull((select sdescripcion from anexos where sTipo="BARCO" limit 1)'+
                                ',"No existe") as Titulo, ' +
                                 'ifnull((select sAnexo from anexos where sTipo="BARCO" limit 1)'+
                                ',"No existe") as sAnexo, ' +
                                 'ifnull((select sMedida from anexos where sTipo="BARCO" limit 1)'+
                                ',"No existe") as sMedidaAnexo, ' +
                                'tm.dVentaMn,tm.dVentaDll,' +
                                '(SELECT ifnull(SUM( mf2.sFactor ),0)' + #10 +
                                'FROM movimientosdeembarcacion AS me2' + #10 +
                                'inner join movimientosxfolios mf2 on (me2.sContrato = mf2.sContrato' + #10 +
                                'and me2.sOrden=mf2.sNumeroOrden and me2.didfecha=mf2.dIdFecha and' + #10 +
                                'me2.iIdDiario = mf2.iIdDiario)' + #10 +
                                'WHERE me2.sContrato =:ContratoBarco and' + #10 +
                                '((:Orden=-1) or (:Orden<>-1 and me2.sOrden =me.sOrden))' + #10 +
                                'and ((:Folio=-1) or (:Folio<>-1 and mf2.sFolio=mf.sFolio))' + #10 +
                                'AND me2.sClasificacion =me.sClasificacion  and me2.didfecha < :Inicio ) AS Anterior '  +
                                'FROM tiposdemovimiento AS tm '  +
                                'inner join movimientosdeembarcacion AS me '   +
                                'on(me.sContrato = tm.sContrato AND me.sClasificacion = tm.sIdTipoMovimiento) ' + #10 +
                                'inner join movimientosxfolios mf on (me.sContrato = mf.sContrato and me.sOrden=mf.sNumeroOrden' + #10 +
                                'and me.didfecha=mf.dIdFecha and me.iIdDiario = mf.iIdDiario) ' +
                                'WHERE tm.sContrato =:ContratoBarco and ' + #10 +
                                '((:Orden=-1) or (:Orden<>-1 and me.sOrden =:Orden)) '+
                                'and ((:Folio=-1) or (:Folio<>-1 and mf.sFolio=:Folio)) ' + #10 +
                                'AND tm.sClasificacion = "Movimiento de Barco" and' + #10 +
                                'tm.lGenera ="Si" and me.didfecha between :Inicio and :termino ' +
                                'group by tm.sIdTipoMovimiento '+
                                'order by tm.sIdTipoMovimiento',
                                'group by me.sOrden,mf.sFolio,',
                                'order by me.sOrden,mf.sFolio,')
                                //'group by me.sOrden,mf.sFolio,tm.sIdTipoMovimiento,me.didfecha '+
                                //'order by me.sOrden,mf.sFolio,tm.sIdTipoMovimiento,me.didfecha'
                              ,
                                ('(SELECT bp.sContrato as Contrato,bp.sNumeroOrden as Folio,bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,a.sAnexo, ' +
                                'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                '(IF(SUM(bp.dCanthh) > 0, '+
                                '(SUM(bp.dCanthh) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = bp2.sContrato AND ba.sNumeroOrden = bp2.sNumeroOrden AND ' +
                                '           ba.sNumeroActividad = bp2.sNumeroActividad AND ' +
                                '           ba.sIdPartidaAnexo = bp2.sIdPersonal AND ba.dFecha = bp2.dIdFecha), 0)), ' +
                                'SUM(bp.dCantidad))) AS Total,bp.dIdFecha,p.iItemOrden,a.iOrden,'+
                                'p.dVentaMn,p.dVentaDll,a.sMedida as sMedidaAnexo,'+
                                '(SELECT IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) + ifnull((IF(SUM(bp2.dCanthh) > 0,' + #10 +
                                '(SUM(bp2.dCanthh) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = bp2.sContrato AND ba.sNumeroOrden = bp2.sNumeroOrden AND ' +
                                '           ba.sNumeroActividad = bp2.sNumeroActividad AND ' +
                                '           ba.sIdPartidaAnexo = bp2.sIdPersonal AND ba.dFecha = bp2.dIdFecha), 0)), ' +
                                'SUM(bp2.dCantidad))),0) FROM bitacoradepersonal bp2' + #10 +
                                'WHERE  ((:Orden=-1) or (:Orden<>-1 and bp2.scontrato =bp.scontrato)) AND' + #10 +
                                '((:Folio=-1) or(:Folio<>-1 and bp2.sNumeroOrden =bp.sNumeroOrden))' + #10 +
                                'and bp2.dIdFecha < :Inicio and bp2.sIdpersonal=bp.sIdpersonal) '+
                                'as Anterior FROM bitacoradepersonal bp '+
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" ) '+    //and a.sTierra="No"
                                'WHERE  ((:Orden=-1) or (:Orden<>-1 and bp.scontrato =:Orden)) AND ' +
                                '((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino and ' +
                                '((:Tipo="TE" and p.sIdTipoPersonal =:Tipo) or (:Tipo<>"TE" and p.sIdTipoPersonal<>"TE" )) '  +
                                'and p.lCobro=:Cobro '  +
                                'group by bp.sIdPersonal,bp.didfecha) '+
                                'union all '+
                                '(SELECT bp.sContrato as Contrato,bp.sNumeroOrden as Folio,bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,a.sAnexo, ' +
                                'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                '(SUM(bp.dCanthh) + ' +
                                ' IFNULL((SELECT ' +
                                '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '          FROM ' +
                                '            bitacoradeajustes ba ' +
                                '          WHERE ' +
                                '            ba.sContrato = bp2.sContrato AND ba.sNumeroOrden = bp2.sNumeroOrden AND ' +
                                '            ba.sNumeroActividad = bp2.sNumeroActividad AND ' +
                                '            ba.sIdPartidaAnexo = bp2.sIdPersonal AND ba.dFecha = bp2.dIdFecha), 0)) AS Total, ' +
                                'bp.dIdFecha, '+
                                'p.iItemOrden,a.iOrden,p.dVentaMn,p.dVentaDll,a.sMedida as sMedidaAnexo,'+
                                'SELECT (IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) + ifnull((IF(SUM(bp2.dCanthh) > 0,' + #10 +
                                'SUM(bp2.dCanthh), SUM(bp2.dCantidad))),0) FROM bitacoradepersonal bp2' + #10 +
                                'WHERE  ((:Orden=-1) or (:Orden<>-1 and bp2.scontrato =bp.scontrato)) AND' + #10 +
                                '((:Folio=-1) or(:Folio<>-1 and bp2.sNumeroOrden =bp.sNumeroOrden))' + #10 +
                                'and bp2.dIdFecha < :Inicio and bp2.sIdpersonal=bp.sIdpersonal) '+
                                'as Anterior FROM bitacoradepersonal bp inner join personal p on(p.sContrato=:ContratoBarco and '+
                                'p.sIdPersonal=bp.sIdPersonal) inner join anexos a on(a.sAnexo=p.sAnexo and '+
                                'a.sTipo="PERSONAL" and a.sTierra="Si") '+
                                'WHERE  ((:Orden=-1) or (:Orden<>-1 and bp.scontrato =:Orden)) AND ' +
                                '((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) ' +
                                'and bp.dIdFecha between :Inicio and :termino and ' +
                                '((:Tipo="TE" and p.sIdTipoPersonal =:Tipo) or (:Tipo<>"TE" and p.sIdTipoPersonal<>"TE" )) '  +
                                'and p.lCobro=:Cobro '  +
                                'group by bp.sIdPersonal) '+
                                'order by iOrden,iItemOrden,sIdRecurso',
                                'group by bp.sContrato,bp.sNumeroOrden,',
                                'order by Contrato,Folio,'
                                )
                              , ('select be.sContrato as Contrato,be.sNumeroOrden as Folio,be.sIdEquipo as sIdRecurso,be.sDescripcion as Descripcion,e.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="EQUIPO" limit 1),"No existe") as Titulo,'+
                                'ifnull((select sAnexo from anexos where sTipo="EQUIPO" limit 1),"N/E") as sAnexo,'+
                                'ifnull((select sMedida from anexos where sTipo="EQUIPO" limit 1),"N/E") as sMedidaAnexo,'+
                                '(sum(be.dCantHH) + ' +
                                ' IFNULL((SELECT ' +
                                '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '          FROM ' +
                                '            bitacoradeajustes ba ' +
                                '          WHERE ' +
                                '            ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                                '            ba.sNumeroActividad = be.sNumeroActividad AND ' +
                                '            ba.sIdPartidaAnexo = be.sIdPersonal AND ba.dFecha = be.dIdFecha), 0)) as total, ' +
                                'be.didfecha,e.dVentaMn,e.dVentaDll, ' +
                                '(select (sum(be2.dCantHH) + ' +
                                ' IFNULL((SELECT ' +
                                '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '          FROM ' +
                                '            bitacoradeajustes ba ' +
                                '          WHERE ' +
                                '            ba.sContrato = be2.sContrato AND ba.sNumeroOrden = be2.sNumeroOrden AND ' +
                                '            ba.sNumeroActividad = be2.sNumeroActividad AND ' +
                                '            ba.sIdPartidaAnexo = be2.sIdEquipo AND ba.dFecha = be2.dIdFecha), 0)) ' +
                                'from bitacoradeequipos be2' + #10 +
                                'where ((:Orden=-1) or (:Orden<>-1 and be2.sContrato =be.sContrato)) and' + #10 +
                                '(select sum(be2.dCantHH) from bitacoradeequipos be2' + #10 +
                                '((:folio=-1) or (:folio<>-1 and be2.sNumeroOrden=be.sNumeroOrden))' + #10 +
                                'and be2.dIdFecha < :Inicio and be2.sIdEquipo=be.sIdEquipo) ' +
                                'as Anterior from bitacoradeequipos be inner join equipos e on '+
                                '(e.sContrato=:contratoBarco and e.sIdEquipo=be.sIdEquipo) inner join bitacoradeactividades '+
                                'ba on(ba.sContrato = be.sContrato and ba.dIdFecha = be.didfecha and ba.sNumeroOrden = '+
                                'be.sNumeroOrden and ba.iIdDiario = be.iIdDiario and ba.iIdTarea = be.iIdTarea and '+
                                'ba.iIdActividad=be.iIdActividad) '+
                                'where ((:Orden=-1) or (:Orden<>-1 and be.sContrato =:Orden)) and '+
                                '((:folio=-1) or (:folio<>-1 and be.sNumeroOrden=:folio)) ' +
                                'and be.dIdFecha between :Inicio and :termino and ' +
                                '((:Tipo="TE" and e.sIdTipoEquipo =:Tipo) or (:Tipo<>"TE" and e.sIdTipoEquipo<>"TE" )) '  +
                                'and e.lCobro=:Cobro ' +
                                'group by e.sIdEquipo '+
                                'order by e.iItemOrden,e.sIdEquipo',
                                'group by be.sContrato,be.sNumeroOrden,',
                                'order by be.sContrato,be.sNumeroOrden,'

                                )
                              ,
                                ('select b.sContrato as Contrato,b.sNumeroOrden as Folio,i.sIdInsumo as sIdRecurso,'+
                                'i.mDescripcion as descripcion,i.sMedida as medida, sum(b.dCantidad) as total,i.sIdProveedor,'+
                                'b.dFechaSalida as dIdFecha,ifnull((select sdescripcion from anexos where sTipo="MATERIAL" '+
                                'and sTierra="No" order by iorden limit 1),"No existe") as Titulo, ' +
                                'ifnull((select sAnexo from anexos where sTipo="MATERIAL" and sTierra="No" '+
                                'order by iorden limit 1),"No existe") as sAnexo,i.dVentaMn,i.dVentaDll, ' +
                                 'ifnull((select sMedida from anexos where sTipo="MATERIAL" and sTierra="No" '+
                                'order by iorden limit 1),"No existe") as sMedidaAnexo,'+
                                '(select sum(b2.dCantidad) from bitacoradesalida b2' + #10 +
                                'where ((:orden=-1) or (:orden<>-1 and b2.sContrato =b.sContrato)) and' + #10 +
                                '((:Folio=-1) or (:Folio<>-1 and b2.sNumeroOrden=b.sNumeroOrden))' + #10 +
                                'and b2.dFechaSalida < :Inicio and b2.sIdInsumo=b.sIdInsumo) '  +
                                'as Anterior '+
                                'from almacen_salida s '+
                                'inner join bitacoradesalida b '+
                                'on(b.sContrato = s.sContrato and b.iFolioSalida = s.iFolioSalida and b.sext = s.sext) '+
                                'inner join insumos i '+
                                'on (i.sContrato =:ContratoBarco and i.sIdInsumo = b.sIdInsumo and i.sTrazabilidad = b.sTrazabilidad ) '+
                                'where ((:orden=-1) or (:orden<>-1 and s.sContrato =:orden)) and '+
                                '((:Folio=-1) or (:Folio<>-1 and s.sNumeroOrden=:Folio)) '+
                                'and b.dFechaSalida between :Inicio and :termino '+
                                'group by b.sIdInsumo '+
                                'Order by b.sIdInsumo',
                                'group by b.sContrato,b.sNumeroOrden,',
                                'Order by b.sContrato,b.sNumeroOrden,'
                                )
                              , ('select bp.sContrato as contrato,bp.sNumeroOrden as Folio,c.sIdPernocta as sIdRecurso,c.sDescripcion as Descripcion,c.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="PERNOCTA" limit 1),"No existe") as Titulo,'+
                                'ifnull((select sAnexo from anexos where sTipo="PERNOCTA" limit 1),"N/E") as sAnexo,'+
                                'ifnull((select sMedida from anexos where sTipo="PERNOCTA" limit 1),"N/E") as sMedidaAnexo,'+
                                '(sum(bp.dCantHH) - ifnull((select ifnull(sum(bpr.dCantidad),0) from bitacoradepernocta bpr '+
                                'where bpr.sContrato = ba.sContrato and ba.didfecha=bpr.didfecha and (bpr.sNumeroOrden <> "@" '+
                                'and bpr.sNumeroOrden=ba.snumeroorden) and bpr.sIdCuenta="" group by bpr.sContrato),0)) as Total,'+
                                'ifnull((select ifnull(sum(bpr.dCantidad),0) from bitacoradepernocta bpr '+
                                'where bpr.sContrato = ba.sContrato and ba.didfecha=bpr.didfecha and (bpr.sNumeroOrden <> "@" '+
                                'and bpr.sNumeroOrden=ba.snumeroorden) and bpr.sIdCuenta=c.sIdCuenta '+
                                'group by bpr.sContrato,bpr.sNumeroOrden),0) as ajuste, '+
                                'bp.didfecha,c.dVentaMn,c.dVentaDll, 0 as Anterior from cuentas c '+
                                'left join bitacoradeactividades ba on(((:Orden =-1) or (:Orden<>-1 and ba.sContrato=:Orden)) '+
                                'and ((:Folio=-1) or (:Folio<>-1 and ba.sNumeroOrden=:Folio)) '+
                                ') left join bitacoradepersonal bp on (bp.sContrato =ba.sContrato and '+
                                'ba.sNumeroOrden=bp.sNumeroOrden and  bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                                'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad and bp.sTipoPernocta = c.sIdPernocta '+
                                ') left join personal p on (p.sContrato = :contratoBarco '+
                                'and p.sIdPersonal = bp.sIdPersonal) left join tiposdepersonal tp on ( p.sIdTipoPersonal = '+
                                'tp.sIdTipoPersonal) left join ordenesdetrabajo ot '+
                                'on (ot.sContrato =ba.sContrato and ot.sNumeroOrden = ba.sNumeroOrden) left join pernoctan pr '+
                                'on (pr.sIdPernocta = ot.sIdPernocta) left join plataformas pl on(pl.sIdPlataforma=ot.sIdPlataforma) '+
                                'where ((:Orden =-1) or (:Orden<>-1 and bp.sContrato=:Orden)) and ((:Folio=-1) or (:Folio<>-1 and bp.snumeroorden=:Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'group by c.sIdPernocta '+
                                'order by c.sIdPernocta ',
                                'group by bp.sContrato,bp.sNumeroOrden,',
                                'order by bp.sContrato,bp.sNumeroOrden,'
                                )
                              ,
                                ('select ba.sContrato as contrato,ba.sNumeroOrden as Folio,ba.sNumeroActividad as sIdRecurso,' + #10 +
                                'ao.mDescripcion as Descripcion,ao.sMedida as medida,sum(ba.dCantidad) as total,' + #10 +
                                'a.sDescripcion as Titulo,a.sAnexo,a.sMedida as sMedidaAnexo,ba.didfecha,ao.dVentaMn,ao.dVentaDll,'+
                                '(select sum(ba2.dCantidad) from bitacoradeactividades ba2' + #10 +
                                'where ((:Orden =-1) or (:Orden<>-1 and ba2.sContrato=ba.sContrato)) and' + #10 +
                                '((:Folio=-1) or (:Folio<>-1 and ba2.snumeroorden=ba.snumeroorden)) and ba2.dIdFecha < :Inicio and' + #10 +
                                'ba2.sNumeroActividad=ba.sNumeroActividad) ' +
                                'as Anterior from bitacoradeactividades ba' + #10 +
                                'inner join actividadesxorden ao on(ao.sContrato=ba.sContrato and ao.sNumeroOrden=ba.sNumeroOrden' + #10 +
                                'and ao.swbs=ba.swbs and ao.sNumeroACtividad=ba.sNumeroActividad) left join anexos a' + #10 +
                                'on(a.sAnexo=ao.sAnexo) where ((:Orden =-1) or (:Orden<>-1 and ba.sContrato=:Orden)) and' + #10 +
                                '((:Folio=-1) or (:Folio<>-1 and ba.snumeroorden=:Folio)) and ba.dIdFecha between :Inicio and' + #10 +
                                ':termino and ba.sIdTipoMovimiento = "E" and ao.sTipoAnexo= "PU"' + #10 +
                                'group by ba.swbs' + #10 +
                                'order by ao.iItemOrden,ba.swbs',
                                'group by ba.sContrato,ba.sNumeroOrden,',
                                'order by ba.sContrato,ba.sNumeroOrden,'
                                )
                              );

  //FtGenerador=(FtGBarco,FtGPersonal,FtGEquipo,FtGMaterial,FtGPernocta);
{$EndRegion}
var
  QrConsulta,QrFolios:TzREadOnlyQUery;
  MryRecursos:TjvMemoryData;
  dsGeneradorGeneral:TfrxDBDataset;
  IPos,iPosTmp,
  iGrupo:Integer;
  sIdRecurso:String;
  ValTmp:Variant;
  CantTmp,dTotal:Double;
  sDia,sConFol:String;
  TmpMantisa,iCiclo,
  MaxCiclo:Integer;
begin
  QrFolios:=TzREadOnlyQUery.Create(nil);
  QrConsulta:=TzREadOnlyQUery.Create(nil);
  try
    try
      QrConsulta.Connection:=Connection.zConnection;
      QrFolios.Connection:=Connection.zConnection;
      dsGeneradorGeneral:=TfrxDBDataset.Create(nil);
      MryRecursos:=TjvMemoryData.Create(nil);

      dsGeneradorGeneral.UserName:='dsGeneradorGeneral';

      QrFolios.Active:=False;
      QrFolios.SQL.Text:= 'select ot.sContrato,ot.mdescripcion as descFolio from ordenesdetrabajo ot '+
                          'where ot.sContrato=:Contrato and ot.sNumeroOrden=:Folio ';


      with MryRecursos do
      begin
        Active:=False;
        FieldDefs.Add('sContrato', ftString, 100, True);
        FieldDefs.Add('sFolio', ftString, 100, True);
        FieldDefs.Add('sDescFolio', ftString, 500, True);
        FieldDefs.Add('sidrecurso', ftString, 100, True);
        FieldDefs.Add('sdescripcion', ftString, 500, True);
        FieldDefs.Add('sAnexo', ftString, 50, false);
        FieldDefs.Add('sTitulo', ftString, 500, false);
        FieldDefs.Add('smedida', ftString, 100, True);
        FieldDefs.Add('smedidaanexo', ftString, 100, True);
        FieldDefs.Add('dTotal', FtFloat, 0, True);
        FieldDefs.Add('dAcumulado', FtFloat, 0, True);
        if isMonto then
        begin
          FieldDefs.Add('dVentaMn', FtCurrency, 0, True);
          FieldDefs.Add('dVentaDll', FtCurrency, 0, True);
          FieldDefs.Add('dTotalMn', FtCurrency, 0, True);
          FieldDefs.Add('dTotalDll', FtCurrency, 0, True);
          FieldDefs.Add('dAcumuladoMn', FtCurrency, 0, True);
        end;
        Active:=true;
      end;


      if ord(TipoEstimacion)=0 then
      begin
        iCiclo:=1;
        MaxCiclo:=length(Asql)
      end
      else
      begin
        iCiclo := ord(TipoEstimacion);
        MaxCiclo:=iCiclo;
      end;




      for iGrupo := iCiclo to MaxCiclo do
      begin
        if iGrupo=1 then
          TmpMantisa:=4
        else
          TmpMantisa:=2;

        QrConsulta.Active:=fALSE;

        if  ListaParams.Values['TIPO']='GENERAL' then
          QrConsulta.SQL.text:=aSql[iGrupo,1]
        else
        begin
          QrConsulta.SQL.text:=AnsiReplaceText(aSql[iGrupo,1],'group by',aSql[iGrupo,2]);
          QrConsulta.SQL.text:=AnsiReplaceText(QrConsulta.SQL.text,'order by',aSql[iGrupo,3]);
        end;

        QrConsulta.ParamByName('Orden').AsString    :=ListaParams.Values['CONTRATO'];
        QrConsulta.ParamByName('Folio').AsString    := ListaParams.Values['FOLIO'];
        QrConsulta.ParamByName('Inicio').AsDate:=strToDate(ListaParams.Values['INICIO']);
        QrConsulta.ParamByName('Termino').AsDate:=strToDate(ListaParams.Values['TERMINO']);
        if (iGrupo<>ord(FtGAnexoPu)) then
          QrConsulta.ParamByName('contratoBarco').AsString := ListaParams.Values['CONTRATO_BARCO'];
        if (iGrupo=ord(FtGPersonal)) or (iGrupo=Ord(FtGEquipo)) then
        begin
          QrConsulta.ParamByName('Tipo').AsString    :=ListaParams.Values['TIEMPO'];
          QrConsulta.ParamByName('Cobro').AsString    := ListaParams.Values['COBRO'];
        end;

        QrConsulta.Open;

        sIdRecurso:='';
        sConFol:='';
        while not QrConsulta.Eof do
        begin
          if sConFol<>(QrConsulta.FieldByName('Contrato').AsString + QrConsulta.FieldByName('Folio').AsString) then
          begin
            QrFolios.Active:=false;
            QrFolios.ParamByName('Contrato').AsString:=QrConsulta.FieldByName('Contrato').AsString;
            QrFolios.ParamByName('Folio').AsString:=QrConsulta.FieldByName('Folio').AsString;
            QrFolios.open;
            sConFol:=(QrConsulta.FieldByName('Contrato').AsString + QrConsulta.FieldByName('Folio').AsString);
          end;


          if ((iGrupo<>5) and (sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString)) or
            ((iGrupo=5) and (ListaParams.Values['TIPO']='GENERAL') and (sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString))
             or ((iGrupo=5) and (ListaParams.Values['TIPO']<>'GENERAL') and
             (sIdRecurso<>QrConsulta.FieldByName('Folio').AsString + QrConsulta.FieldByName('sIdRecurso').AsString)) then
          begin
            if MryRecursos.State=dsInsert then
            begin
              if dTotal<>xRound(ValTmp,TmpMantisa) then
              begin
                MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,TmpMantisa);
              end
              else
                MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;

              MryRecursos.FieldByName('dAcumulado').AsFloat:= MryRecursos.FieldByName('dTotal').AsFloat +
                                                              xRound(QrConsulta.FieldByName('Anterior').Value,TmpMantisa);
              if IsMonto then
              begin
                MryRecursos.FieldByName('dTotalMn').AsCurrency:=(MryRecursos.FieldByName('dTotal').AsFloat * MryRecursos.FieldByName('dventaMn').AsCurrency);
                MryRecursos.FieldByName('dTotalDll').AsCurrency:=(MryRecursos.FieldByName('dTotal').AsFloat * MryRecursos.FieldByName('dventaDll').AsCurrency);
                MryRecursos.FieldByName('dAcumuladoMn').AsCurrency:=(MryRecursos.FieldByName('dAcumulado').AsFloat * MryRecursos.FieldByName('dventaMn').AsCurrency);


              end;

              MryRecursos.Post;
            end;

            ValTmp:=0;
            CantTmp:=0;
            iPosTmp:=0;
            dTotal:=0;

            if iGrupo=5 then
            begin
              if  ListaParams.Values['TIPO']<>'GENERAL' then
                sIdRecurso:=QrConsulta.FieldByName('Folio').AsString + QrConsulta.FieldByName('sIdRecurso').AsString
              else
                sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
            end
            else
              sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.Append;
            MryRecursos.FieldByName('sContrato').AsString:=QrConsulta.FieldByName('Contrato').AsString;

            if  ListaParams.Values['TIPO']='GENERAL' then
              MryRecursos.FieldByName('sFolio').AsString:=QrConsulta.FieldByName('Contrato').AsString
            else
              MryRecursos.FieldByName('sFolio').AsString:=QrConsulta.FieldByName('Folio').AsString;

            MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
            MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
            MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
            MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
            MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
            MryRecursos.FieldByName('smedidaanexo').AsString:=QrConsulta.FieldByName('smedidaanexo').AsString;
            MryRecursos.FieldByName('dTotal').AsFloat:=0;
            MryRecursos.FieldByName('dAcumulado').AsFloat:=0;

            if isMonto then
            begin
              MryRecursos.FieldByName('dVentaMn').AsCurrency:=QrConsulta.FieldByName('dventaMn').AsCurrency;
              MryRecursos.FieldByName('dVentaDll').AsCurrency:=QrConsulta.FieldByName('dventaDll').AsCurrency;
              MryRecursos.FieldByName('dTotalMn').AsCurrency:=0;
              MryRecursos.FieldByName('dTotalDll').AsCurrency:=0;
              MryRecursos.FieldByName('dAcumuladoMn').AsFloat:=0;
            end;
          end;


          if (iGrupo=2) or (iGrupo=5) then
          begin
              ValTmp:=ValTmp + (QrConsulta.FieldByName('total').Value + QrConsulta.FieldByName('Ajuste').AsFloat);
              if CantTmp<(xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat) then
                CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat;
          end
          else
          begin
           
            ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
            if CantTmp<xRound(QrConsulta.FieldByName('total').Value,TmpMantisa) then
            begin
              CantTmp:=xRound(QrConsulta.FieldByName('total').Value,TmpMantisa);
              
            end;
          end;

          //dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;

          QrConsulta.Next;
        end;

        if MryRecursos.State=dsInsert then
        begin
          if dTotal<>xRound(ValTmp,TmpMantisa) then
          begin
            MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,TmpMantisa);
          end
          else
          begin
            MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
          end;
          MryRecursos.FieldByName('dAcumulado').AsFloat:= MryRecursos.FieldByName('dTotal').AsFloat +
                                                              xRound(QrConsulta.FieldByName('Anterior').Value,TmpMantisa);

          if IsMonto then
          begin
            MryRecursos.FieldByName('dTotalMn').AsCurrency:=(MryRecursos.FieldByName('dTotal').AsFloat * MryRecursos.FieldByName('dventaMn').AsCurrency);
            MryRecursos.FieldByName('dTotalDll').AsCurrency:=(MryRecursos.FieldByName('dTotal').AsFloat * MryRecursos.FieldByName('dventaDll').AsCurrency);
            MryRecursos.FieldByName('dAcumuladoMn').AsCurrency:=(MryRecursos.FieldByName('dAcumulado').AsFloat * MryRecursos.FieldByName('dventaMn').AsCurrency);

          end;

          MryRecursos.Post;
        end;
      end;
      dsGeneradorGeneral.DataSet:=MryRecursos;
      dsGeneradorGeneral.FieldAliases.Clear;
      Reporte.DataSets.Add(dsGeneradorGeneral);
    except
      on e:Exception do
        Showmessage(e.Message + ', ' + e.ClassName);
    end;

  finally
    QrConsulta.Destroy;
  end;
end;


Procedure TdProrrateaPernocta(const ParamContrato:String;Const ParamFecha:TDatetime);
var
  QrFolios,QrHoras,QrAdicional,
  QrPernoctas:TzReadOnlyQuery;
  QCmd:TzQuery;
  HorasFolio:TstringList;
  TotalHoras:Double;
  sHoras,sTotalHrs,HrsTotales,
  sFolio:String;
  i:Integer;
  dCant:Extended;
  Factor,FactorT:VAriant;
begin
  QrFolios:=TzReadOnlyQuery.create(nil);
  QrHoras:=TzReadOnlyQuery.create(nil);
  QrAdicional:=TzReadOnlyQuery.create(nil);
  QrPernoctas:=TzReadOnlyQuery.create(nil);
  QCmd:=TzQuery.Create(nil);
  HorasFolio:=TstringList.Create;
  try
    try
      QCmd.Connection:=Connection.zConnection;
      QrPernoctas.Connection:=Connection.zConnection;
      QrPernoctas.SQL.Text:='select * from cuentas';
// Rangel
      QrAdicional.Connection:=Connection.zConnection;
      QrAdicional.SQL.Text:='select tp.sDescripcion as sDescripcionTipo, p.sIdTipoPersonal, p.sDescripcion, sum(bpc.dCantidad) as dCantidad' + #10 +
                            'from bitacoradepersonal_cuadre bpc' + #10 +
                            'inner join personal p' + #10 +
                            'on ( p.sContrato = :ContratoBarco and bpc.sIdPersonal = p.sIdPersonal )' + #10 +
                            'inner join tiposdepersonal tp' + #10 +
                            'on (tp.sIdTipoPersonal = p.sIdTipoPersonal )' + #10 +
                            'where bpc.sContrato =:Contrato and bpc.snumeroorden=:Folio' + #10 +
                            'and bpc.dIdFecha =:Fecha and tp.lPernocta="Si" and bpc.sTipoPernocta=:Pernocta' + #10 +
                            'order by tp.sIdTipoPersonal, bpc.sIdPersonal' ;

      QrHoras.Connection:=connection.zConnection;
      QrHoras.SQL.Text:='select * from bitacoradeactividades where sContrato=:Contrato and ' +
                        'dIdFecha=:Fecha and sNumeroOrden=:Folio and sIdTipoMovimiento="ED"';


      QrFolios.Connection:=Connection.zConnection;
      QrFolios.SQL.Text:= 'select * from bitacoradeactividades where scontrato=:Contrato' + #10 +
                          'and didfecha=:Fecha group by snumeroorden';
      QrFolios.ParamByName('Contrato').AsString:=ParamContrato;
      QrFolios.ParamByName('Fecha').AsDate:=ParamFecha;
      QrFolios.Open;
      HrsTotales:='00:00';
      while not QrFolios.Eof do
      begin
        QrHoras.Active:=false;
        QrHoras.ParamByName('Contrato').AsString:= QrFolios.FieldByName('sContrato').AsString;
        QrHoras.ParamByName('Fecha').AsDate:= QrFolios.FieldByName('dIdFecha').AsDateTime;
        QrHoras.ParamByName('Folio').AsString:= QrFolios.FieldByName('snumeroorden').AsString;
        QrHoras.Open;
        sTotalHrs:='00:00';
        while not QrHoras.Eof do
        begin
          sHoras:=sfnRestaHoras(QrHoras.FieldByName('sHoraFinal').AsString,QrHoras.FieldByName('sHoraInicio').AsString);
          sTotalHrs:=sfnSumaHoras(sTotalHrs,sHoras);
          QrHoras.Next;
        end;

        if HrsTotales<>'00:00' then
          HrsTotales:=sfnSumaHoras(HrsTotales,sTotalHrs)
        else
          HrsTotales:=sTotalHrs;
        HorasFolio.Add(QrFolios.FieldByName('sNumeroOrden').AsString + '=' + sTotalHrs);
        QrFolios.Next;
      end;

      QrPernoctas.Open;
      FactorT:=sHrsToExt(HrsTotales);//sfnHora(HrsTotales);
      QrFolios.First;
      while not QrFolios.Eof do
      begin
        QrPernoctas.First;
        while not QrPernoctas.Eof do
        begin
          dCant:=0;
          QCmd.Active:=false;
          QCmd.SQL.Text:='delete from bitacoradepernocta where sContrato=:Contrato and dIdFecha=:Fecha and '+
                         'sNumeroOrden=:Folio and sIdCuenta=:Pernocta';
          QCmd.ParamByName('Contrato').AsString:=QrFolios.FieldByName('sContrato').AsString;
          QCmd.ParamByName('Fecha').AsDate:= QrFolios.FieldByName('dIdFecha').AsDateTime;
          QCmd.ParamByName('Folio').AsString:= QrFolios.FieldByName('sNumeroOrden').AsString;
          QCmd.ParamByName('Pernocta').AsString:= QrPernoctas.FieldByName('sIdCuenta').AsString;
          QCmd.ExecSQL;

          for I := 1 to 2 do
          begin
            if i=1 then
              sFolio:='@'
            else
              sFolio:=QrFolios.FieldByName('sNumeroOrden').AsString;

            QrAdicional.Active:=false;
            QrAdicional.ParamByName('ContratoBarco').AsString:= global_contrato_barco;
            QrAdicional.ParamByName('Contrato').AsString:=QrFolios.FieldByName('sContrato').AsString;
            QrAdicional.ParamByName('Folio').AsString:=sFolio;
            QrAdicional.ParamByName('Fecha').Asdate:=QrFolios.FieldByName('dIdFecha').AsDateTime;
            QrAdicional.ParamByName('Pernocta').AsString:= QrPernoctas.FieldByName('sIdCuenta').AsString;
            QrAdicional.Open;
            sHoras:='';
            if QrAdicional.RecordCount=1 then
            begin
              sHoras:=HorasFolio.Values[QrFolios.FieldByName('sNumeroOrden').AsString];
              if sHoras<>'' then
              begin
                Factor:=sHrsToExt(sHoras);
                if Factor>0 then
                begin
                  if i=1 then
                    dCant:=dCant + (QrAdicional.FieldByName('dCantidad').AsFloat * Factor)/FactorT
                  else
                    dCant:=dCant + (QrAdicional.FieldByName('dCantidad').AsFloat * Factor);
                end;
              end;
            end;
          end;

          QCmd.Active:=false;
          QCmd.SQL.Text:='insert into bitacoradepernocta(sContrato,dIdFecha,sNumeroOrden,sIdCuenta,dCantidad) '+
                         'values(:Contrato,:Fecha,:Folio,:Pernocta,:Cantidad)';
          QCmd.ParamByName('Contrato').AsString:=QrFolios.FieldByName('sContrato').AsString;
          QCmd.ParamByName('Fecha').AsDate:= QrFolios.FieldByName('dIdFecha').AsDateTime;
          QCmd.ParamByName('Folio').AsString:= QrFolios.FieldByName('snumeroorden').AsString;
          QCmd.ParamByName('Cantidad').value:=dCant;
          QCmd.ParamByName('Pernocta').AsString:= QrPernoctas.FieldByName('sIdCuenta').AsString;
          QCmd.ExecSQL;

          QrPernoctas.Next;
        end;
        QrFolios.Next;
      end;


    except
      on e:Exception do
      begin
        showmessage(e.Message + ', ' + e.ClassName);

      end;

    end;
  finally
    QrFolios.Destroy;
    HorasFolio.Destroy;
    QrHoras.Destroy;
    QrAdicional.Destroy;
    QrPernoctas.Destroy;
    QCmd.Destroy;

  end;



end;

Function TdProrrateoFolio(const ParamContrato:String;Const ParamFecha:TDatetime;Const IdDiario:Integer):Boolean;
var
  QrMovimiento,QrRecursos:TzReadOnlyQuery;
  QFolios:TzQuery;
  TotalPersonal:Double;
  res:Boolean;
  sCadena : string;
begin
  Res:=true;
  QrMovimiento:=TzREadOnlyQuery.Create(nil);
  QFolios:=TzQuery.Create(nil);
  QrRecursos:=TzReadOnlyQuery.Create(nil);
  try
    try
      QrMovimiento.connection:=Connection.zConnection;
      QFolios.connection:=Connection.zConnection;
      QrRecursos.Connection:=Connection.zConnection;
// Rangel
      QrRecursos.Active := False;
      QrRecursos.SQL.Clear;
      QrRecursos.SQL.Add('SELECT bp.sIdPersonal, bp.sDescripcion, ' +
                         'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                         'IF(SUM(bp.dCanthh) > 0,(round(SUM(bp.dCanthh),2) + ' +
                         ' IFNULL((SELECT ' +
                         '            SUM(ROUND(CAST(ba.dAjuste AS CHAR), 2)) ' +
                         '          FROM ' +
                         '            bitacoradeajustes ba ' +
                         '          WHERE ' +
                         '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                         '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                         '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)), SUM(bp.dCantidad)) AS Total '+
                         'FROM bitacoradepersonal bp ' +
                         'inner join bitacoradeactividades ba ' +
                         'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                         'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                         'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) ' +
                         'WHERE bp.scontrato = :Orden AND bp.sNumeroOrden = :Folio AND bp.didfecha = :Fecha ' +
                         'GROUP by bp.sNumeroOrden, bp.sIdPersonal ');

      TotalPersonal:=0;
      with QFolios do
      begin
        sql.text:='select * from movimientosxfolios where sContrato=:Contrato and dIdFecha=:fecha and iIdDiario=:Diario';
        ParamByname('Contrato').AsString :=ParamContrato;
        ParamByname('Fecha').AsDate      :=ParamFecha;
        ParamByname('Diario').AsInteger  :=IdDiario;
        Open;

        while not eof do
        begin
          QrRecursos.Active := False;
          QrRecursos.ParamByName('Orden').AsString      := FieldByName('sNumeroOrden').AsString;
          QrRecursos.ParamByName('Folio').AsString      := FieldByName('sFolio').AsString;
          QrRecursos.ParamByName('Fecha').AsDate        := FieldByName('dIdFecha').AsDateTime;
          QrRecursos.Open;
          if QrRecursos.recordcount > 0 then
          begin
              while not QrRecursos.Eof do
              begin
                  TotalPersonal:=TotalPersonal + (QrRecursos.fieldByName('Total').asfloat + QrRecursos.fieldByName('ajuste').asfloat);
                  QrRecursos.Next;
              end;
          end;

          next;
        end;
      end;

      with QrMovimiento do
      begin
        sql.text:= 'select * from movimientosdeembarcacion where sContrato=:Contrato and '  +
                  'dIdFecha=:Fecha and iIdDiario=:Diario';
        ParamByname('Contrato').AsString :=ParamContrato;
        ParamByname('Fecha').AsDate      :=ParamFecha;
        ParamByname('Diario').AsInteger  :=IdDiario;
        Open;

        QFolios.First;
        while not eof do
        begin
          while not QFolios.Eof do
          begin
            QrRecursos.Active := False;
            QrRecursos.ParamByName('Orden').AsString      := QFolios.FieldByName('sNumeroOrden').AsString;
            QrRecursos.ParamByName('Folio').AsString      := QFolios.FieldByName('sFolio').AsString;
            QrRecursos.ParamByName('Fecha').AsDate        := QFolios.FieldByName('dIdFecha').AsDateTime;
            QrRecursos.Open;
            if QrRecursos.recordcount > 0 then
            begin
              if TotalPersonal>0 then
              begin
                while not QrRecursos.Eof do
                begin
                    QFolios.Edit;
                    QFolios.FieldByName('dCanthh').value:= QFolios.FieldByName('dCanthh').value + (QrRecursos.fieldByName('Total').asfloat + QrRecursos.fieldByName('ajuste').asfloat);
                    QFolios.post;
                    QrRecursos.Next;
                end;
                sCadena :=  FloatToStr( (QFolios.FieldByName('dCanthh').value )
                                                  * StrToFloat(QrMovimiento.FieldByName('sFactor').AsString) /TotalPersonal);

                QFolios.Edit;
                QFolios.FieldByName('sFactor').value:= (StrToFloat(sCadena));
                QFolios.Post;
              end;
            end;

            QFolios.next;
          end;
          next;
        end;
      end;

    except
      Res:=false;
    end;


  finally
    QrMovimiento.destroy;
    QFolios.destroy;
    QrRecursos.destroy;
  end;
  Result:=res;
end;

Function TdAjustaFolio(const ParamContrato:String;Const ParamFecha:TDatetime):Boolean;
var
  QrRecursos:TzReadOnlyQuery;
  QFolios:TzQuery;
  TotalPersonal:Double;
  res:Boolean;
begin
  Res:=true;
  QFolios:=TzQuery.Create(nil);
  QrRecursos:=TzReadOnlyQuery.Create(nil);
  try
    try
      QFolios.connection:=Connection.zConnection;
      QrRecursos.Connection:=Connection.zConnection;

      QrRecursos.Active := False;
      QrRecursos.SQL.Clear;
      QrRecursos.SQL.Add('select sum(sFactor) as suma from movimientosxfolios '+
                         'where sContrato =:Contrato and dIdFecha =:Fecha group by sContrato ');
      QrRecursos.ParamByName('Contrato').AsString      := ParamContrato;
      QrRecursos.ParamByName('Fecha').AsDate        := ParamFecha;
      QrRecursos.Open;

      if QrRecursos.RecordCount > 0 then
         if QrRecursos.FieldByName('suma').AsFloat > 1 then
         begin
            QFolios.Active := False;
            QFolios.sql.text:='select * from movimientosxfolios where sContrato=:Contrato and dIdFecha=:Fecha '+
                              'order by dCantHH DESC, sFactor DESC';
            QFolios.ParamByname('Contrato').AsString :=ParamContrato;
            QFolios.ParamByname('Fecha').AsDate      :=ParamFecha;
            QFolios.Open;

            QFolios.Edit;
            QFolios.FieldByName('sFactor').AsFloat := QFolios.FieldByName('sFactor').AsFloat - (QrRecursos.FieldByName('suma').AsFloat - 1);
            QFolios.Post;
         end;

    except
      Res:=false;
    end;


  finally
    QFolios.destroy;
    QrRecursos.destroy;
  end;
  Result:=res;
end;


procedure ReportePDF_ClearDataset(Var Reporte: TfrxReport);
var
  i:Integer;
begin
  for I := 0 to Reporte.DataSets.Count - 1 do
    Reporte.DataSets.Items[i].DataSet.Destroy;


  Reporte.DataSets.Clear;
end;

Procedure ReportePDF_MovimientosLogisticos2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
Var
  zqProrrateo,
  zqMovimientos,
  zqFactor,
  zqArribos,
  ZMovimientos : TZReadOnlyQuery;

  Td_Movimientos,
  Td_prorrateoBarco,
  Td_Arribos,
  Td_ImpMovimientos,
  Td_MovFolios       : TfrxDBDataset;
  rxMovimientos      : TClientDataSet;
  mDatos:TJvMemoryData;
  mMovimientos:TJvMemoryData;
  iIdDiarioTmp:Integer;
  useOrigen:Boolean;
  CuentaCol, CuentaFolios, num : integer;
  sMovimientos, sTipo : string;
  dSuma : double;
begin
  Try
    zqProrrateo   := TZReadOnlyQuery.Create(nil);
    zqMovimientos := TZReadOnlyQuery.Create(nil);
    zqArribos     := TZReadOnlyQuery.Create(nil);
    zMovimientos  := TZReadOnlyQuery.Create(nil);

    mMovimientos:=TJvMemoryData.Create(nil);
    useOrigen:=false;
    Td_Movimientos:= TfrxDBDataset.Create(nil);
    Td_Movimientos.UserName:='Td_Movimientos';

    Td_ImpMovimientos:= TfrxDBDataset.Create(nil);
    Td_ImpMovimientos.UserName:='Td_ImpMovimientos';

    Td_prorrateoBarco:= TfrxDBDataset.Create(nil);
    Td_prorrateoBarco.UserName:='Td_Prorrateobarco';

    Td_Arribos:= TfrxDBDataset.Create(nil);
    Td_Arribos.UserName:='Td_Arribos';

    Td_MovFolios:= TfrxDBDataset.Create(nil);
    Td_MovFolios.UserName:='Td_MovFolios';

    mDatos:=TJvMemoryData.Create(nil);
    with mDatos do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidrecurso', ftString, 10, True);
      FieldDefs.Add('sdescripcion', ftString, 100, True);
      FieldDefs.Add('smedida', ftString, 100, true);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      for CuentaCol:=1 to 7 do
      begin
        FieldDefs.Add('dcantvew' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sTipo'+ Inttostr(CuentaCol), ftString, 4, false);
      end;
      Active:=true;
    end;

    if (TImpresion=FtsCaratula) or (TImpresion=FtsAll) then
    begin
        {$REGION 'CONSULTA'}
        zqProrrateo.Connection := Connection.zConnection;
        zqProrrateo.SQL.Add('SELECT ' +
                            '  tm.sIdTipoMovimiento, ' +
                            '  tm.sDescripcion, ' +
                            '  tm.sTipo, ' +
                            '  ( ' +
                            '    SELECT ' +
                            '      ifnull(SUM( me.sFactor ),0) ' +
                            '    FROM movimientosdeembarcacion AS me ' +
                            '    WHERE me.sContrato = tm.sContrato ' +
                            '    AND me.sOrden = :Orden ' +
                            '    AND me.dIdFecha = :Fecha ' +
                            '    AND me.sIdFase = "OPER" ' +
                            '    AND me.sClasificacion = tm.sIdTipoMovimiento ' +
                            '  ) AS sFactor,tm.lImprimeCeros ' +
                            'FROM tiposdemovimiento AS tm ' +
                            'WHERE ' +
                            '  tm.sContrato = :ContratoBarco ' +
                            '  AND tm.sClasificacion = "Movimiento de Barco" '+
                            'having ((sFactor<>0) or ((tm.lImprimeCeros ="Si") and (sFactor=0)))');
        zqProrrateo.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        zqProrrateo.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
        zqProrrateo.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
        zqProrrateo.Open;
        {$ENDREGION}

        {$REGION 'MOVIMIENTOS DE BARCO'}
        zqMovimientos.Connection := connection.zconnection;
        zqMovimientos.SQL.Add('select me.shorainicio, me.shorafinal, concat(me.sHoraInicio, " - " ,  me.sHoraFinal) as sHorario, '+
                              'me.mDescripcion, tm.sIdTipoMovimiento , me.sClasificacion, '+
                              'me.sFactor, me.iiddiario '+
                              'from movimientosdeembarcacion me '+
                              'inner join tiposdemovimiento tm '+
                              'on (me.sContrato = tm.sContrato and me.sClasificacion = tm.sIdTipoMovimiento) '+
                              'where me.sContrato =:ContratoBarco '+
                              'and me.dIdFecha = :fecha '+
                              'and me.sOrden =:orden ' +
                              'and me.sActividades = "MOV" '+
                              'group by me.iIddiario order by me.shorainicio,me.iIdDiario');
        zqMovimientos.ParamByName('contratoBarco').AsString := global_contrato_barco;
        zqMovimientos.ParamByName('fecha').AsDate           := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        zqmovimientos.ParamByName( 'orden' ).AsString       := global_contrato;
        zqMovimientos.Open;


        zqArribos.Connection := connection.zconnection;
        zqArribos.SQL.Add('select concat( me.sHoraInicio , " - ", me.sHoraFinal ) as sHorario, me.sHoraInicio, me.sHoraFinal, '+
                  'me.mDescripcion '+
                  'from movimientosdeembarcacion me '+
                  'where me.sContrato = :contratoBarco '+
                  'and me.dIdFecha = :fecha '+
                  'and me.sOrden = :orden '+
                  'and me.sClasificacion = "" and sTipo <> "MOVIMIENTO" '+
                  'order by me.sHoraInicio');
        zqArribos.ParamByName('contratoBarco').AsString := global_contrato_barco;
        zqArribos.ParamByName('fecha').AsDate := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        zqArribos.ParamByName( 'orden' ).AsString := global_contrato;
        zqArribos.Open;

        if zqArribos.RecordCount = 0 then
        begin
            zqArribos.Connection := connection.zconnection;
            zqArribos.SQL.Clear;
            zqArribos.SQL.Add('select " " as sHorario, " " as sHoraInicio, " " as sHoraFinal, '+
                      ' " " as mDescripcion '+
                      'from movimientosdeembarcacion group by sContrato ');
            zqArribos.Open;
        end;

        {$ENDREGION}

        {$REGION 'Resumen de movimientos x folio'}

        //Buscamos los movimientos de barco
        zMovimientos.Connection := connection.zconnection;
        zMovimientos.SQL.Add('select me.sContrato, me.sOrden, t.sIdTipoMovimiento, t.sDescripcion, t.sMedida, ot.sIdFolio, ot.sTipo, SUM(mf.sFactor) as Factor, sum(mf.sFactorBarco) as FactorBarco from tiposdemovimiento t '+
                    'inner  join movimientosdeembarcacion me on (me.sContrato = t.sContrato and me.dIdFecha =:Fecha and me.sClasificacion  = t.sIdTipoMovimiento) '+
                    'inner join movimientosxfolios mf on (mf.sContrato = me.sContrato and me.iIdDiario = mf.iIdDiario and mf.dIdFecha = me.dIdFecha and me.sOrden = mf.sNumeroOrden and mf.sFactor > 0) '+
                    'inner join ordenesdetrabajo ot on (ot.sContrato = mf.sNumeroOrden and ot.sNumeroOrden = mf.sFolio) '+
                    'where t.sContrato =:ContratoBarco and t.sClasificacion = "Movimiento de Barco" '+
                    'group by t.sIdTipoMovimiento, mf.sFolio order by t.sIdTipoMovimiento, ot.iOrden');
        zMovimientos.ParamByName('contratoBarco').AsString := global_contrato_barco;
        zMovimientos.ParamByName('Fecha').AsDate           := ReporteDiario.FieldByName('dIdFecha').AsDateTime;;
        zMovimientos.Open;

        CuentaFolios := 1;
        if zMovimientos.RecordCount > 7 then
           CuentaFolios := 2;

       { mDatos:=TJvMemoryData.Create(nil);
        with mDatos do
        begin
          Active:=false;
          FieldDefs.Add('iGrupo', ftInteger, 0, True);
          FieldDefs.Add('sidrecurso', ftString, 10, True);
          FieldDefs.Add('sdescripcion', ftString, 100, True);
          FieldDefs.Add('smedida', ftString, 100, true);
          FieldDefs.Add('dcantTotal', FtFloat, 0, True);
          for CuentaCol:=1 to 7 do
          begin
            FieldDefs.Add('dcantvew' + Inttostr(CuentaCol), FtFloat, 0, false);
            FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
            FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
            FieldDefs.Add('sTipo'+ Inttostr(CuentaCol), ftString, 4, false);
          end;
          Active:=true;
        end;      }

        //Primero buscamos los movimeintos de barco
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select sIdTipoMovimiento, sDescripcion, sMedida from tiposdemovimiento where sContrato =:Contrato '+
                                    'and sClasificacion = "Movimiento de Barco" order by sIdTipoMovimiento ');
        connection.zCommand.ParamByName('Contrato').AsString := global_contrato_barco;
        connection.zCommand.Open;

        for num := 1 to CuentaFolios do
        begin
            connection.zCommand.First;
            while not connection.zCommand.Eof do
            begin
               mDatos.Append;
               mDatos.FieldByName('iGrupo').AsInteger      := num;
               mDatos.FieldByName('sidrecurso').AsString   := connection.zCommand.FieldByName('sIdTipoMovimiento').AsString;
               mDatos.FieldByName('sdescripcion').AsString := connection.zCommand.FieldByName('sDescripcion').AsString;
               mDatos.FieldByName('smedida').AsString      := connection.zCommand.FieldByName('sMedida').AsString;
               mDatos.FieldByName('dcantTotal').AsFloat    := 0;
               for CuentaCol := 1 to 7 do
               begin
                  mDatos.FieldByName('sNumeroOrden'+Inttostr(CuentaCol)).AsString := '';
                  mDatos.FieldByName('sTipo'+Inttostr(CuentaCol)).AsString        := '';
                  mDatos.FieldByName('dCantidad'+Inttostr(CuentaCol)).AsFloat     := 0;
               end;
               mDatos.Post;
               connection.zCommand.Next;
            end;
        end;

        sMovimientos := '';
        zmovimientos.First;
        mDatos.First;
        num := 1;
        while not zmovimientos.Eof do
        begin
            if sMovimientos <> zMovimientos.FieldByName('sIdTipoMovimiento').AsString then
            begin
                CuentaCol := 1;
                dSuma     := 0;
                mDatos.Filtered := False;
                mDatos.Filter   := 'sidrecurso = '+QuotedStr(zMovimientos.FieldByName('sIdTipoMovimiento').AsString);
                mDatos.Filtered := True;
                sMovimientos := zMovimientos.FieldByName('sIdTipoMovimiento').AsString;
            end;

            mDatos.Edit;
            mDatos.FieldByName('sNumeroOrden'+Inttostr(CuentaCol)).AsString := zMovimientos.FieldByName('sIdFolio').AsString;
            mDatos.FieldByName('sTipo'+Inttostr(CuentaCol)).AsString        := zMovimientos.FieldByName('sTipo').AsString;;
            mDatos.FieldByName('dCantidad'+Inttostr(CuentaCol)).AsFloat     := xRound(zMovimientos.FieldByName('Factor').AsFloat + zMovimientos.FieldByName('FactorBarco').AsFloat,6);
            dSuma := dSuma + xRound(zMovimientos.FieldByName('Factor').AsFloat + zMovimientos.FieldByName('FactorBarco').AsFloat,6);
            mDatos.FieldByName('dcantTotal').AsFloat   := dSuma;
            mDatos.Post;

            inc(CuentaCol);

            if CuentaCol = 8 then
            begin
               CuentaCol := 1;
               inc(num);
               mDatos.Locate('iGrupo', num, []);
            end;

            zMovimientos.Next;
        end;
        mDatos.Filtered := False;

        {$ENDREGION}

        {Esta secccion es para mostrar la hoja vacia sino existen datos.}
        useOrigen:=True;
        if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
        begin
            if zqMovimientos.RecordCount = 0 then
            begin
              zqMovimientos.SQL.Clear;
              zqMovimientos.Connection := connection.zconnection;
              zqMovimientos.SQL.Add('select "" as shorainicio, "" as shorafinal, "" as sHorario, '+
                            '"" as mDescripcion, "" as sIdTipoMovimiento , "" as sClasificacion, "" as sIdFolio, '+
                            '"" as sFactor, "" as iiddiario, "" as sIdPlataforma, "" as sdescripcion ');
              zqMovimientos.Open;
              useOrigen:=False;

              zqProrrateo.Connection := Connection.zConnection;
              zqProrrateo.SQL.Clear;
              zqProrrateo.SQL.Add('SELECT "" as sIdTipoMovimiento, "" as sDescripcion, "" as sTipo, ' +
                                  ' ""  AS sFactor ');
              zqProrrateo.Open;

              zqMovimientos.Connection := connection.zconnection;
              zqMovimientos.SQL.Clear;
              zqMovimientos.SQL.Add('select "" as shorainicio, "" as shorafinal, "" as sHorario, '+
                            '"" as mDescripcion, "" as sIdTipoMovimiento , "" as sClasificacion, "" as sIdFolio, '+
                            '"" as sFactor, "" as iiddiario, "" as sIdPlataforma, "" as sdescripcion ');
              zqMovimientos.Open;

              zqArribos.Connection := connection.zconnection;
              zqArribos.SQL.Clear;
              zqArribos.SQL.Add('select "" as sHoraInicio, "" as sHoraFinal, "" as sHorario, "" as mDescripcion ');
              zqArribos.Open;
            end;
        end;


    end
    else
    begin
       zqProrrateo.Connection := Connection.zConnection;
       zqProrrateo.SQL.Add('SELECT "" as sIdTipoMovimiento, "" as sDescripcion, "" as sTipo, ' +
                            ' ""  AS sFactor ');
       zqProrrateo.Open;

       zqMovimientos.Connection := connection.zconnection;
       zqMovimientos.SQL.Add('select "" as shorainicio, "" as shorafinal, "" as sHorario, '+
                      '"" as mDescripcion, "" as sIdTipoMovimiento , "" as sClasificacion, "" as sIdFolio, '+
                      '"" as sFactor, "" as iiddiario, "" as sIdPlataforma, "" as sdescripcion ');
       zqMovimientos.Open;

       zqArribos.Connection := connection.zconnection;
       zqArribos.SQL.Add('select "" as sHorario, "" as mDescripcion ');
       zqArribos.Open;

       rxMovimientos := TClientDataSet.Create(nil);
       rxMovimientos.Active := False;
       rxMovimientos.FieldDefs.Add('shorainicio', ftString, 0, True);
       rxMovimientos.CreateDataSet;
       rxMovimientos.Active := True;

       Td_ImpMovimientos.DataSet    := rxMovimientos;
       Td_ImpMovimientos.FieldAliases.Clear;
    end;
    if not Assigned(Td_ImpMovimientos.DataSet) then
    begin
      Td_ImpMovimientos.DataSet    := zqMovimientos;
      Td_ImpMovimientos.FieldAliases.Clear;
    end;

    Td_ProrrateoBarco.DataSet := zqProrrateo;
    Td_ProrrateoBarco.FieldAliases.Clear;

    Td_Movimientos.DataSet    := zqMovimientos;
    Td_Movimientos.FieldAliases.Clear;

    Td_Arribos.DataSet          := zqArribos;
    Td_Arribos.FieldAliases.Clear;

    Td_MovFolios.DataSet        := mDatos;
    Td_MovFolios.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_ImpMovimientos);
    Reporte.DataSets.Add(Td_prorrateobarco);
    Reporte.DataSets.Add(Td_Movimientos);
    Reporte.DataSets.Add(Td_Arribos);
    Reporte.DataSets.Add(Td_MovFolios);

  Finally

  End;
end;

Procedure ReportePDF_BalanceGeneral(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
Var
  sFolios: String;

  iFilaIniciaMovimientos: Integer;

  dExistenciaAnteriorAux,
  dConsumoTotal,
  dConsumoSinEquipos,
  dRecibido,
  dConsumoCIA,
  dConsumoNAV,
  dConsumoSDP,
  dConsumoCDP,
  dConsumoESP,
  dConsumoEquipos,
  dPrestamo,
  dProduccion : Real;

  lProducido, lTrasiego, lRecibido : Boolean;

  QryArribos,
  QryConsumos,
  QryCondiciones,
  QryEquipos,
  QryConsumoEmbarcaciones,
  QryProrrateoBarco,
  QryConsumosDiesel : TZQuery;
  QrySinEquipos: TZQuery;
  QryProrratipo: TZQuery;
  QryOrdenamiento: TZQuery;

  ConsumosAnt: array[1..5,1..2] of double;
  indice : integer;
  dPorcionDia: Real;
  rxBalance, rxBalanceResumen : TClientDataSet;

  Td_balance,
  Td_balanceResumen,
  Td_ImpBalance,
  Td_consumoEquipos  : TfrxDBDataset;
begin
  Try

    Td_balance:= TfrxDBDataset.Create(nil);
    Td_balance.UserName:='Td_balance';

    Td_ImpBalance:= TfrxDBDataset.Create(nil);
    Td_ImpBalance.UserName:='Td_ImpBalance';

    Td_balanceResumen:= TfrxDBDataset.Create(nil);
    Td_balanceResumen.UserName:='Td_balanceResumen';

    Td_consumoEquipos:= TfrxDBDataset.Create(nil);
    Td_consumoEquipos.UserName:='Td_consumoEquipos';

    QryConsumoEmbarcaciones:= TZQuery.Create(nil);
    QryConsumoEmbarcaciones.Connection := Connection.zConnection;

    QryProrrateoBarco:= TZQuery.Create(nil);
    QryProrrateoBarco.Connection := Connection.zConnection;

    QryOrdenamiento:=TZQuery.Create (nil);
    QryOrdenamiento.connection:= connection.zConnection;

    QryConsumos:=TZQuery.Create (nil);
    QryConsumos.connection:= connection.zConnection;

    QryEquipos:= TZQuery.Create (nil);
    QryEquipos.connection:= Connection.zConnection;

    QryProrratipo := TZQuery.Create(nil);
    QryProrratipo.Connection := Connection.zConnection;

    QryConsumosDiesel := TZQuery.Create(nil);
    QryConsumosDiesel.Connection := Connection.zConnection;

    rxBalance := TClientDataSet.Create(nil);
    rxBalance.Active := False;
    rxBalance.FieldDefs.Add('sIdEmbarcacion', ftString, 10, True);
    rxBalance.FieldDefs.Add('sEmbarcacion', ftString, 50, True);
    rxBalance.FieldDefs.Add('sConcepto', ftString, 25, True);
    rxBalance.FieldDefs.Add('sMedida', ftString, 25, True);
    rxBalance.FieldDefs.Add('dExistenciaAnt', ftFloat, 0, True);
    rxBalance.FieldDefs.Add('dRecepcion', ftFloat, 0, True);
    rxBalance.FieldDefs.Add('dProducido', ftFloat, 0, True);
    rxBalance.FieldDefs.Add('dConsumo', ftFloat, 0, True);
    rxBalance.FieldDefs.Add('dTransferido', ftFloat, 0, True);
    rxBalance.FieldDefs.Add('dConsumoOT', ftFloat, 0, True);
    rxBalance.FieldDefs.Add('dExistenciaAct', ftFloat, 0, True);
    rxBalance.CreateDataSet;
    rxBalance.Active := True;

    rxBalanceResumen := TClientDataSet.Create(nil);
    rxBalanceResumen.Active := False;
    rxBalanceResumen.FieldDefs.Add('sIdEmbarcacion', ftString, 10, True);
    rxBalanceResumen.FieldDefs.Add('sEmbarcacion', ftString, 50, True);
    rxBalanceResumen.FieldDefs.Add('dOperacion', ftFloat, 0, True);
    rxBalanceResumen.FieldDefs.Add('dEspera', ftFloat, 0, True);
    rxBalanceResumen.FieldDefs.Add('dNavegando', ftFloat, 0, True);
    rxBalanceResumen.FieldDefs.Add('dCia', ftFloat, 0, True);
    rxBalanceResumen.FieldDefs.Add('dTotal', ftFloat, 0, True);
    rxBalanceResumen.CreateDataSet;
    rxBalanceResumen.Active := True;

    if (TImpresion=FtsBGeneral) or (TImpresion=FtsAll) then
    begin
        {$REGION 'CONSULTA DE EMBARCACIONES'}
        QryConsumoEmbarcaciones.Active := False;
        QryConsumoEmbarcaciones.SQL.Add('SELECT e.sContrato, r.sIdEmbarcacion, 	e.sDescripcion '+
                                        'FROM recursos AS r ' +
                                        'INNER JOIN embarcaciones AS e ON(e.sIdEmbarcacion = r.sIdEmbarcacion) ' +
                                        'WHERE r.dIdFecha = :Fecha GROUP BY r.sIdEmbarcacion order by e.sTipo, e.sIdEmbarcacion ');
        QryConsumoEmbarcaciones.Params.ParamByName('Fecha').AsDateTime := reportediario.FieldByName('dIdFecha').AsDateTime;
        QryConsumoEmbarcaciones.Open;
        {$ENDREGION}

        {$REGION 'CONSULTA DE APLICA O NO PRODUCIDO, TRAIEGO, RECIBIDO AGUA, DIESEL'}
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('select * from recursosordenados_orden where dIdFecha =:Fecha ');
        connection.QryBusca.Params.ParamByName('Fecha').AsDate := reportediario.FieldValues['dIdFecha'];
        connection.QryBusca.Open;

        lProducido := False;
        lTrasiego  := False;
        lRecibido  := False;
        while not connection.QryBusca.Eof do
        begin
             if lProducido = False then
                if connection.QryBusca.FieldValues['lAplicaProducidoAgua'] = 'Si' then
                   lProducido := True;

             if lTrasiego = False then
                if connection.QryBusca.FieldValues['lAplicaTrasiegoAgua'] = 'Si' then
                   lTrasiego := True;

             if lRecibido = False then
                if connection.QryBusca.FieldValues['lAplicaRecibidoAgua'] = 'Si' then
                   lRecibido := True;

             connection.QryBusca.Next;
        end;
        {$ENDREGION}

        if QryConsumoEmbarcaciones.RecordCount > 0 then
        begin
            {$REGION 'CONSULTA DE CONSUMOS POR EMBARCACION'}
            QryOrdenamiento.SQL.Clear;
            QryOrdenamiento.SQL.Add('select * from recursosordenados_orden where sOrden =:Orden and dIdFecha =:fecha order by iIdOrden ');
            QryOrdenamiento.ParamByName('Fecha').AsDateTime  := Reportediario.FieldByName('dIdFecha').AsDateTime;
            QryOrdenamiento.ParamByName('Orden').AsString    := Reportediario.FieldByName('sOrden').AsString;
            QryOrdenamiento.Open;
            {$ENDREGION}

            while not QryOrdenamiento.Eof do
            begin
                QryConsumoEmbarcaciones.First;
                while Not QryConsumoEmbarcaciones.Eof do
                begin
                    {$REGION 'CONSULTA DE CONSUMOS POR EMBARCACION'}
                    QryConsumos.SQL.Clear;
                    QryConsumos.SQL.Add('SELECT re.iIdRecursoExistencia, re.lCombustible, re.sDescripcion, re.sMedida, r.dExistenciaAnterior, r.dRecibido, r.dPrestamos, r.dConsumo AS dConsumoReal, ' +
                                        'IFNULL(SUM(ce.dCantidad), 0) AS dConsumoEq, r.dExistenciaActual, r.dConsumoEquipos, '+
                                        '(r.dConsumo - IFNULL(SUM(ce.dCantidad), 0)) AS dConsumo, r.dProduccion, ' +
                                        '(select sum(dCantidad) from consumosdecombustibleporequipo where dIdFecha =:Fecha) as dConsumoTotalEquipos '+
                                        'FROM  recursos AS r '+
                                        'INNER JOIN recursosdeexistencias AS re ON (re.iIdRecursoExistencia = r.iIdRecursoExistencia) ' +
                                        'LEFT JOIN consumosdecombustibleporequipo AS ce ON (ce.dIdFecha = r.dIdFecha AND ce.iIdTipoConsumo = r.iIdRecursoExistencia AND ce.sContrato = :Contrato) ' +
                                        'WHERE 	r.dIdFecha = :Fecha AND r.sIdEmbarcacion = :Embarcacion GROUP BY re.iIdRecursoExistencia ');
                    QryConsumos.ParamByName('Contrato').AsString    := QryOrdenamiento.FieldByName('sOrden').AsString;
                    QryConsumos.ParamByName('Embarcacion').AsString := QryConsumoEmbarcaciones.FieldByName('sIdEmbarcacion').AsString;
                    QryConsumos.ParamByName('Fecha').AsDateTime     := Reportediario.FieldByName('dIdFecha').AsDateTime;
                    QryConsumos.Open;
                    {$ENDREGION}

                    {$REGION 'IMPRESIÓN DE REGISTROS'}
                    if QryConsumos.RecordCount > 0 then
                    begin
                        indice := 1;
                        while not QryConsumos.Eof do
                        begin
                            ConsumosAnt[indice, 1] := QryConsumos.FieldValues['iIdRecursoExistencia'];
                            inc(indice);
                            QryConsumos.Next;
                        end;

                        QryConsumos.First;
                        indice := 1;
                        while Not QryConsumos.Eof do
                        begin
                            dConsumoTotal := QryConsumos.FieldByName('dConsumoReal').AsFloat;
                            dConsumoTotal := RoundTo(dConsumoTotal, -3);
                            if QryConsumos.FieldValues['lCombustible'] = 'Si' then {Estos son recursos que son combustibles..}
                            begin
                               if QryOrdenamiento.FieldValues['lAplicaRecibidoDiesel'] = 'Si' then
                                  dRecibido := QryConsumos.FieldByName('dRecibido').AsFloat
                               else
                                  dRecibido := 0;
                            end
                            else
                            begin
                                dRecibido   := 0;
                                dProduccion := 0;
                                dPrestamo   := 0;

                                dRecibido   := QryConsumos.FieldByName('dRecibido').AsFloat;
                                dPrestamo   := QryConsumos.FieldByName('dPrestamos').AsFloat;
                                dProduccion := QryConsumos.FieldByName('dProduccion').AsFloat;
                            end;

                            dExistenciaAnteriorAux := QryConsumos.FieldByName('dExistenciaAnterior').AsFloat;

                            //Estos son los datos iniciales anterior, recibido y las descripciones..
                            if QryOrdenamiento.FieldValues['sOrden'] = Reportediario.FieldValues['sOrden'] then
                            begin
                                rxBalance.Append;
                                rxBalance.FieldValues['sIdEmbarcacion'] := QryConsumoEmbarcaciones.FieldByName('sIdEmbarcacion').AsString;
                                rxBalance.FieldValues['sEmbarcacion']   := QryConsumoEmbarcaciones.FieldByName('sDescripcion').AsString;
                                rxBalance.FieldValues['sConcepto']      := QryConsumos.FieldByName('sDescripcion').AsString;
                                rxBalance.FieldValues['sMedida']        := QryConsumos.FieldByName('sMedida').AsString;
                                rxBalance.FieldValues['dExistenciaAnt'] := QryConsumos.FieldByName('dExistenciaAnterior').AsFloat;
                                rxBalance.FieldValues['dConsumo']       := QryConsumos.FieldByName('dConsumoReal').AsFloat;
                                rxBalance.FieldValues['dExistenciaAct'] := QryConsumos.FieldByName('dExistenciaActual').AsFloat;
                            end;

                            {$REGION 'PRORRATEA LOS CONSUMOS DE EMBARCACION'}

                            {$REGION 'CONSULTA PRORRATEOS'}
                            QryProrrateoBarco.SQL.Clear;
                            QryProrrateoBarco.SQL.Add('SELECT tm.sIdTipoMovimiento, tm.sDescripcion, tm.sTipo, '+
                                                      'IFNULL(( SELECT SUM( me.sFactor ) FROM movimientosdeembarcacion AS me ' +
                                                      'WHERE me.sContrato = tm.sContrato AND me.dIdFecha = :Fecha  AND me.sIdFase = "OPER" AND me.sClasificacion = tm.sIdTipoMovimiento ), 0.000000) AS sFactor ' +
                                                      'FROM tiposdemovimiento AS tm '+
                                                      'WHERE tm.sContrato = :ContratoBarco AND tm.sClasificacion = "Movimiento de Barco" ');
                            QryProrrateoBarco.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                            QryProrrateoBarco.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
                            QryProrrateoBarco.Open;
                            {$ENDREGION}

                            {$REGION 'CONSULTA MOVIMIENTOS POR OT'}
                            // Consulta Para sacar valores CIA, SDP, CDP  y aqui mismo se hace el guardado de valores
                            QryProrratipo.SQL.Clear;
                            QryProrratipo.SQL.Add('SELECT tm.sIdTipoMovimiento, tm.sDescripcion, tm.sTipo, ' +
                                                  ' (SELECT SUM( me.sFactor ) FROM movimientosdeembarcacion AS me '+
                                                  '  WHERE me.sContrato = tm.sContrato AND me.sOrden = :Contrato AND me.dIdFecha = :Fecha '+
                                                  '  AND me.sIdFase = "OPER" AND me.sClasificacion = tm.sIdTipoMovimiento) AS sFactor ' +
                                                  'FROM tiposdemovimiento AS tm '+
                                                  'WHERE tm.sContrato = :ContratoBarco AND tm.sClasificacion = "Movimiento de Barco" ');
                            QryProrratipo.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                            QryProrratipo.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
                            QryProrratipo.ParamByName('Contrato').AsString      := QryOrdenamiento.FieldByName('sOrden').AsString;
                            QryProrratipo.Open;
                            {$ENDREGION}

                            dConsumoCIA := 0;
                            dConsumoSDP := 0;
                            dConsumoCDP := 0;
                            QryProrratipo.First;
                            while Not QryProrratipo.Eof do
                            begin
                                if QryProrratipo.FieldByName('sTipo').AsString = 'CIA' then
                                begin
                                    Try
                                       dConsumoCIA := (QryProrratipo.FieldByName('sFactor').AsFloat * dConsumoSinEquipos) ;
                                    Except
                                    End;
                                end;

                                if (QryProrratipo.FieldByName('sTipo').AsString = 'SDP') OR (QryProrratipo.FieldByName('sTipo').AsString = 'ESP') then
                                begin
                                    Try
                                       dConsumoSDP := (QryProrratipo.FieldByName('sFactor').AsFloat * dConsumoSinEquipos);
                                    Except
                                    End;
                                end;

                                if (QryProrratipo.FieldByName('sTipo').AsString = 'CDP')
                                    OR (QryProrratipo.FieldByName('sTipo').AsString = 'OPE')then
                                begin
                                   Try
                                      dConsumoCDP := dConsumoTotal;
                                      //dConsumoCDP := dConsumoCDP + (QryProrratipo.FieldByName('sFactor').AsFloat * dConsumoSinEquipos);
                                   Except
                                   End;
                                end;
                                QryProrratipo.Next;
                            end;
                           {$ENDREGION}

                            {$REGION 'IMPRESIÓN EN EXCEL'}
                            if QryOrdenamiento.FieldValues['sOrden'] = Reportediario.FieldValues['sOrden'] then
                            begin
                                //Esta condicion es cuando la OT no tiene movimientos de barco y solo se va a descontar los equipos JJF by ivan. 28 oct 2013

                                if (dConsumoCIA + dConsumoSDP + dConsumoCDP) = 0 then
                                    //Excel.Selection.Value := 0
                                else
                                begin
                                    {Aqui prorratemos la cantidad de Agua recibida por el factor del barco. JJF by ivan 18 Nov 2013}
                                    if lRecibido = False then
                                      // dRecibido := (QryConsumos.FieldByName('dRecibido').AsFloat / 100) * (dPorcionDia * 100)
                                    else
                                    begin
                                        if QryOrdenamiento.FieldValues['lAplicaRecibidoAgua'] = 'Si' then
                                            dRecibido := QryConsumos.FieldByName('dRecibido').AsFloat
                                    end;
                                    rxBalance.FieldValues['dRecepcion'] := dRecibido;
                                end;

                                rxBalance.FieldValues['dProducido']   := dProduccion;
                                dPrestamo                             := QryConsumos.FieldByName('dPrestamos').AsFloat;
                                rxBalance.FieldValues['dTransferido'] := dPrestamo;
                                rxBalance.FieldValues['dConsumoOT']   := QryConsumos.FieldByName('dConsumoEquipos').AsFloat;
                                rxBalance.Post;
                            end;

                            {Ahora llenamos el desgloce de consumo de diesel.}
                            if QryConsumos.FieldValues['lCombustible'] = 'Si' then
                            begin
                                if rxBalanceResumen.RecordCount = 0 then
                                begin
                                    rxBalanceResumen.Append;
                                    rxBalanceResumen.FieldValues['sIdEmbarcacion'] := QryConsumoEmbarcaciones.FieldByName('sIdEmbarcacion').AsString;
                                    rxBalanceResumen.FieldValues['sEmbarcacion']   := QryConsumoEmbarcaciones.FieldByName('sDescripcion').AsString;
                                    rxBalanceResumen.FieldValues['dOperacion']     := 0;
                                    rxBalanceResumen.FieldValues['dEspera']        := 0;
                                    rxBalanceResumen.FieldValues['dNavegando']     := 0;
                                    rxBalanceResumen.FieldValues['dCia']           := 0;
                                    rxBalanceResumen.FieldValues['dTotal']         := 0;
                                    rxBalanceResumen.Post;
                                end;

                                rxBalanceResumen.Edit;

                                dConsumoNAV := 0;
                                //Obtenemos el consumo en operacion
                                connection.zCommand.SQL.Clear;
                                connection.zCommand.SQL.Text := 'select '+
                                                                'tm.sIdTipoMovimiento, '+
                                                                'tm.sDescripcion, '+
                                                                'tm.sTipo, '+
                                                                'sum(( select ifnull( sum( me.sFactor ), 0 ) '+
                                                                  'from movimientosdeembarcacion as me '+
                                                                  'where me.sContrato = tm.sContrato and '+
                                                                        'me.sorden = :Contrato and '+
                                                                        'me.dIdFecha = :Fecha and '+
                                                                        'me.sIdFase = "OPER" and '+
                                                                        'me.sClasificacion = tm.sIdTipoMovimiento and '+
                                                                        'tm.sTipo = me.sTipo and '+
                                                                        'tm.sTipo = "OPE" and '+
                                                                        'me.eNavegando = "No" ) ) as sFactorOP, '+
                                                                'sum(( select ifnull( sum( me.sFactor ), 0 ) '+
                                                                  'from movimientosdeembarcacion as me '+
                                                                  'where me.sContrato = tm.sContrato and '+
                                                                        'me.sorden = :Contrato and '+
                                                                        'me.dIdFecha = :Fecha and '+
                                                                        '( me.sIdFase = "OPER" or me.sTipo = "ESP" ) and '+
                                                                        'me.sClasificacion = tm.sIdTipoMovimiento and '+
                                                                        'tm.sTipo = me.sTipo and '+
                                                                        'me.eNavegando = "Si" )) as sFactorNAV, '+
                                                                'sum(( select ifnull( sum( me.sFactor ), 0 ) '+
                                                                  'from movimientosdeembarcacion as me '+
                                                                  'where me.sContrato = tm.sContrato and '+
                                                                  'me.sorden = :Contrato and '+
                                                                  'me.dIdFecha = :Fecha and '+
                                                                  'me.sIdFase = "OPER" and '+
                                                                  'me.sTipo = "ESP" and '+
                                                                  'tm.sTipo = "ESP" and '+
                                                                  'me.eNavegando <> "Si" )) as sFactorESP, '+
                                                                'sum(( select ifnull( sum( me.sFactor ), 0 ) '+
                                                                  'from movimientosdeembarcacion as me '+
                                                                  'where me.sContrato = tm.sContrato and '+
                                                                  'me.sorden = :Contrato and '+
                                                                  'me.dIdFecha = :Fecha and '+
                                                                  'me.sIdFase = "OPER" and '+
                                                                  'me.sTipo = "CIA" )) as sFactorCIA '+
                                                              'from tiposdemovimiento as tm '+
                                                              'where '+
                                                                'tm.sContrato = :ContratoBarco and '+
                                                                'tm.sClasificacion = "Movimiento de Barco"; ';
                                connection.zCommand.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                                connection.zCommand.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
                                connection.zCommand.ParamByName('Contrato').AsString      := QryOrdenamiento.FieldByName('sOrden').AsString;
                                connection.zCommand.Open;

                                if connection.zCommand.RecordCount > 0 then
                                begin
                                    dConsumoCDP := connection.zCommand.FieldByName('sFactorOP').AsFloat;
                                    dConsumoNAV := connection.zCommand.FieldByName('sFactorNAV').AsFloat;
                                    dConsumoCIA := connection.zCommand.FieldByName('sFactorCIA').AsFloat;
                                    dConsumoESP := connection.zCommand.FieldByName('sFactorESP').AsFloat;
                                end;

                                rxBalanceResumen.FieldValues['dOperacion']   := rxBalanceResumen.FieldValues['dOperacion'] + dConsumoCDP * rxBalance.FieldValues['dConsumoOT'];
                                rxBalanceResumen.FieldValues['dEspera']      := rxBalanceResumen.FieldValues['dEspera']    + dConsumoESP * rxBalance.FieldValues['dConsumoOT'];
                                rxBalanceResumen.FieldValues['dNavegando']   := rxBalanceResumen.FieldValues['dNavegando'] + dConsumoNAV * rxBalance.FieldValues['dConsumoOT'];
                                rxBalanceResumen.FieldValues['dCia']         := rxBalanceResumen.FieldValues['dCia']       + dConsumoCIA * rxBalance.FieldValues['dConsumoOT'];
                                rxBalanceResumen.FieldValues['dTotal']       := rxBalanceResumen.FieldValues['dTotal']     + (dConsumoCDP + dConsumoCIA + dConsumoESP + dConsumoNAV) * rxBalance.FieldValues['dConsumoOT'];
                                rxBalanceResumen.Post;
                            end;

                            {$ENDREGION}
                            QryConsumos.Next;
                        end;
                    end;
                    {$ENDREGION}

                    QryConsumoEmbarcaciones.Next;
                end;
                QryOrdenamiento.Next;
            end;
        end;
        {$ENDREGION}

        {$REGION 'INSERCION DE DATOS'}
        QryConsumosDiesel.SQL.Clear;
        QryConsumosDiesel.SQL.Add('SELECT '+
                                  '  eq.sIdEquipo, '+
                                  '	eq.sDescripcion, '+
                                  '	ce.sNumerosDeSerie, '+
                                  '	re.sDescripcion AS sConsumible, '+
                                  '	eq.sMedida, '+
                                  '	ce.dCantidad, '+
                                  ' ce.sNumeroOrden, '+
                                  ' o.sIdFolio, '+
                                  ' ce.sIdPlataforma,  '+
                                  ' (select sNotaGeneral from notas_generales where sContrato =:Orden and dIdFecha =:Fecha and lAplicaConsumos = "si") as NotaGeneral, '+
                                  ' (select count(sNumeroOrden) from consumosdecombustibleporequipo where sContrato =:Orden and dIdFecha =:Fecha and sNumeroOrden = "vacio") as TotalVacios '+
                                  'FROM equipos  eq '+
                                  '	inner JOIN consumosdecombustibleporequipo AS ce '+
                                  '		ON (eq.sIdEquipo = ce.sIdEquipo AND ce.dIdFecha = :Fecha AND ce.sContrato =:Orden ) '+
                                  '	left JOIN recursosdeexistencias AS re '+
                                  '		ON (re.iIdRecursoExistencia = ce.iIdTipoConsumo) '+
                                  ' left join ordenesdetrabajo o on (o.sContrato = ce.sContrato and o.sNumeroOrden = ce.sNumeroOrden) '+
                                  'WHERE '+
                                  '     eq.sContrato = :ContratoBarco and lAplicadiesel = "Si" '+
                                  ' order by ce.sNumeroOrden, ce.iOrden');
        QryConsumosDiesel.ParamByName('Fecha').AsDate      := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        QryConsumosDiesel.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
        QryConsumosDiesel.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
        QryConsumosDiesel.Open;
        {$ENDREGION}

        {Esta secccion es para mostrar la hoja vacia sino existen datos.}
        if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
        begin
            if rxBalance.RecordCount = 0 then
            begin
                rxBalance.Append;
                rxBalance.FieldValues['sIdEmbarcacion'] := '';
                rxBalance.Post;
            end;
        end;
    end;

    Td_balance.DataSet:= rxBalance;
    Td_balance.FieldAliases.Clear;
    Td_ImpBalance.DataSet:= rxBalance;
    Td_ImpBalance.FieldAliases.Clear;
    Td_balanceResumen.DataSet:= rxBalanceResumen;
    Td_balanceResumen.FieldAliases.Clear;
    Td_consumoEquipos.DataSet    := QryConsumosDiesel;
    Td_consumoEquipos.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_balance);
    Reporte.DataSets.Add(Td_ImpBalance);
    Reporte.DataSets.Add(Td_balanceResumen);
    Reporte.DataSets.Add(Td_consumoEquipos);

  Finally

  End;

end;

Procedure ReportePDF_ActividadesPorFolio(sFiltroFolio:string; ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
Var
  QryInfoFolio,
  QryPartidas,
  QryDetallePartidas,
  QryConsultaAvanceAnterior,
  QryConsultaAvancesAcumulados,
  QryPersonal,
  QryEquipo,
  QryMateriales,
  QryFolios,
  QryPU,
  QryAnexos,
  QryTiemposAct     : TZQuery;

  Td_partidas,
  Td_ImpPartidas,
  Td_ClasificacionTiempo : TfrxDBDataset;
  dDiaSiguiente  : TDateTime;
  rxPartidas: TClientDataSet;
  rxTiemposAct: TClientDataSet;
  iTotalPersonal, indice, iTotalEsp, num : integer;
  lContinua      : boolean;
  QrAux:TzReadOnlyQuery;
  sIidDiario,
  siIdTarea,
  siIdActividad:String;
begin
    QrAux:=TzReadOnlyQuery.create(nil);
    Try
      {Creacion de objetos de consulta}
      QryConsultaAvancesAcumulados := TZQuery.Create(nil);
      QryConsultaAvancesAcumulados.Connection := Connection.zConnection;
      QryConsultaAvanceAnterior := TZQuery.Create(nil);
      QryConsultaAvanceAnterior.Connection := connection.zConnection;
      QryFolios := TZQuery.Create(nil);
      QryFolios.Connection := Connection.zConnection;
      QryPartidas := TZQuery.Create(nil);
      QryPartidas.Connection := Connection.zConnection;
      QryDetallePartidas := TZQuery.Create(nil);
      QryDetallePartidas.Connection := Connection.zConnection;
      QryInfoFolio := TZQuery.Create(nil);
      QryInfoFolio.Connection := Connection.zConnection;
      QryPersonal := TZQuery.Create(nil);
      QryPersonal.Connection := Connection.zConnection;
      QryEquipo := TZQuery.Create(nil);
      QryEquipo.Connection := Connection.zConnection;
      QryMateriales := TZQuery.Create(nil);
      QryMateriales.Connection := Connection.zConnection;
      QryPU := TZQuery.Create(nil);
      QryPU.Connection := Connection.zConnection;
      QryAnexos := TZQuery.Create(nil);
      QryAnexos.Connection := Connection.zConnection;
      QryTiemposAct := TZQuery.Create(nil);
      QryTiemposAct.Connection := Connection.zConnection;

      Td_partidas:= TfrxDBDataset.Create(nil);
      Td_partidas.UserName:='td_partidas';

      Td_ImpPartidas:= TfrxDBDataset.Create(nil);
      Td_ImpPartidas.UserName:='Td_ImpPartidas';

      QrAux.connection:=Connection.zconnection;
      rxPartidas := TClientDataSet.Create(nil);
      rxPartidas.Active := False;
      rxPartidas.FieldDefs.Add('sTipo', ftString, 25, False);
      rxPartidas.FieldDefs.Add('sOT', ftString, 15, False);
      rxPartidas.FieldDefs.Add('sFolio', ftString, 50, False);
      rxPartidas.FieldDefs.Add('sPlataforma', ftString, 50, False);
      rxPartidas.FieldDefs.Add('mDescripcionFolio', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sNumeroActividad', ftString, 20, False);
      rxPartidas.FieldDefs.Add('mDescripcion', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('iIdDiario', ftInteger, 0, False);
      rxPartidas.FieldDefs.Add('sClasificacion', ftString, 10, False);
      rxPartidas.FieldDefs.Add('sHoraInicio', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sHoraFin', ftString, 5, False);
      rxPartidas.FieldDefs.Add('mDescripcionHorario', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('sHoraInicioPer', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sHoraFinPer', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sPartidaPer', ftString, 20, False);
      rxPartidas.FieldDefs.Add('dCantidadPer', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dTotalPer', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sDescripcionPer', ftString, 250, False);
      rxPartidas.FieldDefs.Add('sHoraInicioEq', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sHoraFinEq', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sPartidaEq', ftString, 10, False);
      rxPartidas.FieldDefs.Add('dCantidadEq', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dTotalEq', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sDescripcionEq', ftString, 350, False);
      rxPartidas.FieldDefs.Add('sMaterial', ftString, 25, False);
      rxPartidas.FieldDefs.Add('dCantidadMat', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dCantidadMatAnt', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sMedidaMat', ftString, 8, False);
      rxPartidas.FieldDefs.Add('sSuministra', ftString, 25, False);
      rxPartidas.FieldDefs.Add('mDescripcionMaterial', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('sNumeroActividadPU', ftString, 20, False);
      rxPartidas.FieldDefs.Add('dCantidadPU', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sMedidaPU', ftString, 10, False);
      rxPartidas.FieldDefs.Add('sAnexoPU', ftString, 25, False);
      rxPartidas.FieldDefs.Add('mDescripcionPU', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dRestaEspacio', ftFloat, 0, False);
      rxPartidas.CreateDataSet;
      rxPartidas.Active := True;

      if (TImpresion=FtsActividades) or (TImpresion=FtsAll) then
      begin
          {$REGION 'FOLIOS REPORTADOS'}
          QryFolios.SQL.Clear;
          QryFolios.SQL.Add('SELECT ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.mDescripcion, p.sDescripcion as plataforma, cv.sIdConvenio as Convenio, cv.iGerencialInicio, cv.iGerencialFinal, ot.lAplicaJornada '+
                            'FROM ordenesdetrabajo AS ot ' +
                            ' inner join convenios cv on (cv.sContrato = ot.sContrato and ot.sNumeroOrden = cv.sNumeroOrden) '+
                            '	INNER JOIN contratos AS c ' +
                            '		ON (ot.sContrato=c.sContrato) ' +
                            '	INNER JOIN bitacoradeactividades AS ba ' +
                            '		ON (ba.sContrato=c.sContrato AND ba.sNumeroOrden=ot.sNumeroOrden) ' +
                            '	LEFT JOIN tiposdemovimiento AS tm ' +
                            '		ON (tm.sContrato= :ContratoBarco AND tm.sIdTipoMovimiento=ba.sIdTipoMovimiento AND tm.sClasificacion="Tarifa Diaria") ' +
                            'inner join plataformas  p on (p.sIdPlataforma = ot.sIdPlataforma) '+
                            'WHERE c.sContrato= :Orden AND ba.dIdFecha= :Fecha and ot.sNumeroOrden like :Folio ' +
                            'GROUP BY cv.sNumeroorden  order by ot.iOrden ');
          QryFolios.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
          QryFolios.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          QryFolios.ParamByName('Folio').AsString         := sFiltroFolio;
          QryFolios.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
          QryFolios.Open;

          while Not QryFolios.Eof do
          begin
              {$REGION 'AVANCES DEL FOLIO'}
              {Avances anteriores}
              QryConsultaAvanceAnterior.SQL.Clear;
              QryConsultaAvanceAnterior.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', ReporteDiario.FieldByName('dIdFecha').AsDateTime)+'", :Orden, :Folio), 2) AS dAvanceAnterior;';
              QryConsultaAvanceAnterior.ParamByName('Orden').AsString := ReporteDiario.FieldByName('sOrden').AsString;
              QryConsultaAvanceAnterior.ParamByName('Folio').AsString := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvanceAnterior.Open;

              //Avances Acumulados
              dDiaSiguiente := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              dDiaSiguiente := IncDay(dDiaSiguiente);

              QryConsultaAvancesAcumulados.SQL.Clear;
              QryConsultaAvancesAcumulados.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 2) AS dAvanceAcumulado;';
              QryConsultaAvancesAcumulados.ParamByName('Orden').AsString := ReporteDiario.FieldByName('sOrden').AsString;
              QryConsultaAvancesAcumulados.ParamByName('Folio').AsString := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvancesAcumulados.Open;
              {$ENDREGION}

              {$REGION 'CONSULTA - PARTIDAS REPORTADAS'}
              QryPartidas.SQL.Clear;
              QryPartidas.SQL.Add('SELECT  a.sContrato, a.iIdDiario,	a.sNumeroActividad, o.dPonderado, a.sWbs, '+
                                  ' (SELECT   SUM(dAvance) ' +
                                  ' FROM bitacoradeactividades ' +
                                  ' WHERE ' +
                                  '   dIdFecha < a.dIdFecha ' +
                                  '   AND sIdTipoMovimiento = "E" ' +
                                  '   AND sWbs = a.sWbs ' +
                                  '   AND sContrato = a.sContrato ' +
                                  '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceAnterior_GeneralPartida, '+

                                  ' (SELECT  SUM(dAvance) ' +
                                  ' FROM bitacoradeactividades ' +
                                  ' WHERE ' +
                                  '   dIdFecha = a.dIdFecha ' +
                                  '   AND sIdTipoMovimiento = "E" ' +
                                  '   AND sWbs = a.sWbs ' +
                                  '   AND sContrato = a.sContrato ' +
                                  '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceActual_GeneralPartida, ' +

                                  '	CONCAT(a.sHoraInicio,"-",a.sHoraFinal)  AS Horario, ' +
                                  '	a.sIdClasificacion,	a.sHoraInicio, a.sHorafinal, ' +
                                  '	o.mdescripcion as sDescripcionActividad ' +
                                  'FROM bitacoradeactividades a ' +
                                  'INNER JOIN actividadesxorden o ' +
                                  '		ON (o.sContrato = a.sContrato AND o.sIdConvenio = a.sIdConvenio AND o.sNumeroOrden = a.sNumeroOrden AND a.sWbs = o.sWbs AND o.sNumeroActividad = a.sNumeroActividad AND o.sTipoActividad = "Actividad" and o.sTipoAnexo = "ADM") ');

              if ReporteDiario.FieldByName('lOrdenaxHorario').AsString = 'No' then
              begin
                 QryPartidas.SQL.Add( 'WHERE a.sContrato = :Orden AND a.dIdFecha = :Fecha AND a.sNumeroOrden = :Folio AND a.sIdTurno = :Turno AND a.sIdTipoMovimiento = "E" and a.lImprime ="Si" ' +
                                 'GROUP BY a.sNumeroActividad '+
                                 'ORDER BY a.sContrato, o.iItemOrden, a.sNumeroActividad, a.sHoraInicio ');
              end
              else
              begin
                 QryPartidas.SQL.Add( 'WHERE a.sContrato = :Orden AND a.dIdFecha = :Fecha AND a.sNumeroOrden = :Folio AND a.sIdTurno = :Turno AND a.sIdTipoMovimiento <> "E" and a.lImprime ="Si" ' +
                                 'GROUP BY a.sNumeroActividad '+
                                 'ORDER BY a.sContrato, a.sHoraInicio, a.sNumeroActividad ');
              end;
              QryPartidas.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sOrden').AsString;
              QryPartidas.ParamByName('Fecha').AsDateTime  := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryPartidas.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryPartidas.ParamByName('Turno').AsString    := ReporteDiario.FieldByName('sIdTurno').AsString;
              QryPartidas.Open;
              {$ENDREGION}

              while Not QryPartidas.Eof do
              begin
                  {Insertar Datos en RxPartidas}
                  rxPartidas.Append;
                  rxPartidas.FieldValues['iIdDiario'] := 0;
                  rxPartidas.FieldValues['sTipo']       := 'General';
                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                  rxPartidas.FieldValues['mDescripcionFolio'] := QryFolios.FieldValues['mDescripcion'];
                  rxPartidas.FieldValues['dAvanceAntFolio']   := QryConsultaAvanceAnterior.FieldByName('dAvanceAnterior').AsFloat;
                  rxPartidas.FieldValues['dAvanceActFolio']   := QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').AsFloat - QryConsultaAvanceAnterior.FieldByName('dAvanceAnterior').AsFloat;
                  rxPartidas.FieldValues['dAvanceAcumFolio']  := QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').AsFloat;
                  rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                  rxPartidas.FieldValues['mDescripcion']     := QryPartidas.FieldByName('sDescripcionActividad').AsString;
                  rxPartidas.FieldValues['dAvanceAntPart']   := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat;
                  rxPartidas.FieldValues['dAvanceActPart']   := QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                  rxPartidas.FieldValues['dAvanceAcumPart']  := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat + QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                  rxPartidas.Post;

                  {$REGION 'CONSULTA - DETALLE DE PARTIDAS'}
                  QryDetallePartidas.SQL.Clear;
                  QryDetallePartidas.SQL.Add( 'SELECT b.dIdFecha, b.sIdClasificacion, b.sHoraInicio, b.sHoraFinal, b.iIdDiario, b.sNumeroActividad, b.iIdTarea,b.iIdActividad,b.iHermano, '+
                                              '	( SELECT (ifnull(sum(ba.dAvance), 0)) ' +
                                              '		FROM ' +
                                              '			bitacoradeactividades AS ba ' +
                                              '		WHERE ' +
                                              '			ba.sContrato = b.sContrato ' +
                                              '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                              '		AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento ' +
                                              '		AND ba.swbs = b.swbs ' +
                                              '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                              '		AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
                                              '   < cast(b.sHoraInicio AS Time)) ' +
                                              '   )	) AS dAvanceAnterior, ' +
                                              '	b.dAvance, b.mDescripcion AS descAct, b.sTipoObra, b.dCantidadAjuste, b.dCantidadAjusteNext, b.dCantidadAjusteNext2, b.dRestaEspacio ' +
                                              'FROM ' +
                                              '	bitacoradeactividades b ' +
                                              'INNER JOIN tiposdemovimiento tm ON ( ' +
                                              '	tm.sContrato = :ContratoBarco ' +
                                              '	AND tm.sIdTipoMovimiento = b.sIdClasificacion ) ' +
                                              'INNER JOIN actividadesxorden ao ON ( ' +
                                              '	ao.sContrato = b.sContrato and ao.sIdConvenio =b.sIdConvenio and ao.sNumeroOrden = b.sNumeroOrden and ' +
                                              '	ao.sWbs = b.sWbs and ao.sTipoActividad = "Actividad" ) ' +
                                              ' inner join tipos_actividades ta on(ta.sIdTipo = b.sTipoObra) '+
                                              'WHERE	b.sContrato = :Orden ' +
                                              'AND b.sNumeroOrden = :Folio ' +
                                              'AND b.sIdTipoMovimiento = "ED" ' +
                                              'AND b.sNumeroActividad = :Actividad ' +
                                              'AND dIdFecha = :Fecha ' +
                                              'and ((b.iHermano=-1) or (b.iHermano=b.iIdActividad)) AND lImprime = "Si" '+
                                              ' group by b.sNumeroActividad, b.sHoraInicio, ta.sAgrupaRd '+
                                              ' ORDER BY b.didFecha, ao.iItemOrden, Time(b.sHoraInicio), ta.iOrden');
                  QryDetallePartidas.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
                  QryDetallePartidas.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
                  QryDetallePartidas.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                  QryDetallePartidas.ParamByName('Folio').AsString         := QryFolios.FieldByName('sNumeroOrden').AsString;
                  QryDetallePartidas.ParamByName('Actividad').AsString     := QryPartidas.FieldByName('sNumeroActividad').AsString;
                  QryDetallePartidas.Open;
                  {$ENDREGION}

                  QryPersonal.Active := False;
                  QryEquipo.Active   := False;

                  {$REGION 'INSERCIÓN DE DATOS'}
                  while Not QryDetallePartidas.Eof do
                  begin

                      if QryDetallePartidas.FieldByName('dCantidadAjuste').AsFloat > 0 then
                      begin
                          iTotalEsp :=  QryDetallePartidas.FieldByName('dCantidadAjuste').AsInteger;
                          for num := 1 to iTotalEsp do
                          begin
                              rxPartidas.Append;
                              rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                              rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                              rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                              rxPartidas.Post;
                          end;
                      end;

                      if QryDetallePartidas.FieldByName('dCantidadAjuste').AsFloat < 0 then
                      begin
                          iTotalEsp :=  QryDetallePartidas.FieldByName('dCantidadAjuste').AsInteger;
                          for num := iTotalEsp + 1 to 0  do
                          begin
                              rxPartidas.Delete;
                          end;
                      end;

                      if QryPersonal.Active then
                      begin
                          if (QryPersonal.RecordCount > 0) and (QryEquipo.RecordCount > 0) then
                          begin
                              if QryPersonal.RecordCount < QryEquipo.RecordCount then
                              begin
                                  rxPartidas.Append;
                                  rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                                  rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                                  rxPartidas.Post;
                              end;
                          end;
                      end;

                      rxPartidas.Append;
                      rxPartidas.FieldValues['sTipo']       := 'TituloAvances';
                      rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                      rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                      rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                      rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                      rxPartidas.Post;

                      if QryPersonal.Active then
                      begin
                          if (QryPersonal.RecordCount > 0) and (QryEquipo.RecordCount > 0) then
                          begin
                              if QryPersonal.RecordCount >= QryEquipo.RecordCount then
                              begin
                                  rxPartidas.Append;
                                  rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                                  rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                                  rxPartidas.Post;
                              end;
                          end;
                      end;

                      if rxPartidas.FieldValues['sTipo'] <> 'TituloHorario' then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloHorario';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString ;
                          rxPartidas.FieldValues['sHoraInicio']  := QryDetallePartidas.FieldByName('sHoraInicio').AsString;
                          rxPartidas.FieldValues['mDescripcionHorario']  := QryDetallePartidas.FieldByName('descAct').AsString;//GHH New
                          rxPartidas.FieldValues['dRestaEspacio']    := QryDetallePartidas.FieldByName('dRestaEspacio').AsFloat;
                          rxPartidas.Post;
                      end;

                      if QryDetallePartidas.FieldByName('iHermano').AsInteger=-1 then
                      begin
                        rxPartidas.Append;
                        rxPartidas.FieldValues['sTipo']       := 'DatosHorario';
                        rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                        rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                        rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                        rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        rxPartidas.FieldValues['sClasificacion']   := QryDetallePartidas.FieldByName('sIdClasificacion').AsString ;
                        rxPartidas.FieldValues['sHoraInicio']  := QryDetallePartidas.FieldByName('sHoraInicio').AsString;
                        rxPartidas.FieldValues['sHoraFin']     := QryDetallePartidas.FieldByName('sHoraFinal').AsString;
                        rxPartidas.FieldValues['dAvanceAntPart']   := QryDetallePartidas.FieldByName('dAvanceAnterior').AsFloat;
                        rxPartidas.FieldValues['dAvanceActPart']   := QryDetallePartidas.FieldByName('dAvance').AsFloat;
                        rxPartidas.FieldValues['dAvanceAcumPart']  := QryDetallePartidas.FieldByName('dAvanceAnterior').AsFloat + QryDetallePartidas.FieldByName('dAvance').AsFloat;
                        rxPartidas.FieldValues['dRestaEspacio']    := QryDetallePartidas.FieldByName('dRestaEspacio').AsFloat;
                        rxPartidas.Post;
                      end
                      else
                      begin
                        QrAux.active:=False;
                        QrAux.sql.text:='select * from bitacoradeactividades where scontrato=:Orden and '+
                                        'didfecha=:fecha and sNumeroOrden=:Folio and sIdTipoMovimiento = "ED" and ' +
                                        'sNumeroActividad = :Actividad and iHermano=:IdVinculo '+
                                        'group by iIdTarea,iIdActividad ' +
                                        'ORDER BY Time(sHoraInicio) ';
                        QrAux.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
                        QrAux.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                        QrAux.ParamByName('Folio').AsString         := QryFolios.FieldByName('sNumeroOrden').AsString;
                        QrAux.ParamByName('Actividad').AsString     := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        QrAux.ParamByName('IdVinculo').AsInteger     := QryDetallePartidas.FieldByName('iHermano').AsInteger;
                        QrAux.open;

                        rxPartidas.Append;
                        rxPartidas.FieldValues['sTipo']       := 'DatosHorario';
                        rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                        rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                        rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                        rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        rxPartidas.FieldValues['sClasificacion']   := QryDetallePartidas.FieldByName('sIdClasificacion').AsString ;
                        rxPartidas.FieldValues['sHoraInicio']  := QrAux.FieldByName('sHoraInicio').AsString;
                        rxPartidas.FieldValues['mDescripcionHorario']  := QryDetallePartidas.FieldByName('descAct').AsString;
                        rxPartidas.FieldValues['dAvanceAntPart']   := QryDetallePartidas.FieldByName('dAvanceAnterior').AsFloat;
                        rxPartidas.FieldValues['dAvanceActPart']   := QryDetallePartidas.FieldByName('dAvance').AsFloat;
                        rxPartidas.FieldValues['dAvanceAcumPart']  := QryDetallePartidas.FieldByName('dAvanceAnterior').AsFloat + QryDetallePartidas.FieldByName('dAvance').AsFloat;
                        rxPartidas.FieldValues['dRestaEspacio']    := QryDetallePartidas.FieldByName('dRestaEspacio').AsFloat;
                        sIidDiario:='';
                        siIdTarea:='';
                        siIdActividad:='';
                        while not QrAux.eof do
                        begin
                          if sIidDiario='' then
                            sIidDiario:=QrAux.FieldByName('iIdDiario').AsString
                          else
                            sIidDiario:= sIidDiario + ',' + QrAux.FieldByName('iIdDiario').AsString;

                          if siIdTarea='' then
                            siIdTarea:=QrAux.FieldByName('iIdTarea').AsString
                          else
                            siIdTarea:= siIdTarea + ',' + QrAux.FieldByName('iIdTarea').AsString;

                          if siIdActividad='' then
                            siIdActividad:=QrAux.FieldByName('iIdActividad').AsString
                          else
                            siIdActividad:= siIdActividad + ',' + QrAux.FieldByName('iIdActividad').AsString;

                          QrAux.Next;
                        end;
                        rxPartidas.FieldValues['sHoraFin']     := QrAux.FieldByName('sHoraFinal').AsString;
                        rxPartidas.Post;

                      end;

                      {$REGION 'MATERIAL X PARTIDA'}
                      {Ahora insertamos el material..}
                      QryMateriales.SQL.Clear;
                      QryMateriales.SQL.Add('select m.sIdMaterial, sum(m.dCantidad) as dCantidad, aa.sMedida, aa.mDescripcion, '+
                                  '(select sum(m1.dCantidad) as cantidad from bitacorademateriales m1 '+
                                  'where m1.sContrato = m.sContrato and m1.sNumeroOrden = m.sNumeroOrden and m1.dIdFecha < m.dIdFecha and m1.sIdMaterial = m.sIdMaterial) as dCantidadAntDia, '+
                                  '(select sum(m1.dCantidad) as cantidad from bitacorademateriales m1 '+
                                  'where m1.sContrato = m.sContrato and m1.sNumeroOrden = m.sNumeroOrden and m1.dIdFecha = m.dIdFecha and m1.sIdMaterial = m.sIdMaterial and m1.iIddiario < m.iIdDiario ) as dCantidadAnt '+
                                  'from bitacorademateriales m '+
                                  'inner join actividadesxanexo aa on (aa.scontrato =:Contrato and aa.sNumeroActividad =m.sIdMaterial and aa.sTipoActividad = "Actividad") '+
                                  'where m.sContrato =:Orden and m.dIdFecha =:Fecha and m.iIdDiario = :Id '+
                                  'group by m.sIdMaterial order by aa.sWbs');
                      QryMateriales.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryMateriales.ParamByName('Orden').AsString     := ReporteDiario.FieldByName('sOrden').AsString;
                      QryMateriales.ParamByName('Id').AsInteger       := QryDetallePartidas.FieldByName('iIdDiario').AsInteger;
                      QryMateriales.ParamByName('Contrato').AsString  := global_Contrato_Barco;
                      QryMateriales.Open;

                      if QryMateriales.RecordCount > 0 then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;

                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloMaterial';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;
                      end;

                      {M A T E R I A L...}
                      while not QryMateriales.Eof do
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'DatosMaterial';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['sMaterial']    := QryMateriales.FieldByName('sIdMaterial').AsString;
                          rxPartidas.FieldValues['dCantidadMat'] := QryMateriales.FieldByName('dCantidad').AsFloat;
                          rxPartidas.FieldValues['sMedidaMat']   := QryMateriales.FieldByName('sMedida').AsString;
                          rxPartidas.FieldValues['dCantidadMatAnt']:= (QryMateriales.FieldByName('dCantidadAnt').AsFloat + QryMateriales.FieldByName('dCantidadAntDia').AsFloat);
                          rxPartidas.FieldValues['mDescripcionMaterial']  := QryMateriales.FieldValues['mDescripcion'];
                          rxPartidas.Post;

                          QryMateriales.Next;
                      end;
                      {$ENDREGION}

                      {$REGION 'PERSONAL X HORARIOS'}
                      {Ahora insertamos el personal..}
                      QryPersonal.SQL.Clear;
                      if QryDetallePartidas.FieldByName('iHermano').AsInteger=-1 then
                      begin
                        // Cambio realizado por E. Rangel 20/08/2016
                        {QryPersonal.SQL.Add('select ' +
                                            '  bp.*, ' +
                                            '  sum(bp.dCantidad) as Cantidad, ' +
                                            '  (IFNULL(sum(CAST(bp.dCantHH AS CHAR)), 0) + ' +
                                            '   IFNULL((SELECT ' +
                                            '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                            '           FROM ' +
                                            '             bitacoradeajustes ba ' +
                                            '               WHERE ba.sContrato = bp.sContrato AND ' +
                                            '                     ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                            '                     ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                            '                     ba.dFecha = bp.dIdFecha), 0)) as dCantHHG ' +
                                            'from ' +
                                            '  bitacoradepersonal bp ' +
                                            'inner join ' +
                                            '  tipos_actividades ta ' +
                                            '    on (ta.sIdTipo = bp.sTipoObra ) ' +
                                            'where ' +
                                            '  bp.sContrato =:Orden and ' +
                                            '  bp.dIdFecha =:fecha and ' +
                                            '  bp.sNumeroOrden =:Folio and ' +
                                            '  bp.sNumeroActividad =:Actividad and ' +
                                            '  bp.sHoraInicio =:sInicio and ' +
                                            '  bp.sTipoObra =:sTipo and ' +
                                            '  bp.lImprime = "Si" and bp.dCantHH > 0 ' +
                                            'group by ' +
                                            '  ta.sAgrupaRd, ' +
                                            '  bp.sHoraInicioG, ' +
                                            '  bp.sIdPersonal ' +
                                            'order by ' +
                                            '  bp.iItemOrden');}
                        QryPersonal.SQL.Add('select bp.*, sum( bp.dCantidad ) as Cantidad, ' +
                                            '(sum(bp.dCantHH) + ' +
                                            ' IFNULL((SELECT ' +
                                            '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                            '          FROM ' +
                                            '            bitacoradeajustes ba ' +
                                            '          WHERE ' +
                                            '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                            '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                            '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantHHG from bitacoradepersonal bp '+
                                            ' inner join tipos_actividades ta on (ta.sIdTipo = bp.sTipoObra ) '+
                                            'where bp.sContrato =:Orden and bp.dIdFecha =:fecha and bp.sNumeroOrden =:Folio and bp.sNumeroActividad =:Actividad and bp.sHoraInicio =:sInicio and bp.sTipoObra =:sTipo and bp.lImprime = "Si" and bp.dCantHH > 0 '+
                                            'group by ta.sAgrupaRd, bp.sHoraInicioG, bp.sIdPersonal order by bp.iItemOrden ');
                        QryPersonal.ParamByName('sInicio').AsString  := QryDetallePartidas.FieldValues['sHoraInicio'];
                        QryPersonal.ParamByName('sTipo').AsString    := QryDetallePartidas.FieldValues['sTipoObra'];

                      end
                      else
                      begin
                        // Modificación por E. Rangel - 20/08/2016
                        {QryPersonal.SQL.Add('select ' +
                                            '  *, ' +
                                            '  sum( dCantidad ) as Cantidad, ' +
                                            '  (IFNULL(sum(CAST(dCantHH AS CHAR)), 0) + ' +
                                            '   IFNULL((SELECT ' +
                                            '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                            '          FROM ' +
                                            '            bitacoradeajustes ba ' +
                                            '          WHERE ' +
                                            '            ba.sContrato = bp.sContrato AND ' +
                                            '            ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                            '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                            '            ba.dFecha = bp.dIdFecha), 0)) as dCantHHG ' +
                                            'from ' +
                                            '  bitacoradepersonal ' +
                                            'where ' +
                                            '  sContrato =:Orden and ' +
                                            '  dIdFecha =:fecha and ' +
                                            '  sNumeroActividad =:Actividad and ' +
                                            '  iIdTarea in ('+siIdTarea+') and ' +
                                            '  iIdActividad in ('+siIdActividad+') and ' +
                                            '  lImprime = "Si" ' +
                                            'group by ' +
                                            '  sHoraInicioG, ' +
                                            '  sHoraFinalG, ' +
                                            '  sIdPersonal ' +
                                            'order by ' +
                                            '  iItemOrden, ' +
                                            '  time(sHoraInicio)');}
                        QryPersonal.SQL.Add('select *, sum(bp.dCantidad ) as Cantidad, (sum(bp.dCantHH) + ' +
                                            ' IFNULL((SELECT ' +
                                            '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                            '         FROM ' +
                                            '           bitacoradeajustes ba ' +
                                            '         WHERE ' +
                                            '           ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                            '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                            '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantHHG ' +
                                            'from bitacoradepersonal bp where bp.sContrato =:Orden '+
                                            'and bp.dIdFecha =:fecha and bp.sNumeroActividad =:Actividad and '+
                                            'bp.iIdTarea in ('+siIdTarea+') and bp.iIdActividad in ('+siIdActividad+') and bp.lImprime = "Si" '+
                                            'group by bp.sHoraInicioG, bp.sHoraFinalG, bp.sIdPersonal order by bp.iItemOrden,time(bp.sHoraInicio)');

                      end;
                      QryPersonal.ParamByName('Folio').AsString     := QryFolios.FieldByName('sNumeroOrden').AsString;
                      QryPersonal.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryPersonal.ParamByName('Orden').AsString     := ReporteDiario.FieldByName('sOrden').AsString;
                      QryPersonal.ParamByName('Actividad').AsString := QryDetallePartidas.FieldByName('sNumeroActividad').AsString;

                      QryPersonal.Open;

                      rxPartidas.Append;
                      rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                      rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                      rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                      rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                      rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString ;
                      rxPartidas.Post;

                      if QryDetallePartidas.FieldByName('dCantidadAjusteNext').AsFloat > 0 then
                      begin
                          iTotalEsp :=  QryDetallePartidas.FieldByName('dCantidadAjusteNext').AsInteger;
                          for num := 1 to iTotalEsp do
                          begin
                              rxPartidas.Append;
                              rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                              rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                              rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                              rxPartidas.Post;
                          end;
                      end;

                      if QryDetallePartidas.FieldByName('dCantidadAjusteNext').AsFloat < 0 then
                      begin
                          iTotalEsp :=  QryDetallePartidas.FieldByName('dCantidadAjusteNext').AsInteger;
                          for num := iTotalEsp + 1 to 0  do
                          begin
                              rxPartidas.Delete;
                          end;
                      end;


                      if QryPersonal.RecordCount > 0 then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;

                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloDatosPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString ;
                          rxPartidas.Post;
                      end;

                      {P E R S O N A L...}
                      while not QryPersonal.Eof do
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'DatosPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['sHoraInicioPer']   := QryPersonal.FieldByName('sHoraInicioG').AsString;
                          rxPartidas.FieldValues['sHoraFinPer']      := QryPersonal.FieldByName('sHoraFinalG').AsString;
                          rxPartidas.FieldValues['sPartidaPer']      := QryPersonal.FieldByName('sIdPersonal').AsString;
                          rxPartidas.FieldValues['dCantidadPer']     := QryPersonal.FieldByName('Cantidad').AsFloat;
                          if QryFolios.FieldByName('lAplicaJornada').AsString = 'Si' then
                             rxPartidas.FieldValues['dTotalPer']        := QryPersonal.FieldByName('dCantHHG').AsFloat
                          else
                             rxPartidas.FieldValues['dTotalPer']        := QryPersonal.FieldByName('dCantidad').AsFloat;
                          rxPartidas.FieldValues['sDescripcionPer']  := QryPersonal.FieldByName('sDescripcion').AsString;
                          rxPartidas.Post;
                          QryPersonal.Next;
                      end;
                      {$ENDREGION}

                      {$REGION 'EQUIPO X HORARIOS'}
                      {Ahora insertamos el equipo..}
                      QryEquipo.SQL.Clear;
                      if QryDetallePartidas.FieldByName('iHermano').AsInteger=-1 then
                      begin
                        QryEquipo.SQL.Add('select be.*, sum( be.dCantidad ) as Cantidad, (sum(be.dCantHH) + ' +
                                          ' IFNULL((SELECT ' +
                                          '          SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                          '        FROM ' +
                                          '          bitacoradeajustes ba ' +
                                          '        WHERE ' +
                                          '          ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                                          '          ba.sNumeroActividad = be.sNumeroActividad AND ' +
                                          '          ba.sIdPartidaAnexo = be.sIdEquipo AND ba.dFecha = be.dIdFecha), 0)) as dCantHHG from bitacoradeequipos be '+
                                          ' inner join tipos_actividades ta on (ta.sIdTipo = be.sTipoObra ) '+
                                          'where be.sContrato =:Orden and be.dIdFecha =:fecha and be.sNumeroOrden =:Folio and be.sNumeroActividad =:Actividad and be.sHoraInicio =:sInicio and be.sTipoObra =:sTipo and be.lImprime ="Si" and be.dCantHH > 0 '+
                                          'group by ta.sAgrupaRd, be.sHoraInicioG, be.sIdEquipo order by be.iItemOrden ');
                        QryEquipo.ParamByName('sInicio').AsString  := QryDetallePartidas.FieldValues['sHoraInicio'];
                        QryEquipo.ParamByName('sTipo').AsString    := QryDetallePartidas.FieldValues['sTipoObra'];
                      end
                      else
                      begin
                        QryEquipo.SQL.Add('select *, (sum(dCantHH) + ' +
                                          ' IFNULL((SELECT ' +
                                          '          SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                          '        FROM ' +
                                          '          bitacoradeajustes ba ' +
                                          '        WHERE ' +
                                          '          ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                                          '          ba.sNumeroActividad = be.sNumeroActividad AND ' +
                                          '          ba.sIdPartidaAnexo = be.sIdEquipo AND ba.dFecha = be.dIdFecha), 0)) as dCantHHG from bitacoradeequipos where sContrato =:Orden '+
                                          'and dIdFecha =:fecha and sNumeroActividad =:Actividad and '+
                                          'iIdTarea in ('+siIdTarea+') and iIdActividad in ('+siIdActividad+') and lImprime = "Si" '+
                                          'group by sHoraInicioG, sIdEquipo order by iItemOrden,time(sHoraInicio)');
                      end;
                      QryEquipo.ParamByName('Folio').AsString     := QryFolios.FieldByName('sNumeroOrden').AsString;
                      QryEquipo.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryEquipo.ParamByName('Orden').AsString     := ReporteDiario.FieldByName('sOrden').AsString;
                      QryEquipo.ParamByName('Actividad').AsString := QryDetallePartidas.FieldByName('sNumeroActividad').AsString;
                      QryEquipo.Open;

                      {Verificamos cuanto personal se inserto para retroceder el memory.}
                      iTotalPersonal := QryPersonal.RecordCount;

                      if (QryEquipo.RecordCount > 0) and (iTotalPersonal = 0) then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;

                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloDatosPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;
                      end;

                      for indice := 1 to iTotalPersonal - 1 do
                          rxPartidas.Prior;

                      indice := 1;
                      lContinua := True;
                      {E Q U I P O...}
                      while not QryEquipo.Eof do
                      begin
                          if iTotalPersonal > 0 then
                          begin
                              if indice <= iTotalPersonal then
                              begin
                                  rxPartidas.Edit;
                                  rxPartidas.FieldValues['sHoraInicioEq']   := QryEquipo.FieldByName('sHoraInicioG').AsString;
                                  rxPartidas.FieldValues['sHoraFinEq']      := QryEquipo.FieldByName('sHoraFinalG').AsString;
                                  rxPartidas.FieldValues['sPartidaEq']      := QryEquipo.FieldByName('sIdEquipo').AsString;
                                  rxPartidas.FieldValues['dCantidadEq']     := QryEquipo.FieldByName('dCantidad').AsFloat;
                                  if QryFolios.FieldByName('lAplicaJornada').AsString = 'Si' then
                                     rxPartidas.FieldValues['dTotalEq']        := QryEquipo.FieldByName('dCantHHG').AsFloat
                                  else
                                     rxPartidas.FieldValues['dTotalEq']        := QryEquipo.FieldByName('dCantidad').AsFloat;
                                  rxPartidas.FieldValues['sDescripcionEq']  := QryEquipo.FieldByName('sDescripcion').AsString;
                                  rxPartidas.Post;
                                  rxPartidas.Next;
                                  inc(indice);
                                  lContinua := False;
                              end
                              else
                                  lContinua := True;
                          end;

                          if lContinua then
                          begin
                              rxPartidas.Append;
                              rxPartidas.FieldValues['sTipo']       := 'DatosPersonal';
                              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                              rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                              rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                              rxPartidas.FieldValues['sHoraInicioEq']   := QryEquipo.FieldByName('sHoraInicioG').AsString;
                              rxPartidas.FieldValues['sHoraFinEq']      := QryEquipo.FieldByName('sHoraFinalG').AsString;
                              rxPartidas.FieldValues['sPartidaEq']      := QryEquipo.FieldByName('sIdEquipo').AsString;
                              rxPartidas.FieldValues['dCantidadEq']     := QryEquipo.FieldByName('dCantidad').AsFloat;
                              if QryFolios.FieldByName('lAplicaJornada').AsString = 'Si' then
                                 rxPartidas.FieldValues['dTotalEq']        := QryEquipo.FieldByName('dCantHHG').AsFloat
                              else
                                 rxPartidas.FieldValues['dTotalEq']        := QryEquipo.FieldByName('dCantidad').AsFloat;
                              rxPartidas.FieldValues['sDescripcionEq']  := QryEquipo.FieldByName('sDescripcion').AsString;
                              rxPartidas.Post;
                          end;
                          QryEquipo.Next;
                      end;

                      if QryEquipo.RecordCount > 0 then
                      begin
                          rxPartidas.Edit;
                          rxPartidas.FieldByName('sSuministra').AsString := 'Oculta';
                          rxPartidas.Post;
                      end;
                      {$ENDREGION}

                      {$ENDREGION}
                      QryDetallePartidas.Next;
                  end;

                  rxPartidas.Append;
                  rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                  rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                  rxPartidas.Post;

                  if QryDetallePartidas.FieldByName('dCantidadAjusteNext2').AsFloat > 0 then
                  begin
                      iTotalEsp :=  QryDetallePartidas.FieldByName('dCantidadAjusteNext2').AsInteger;
                      for num := 1 to iTotalEsp do
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;
                      end;
                  end;

                  if QryDetallePartidas.FieldByName('dCantidadAjusteNext2').AsFloat < 0 then
                  begin
                      iTotalEsp :=  QryDetallePartidas.FieldByName('dCantidadAjusteNext2').AsInteger;
                      for num := iTotalEsp + 1 to 0  do
                      begin
                          rxPartidas.Delete;
                      end;
                  end;

                  {$REGION 'PARTIDA DE PU X ACTIVIAD'}

                  QryAnexos.Active := False;
                  QryAnexos.SQL.Clear;
                  QryAnexos.SQL.Add('select * from anexos where (sTipo = "Anexo" or sTipo = "AnexoEXT" or sTipo = "AnexoC1" or sTipo = "AnexoC2" or sTipo = "AnexoC3" or sTipo ="AnexoC4") order by iOrden ');
                  QryAnexos.Open;

                  while not QryAnexos.Eof do
                  begin
                      {Ahora insertamos el pu..}
                      QryPU.SQL.Clear;
                      QryPU.SQL.Add('select b.sNumeroActividad, b.dCantidad, a.sMedida, a.sAnexo, b.mDescripcion '+
                                    'from bitacoradeactividades b '+
                                    'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sIdConvenio = b.sIdConvenio and a.sNumeroOrden  = a.sNumeroOrden '+
                                    'and a.sTipoAnexo= "PU" and a.sNumeroActividad = b.sNumeroActividad and a.sWbs = b.sWbs and a.sTipoActividad = "Actividad" and sAnexo =:Anexo) '+
                                    'where b.sContrato = :Contrato and b.sNumeroOrden = :Orden and b.dIdFecha = :Fecha and sIdTipoMovimiento = "E" and b.sWbs_ADM = :Wbs ');
                      QryPU.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryPU.ParamByName('Contrato').AsString  := ReporteDiario.FieldByName('sOrden').AsString;
                      QryPU.ParamByName('Wbs').AsString       := QryPartidas.FieldByName('sWbs').AsString;
                      QryPU.ParamByName('Anexo').AsString     := QryAnexos.FieldByName('sAnexo').AsString;
                      QryPU.Open;

                      if QryPU.RecordCount > 0 then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloPUAnexo';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['mDescripcionPU']   := QryAnexos.FieldValues['sDescripcion'];
                          rxPartidas.Post;

                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloPU';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;
                      end;

                      {P A R T I D A  D E  P U...}
                      while not QryPU.Eof do
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'DatosPU';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad']   := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['sNumeroActividadPU'] := QryPU.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['dCantidadPU'] := QryPU.FieldByName('dCantidad').AsFloat;
                          rxPartidas.FieldValues['sMedidaPU']   := QryPU.FieldByName('sMedida').AsString;
                          rxPartidas.FieldValues['sAnexoPU']    := QryPU.FieldByName('sAnexo').AsString;
                          rxPartidas.FieldValues['mDescripcionPU']  := QryPU.FieldValues['mDescripcion'];
                          rxPartidas.Post;

                          QryPU.Next;
                      end;

                      if QryPU.RecordCount > 0 then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;
                      end;
                      QryAnexos.Next;
                  end;

                  {$ENDREGION}
                 QryPartidas.Next;
              end;

              QryFolios.Next;
          end;


          {Esta secccion es para mostrar la hoja vacia sino existen datos.}
          if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
          begin
             if QryFolios.RecordCount = 0 then
             begin
                 rxPartidas.Append;
                 rxPartidas.FieldValues['sOT'] := '';
                 rxPartidas.Post;
             end;
          end;
      end;

      {$ENDREGION}
      Td_Partidas.DataSet:= rxPartidas;
      Td_Partidas.FieldAliases.Clear;

      Td_ImpPartidas.DataSet:= rxPartidas;
      Td_ImpPartidas.FieldAliases.Clear;


      Reporte.DataSets.Add(Td_partidas);
      Reporte.DataSets.Add(Td_ImpPartidas);

    Finally
      QrAux.destroy;
    End;

end;

Procedure ReportePDF_ActividadesPorFolio(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
Var
  QryInfoFolio,
  QryPartidas,
  QryDetallePartidas,
  QryConsultaAvanceAnterior,
  QryConsultaAvancesAcumulados,
  QryPersonal,
  QryEquipo,
  QryMateriales,
  QryFolios,
  QryPU,
  QryAnexos,
  QryTiemposAct     : TZQuery;

  Td_partidas,
  Td_ImpPartidas,
  Td_ClasificacionTiempo : TfrxDBDataset;
  dDiaSiguiente  : TDateTime;
  rxPartidas,
  rxTiemposAct   : TClientDataSet;
  iTotalPersonal, indice : integer;
  lContinua      : boolean;
  QrAux:TzReadOnlyQuery;
  sIidDiario,
  siIdTarea,
  siIdActividad:String;
begin
    QrAux:=TzReadOnlyQuery.create(nil);
    Try
      {Creacion de objetos de consulta}
      QryConsultaAvancesAcumulados := TZQuery.Create(nil);
      QryConsultaAvancesAcumulados.Connection := Connection.zConnection;
      QryConsultaAvanceAnterior := TZQuery.Create(nil);
      QryConsultaAvanceAnterior.Connection := connection.zConnection;
      QryFolios := TZQuery.Create(nil);
      QryFolios.Connection := Connection.zConnection;
      QryPartidas := TZQuery.Create(nil);
      QryPartidas.Connection := Connection.zConnection;
      QryDetallePartidas := TZQuery.Create(nil);
      QryDetallePartidas.Connection := Connection.zConnection;
      QryInfoFolio := TZQuery.Create(nil);
      QryInfoFolio.Connection := Connection.zConnection;
      QryPersonal := TZQuery.Create(nil);
      QryPersonal.Connection := Connection.zConnection;
      QryEquipo := TZQuery.Create(nil);
      QryEquipo.Connection := Connection.zConnection;
      QryMateriales := TZQuery.Create(nil);
      QryMateriales.Connection := Connection.zConnection;
      QryPU := TZQuery.Create(nil);
      QryPU.Connection := Connection.zConnection;
      QryAnexos := TZQuery.Create(nil);
      QryAnexos.Connection := Connection.zConnection;
      QryTiemposAct := TZQuery.Create(nil);
      QryTiemposAct.Connection := Connection.zConnection;

      Td_partidas:= TfrxDBDataset.Create(nil);
      Td_partidas.UserName:='td_partidas';

      Td_ImpPartidas:= TfrxDBDataset.Create(nil);
      Td_ImpPartidas.UserName:='Td_ImpPartidas';

      QrAux.connection:=Connection.zconnection;
      rxPartidas := TClientDataSet.Create(nil);
      rxPartidas.Active := False;
      rxPartidas.FieldDefs.Add('sTipo', ftString, 25, False);
      rxPartidas.FieldDefs.Add('sOT', ftString, 15, False);
      rxPartidas.FieldDefs.Add('sFolio', ftString, 50, False);
      rxPartidas.FieldDefs.Add('sPlataforma', ftString, 50, False);
      rxPartidas.FieldDefs.Add('mDescripcionFolio', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sNumeroActividad', ftString, 20, False);
      rxPartidas.FieldDefs.Add('mDescripcion', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('iIdDiario', ftInteger, 0, False);
      rxPartidas.FieldDefs.Add('sClasificacion', ftString, 10, False);
      rxPartidas.FieldDefs.Add('sHoraInicio', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sHoraFin', ftString, 5, False);
      rxPartidas.FieldDefs.Add('mDescripcionHorario', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('sHoraInicioPer', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sHoraFinPer', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sPartidaPer', ftString, 20, False);
      rxPartidas.FieldDefs.Add('dCantidadPer', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dTotalPer', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sDescripcionPer', ftString, 250, False);
      rxPartidas.FieldDefs.Add('sHoraInicioEq', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sHoraFinEq', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sPartidaEq', ftString, 10, False);
      rxPartidas.FieldDefs.Add('dCantidadEq', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dTotalEq', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sDescripcionEq', ftString, 250, False);
      rxPartidas.FieldDefs.Add('sMaterial', ftString, 25, False);
      rxPartidas.FieldDefs.Add('dCantidadMat', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dCantidadMatAnt', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sMedidaMat', ftString, 8, False);
      rxPartidas.FieldDefs.Add('sSuministra', ftString, 25, False);
      rxPartidas.FieldDefs.Add('mDescripcionMaterial', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('sNumeroActividadPU', ftString, 20, False);
      rxPartidas.FieldDefs.Add('dCantidadPU', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sMedidaPU', ftString, 10, False);
      rxPartidas.FieldDefs.Add('sAnexoPU', ftString, 25, False);
      rxPartidas.FieldDefs.Add('mDescripcionPU', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('iHermano', ftInteger, 0, False);
      rxPartidas.CreateDataSet;
      rxPartidas.Active := True;

      if (TImpresion=FtsActividades) or (TImpresion=FtsAll) then
      begin
          {$REGION 'FOLIOS REPORTADOS'}
          QryFolios.SQL.Clear;
          QryFolios.SQL.Add('SELECT ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.mDescripcion, p.sDescripcion as plataforma, cv.sIdConvenio as Convenio, cv.iGerencialInicio, cv.iGerencialFinal '+
                            'FROM ordenesdetrabajo AS ot ' +
                            ' inner join convenios cv on (cv.sContrato = ot.sContrato and ot.sNumeroOrden = cv.sNumeroOrden '+
                            '       and (:fecha between cv.dFechaInicio and cv.dFechaFinal)) '+
                            '	INNER JOIN contratos AS c ' +
                            '		ON (ot.sContrato=c.sContrato) ' +
                            '	INNER JOIN bitacoradeactividades AS ba ' +
                            '		ON (ba.sContrato=c.sContrato AND ba.sNumeroOrden=ot.sNumeroOrden) ' +
                            '	LEFT JOIN tiposdemovimiento AS tm ' +
                            '		ON (tm.sContrato= :ContratoBarco AND tm.sIdTipoMovimiento=ba.sIdTipoMovimiento AND tm.sClasificacion="Tarifa Diaria") ' +
                            'inner join plataformas  p on (p.sIdPlataforma = ot.sIdPlataforma) '+
                            'WHERE c.sContrato= :Orden AND ba.dIdFecha= :Fecha ' +
                            'GROUP BY cv.sNumeroorden  order by ot.iOrden ');
          QryFolios.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
          QryFolios.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          QryFolios.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
          QryFolios.Open;

          while Not QryFolios.Eof do
          begin
              {$REGION 'AVANCES DEL FOLIO'}
              {Avances anteriores}
              QryConsultaAvanceAnterior.SQL.Clear;
              QryConsultaAvanceAnterior.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', ReporteDiario.FieldByName('dIdFecha').AsDateTime)+'", :Orden, :Folio), 2) AS dAvanceAnterior;';
              QryConsultaAvanceAnterior.ParamByName('Orden').AsString := ReporteDiario.FieldByName('sOrden').AsString;
              QryConsultaAvanceAnterior.ParamByName('Folio').AsString := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvanceAnterior.Open;

              //Avances Acumulados
              dDiaSiguiente := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              dDiaSiguiente := IncDay(dDiaSiguiente);

              QryConsultaAvancesAcumulados.SQL.Clear;
              QryConsultaAvancesAcumulados.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 2) AS dAvanceAcumulado;';
              QryConsultaAvancesAcumulados.ParamByName('Orden').AsString := ReporteDiario.FieldByName('sOrden').AsString;
              QryConsultaAvancesAcumulados.ParamByName('Folio').AsString := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvancesAcumulados.Open;
              {$ENDREGION}

              {$REGION 'CONSULTA - PARTIDAS REPORTADAS'}
              QryPartidas.SQL.Clear;
              QryPartidas.SQL.Add('SELECT  a.sContrato, a.iIdDiario,	a.sNumeroActividad, o.dPonderado, a.sWbs, '+
                                  ' (SELECT   SUM(dAvance) ' +
                                  ' FROM bitacoradeactividades ' +
                                  ' WHERE ' +
                                  '   dIdFecha < a.dIdFecha ' +
                                  '   AND sIdTipoMovimiento = "E" ' +
                                  '   AND sWbs = a.sWbs ' +
                                  '   AND sContrato = a.sContrato ' +
                                  '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceAnterior_GeneralPartida, '+

                                  ' (SELECT  SUM(dAvance) ' +
                                  ' FROM bitacoradeactividades ' +
                                  ' WHERE ' +
                                  '   dIdFecha = a.dIdFecha ' +
                                  '   AND sIdTipoMovimiento = "E" ' +
                                  '   AND sWbs = a.sWbs ' +
                                  '   AND sContrato = a.sContrato ' +
                                  '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceActual_GeneralPartida, ' +

                                  '	CONCAT(a.sHoraInicio,"-",a.sHoraFinal)  AS Horario, ' +
                                  '	a.sIdClasificacion,	a.sHoraInicio, a.sHorafinal, ' +
                                  '	o.mdescripcion as sDescripcionActividad ' +
                                  'FROM bitacoradeactividades a ' +
                                  'INNER JOIN actividadesxorden o ' +
                                  '		ON (o.sContrato = a.sContrato AND o.sIdConvenio = a.sIdConvenio AND o.sNumeroOrden = a.sNumeroOrden AND a.sWbs = o.sWbs AND o.sNumeroActividad = a.sNumeroActividad AND o.sTipoActividad = "Actividad" and o.sTipoAnexo = "ADM") ' +

                                  'WHERE a.sContrato = :Orden AND a.dIdFecha = :Fecha AND a.sNumeroOrden = :Folio AND a.sIdTurno = :Turno AND a.sIdTipoMovimiento = "E" and a.lImprime ="Si" ' +
                                  'GROUP BY	a.sNumeroActividad ' +
                                  'ORDER BY a.sContrato, o.iItemOrden, a.sNumeroActividad, a.sHoraInicio ');
              QryPartidas.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sOrden').AsString;
              QryPartidas.ParamByName('Fecha').AsDateTime  := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryPartidas.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryPartidas.ParamByName('Turno').AsString    := ReporteDiario.FieldByName('sIdTurno').AsString;
              QryPartidas.Open;
              {$ENDREGION}

              while Not QryPartidas.Eof do
              begin
                  {Insertar Datos en RxPartidas}
                  rxPartidas.Append;
                  rxPartidas.FieldValues['sTipo']       := 'General';
                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                  rxPartidas.FieldValues['mDescripcionFolio'] := QryFolios.FieldValues['mDescripcion'];
                  rxPartidas.FieldValues['dAvanceAntFolio']   := QryConsultaAvanceAnterior.FieldByName('dAvanceAnterior').AsFloat;
                  rxPartidas.FieldValues['dAvanceActFolio']   := QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').AsFloat - QryConsultaAvanceAnterior.FieldByName('dAvanceAnterior').AsFloat;
                  rxPartidas.FieldValues['dAvanceAcumFolio']  := QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').AsFloat;
                  rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                  rxPartidas.FieldValues['mDescripcion']     := QryPartidas.FieldByName('sDescripcionActividad').AsString;
                  rxPartidas.FieldValues['dAvanceAntPart']   := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat;
                  rxPartidas.FieldValues['dAvanceActPart']   := QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                  rxPartidas.FieldValues['dAvanceAcumPart']  := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat + QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                  rxPartidas.Post;

                  {$REGION 'CONSULTA - DETALLE DE PARTIDAS'}
                  QryDetallePartidas.SQL.Clear;
                  QryDetallePartidas.SQL.Add( 'SELECT b.dIdFecha, b.sIdClasificacion, b.sHoraInicio, b.sHoraFinal, b.iIdDiario, b.sNumeroActividad, b.iIdTarea,b.iIdActividad,b.iHermano, '+
                                              '	( SELECT (ifnull(sum(ba.dAvance), 0)) ' +
                                              '		FROM ' +
                                              '			bitacoradeactividades AS ba ' +
                                              '		WHERE ' +
                                              '			ba.sContrato = b.sContrato ' +
                                              '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                              '		AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento ' +
                                              '		AND ba.swbs = b.swbs ' +
                                              '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                              '		AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
                                              '   < cast(b.sHoraInicio AS Time))  '+
                                              '   or (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) = cast(b.sHoraInicio AS Time) '+
                                              '   AND cast(ba.sHoraFinal AS Time) < cast(b.sHoraFinal AS Time) ) ' +
                                              ')	) AS dAvanceAnterior, ' +
                                              '	b.dAvance, b.mDescripcion AS descAct ' +
                                              'FROM ' +
                                              '	bitacoradeactividades b ' +
                                              'INNER JOIN tiposdemovimiento tm ON ( ' +
                                              '	tm.sContrato = :ContratoBarco ' +
                                              '	AND tm.sIdTipoMovimiento = b.sIdClasificacion ) ' +
                                              'INNER JOIN actividadesxorden ao ON ( ' +
                                              '	ao.sContrato = b.sContrato and ao.sIdConvenio =b.sIdConvenio and ao.sNumeroOrden = b.sNumeroOrden and ' +
                                              '	ao.sWbs = b.sWbs and ao.sTipoActividad = "Actividad" ) ' +
                                              'WHERE	b.sContrato = :Orden ' +
                                              'AND b.sNumeroOrden = :Folio ' +
                                              'AND Find_In_Set(b.sIdTipoMovimiento,"ED,EN") ' +
                                              'AND b.sNumeroActividad = :Actividad ' +
                                              'AND dIdFecha = :Fecha ');

                                              if tipo=FtAbordo then
                                                QryDetallePartidas.SQL.Add('and ((b.iHermano=-1) or (b.iHermano=b.iIdActividad)) ');

                  QryDetallePartidas.SQL.Add('AND lImprime = "Si" group by b.iIdTarea,b.iIdActividad ' +
                                              'ORDER BY b.didFecha, ao.iItemOrden, Time(b.sHoraInicio), Time(b.sHorafinal) ');
                  QryDetallePartidas.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
                  QryDetallePartidas.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
                  QryDetallePartidas.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                  QryDetallePartidas.ParamByName('Folio').AsString         := QryFolios.FieldByName('sNumeroOrden').AsString;
                  QryDetallePartidas.ParamByName('Actividad').AsString     := QryPartidas.FieldByName('sNumeroActividad').AsString;
                  QryDetallePartidas.Open;
                  {$ENDREGION}

                  QryPersonal.Active := False;
                  QryEquipo.Active   := False;

                  {$REGION 'INSERCIÓN DE DATOS'}
                  while Not QryDetallePartidas.Eof do
                  begin
                      rxPartidas.Append;
                      rxPartidas.FieldValues['sTipo']       := 'TituloAvances';
                      rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                      rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                      rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                      rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                      if (QryDetallePartidas.FieldByName('iHermano').AsInteger=-1) or
                            (QryDetallePartidas.FieldByName('iHermano').AsInteger=QryDetallePartidas.FieldByName('iIdActividad').AsInteger)
                      then
                        rxPartidas.FieldValues['iHermano'] :=-1;
                      rxPartidas.Post;

                      if QryPersonal.Active then
                      begin
                          if QryPersonal.RecordCount >= QryEquipo.RecordCount then
                          begin
                              rxPartidas.Append;
                              rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                              rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                              rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                              rxPartidas.Post;
                          end;
                      end;

                      if rxPartidas.FieldValues['sTipo'] <> 'TituloHorario' then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloHorario';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString ;
                          rxPartidas.FieldValues['sHoraInicio']  := QryDetallePartidas.FieldByName('sHoraInicio').AsString;
                          rxPartidas.FieldValues['mDescripcionHorario']  := QryDetallePartidas.FieldByName('descAct').AsString;//GHH New
                          if (QryDetallePartidas.FieldByName('iHermano').AsInteger=-1) or
                            (QryDetallePartidas.FieldByName('iHermano').AsInteger=QryDetallePartidas.FieldByName('iIdActividad').AsInteger)
                          then
                            rxPartidas.FieldValues['iHermano'] :=-1;
                          rxPartidas.Post;
                      end;

                      if QryDetallePartidas.FieldByName('iHermano').AsInteger=-1 then
                      begin
                        rxPartidas.Append;
                        rxPartidas.FieldValues['sTipo']       := 'DatosHorario';
                        rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                        rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                        rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                        rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        rxPartidas.FieldValues['sClasificacion']   := QryDetallePartidas.FieldByName('sIdClasificacion').AsString ;
                        rxPartidas.FieldValues['sHoraInicio']  := QryDetallePartidas.FieldByName('sHoraInicio').AsString;
                        rxPartidas.FieldValues['sHoraFin']     := QryDetallePartidas.FieldByName('sHoraFinal').AsString;
                        //rxPartidas.FieldValues['mDescripcionHorario']  := QryDetallePartidas.FieldByName('descAct').AsString; //GHH Orig
                        rxPartidas.FieldValues['dAvanceAntPart']   := QryDetallePartidas.FieldByName('dAvanceAnterior').AsFloat;
                        rxPartidas.FieldValues['dAvanceActPart']   := QryDetallePartidas.FieldByName('dAvance').AsFloat;
                        rxPartidas.FieldValues['dAvanceAcumPart']  := QryDetallePartidas.FieldByName('dAvanceAnterior').AsFloat + QryDetallePartidas.FieldByName('dAvance').AsFloat;
                        rxPartidas.Post;
                      end
                      else
                      begin
                        QrAux.active:=False;
                        QrAux.sql.text:='select * from bitacoradeactividades where scontrato=:Orden and '+
                                        'didfecha=:fecha and sNumeroOrden=:Folio and sIdTipoMovimiento = "ED" and ' +
                                        'sNumeroActividad = :Actividad and iHermano=:IdVinculo '+
                                        'group by iIdTarea,iIdActividad ' +
                                        'ORDER BY Time(sHoraInicio) ';
                        QrAux.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
                        QrAux.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                        QrAux.ParamByName('Folio').AsString         := QryFolios.FieldByName('sNumeroOrden').AsString;
                        QrAux.ParamByName('Actividad').AsString     := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        QrAux.ParamByName('IdVinculo').AsInteger     := QryDetallePartidas.FieldByName('iHermano').AsInteger;
                        QrAux.open;

                        rxPartidas.Append;
                        if QryDetallePartidas.FieldByName('iHermano').AsInteger=QryDetallePartidas.FieldByName('iIdActividad').AsInteger then
                          rxPartidas.FieldValues['sTipo']       := 'DatosHorario'
                        else
                          rxPartidas.FieldValues['sTipo']       := 'DatosHorario2';
                        rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                        rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                        rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                        rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        rxPartidas.FieldValues['sClasificacion']   := QryDetallePartidas.FieldByName('sIdClasificacion').AsString ;
                        rxPartidas.FieldValues['sHoraInicio']  := QrAux.FieldByName('sHoraInicio').AsString;
                        //rxPartidas.FieldValues['mDescripcionHorario']  := QryDetallePartidas.FieldByName('descAct').AsString;   GHH Orig
                        rxPartidas.FieldValues['dAvanceAntPart']   := QryDetallePartidas.FieldByName('dAvanceAnterior').AsFloat;
                        rxPartidas.FieldValues['dAvanceActPart']   := QryDetallePartidas.FieldByName('dAvance').AsFloat;
                        rxPartidas.FieldValues['dAvanceAcumPart']  := QryDetallePartidas.FieldByName('dAvanceAnterior').AsFloat + QryDetallePartidas.FieldByName('dAvance').AsFloat;
                        sIidDiario:='';
                        siIdTarea:='';
                        siIdActividad:='';
                        while not QrAux.eof do
                        begin
                          if sIidDiario='' then
                            sIidDiario:=QrAux.FieldByName('iIdDiario').AsString
                          else
                            sIidDiario:= sIidDiario + ',' + QrAux.FieldByName('iIdDiario').AsString;

                          if siIdTarea='' then
                            siIdTarea:=QrAux.FieldByName('iIdTarea').AsString
                          else
                            siIdTarea:= siIdTarea + ',' + QrAux.FieldByName('iIdTarea').AsString;

                          if siIdActividad='' then
                            siIdActividad:=QrAux.FieldByName('iIdActividad').AsString
                          else
                            siIdActividad:= siIdActividad + ',' + QrAux.FieldByName('iIdActividad').AsString;

                          QrAux.Next;
                        end;
                        rxPartidas.FieldValues['sHoraFin']     := QrAux.FieldByName('sHoraFinal').AsString;
                        rxPartidas.Post;



                      end;

                      {$REGION 'PERSONAL X HORARIOS'}
                      {Ahora insertamos el personal..}
                      QryPersonal.SQL.Clear;
                      if (QryDetallePartidas.FieldByName('iHermano').AsInteger=-1) or (tipo=FtTierra) then
                      begin
                        // Cambios realizados por E. Rangel 20/08/2016
                        {QryPersonal.SQL.Add('select ' +
                                            '  bp.*, ' +
                                            '  sum( bp.dCantidad ) as Cantidad, ' +
                                            '  (IFNULL(sum(CAST(bp.dCantHH AS CHAR)), 0) + ' +
                                            '   IFNULL((SELECT ' +
                                            '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                            '          FROM ' +
                                            '            bitacoradeajustes ba ' +
                                            '          WHERE ' +
                                            '            ba.sContrato = bp.sContrato AND ' +
                                            '            ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                            '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                            '            ba.dFecha = bp.dIdFecha), 0)) as dCantHHG ' +
                                            'from ' +
                                            '  bitacoradepersonal bp ' +
                                            'where ' +
                                            '  bp.sContrato =:Orden and ' +
                                            '  bp.dIdFecha =:fecha and ' +
                                            '  bp.sNumeroActividad =:Actividad and ' +
                                            '  bp.iIdTarea =:IdTarea and ' +
                                            '  bp.iIdActividad=:IdActividad and ' +
                                            '  bp.lImprime = "Si" ' +
                                            'group by ' +
                                            '  sHoraInicioG, ' +
                                            '  sHoraFinalG, ' +
                                            '  sIdPersonal ' +
                                            'order by ' +
                                            '  iItemOrden');}
                        QryPersonal.SQL.Add('select *, sum( dCantidad ) as Cantidad, (sum(dCantHH) + ' +
                                            ' IFNULL((SELECT ' +
                                            '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                            '         FROM ' +
                                            '           bitacoradeajustes ba ' +
                                            '         WHERE ' +
                                            '           ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                            '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                            '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantHHG from bitacoradepersonal where sContrato =:Orden '+
                                            'and dIdFecha =:fecha and sNumeroActividad =:Actividad and iIdTarea =:IdTarea and iIdActividad=:IdActividad and lImprime = "Si" group by sHoraInicioG, sHoraFinalG, sIdPersonal order by iItemOrden ');
                        QryPersonal.ParamByName('IdTarea').AsInteger  := QryDetallePartidas.FieldValues['iIdTarea'];
                        QryPersonal.ParamByName('IdActividad').AsInteger  := QryDetallePartidas.FieldValues['iIdActividad'];

                      end
                      else
                      begin
                        // Cambios realizados por E. Rangel 20/08/2016
                        {QryPersonal.SQL.Add('select ' +
                                            '  bp.*, ' +
                                            '  sum( bp.dCantidad ) as Cantidad, ' +
                                            '  (IFNULL(sum(CAST(bp.dCantHH AS CHAR)), 0) + ' +
                                            '   IFNULL((SELECT ' +
                                            '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                            '          FROM ' +
                                            '            bitacoradeajustes ba ' +
                                            '          WHERE ' +
                                            '            ba.sContrato = bp.sContrato AND ' +
                                            '            ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                            '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                            '            ba.dFecha = bp.dIdFecha), 0))as dCantHHG ' +
                                            'from ' +
                                            '  bitacoradepersonal bp ' +
                                            'where ' +
                                            '  bp.sContrato =:Orden and ' +
                                            '  bp.dIdFecha =:fecha and ' +
                                            '  bp.sNumeroActividad =:Actividad and ' +
                                            '  bp.iIdTarea in ('+siIdTarea+') and ' +
                                            '  bp.iIdActividad in ('+siIdActividad+') and ' +
                                            '  bp.lImprime = "Si" ' +
                                            'group by ' +
                                            '  sHoraInicioG, ' +
                                            '  sHoraFinalG, ' +
                                            '  sIdPersonal ' +
                                            'order by ' +
                                            '  iItemOrden, ' +
                                            '  time(sHoraInicio)');}
                        QryPersonal.SQL.Add('select *, sum( bp.dCantidad ) as Cantidad, (sum(bp.dCantHH) + ' +
                                            ' IFNULL((SELECT ' +
                                          '          SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                          '        FROM ' +
                                          '          bitacoradeajustes ba ' +
                                          '        WHERE ' +
                                          '          ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                          '          ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                          '          ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantHHG from bitacoradepersonal where sContrato =:Orden '+
                                            'and bp.dIdFecha =:fecha and bp.sNumeroActividad =:Actividad and '+
                                            'bp.iIdTarea in ('+siIdTarea+') and bp.iIdActividad in ('+siIdActividad+') and bp.lImprime = "Si" '+
                                            'group by bp.sHoraInicioG, bp.sHoraFinalG, bp.sIdPersonal order by bp.iItemOrden,time(bp.sHoraInicio)');

                      end;
                      QryPersonal.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryPersonal.ParamByName('Orden').AsString     := ReporteDiario.FieldByName('sOrden').AsString;
                      QryPersonal.ParamByName('Actividad').AsString := QryDetallePartidas.FieldByName('sNumeroActividad').AsString;

                      QryPersonal.Open;

                      rxPartidas.Append;
                      rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                      rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                      rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                      rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                      rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString ;
                      rxPartidas.Post;

                      if QryPersonal.RecordCount > 0 then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;

                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloDatosPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString ;
                          if QryPersonal.recordcount>0 then
                            if  AnsiContainsText(QryPersonal.FieldByName('sDescripcion').AsString,'TIEMPO EXTRA') then
                              rxPartidas.FieldValues['sMedidaPU'] := 'HORAS';
                          rxPartidas.Post;
                      end;

                      {P E R S O N A L...}
                      while not QryPersonal.Eof do
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'DatosPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['sHoraInicioPer']   := QryPersonal.FieldByName('sHoraInicioG').AsString;
                          rxPartidas.FieldValues['sHoraFinPer']      := QryPersonal.FieldByName('sHoraFinalG').AsString;
                          rxPartidas.FieldValues['sPartidaPer']      := QryPersonal.FieldByName('sIdPersonal').AsString;
                          rxPartidas.FieldValues['dCantidadPer']     := QryPersonal.FieldByName('Cantidad').AsFloat;
                          rxPartidas.FieldValues['dTotalPer']        := QryPersonal.FieldByName('dCantHHG').AsFloat;
                          rxPartidas.FieldValues['sDescripcionPer']  := QryPersonal.FieldByName('sDescripcion').AsString;
                          rxPartidas.Post;
                          QryPersonal.Next;
                      end;
                      {$ENDREGION}

                      {$REGION 'EQUIPO X HORARIOS'}
                      {Ahora insertamos el equipo..}
                      QryEquipo.SQL.Clear;
                      if (QryDetallePartidas.FieldByName('iHermano').AsInteger=-1)  or (tipo=FtTierra) then
                      begin
                        QryEquipo.SQL.Add('select *, (sum(dCantHH) + ' +
                                          ' IFNULL((SELECT ' +
                                          '          SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                          '        FROM ' +
                                          '          bitacoradeajustes ba ' +
                                          '        WHERE ' +
                                          '          ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                                          '          ba.sNumeroActividad = be.sNumeroActividad AND ' +
                                          '          ba.sIdPartidaAnexo = be.sIdEquipo AND ba.dFecha = be.dIdFecha), 0)) as dCantHHG from bitacoradeequipos where sContrato =:Orden '+
                                          'and dIdFecha =:fecha and sNumeroActividad =:Actividad and iIdTarea =:IdTarea and iIdActividad=:IdActividad and lImprime ="Si" group by sHoraInicioG, sIdEquipo order by iItemOrden ');
                        QryEquipo.ParamByName('IdTarea').AsInteger  := QryDetallePartidas.FieldValues['iIdTarea'];
                        QryEquipo.ParamByName('IdActividad').AsInteger  := QryDetallePartidas.FieldValues['iIdActividad'];
                      end
                      else
                      begin
                        QryEquipo.SQL.Add('select *, (sum(dCantHH) + ' +
                                          ' IFNULL((SELECT ' +
                                          '          SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                          '        FROM ' +
                                          '          bitacoradeajustes ba ' +
                                          '        WHERE ' +
                                          '          ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                                          '          ba.sNumeroActividad = be.sNumeroActividad AND ' +
                                          '          ba.sIdPartidaAnexo = be.sIdEquipo AND ba.dFecha = be.dIdFecha), 0)) as dCantHHG from bitacoradeequipos where sContrato =:Orden '+
                                          'and dIdFecha =:fecha and sNumeroActividad =:Actividad and '+
                                          'iIdTarea in ('+siIdTarea+') and iIdActividad in ('+siIdActividad+') and lImprime = "Si" '+
                                          'group by sHoraInicioG, sIdEquipo order by iItemOrden,time(sHoraInicio)');
                      end;
                      QryEquipo.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryEquipo.ParamByName('Orden').AsString     := ReporteDiario.FieldByName('sOrden').AsString;
                      QryEquipo.ParamByName('Actividad').AsString := QryDetallePartidas.FieldByName('sNumeroActividad').AsString;
                      QryEquipo.Open;

                      {Verificamos cuanto personal se inserto para retroceder el memory.}
                      iTotalPersonal := QryPersonal.RecordCount;

                      if (QryEquipo.RecordCount > 0) and (iTotalPersonal = 0) then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['dCantidadEq']     := 1;
                          rxPartidas.Post;

                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloDatosPersonal';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['dCantidadEq']     := 1;
                          rxPartidas.Post;
                      end;

                      for indice := 1 to iTotalPersonal - 1 do
                          rxPartidas.Prior;

                      if QryEquipo.recordcount>0 then
                      begin
                        rxPartidas.Prior;
                        rxPartidas.Prior;
                        rxPartidas.Edit;
                        rxPartidas.FieldValues['dCantidadEq']     := 1;
                        rxPartidas.Post;
                        rxPartidas.next;
                        rxPartidas.Edit;
                        rxPartidas.FieldValues['dCantidadEq']     := 1;
                        rxPartidas.Post;
                        rxPartidas.next;
                      end;

                      indice := 1;
                      lContinua := True;
                      {E Q U I P O...}
                      while not QryEquipo.Eof do
                      begin
                          if iTotalPersonal > 0 then
                          begin
                              if indice <= iTotalPersonal then
                              begin
                                  rxPartidas.Edit;
                                  rxPartidas.FieldValues['sHoraInicioEq']   := QryEquipo.FieldByName('sHoraInicioG').AsString;
                                  rxPartidas.FieldValues['sHoraFinEq']      := QryEquipo.FieldByName('sHoraFinalG').AsString;
                                  rxPartidas.FieldValues['sPartidaEq']      := QryEquipo.FieldByName('sIdEquipo').AsString;
                                  rxPartidas.FieldValues['dCantidadEq']     := QryEquipo.FieldByName('dCantidad').AsFloat;
                                  rxPartidas.FieldValues['dTotalEq']        := QryEquipo.FieldByName('dCantHHG').AsFloat;
                                  rxPartidas.FieldValues['sDescripcionEq']  := QryEquipo.FieldByName('sDescripcion').AsString;
                                  rxPartidas.Post;
                                  rxPartidas.Next;
                                  inc(indice);
                                  lContinua := False;
                              end
                              else
                                  lContinua := True;
                          end;

                          if lContinua then
                          begin
                              rxPartidas.Append;
                              rxPartidas.FieldValues['sTipo']       := 'DatosPersonal';
                              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                              rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                              rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                              rxPartidas.FieldValues['sHoraInicioEq']   := QryEquipo.FieldByName('sHoraInicioG').AsString;
                              rxPartidas.FieldValues['sHoraFinEq']      := QryEquipo.FieldByName('sHoraFinalG').AsString;
                              rxPartidas.FieldValues['sPartidaEq']      := QryEquipo.FieldByName('sIdEquipo').AsString;
                              rxPartidas.FieldValues['dCantidadEq']     := QryEquipo.FieldByName('dCantidad').AsFloat;
                              rxPartidas.FieldValues['dTotalEq']        := QryEquipo.FieldByName('dCantHHG').AsFloat;
                              rxPartidas.FieldValues['sDescripcionEq']  := QryEquipo.FieldByName('sDescripcion').AsString;
                              rxPartidas.Post;
                          end;
                          QryEquipo.Next;
                      end;
                      {$ENDREGION}

                      {$REGION 'MATERIAL X PARTIDA'}
                      {Ahora insertamos el material..}
                      QryMateriales.SQL.Clear;
                      QryMateriales.SQL.Add('select m.sIdMaterial, sum(m.dCantidad) as dCantidad, aa.sMedida, aa.mDescripcion, '+
                                  '(select sum(m1.dCantidad) as cantidad from bitacorademateriales m1 '+
                                  'where m1.sContrato = m.sContrato and m1.dIdFecha <= m.dIdFecha and m1.sIdMaterial = m.sIdMaterial and m1.iIddiario < m.iIdDiario ) as dCantidadAnt '+
                                  'from bitacorademateriales m '+
                                  'inner join actividadesxanexo aa on (aa.scontrato =:Contrato and aa.sNumeroActividad =m.sIdMaterial and aa.sTipoActividad = "Actividad") '+
                                  'where m.sContrato =:Orden and m.dIdFecha =:Fecha and m.iIdDiario = :Id '+
                                  'group by m.sIdMaterial order by aa.sWbs');
                      QryMateriales.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryMateriales.ParamByName('Orden').AsString     := ReporteDiario.FieldByName('sOrden').AsString;
                      QryMateriales.ParamByName('Id').AsInteger       := QryDetallePartidas.FieldByName('iIdDiario').AsInteger;
                      QryMateriales.ParamByName('Contrato').AsString  := global_Contrato_Barco;
                      QryMateriales.Open;

                      if QryMateriales.RecordCount > 0 then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;

                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloMaterial';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;
                      end;

                      {M A T E R I A L...}
                      while not QryMateriales.Eof do
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'DatosMaterial';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['sMaterial']    := QryMateriales.FieldByName('sIdMaterial').AsString;
                          rxPartidas.FieldValues['dCantidadMat'] := QryMateriales.FieldByName('dCantidad').AsFloat;
                          rxPartidas.FieldValues['sMedidaMat']   := QryMateriales.FieldByName('sMedida').AsString;
                          rxPartidas.FieldValues['dCantidadMatAnt']:= QryMateriales.FieldByName('dCantidadAnt').AsFloat;
                          rxPartidas.FieldValues['mDescripcionMaterial']  := QryMateriales.FieldValues['mDescripcion'];
                          rxPartidas.Post;

                          QryMateriales.Next;
                      end;
                      {$ENDREGION}


                      {$ENDREGION}
                      QryDetallePartidas.Next;
                  end;

                  rxPartidas.Append;
                  rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                  rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                  rxPartidas.Post;

                  {$REGION 'PARTIDA DE PU X ACTIVIAD'}

                  QryAnexos.Active := False;
                  QryAnexos.SQL.Clear;
                  QryAnexos.SQL.Add('select * from anexos where (sTipo = "Anexo" or sTipo = "AnexoEXT" or sTipo = "AnexoC1" or sTipo = "AnexoC2" or sTipo = "AnexoC3" or sTipo ="AnexoC4") order by iOrden ');
                  QryAnexos.Open;

                  while not QryAnexos.Eof do
                  begin
                      {Ahora insertamos el pu..}
                      QryPU.SQL.Clear;
                      QryPU.SQL.Add('select b.sNumeroActividad, b.dCantidad, a.sMedida, a.sAnexo, b.mDescripcion '+
                                    'from bitacoradeactividades b '+
                                    'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sIdConvenio = b.sIdConvenio and a.sNumeroOrden  = a.sNumeroOrden '+
                                    'and a.sTipoAnexo= "PU" and a.sNumeroActividad = b.sNumeroActividad and a.sWbs = b.sWbs and a.sTipoActividad = "Actividad" and sAnexo =:Anexo) '+
                                    'where b.sContrato = :Contrato and b.sNumeroOrden = :Orden and b.dIdFecha = :Fecha and sIdTipoMovimiento = "E" and b.sWbs_ADM = :Wbs ');
                      QryPU.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryPU.ParamByName('Contrato').AsString  := ReporteDiario.FieldByName('sOrden').AsString;
                      QryPU.ParamByName('Wbs').AsString       := QryPartidas.FieldByName('sWbs').AsString;
                      QryPU.ParamByName('Anexo').AsString     := QryAnexos.FieldByName('sAnexo').AsString;
                      QryPU.Open;

                      if QryPU.RecordCount > 0 then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloPUAnexo';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['mDescripcionPU']   := QryAnexos.FieldValues['sDescripcion'];
                          rxPartidas.Post;

                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloPU';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;
                      end;

                      {P A R T I D A  D E  P U...}
                      while not QryPU.Eof do
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'DatosPU';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad']   := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['sNumeroActividadPU'] := QryPU.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.FieldValues['dCantidadPU'] := QryPU.FieldByName('dCantidad').AsFloat;
                          rxPartidas.FieldValues['sMedidaPU']   := QryPU.FieldByName('sMedida').AsString;
                          rxPartidas.FieldValues['sAnexoPU']    := QryPU.FieldByName('sAnexo').AsString;
                          rxPartidas.FieldValues['mDescripcionPU']  := QryPU.FieldValues['mDescripcion'];
                          rxPartidas.Post;

                          QryPU.Next;
                      end;

                      if QryPU.RecordCount > 0 then
                      begin
                          rxPartidas.Append;
                          rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
                          rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                          rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                          rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                          rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                          rxPartidas.Post;
                      end;
                      QryAnexos.Next;
                  end;

                  {$ENDREGION}
                 QryPartidas.Next;
              end;

              QryFolios.Next;
          end;


          {Esta secccion es para mostrar la hoja vacia sino existen datos.}
          if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
          begin
             if QryFolios.RecordCount = 0 then
             begin
                 rxPartidas.Append;
                 rxPartidas.FieldValues['sOT'] := '';
                 rxPartidas.Post;
             end;
          end;
      end;

      {$ENDREGION}
      Td_Partidas.DataSet:= rxPartidas;
      Td_Partidas.FieldAliases.Clear;

      Td_ImpPartidas.DataSet:= rxPartidas;
      Td_ImpPartidas.FieldAliases.Clear;


      Reporte.DataSets.Add(Td_partidas);
      Reporte.DataSets.Add(Td_ImpPartidas);

    Finally
      QrAux.destroy;
    End;

end;

Procedure ReportePDF_NotasGenerales(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
Var
  QryNotas: TZQuery;
  QryNotasGenerales: TZQuery;
  QryFoliosConNotas: TZQuery;
  QryTiemposAct: TZQuery;
  QryArribosMov: TZQuery;
  Td_NotasGrales,
  Td_ImpNotasGrales,
  Td_ClasificacionTiempo : TfrxDBDataset;
  Td_ArribosMov : TfrxDBDataset;
  rxNotas, rxTiemposAct: TClientDataSet;
  indice : integer;
  sGrupo : string;
begin

  {$REGION 'NOTAS'}
  Try

    QryNotas := TZQuery.Create(nil);
    QryNotas.Connection := Connection.zConnection;
    QryFoliosConNotas := TZQuery.Create(nil);
    QryFoliosConNotas.Connection := Connection.zConnection;
    QryNotasGenerales := TZQuery.Create(nil);
    QryNotasGenerales.Connection := Connection.zConnection;
    QryTiemposAct := TZQuery.Create(nil);
    QryTiemposAct.Connection := Connection.zConnection;
    QryArribosMov := TZQuery.Create(nil);
    QryArribosMov.Connection := Connection.zConnection;

    Td_NotasGrales:= TfrxDBDataset.Create(nil);
    Td_NotasGrales.UserName:='Td_NotasGrales';

    Td_ImpNotasGrales:= TfrxDBDataset.Create(nil);
    Td_ImpNotasGrales.UserName:='Td_ImpNotasGrales';

    Td_ClasificacionTiempo:= TfrxDBDataset.Create(nil);
    Td_ClasificacionTiempo.UserName:='Td_ClasificacionTiempo';

    Td_ArribosMov:= TfrxDBDataset.Create(nil);
    Td_ArribosMov.UserName:='Td_ArribosMov';

    rxNotas := TClientDataSet.Create(nil);
    rxNotas.Active := False;
    rxNotas.FieldDefs.Add('sNotaGeneral', ftMemo, 0, False);
    rxNotas.CreateDataSet;
    rxNotas.Active := True;
    rxNotas.Append;
    rxNotas.FieldValues['sNotaGeneral'] := '';
    rxNotas.Post;

    rxTiemposAct := TClientDataSet.Create(nil);
    rxTiemposAct.Active := False;
    rxTiemposAct.FieldDefs.Add('sGrupo', ftString, 30, False);
    rxTiemposAct.FieldDefs.Add('sContrato', ftString, 15, False);
    rxTiemposAct.FieldDefs.Add('sTiempo1', ftString, 10, False);
    rxTiemposAct.FieldDefs.Add('sDescripcion1', ftString, 100, False);
    rxTiemposAct.FieldDefs.Add('sTiempo2', ftString, 10, False);
    rxTiemposAct.FieldDefs.Add('sDescripcion2', ftString, 100, False);
    rxTiemposAct.CreateDataSet;
    rxTiemposAct.Active := True;

    if (TImpresion=FtsNGenerales) or (TImpresion=FtsAll) then
    begin
        QryNotasGenerales.SQL.Clear;
		if tipo=ftAbordo then
		 QryNotasGenerales.SQL.Add('SELECT * FROM notas_generales WHERE sContrato = :Orden AND dIdFecha = :Fecha and eAplicaResumenPersonal = "No" and lAplicaConsumos ="No" ORDER BY iOrden ASC')
        else
		 QryNotasGenerales.SQL.Add('SELECT * FROM notas_generales WHERE sContrato = :Orden AND dIdFecha = :Fecha and eAplicaResumenPersonal = "No" ORDER BY iOrden ASC');
        QryNotasGenerales.ParamByName('Orden').AsString  := ReporteDiario.FieldByName('sOrden').AsString;
        QryNotasGenerales.ParamByName('Fecha').AsDate    := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        QryNotasGenerales.Open;

        while not QryNotasGenerales.Eof do
        begin
			if tipo=ftAbordo then
			begin
			  rxNotas.Append;
              rxNotas.FieldValues['sNotaGeneral'] := 'NOTA '+ IntToStr(QryNotasGenerales.RecNo) +': '+ QryNotasGenerales.FieldByName('sNotaGeneral').AsWideString;
              rxNotas.Post;
            end
			else
			begin
              rxNotas.Edit;
              rxNotas.FieldValues['sNotaGeneral'] := rxNotas.FieldValues['sNotaGeneral'] + QryNotasGenerales.FieldByName('sNotaGeneral').AsWideString;
              rxNotas.Post;
			end;
            QryNotasGenerales.Next;
        end;

        {Recuadro de Clasificacion para tiempo de actividades}
          QryTiemposAct.SQL.Clear;
          if tipo=ftAbordo then
		  begin
            QryTiemposAct.SQL.Add('select * from tiposdemovimiento where sContrato =:Orden and (sClasificacion = "Tiempo de Actividades" OR sClasificacion = "Movimiento de Barco") order by sClasificacion, iOrden' );
            if connection.configuracion.FieldByName('sIdEmbarcacion').AsString ='*' then
               QryTiemposAct.ParamByName('Orden').AsString := global_contrato_barco          
          end
		  else
		    QryTiemposAct.SQL.Add('select * from tiposdemovimiento where sContrato =:Orden and sClasificacion = "Tiempo de Actividades" order by sClasificacion, iOrden' );

          QryTiemposAct.ParamByName('Orden').AsString := global_contrato; // global_contrato_barco;
          QryTiemposAct.Open;

          indice := 1;
          sGrupo := '';
          while not QryTiemposAct.Eof do
          begin
              if sGrupo <> QryTiemposAct.FieldByName('sClasificacion').AsString then
              begin
                  sGrupo := QryTiemposAct.FieldByName('sClasificacion').AsString;
                  indice := 3;
              end;

              if indice = 3 then
                 indice := 1;
              if indice = 1 then
                 rxTiemposAct.Append
              else
                 rxTiemposAct.Edit;
              rxTiemposAct.FieldValues['sGrupo']                        := QryTiemposAct.FieldByName('sClasificacion').AsString;
              rxTiemposAct.FieldValues['sContrato']                     := QryTiemposAct.FieldByName('sContrato').AsString;
              rxTiemposAct.FieldValues['sTiempo'+IntToStr(indice)]      := QryTiemposAct.FieldByName('sIdTipoMovimiento').AsString;
              rxTiemposAct.FieldValues['sDescripcion'+IntToStr(indice)] := QryTiemposAct.FieldByName('sDescripcion').AsString;
              rxTiemposAct.Post;
              inc(indice);      
              QryTiemposAct.Next;
          end;

          if tipo=ftAbordo then
          begin
            QryArribosMov.Connection := connection.zconnection;
            QryArribosMov.SQL.Add('select concat( me.sHoraInicio , " - ", me.sHoraFinal ) as sHorario, me.sHoraInicio, me.sHoraFinal, '+
                    'me.mDescripcion, me.sNumeroActividad '+
                    'from movimientosdeembarcacion me '+
                    'where me.sContrato = :contratoBarco '+
                    'and me.dIdFecha = :fecha '+
                    'and me.sOrden = :orden '+
                    'and me.sClasificacion = "" and sTipo = "MOVIMIENTO" '+
                    'order by me.sHoraInicio');
            QryArribosMov.ParamByName('contratoBarco').AsString := global_contrato_barco;
            QryArribosMov.ParamByName('fecha').AsDate := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
            QryArribosMov.ParamByName('orden' ).AsString := global_contrato;
            QryArribosMov.Open;
          end;

        {Esta secccion es para mostrar la hoja vacia sino existen datos.}
        //if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
         //  rxNotas.Delete;
    end
    else
       rxNotas.Delete;

    Td_NotasGrales.DataSet:= rxNotas;
    Td_NotasGrales.FieldAliases.Clear;

    Td_ImpNotasGrales.DataSet:= rxNotas;
    Td_ImpNotasGrales.FieldAliases.Clear;

    Td_ClasificacionTiempo.DataSet := rxTiemposAct;
    Td_ClasificacionTiempo.FieldAliases.Clear;

    Td_ArribosMov.DataSet := QryArribosMov;
    Td_ArribosMov.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_NotasGrales);
    Reporte.DataSets.Add(Td_ImpNotasGrales);
    Reporte.DataSets.Add(Td_ClasificacionTiempo);
    Reporte.DataSets.Add(Td_ArribosMov);

  Finally
  End;

  {$ENDREGION}
end;

Procedure ReportePDF_ConcentradoDePersonal2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);

Var
  i, x,y,CuentaCol,iGrupo,iCiclo: Integer;
  DCiclo,dTotal, dSubTotal, dSubTotalPU :Double;
  QryConsulta,
  QrMoe,
  QrColumnas,
  QrRecursos,
  QryCondiciones,
  QryEmbarcacion,
  QryOtroPersonal: TZQuery;
  QryAgrupador: TZQuery;

  dContrato_Inicio,
  dContrato_Final: TDateTime;
  mDatos:TJvMemoryData;
  mImprime:TJvMemoryData;
  Td_CImpDistribucionP: TfrxDBDataset;
  Td_CDistribucionP: TfrxDBDataset;
  ValTmp:Variant;
  CantTmp:Double;
  iPosTmp:Integer;
  lImprimeRecurso : string;
begin

  mDatos:=TJvMemoryData.Create(nil);
  mImprime:=TJvMemoryData.Create(nil);
  QrRecursos := TZQuery.Create(nil);
  QrColumnas := TZQuery.Create(nil);
  QrMoe := TZQuery.Create(nil);
  Td_CDistribucionP:=TfrxDBDataset.Create(nil);
  Td_CImpDistribucionP:=TfrxDBDataset.Create(nil);
  try
    QrMoe.Connection := Connection.zConnection;
    QrColumnas.Connection := Connection.zConnection;
    QrRecursos.Connection := Connection.zConnection;
    Td_CDistribucionP.UserName:='Td_CDistribucionP';
    Td_CImpDistribucionP.UserName:='Td_CImpDistribucionP';
    with mDatos do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('iTotalGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidAnexo', ftString, 10, false);
      FieldDefs.Add('sidrecurso', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 250, True);
      FieldDefs.Add('sAnexo', ftString, 250, false);
      FieldDefs.Add('sTitulo', ftString, 250, false);
      FieldDefs.Add('smedida', ftString, 100, True);
      FieldDefs.Add('dcantSol', FtFloat, 0, True);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      FieldDefs.Add('dcantTotalPU', FtFloat, 0, True);
      for CuentaCol:=1 to TotalCol do
      begin
        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPernocta' + Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPlataforma'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sTipo'+ Inttostr(CuentaCol), ftString, 5, false);
      end;
      Active:=true;
    end;
    with mImprime do
    begin
      Active:=false;
      FieldDefs.Add('iCampo', ftInteger, 0, True);
      Active:=true;
    end;

    mImprime.EmptyTable;
    if (TImpresion=FtsCPersonal) or (TImpresion=FtsAll) then
    begin
      QrRecursos.Active := False;
      QrRecursos.SQL.Clear;
      // Cambios realizados por E. Rangel - 20/08/2016
      // Ésta es la sentencia que genera los datos del resumen de personal en el reporte diario
      QrRecursos.SQL.Add('SELECT ' +
                         '  bp.sIdPersonal, ' +
                         '  bp.sDescripcion, ' +
                         '  ifnull(SUM(CAST(bp.dAjuste AS CHAR)),0) AS Ajuste, ' +
                         '  if(:Anexo=1,( IF(SUM(bp.dCanthh) > 0, (SUM(xround(bp.dCanthh, 2)) + ' +
                         '   IFNULL((SELECT ' +
                         '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                         '          FROM ' +
                         '            bitacoradeajustes ba ' +
                         '          WHERE ' +
                         '            ba.sContrato = bp.sContrato AND ' +
                         '            ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                         '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                         '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                         '            ba.dFecha = bp.dIdFecha), 0)) ,' +
                         'SUM(bp.dCantidad))),SUM(bp.dCantidad)) AS Total, ' +
                         '  bp.lImprime ' +
                         'FROM ' +
                         '  bitacoradepersonal bp ' +
                         'inner join ' +
                         '  bitacoradeactividades ba ' +
                         '    on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and ' +
                         '        bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and ' +
                         '        ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) ' +
                         'WHERE ' +
                         '  bp.scontrato = :Orden AND bp.sNumeroOrden = :Folio AND bp.didfecha = :Fecha AND ' +
                         '  bp.sidPernocta = :Pernocta AND bp.sidplataforma = :Plataforma and ' +
                         '  bp.sIdPersonal=:Personal and bp.sTipoObra =:Tipo ' +
                         'GROUP BY ' +
                         '  bp.sIdPersonal');
      {QrRecursos.SQL.Add('SELECT bp.sIdPersonal, bp.sDescripcion, ifnull(SUM(bp.dAjuste),0) AS Ajuste, ' +
                         'if(:Anexo=1,( ' +
                         'IF(SUM(bp.dCanthh) > 0, xround(SUM(bp.dCanthh),2), SUM(bp.dCantidad))'+
                         '),SUM(bp.dCantidad)) AS Total, bp.lImprime ' +
                         'FROM bitacoradepersonal bp ' +
                         'inner join bitacoradeactividades ba ' +
                         'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                         'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                         'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) ' +
                         'WHERE bp.scontrato = :Orden AND bp.sNumeroOrden = :Folio AND bp.didfecha = :Fecha ' +
                         'AND bp.sidPernocta = :Pernocta AND bp.sidplataforma = :Plataforma ' +
                         'and bp.sIdPersonal=:Personal and bp.sTipoObra =:Tipo GROUP BY bp.sIdPersonal ');}



      {$REGION 'CONSULTAS - PARTIDAS'}
      QrColumnas.active:=false;
      QrColumnas.SQL.Add( 'SELECT ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, p.sDescripcion AS pernocta, p.sDescripcion AS pernocta, pf.sDescripcion AS plataforma, '+
                          'bp.sIdpernocta AS PernoctaP, bp.sIdPlataforma AS idPlataforma, ' +
                          'bp.sTipoObra as sTipo '+
                          'FROM ordenesdetrabajo ot ' +
                          'INNER JOIN contratos AS c ON (ot.sContrato=c.sContrato) ' +
                          'INNER JOIN bitacoradepersonal AS bp ON (bp.scontrato=:OT AND bp.sNumeroOrden = ot.sNumeroOrden ) ' +
                          'INNER JOIN pernoctan AS p ON (p.sidPernocta=bp.sIdpernocta) ' +
                          'INNER JOIN plataformas AS pf ON (pf.sidPlataforma=bp.sIdPlataforma) ' +
                          'WHERE (c.sContrato=:OT OR c.sCodigo=:OT)	AND bp.dIdFecha= :Fecha ' +
                          'GROUP BY	ot.sContrato, ot.sNumeroorden, bp.sidPernocta, bp.sIdPlataforma, bp.sTipoObra '+
                          'order by ot.sContrato, ot.iOrden ASC, bp.sTipoObra  DESC');
      QrColumnas.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
      QrColumnas.ParamByName('Fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
      QrColumnas.Open;
      {$ENDREGION}

      {$REGION 'CONSULTAS - TODO EL PERSONAL SOLICITADO QUE SE REGISTRA EN EL MOE'}
      QrMoe.Active := False;
      QrMoe.SQL.Clear;
      QrMoe.SQL.Add('SELECT p.sDescripcion,	mr.*, p.lSumaSolicitado, p.sMedida, ' +
                    'a.sanexo,ifnull(a.sdescripcion,"SIN ANEXO MAR/TIERRA") as anexo, ifnull(a.stitulo,"FALTA CONFIGURAR") as tituloAnexo,a.stierra '+
                    'FROM moe AS m ' +
                    'INNER JOIN moerecursos AS mr ON (mr.iidMoe=m.iidMoe) ' +
                    'INNER JOIN personal AS p ON (p.scontrato=:Contrato AND p.sidpersonal=mr.sidRecurso) ' +
                    'left join anexos a on(a.sAnexo=p.sAnexo) ' +
                    'WHERE m.didfecha = (SELECT max(didfecha) FROM moe WHERE didfecha <=:Fecha AND sContrato = :OT) ' +
                    'AND m.sContrato = :OT AND mr.eTipoRecurso = "Personal" ORDER BY a.iOrden,p.iItemOrden');
      QrMoe.ParamByName('Contrato').AsString := Global_Contrato_Barco;
      QrMoe.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
      QrMoe.ParamByName('Fecha').AsDateTime  := reportediario.FieldByName('dIdFecha').AsDateTime;
      QrMoe.Open;
      {$ENDREGION}

      {$REGION 'INSERTA EL MOE'}


      DCiclo:=QrColumnas.RecordCount/TotalCol;
      iCiclo:=Trunc(DCiclo);
      if (dCiclo -iCiclo)>0 then
        Inc(iCiclo,1);

      iGrupo:=1;
      dSubTotal   := 0;
      dSubTotalPU := 0;
      while iGrupo<=iCiclo do
      begin
        with qrMoe do
        begin

          first;

          while not eof do
          begin
            if iGrupo > 1 then
            begin
               mDatos.Locate('sidrecurso',FieldByName('sidrecurso').asstring, []);
               dSubTotal   := mDatos.FieldByName('dcantTotal').AsFloat;
               dSubTotalPU := mDatos.FieldByName('dcantTotalPU').AsFloat;
            end;

            mDatos.Append;
            mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
            mDatos.FieldByName('iTotalGrupo').AsInteger:=iCiclo;
            mDatos.FieldByName('sidAnexo').AsString:=FieldByName('sanexo').asstring;
            mDatos.FieldByName('sAnexo').AsString:=FieldByName('anexo').asstring;
            mDatos.FieldByName('sTitulo').AsString:=FieldByName('tituloAnexo').asstring;
            mDatos.FieldByName('sidrecurso').AsString:=FieldByName('sidrecurso').asstring;
            mDatos.FieldByName('sdescripcion').AsString:=fieldbyname('sdescripcion').asstring;
            mDatos.FieldByName('smedida').AsString:=fieldbyname('smedida').asstring;
            mDatos.FieldByName('dcantSol').AsFloat:=fieldbyname('dcantidad').AsFloat;
            mDatos.FieldByName('dcantTotalPU').AsFloat:=dSubTotalPU;
            mDatos.FieldByName('dcantTotal').AsFloat  :=dSubTotal;
            dTotal:=0;
            if iGrupo=1 then
              QrColumnas.First
            else
              QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

            ValTmp:=0;
            CantTmp:=0;
            iPosTmp:=0;
            CuentaCol:=1;
            lImprimeRecurso := 'Si';
            while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
            begin

              mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString :=qrColumnas.FieldbyName('sIdFolio').AsString;
              mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString   :=QrColumnas.Fieldbyname('Pernocta').asstring;
              mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  :=QrColumnas.Fieldbyname('Plataforma').Asstring;
              mDatos.FieldByName('sTipo'+ Inttostr(CuentaCol)).AsString        :=QrColumnas.Fieldbyname('sTipo').Asstring;

              QrRecursos.Active := False;
              QrRecursos.ParamByName('Anexo').AsINteger :=1;
              QrRecursos.ParamByName('Orden').AsString      := QrColumnas.FieldByName('sContrato').AsString;
              QrRecursos.ParamByName('Folio').AsString      := QrColumnas.FieldByName('sNumeroOrden').AsString;
              QrRecursos.ParamByName('Fecha').AsDate        := reportediario.FieldByName('dIdFecha').AsDateTime;
              QrRecursos.ParamByName('Pernocta').AsString   := QrColumnas.FieldByName('PernoctaP').AsString;
              QrRecursos.ParamByName('Plataforma').AsString := QrColumnas.FieldByName('idPlataforma').AsString;
              QrRecursos.ParamByName('Personal').AsString   := FieldByName('sIdRecurso').AsString;
              QrRecursos.ParamByName('Tipo').AsString       := QrColumnas.Fieldbyname('sTipo').Asstring;
              QrRecursos.Open;

              if QrRecursos.RecordCount>0 then
              begin
                ValTmp                                                       :=  ValTmp + (QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value);
                mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:= (QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value);
                dTotal                                                       := dTotal+ (QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value);
                if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                begin
                  CantTmp := mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                  iPosTmp := CuentaCol;
                end;

                if QrColumnas.Fieldbyname('sTipo').Asstring = 'TD' then
                   mDatos.FieldByName('dcantTotal').AsFloat   := mDatos.FieldByName('dcantTotal').AsFloat + mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat
                else
                   mDatos.FieldByName('dcantTotalPU').AsFloat := mDatos.FieldByName('dcantTotalPU').AsFloat + mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;

                lImprimeRecurso := QrRecursos.FieldByName('lImprime').AsString;
              end;
              QrColumnas.next;
              Inc(CuentaCol);
            end;

            if dTotal<>xRound(ValTmp,2) then
            begin
              mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
              + (xRound(ValTmp,2)-dTotal);
            end;

            if lImprimeRecurso = 'Si' then
               mDatos.Post
            else
               mDatos.Cancel;

            next;
          end;

        end;
        Inc(iGrupo,1);
      end;

     {Esta secccion es para mostrar la hoja vacia sino existen datos.}
//     if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
//        if MDatos.RecordCount>=0 then
//     else
//        if MDatos.RecordCount>0 then
//     begin
//
//     end;

      mDatos.first;
      mImprime.Append;
      mImprime.FieldByName('iCampo').AsInteger:=1;
      mImprime.Post;

      {$ENDREGION}
    end;
    Td_CDistribucionP.DataSet:=MDatos;
    Td_CDistribucionP.FieldAliases.Clear;
    Td_CImpDistribucionP.DataSet:=MImprime;
    Td_CImpDistribucionP.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_CDistribucionP);
    Reporte.DataSets.Add(Td_CImpDistribucionP);
  finally
    QrRecursos.destroy;
    QrColumnas.destroy;
    QrMoe.destroy;
  end;
end;

Procedure ReportePDF_ConcentradoDePersonalTierra(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);

Var
  i, x,y,CuentaCol,iGrupo,iCiclo: Integer;
  DCiclo,dTotal:Double;
  QryConsulta,
  QrMoe,
  QrColumnas,
  QrRecursos,
  QryCondiciones,
  QryEmbarcacion,
  QryOtroPersonal: TZQuery;
  QryAgrupador: TZQuery;
  QrMHE:TZReadOnlyQuery;

  dContrato_Inicio,
  dContrato_Final: TDateTime;
  mDatos,mDatosTE:TJvMemoryData;
  mImprime:TJvMemoryData;
  Td_CImpDistribucionP: TfrxDBDataset;
  Td_CDistribucionP,
  Td_CDistribucionHE: TfrxDBDataset;
  ValTmp:Variant;
  CantTmp:Double;
  iPosTmp:Integer;
begin

  mDatos:=TJvMemoryData.Create(nil);
  mDatosTE:=TJvMemoryData.Create(nil);
  mImprime:=TJvMemoryData.Create(nil);
  
  QrRecursos := TZQuery.Create(nil);
  QrColumnas := TZQuery.Create(nil);
  QrMoe := TZQuery.Create(nil);
  QrMHE:=TZReadOnlyQuery.Create(nil);
  Td_CDistribucionP:=TfrxDBDataset.Create(nil);
  Td_CImpDistribucionP:=TfrxDBDataset.Create(nil);
  Td_CDistribucionHE:=TfrxDBDataset.Create(nil);
  try
    QrMoe.Connection := Connection.zConnection;
    QrColumnas.Connection := Connection.zConnection;
    QrRecursos.Connection := Connection.zConnection;
    QrMHE.Connection:=connection.zConnection;
    Td_CDistribucionP.UserName:='Td_CDistribucionP';
    Td_CDistribucionHE.UserName:='Td_CDistribucionHE';
    Td_CImpDistribucionP.UserName:='Td_CImpDistribucionP';
    with mDatos do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidAnexo', ftString, 10, false);
      FieldDefs.Add('sidrecurso', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 250, True);
      FieldDefs.Add('sAnexo', ftString, 250, false);
      FieldDefs.Add('sTitulo', ftString, 250, false);
      FieldDefs.Add('smedida', ftString, 100, True);
      FieldDefs.Add('dcantSol', FtFloat, 0, True);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      FieldDefs.Add('dcantTotalPU', FtFloat, 0, True);
      for CuentaCol:=1 to TotalCol do
      begin
        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPernocta' + Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPlataforma'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sTipo'+ Inttostr(CuentaCol), ftString, 5, false);
      end;
      Active:=true;
    end;

    with mDatosTe do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidAnexo', ftString, 10, false);
      FieldDefs.Add('sidrecurso', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 250, True);
      FieldDefs.Add('sAnexo', ftString, 250, false);
      FieldDefs.Add('sTitulo', ftString, 250, false);
      FieldDefs.Add('smedida', ftString, 100, True);
      FieldDefs.Add('dcantSol', FtFloat, 0, True);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      FieldDefs.Add('dcantTotalPU', FtFloat, 0, True);
      for CuentaCol:=1 to TotalCol do
      begin
        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPernocta' + Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPlataforma'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sTipo'+ Inttostr(CuentaCol), ftString, 5, false);
      end;
      Active:=true;
    end;


    with mImprime do
    begin
      Active:=false;
      FieldDefs.Add('iCampo', ftInteger, 0, True);
      Active:=true;
    end;

    mImprime.EmptyTable;
    if (TImpresion=FtsCPersonal) or (TImpresion=FtsAll) then
    begin
      QrRecursos.Active := False;
      QrRecursos.SQL.Clear;
      QrRecursos.SQL.Add('SELECT bp.sIdPersonal, bp.sDescripcion, ' +
                         'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                         'if(:Anexo=1,( ' +
                         'IF(SUM(bp.dCanthh) > 0, (SUM(bp.dCanthh) + ' +
                         '   IFNULL((SELECT ' +
                         '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                         '           FROM ' +
                         '             bitacoradeajustes ba ' +
                         '           WHERE ' +
                         '             ba.sContrato = bp.sContrato AND ' +
                         '             ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                         '             ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                         '             ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                         '             ba.dFecha = bp.dIdFecha), 0)), SUM(bp.dCantidad))'+
                         '),SUM(bp.dCantidad)) AS Total ' +
                         'FROM bitacoradepersonal bp ' +
                         'inner join bitacoradeactividades ba ' +
                         'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                         'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                         'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad) ' +
                         'WHERE bp.scontrato = :Orden AND bp.sNumeroOrden = :Folio AND bp.didfecha = :Fecha ' +
                         'AND bp.sidPernocta = :Pernocta AND bp.sidplataforma = :Plataforma ' +
                         'and bp.sIdPersonal=:Personal GROUP BY bp.sIdPersonal ');

      



      {$REGION 'CONSULTAS - PARTIDAS'}
      QrColumnas.active:=false;
      QrColumnas.SQL.Add( 'SELECT ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, p.sDescripcion AS pernocta, p.sDescripcion AS pernocta, pf.sDescripcion AS plataforma, ot.sTipo, '+
                          'bp.sIdpernocta AS PernoctaP, bp.sIdPlataforma AS idPlataforma ' +
                          'FROM ordenesdetrabajo ot ' +
                          'INNER JOIN contratos AS c ON (ot.sContrato=c.sContrato) ' +
                          'INNER JOIN bitacoradepersonal AS bp ON (bp.scontrato=:OT AND bp.sNumeroOrden = ot.sNumeroOrden ) ' +
                          'INNER JOIN pernoctan AS p ON (p.sidPernocta=bp.sIdpernocta) ' +
                          'INNER JOIN plataformas AS pf ON (pf.sidPlataforma=bp.sIdPlataforma) ' +
                          'WHERE (c.sContrato=:OT OR c.sCodigo=:OT)	AND bp.dIdFecha= :Fecha ' +
                          'GROUP BY	ot.sContrato, ot.sNumeroorden, bp.sidPernocta, bp.sIdPlataforma');
      QrColumnas.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
      QrColumnas.ParamByName('Fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
      QrColumnas.Open;
      {$ENDREGION}

      {$REGION 'CONSULTAS - TODO EL PERSONAL SOLICITADO QUE SE REGISTRA EN EL MOE'}
      QrMoe.Active := False;
      QrMoe.SQL.Clear;
      QrMoe.SQL.Add('SELECT p.sDescripcion,	mr.*, p.lSumaSolicitado, p.sMedida,a.stipo as tipoanexo, ' +
                    'a.sanexo,ifnull(a.sdescripcion,"SIN ANEXO MAR/TIERRA") as anexo, ifnull(a.stitulo,"FALTA CONFIGURAR") as tituloAnexo,a.stierra '+
                    'FROM moe AS m ' +
                    'INNER JOIN moerecursos AS mr ON (mr.iidMoe=m.iidMoe) ' +
                    'INNER JOIN personal AS p ON (p.scontrato=:Contrato AND p.sidpersonal=mr.sidRecurso) ' +
                    'left join anexos a on(a.sAnexo=p.sAnexo) ' +
                    'WHERE m.didfecha = (SELECT max(didfecha) FROM moe WHERE didfecha <=:Fecha AND sContrato = :OT) ' +
                    'AND m.sContrato = :OT AND mr.eTipoRecurso = "Personal" and a.stipo=:tipoAnexo ORDER BY a.iOrden,p.iItemOrden');
      QrMoe.ParamByName('Contrato').AsString := Global_Contrato_Barco;
      QrMoe.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
      QrMoe.ParamByName('Fecha').AsDateTime  := reportediario.FieldByName('dIdFecha').AsDateTime;
      QrMoe.ParamByName('tipoAnexo').AsString:='PERSONAL';
      QrMoe.Open;
      {$ENDREGION}

      {$REGION 'INSERTA EL MOE'}


      DCiclo:=QrColumnas.RecordCount/TotalCol;
      iCiclo:=Trunc(DCiclo);
      if (dCiclo -iCiclo)>0 then
        Inc(iCiclo,1);

      iGrupo:=1;
      while iGrupo<=iCiclo do
      begin
        with qrMoe do
        begin

          first;

          while not eof do
          begin
            mDatos.Append;
            mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
            mDatos.FieldByName('sidAnexo').AsString:=FieldByName('sanexo').asstring;
            mDatos.FieldByName('sAnexo').AsString:=FieldByName('anexo').asstring;
            mDatos.FieldByName('sTitulo').AsString:=FieldByName('tituloAnexo').asstring;
            mDatos.FieldByName('sidrecurso').AsString:=FieldByName('sidrecurso').asstring;
            mDatos.FieldByName('sdescripcion').AsString:=fieldbyname('sdescripcion').asstring;
            mDatos.FieldByName('smedida').AsString:=fieldbyname('smedida').asstring;
            mDatos.FieldByName('dcantSol').AsFloat:=fieldbyname('dcantidad').AsFloat;
            mDatos.FieldByName('dcantTotalPU').AsFloat:=0;
            dTotal:=0;
            if iGrupo=1 then
              QrColumnas.First
            else
              QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

            ValTmp:=0;
            CantTmp:=0;
            iPosTmp:=0;
            CuentaCol:=1;
            while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
            begin

              mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString:=qrColumnas.FieldbyName('sIdFolio').AsString;
              mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('Pernocta').asstring;
              mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('Plataforma').Asstring;
              mDatos.FieldByName('sTipo'+ Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('sTipo').Asstring;

              QrRecursos.Active := False;
              {if fieldbyname('sTierra').asstring='Si' then
                QrRecursos.ParamByName('Anexo').AsINteger :=0
              else}
                QrRecursos.ParamByName('Anexo').AsINteger :=1;
              QrRecursos.ParamByName('Orden').AsString      := QrColumnas.FieldByName('sContrato').AsString;
              QrRecursos.ParamByName('Folio').AsString      := QrColumnas.FieldByName('sNumeroOrden').AsString;
              QrRecursos.ParamByName('Fecha').AsDate        := reportediario.FieldByName('dIdFecha').AsDateTime;
              QrRecursos.ParamByName('Pernocta').AsString   := QrColumnas.FieldByName('PernoctaP').AsString;
              QrRecursos.ParamByName('Plataforma').AsString := QrColumnas.FieldByName('idPlataforma').AsString;
              QrRecursos.ParamByName('Personal').AsString := FieldByName('sIdRecurso').AsString;
              QrRecursos.Open;

              if QrRecursos.RecordCount>0 then
              begin
                ValTmp:=ValTmp + (QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value);
                mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=(xRound(QrRecursos.FieldByName('Total').Value,2) + QrRecursos.FieldByName('Ajuste').Value);
                dTotal:=dTotal+ (xRound(QrRecursos.FieldByName('Total').Value,2) + QrRecursos.FieldByName('Ajuste').Value);
                if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                begin
                  CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                  iPosTmp:=CuentaCol;
                end;
              end;
              QrColumnas.next;
              Inc(CuentaCol);
            end;
            {if dTotal>0 then
            begin }
              if dTotal<>xRound(ValTmp,2) then
              begin
                mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                + (xRound(ValTmp,2)-dTotal);
                if QrColumnas.FieldByName('sTipo').AsString = 'TD' then
                   mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2)
                else
                   mDatos.FieldByName('dcantTotalPU').AsFloat:=xRound(ValTmp,2);
              end
              else
              begin
                 if QrColumnas.FieldByName('sTipo').AsString = 'TD' then
                    mDatos.FieldByName('dcantTotal').AsFloat:=dTotal
                 else
                    mDatos.FieldByName('dcantTotalPU').AsFloat:=dTotal;
              end;
              mDatos.Post;
            {end
            else
              mDatos.Cancel;   }
            next;
          end;

        end;
        Inc(iGrupo,1);
      end;

     {Esta secccion es para mostrar la hoja vacia sino existen datos.}
//     if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
//        if MDatos.RecordCount>=0 then
//     else
//        if MDatos.RecordCount>0 then
//     begin
//
//     end;

      mDatos.first;
      mImprime.Append;
      mImprime.FieldByName('iCampo').AsInteger:=1;
      mImprime.Post;

      {$ENDREGION}


      if Tipo=FtTierra then
      begin
        {QrMHE.Active:=False;
        QrMHE.SQL.Text:='select bte.sContrato,bte.sIdPersonal as sidrecurso,bte.sPuesto as sDescripcion,sum(bte.dTiempoExtra) as dcantidad1, ' +
                        'sum(bte.dTiempoExtra) as dCantTotal ,sum(bte.dTiempoExtra) as dCantSol,p.sMedida '+
                        'from bitacoradetiemposextras bte '+
                        'inner join personal p on(p.sContrato=:Contrato and p.sIdPersonal=bte.sIdPersonal) '+
                        'where bte.sContrato=:Orden and bte.didfecha=:Fecha '+
                        'group by bte.sIdPersonal order by p.iItemorden';

        QrMHE.ParamByName('Contrato').AsString:=ReporteDiario.FieldByName('sContrato').AsString;
        QrMHE.ParamByName('Orden').AsString:=ReporteDiario.FieldByName('sOrden').AsString;
        QrMHE.ParamByName('Fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrMHE.Open;  }
        QrMoe.Close;
        QrMoe.ParamByName('Contrato').AsString := Global_Contrato_Barco;
        QrMoe.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
        QrMoe.ParamByName('Fecha').AsDateTime  := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrMoe.ParamByName('tipoAnexo').AsString:='TIEMPO_EXTRA';
        QrMoe.Open;

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
          Inc(iCiclo,1);

        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
          with qrMoe do
          begin

            first;

            while not eof do
            begin
              mDatosTE.Append;
              mDatosTE.FieldByName('iGrupo').AsInteger:=Igrupo;
              mDatosTE.FieldByName('sidAnexo').AsString:=FieldByName('sanexo').asstring;
              mDatosTE.FieldByName('sAnexo').AsString:=FieldByName('anexo').asstring;
              mDatosTE.FieldByName('sTitulo').AsString:=FieldByName('tituloAnexo').asstring;
              mDatosTE.FieldByName('sidrecurso').AsString:=FieldByName('sidrecurso').asstring;
              mDatosTE.FieldByName('sdescripcion').AsString:=fieldbyname('sdescripcion').asstring;
              mDatosTE.FieldByName('smedida').AsString:=fieldbyname('smedida').asstring;
              mDatosTE.FieldByName('dcantSol').AsFloat:=fieldbyname('dcantidad').AsFloat;
              mDatosTE.FieldByName('dcantTotalPU').AsFloat:=0;
              dTotal:=0;
              if iGrupo=1 then
                QrColumnas.First
              else
                QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

              ValTmp:=0;
              CantTmp:=0;
              iPosTmp:=0;
              CuentaCol:=1;
              while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
              begin

                mDatosTE.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString:=qrColumnas.FieldbyName('sIdFolio').AsString;
                mDatosTE.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('Pernocta').asstring;
                mDatosTE.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('Plataforma').Asstring;
                mDatosTE.FieldByName('sTipo'+ Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('sTipo').Asstring;

                QrRecursos.Active := False;
                {if fieldbyname('sTierra').asstring='Si' then
                  QrRecursos.ParamByName('Anexo').AsINteger :=0
                else}
                  QrRecursos.ParamByName('Anexo').AsINteger :=1;
                QrRecursos.ParamByName('Orden').AsString      := QrColumnas.FieldByName('sContrato').AsString;
                QrRecursos.ParamByName('Folio').AsString      := QrColumnas.FieldByName('sNumeroOrden').AsString;
                QrRecursos.ParamByName('Fecha').AsDate        := reportediario.FieldByName('dIdFecha').AsDateTime;
                QrRecursos.ParamByName('Pernocta').AsString   := QrColumnas.FieldByName('PernoctaP').AsString;
                QrRecursos.ParamByName('Plataforma').AsString := QrColumnas.FieldByName('idPlataforma').AsString;
                QrRecursos.ParamByName('Personal').AsString := FieldByName('sIdRecurso').AsString;
                QrRecursos.Open;

                if QrRecursos.RecordCount>0 then
                begin
                  ValTmp:=ValTmp + (QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value);
                  mDatosTE.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=(xRound(QrRecursos.FieldByName('Total').Value,2) + QrRecursos.FieldByName('Ajuste').Value);
                  dTotal:=dTotal+ (xRound(QrRecursos.FieldByName('Total').Value,2) + QrRecursos.FieldByName('Ajuste').Value);
                  if CantTmp<mDatosTE.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                  begin
                    CantTmp:=mDatosTE.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                    iPosTmp:=CuentaCol;
                  end;
                end;
                QrColumnas.next;
                Inc(CuentaCol);
              end;
              if dTotal>0 then
              begin
                if dTotal<>xRound(ValTmp,2) then
                begin
                  mDatosTE.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatosTE.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                  + (xRound(ValTmp,2)-dTotal);
                  if QrColumnas.FieldByName('sTipo').AsString = 'TD' then
                     mDatosTE.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2)
                  else
                     mDatosTE.FieldByName('dcantTotalPU').AsFloat:=xRound(ValTmp,2);
                end
                else
                begin
                   if QrColumnas.FieldByName('sTipo').AsString = 'TD' then
                      mDatosTE.FieldByName('dcantTotal').AsFloat:=dTotal
                   else
                      mDatosTE.FieldByName('dcantTotalPU').AsFloat:=dTotal;
                end;
                mDatosTE.Post;
              end
              else
                mDatosTE.Cancel;   
              next;
            end;

          end;
          Inc(iGrupo,1);
        end;


















      end;



    end;
    Td_CDistribucionP.DataSet:=MDatos;
    Td_CDistribucionP.FieldAliases.Clear;
    Td_CImpDistribucionP.DataSet:=MImprime;
    Td_CImpDistribucionP.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_CDistribucionP);
    Reporte.DataSets.Add(Td_CImpDistribucionP);
    if Tipo=FtTierra then
    begin
      Td_CDistribucionHE.DataSet:=mDatosTE;
      Td_CDistribucionHE.FieldAliases.Clear;
      Reporte.DataSets.Add(Td_CDistribucionHE);
    end;



  finally
    QrRecursos.destroy;
    QrColumnas.destroy;
    QrMoe.destroy;
  end;
end;

Procedure ReportePDF_ConcentradoDeEquipos2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
var

  QrMoe,
  QrColumnas,
  QrRecursos,
  QryEmbarcacion: TZQuery;

  ZqEqBarcos : tzreadonlyquery;

begin
  ZqEqBarcos := TZReadOnlyQuery.Create(nil);
  ZqEqBarcos.Connection := connection.zConnection;
  ZqEqBarcos.Active := False;
  ZqEqBarcos.SQL.Clear;
  ZqEqBarcos.SQL.Text := ' Select sidequipo,iunificador from equipos '+
                         ' where scontrato = :contrato and Length(iunificador) > 0  ';
  ZqEqBarcos.ParamByName('contrato').AsString := global_Contrato_Barco;
  ZqEqBarcos.Open;

  Try

    {$REGION 'CONSULTAS - TODOS LOS EQUIPOS REGISTRADOS EN MOE'}
      QrMoe := TZQuery.Create(nil);
      QrMoe.Connection := Connection.zConnection;
      QrMoe.Active := False;
      QrMoe.SQL.Clear;
      QrMoe.SQL.Add('SELECT p.lSumaSolicitado, p.iId_AgrupadorEquipo, ge.sGrupo, p.sDescripcion, p.iUnificador ,mr.*, p.sMedida ' +
                    'FROM moerecursos AS mr ' +
                    'INNER JOIN moe AS m ' +
                    'ON (m.iidMoe=mr.iidMoe) ' +
                    'INNER JOIN equipos AS p ' +
                    'ON (p.sContrato=:Contrato AND p.sIdEquipo=mr.sidRecurso) ' +
                    'LEFT JOIN gruposdeequipo AS ge ' +
                    'ON (ge.iId = p.iId_AgrupadorEquipo) ' +
                    'WHERE ' +
                    'dIdFecha = (SELECT max(didfecha) FROM moe WHERE didfecha <=:Fecha AND sContrato = :ContratoNormal) ' +
                    'AND m.sContrato = :ContratoNormal ' +
                    'AND mr.eTipoRecurso="Equipo" ' +
                    'ORDER BY p.iItemOrden');
      QrMoe.ParamByName('Contrato').AsString := Global_Contrato_Barco;
      QrMoe.ParamByName('ContratoNormal').AsString := ReporteDiario.FieldByName('sOrden').AsString;
      QrMoe.ParamByName('Fecha').AsDateTime := reportediario.FieldByName('dIdFecha').AsDateTime;
      QrMoe.Open;

    {$ENDREGION}

    {$REGION 'INSERCION DE DATOS - INFORMACIÓN DEL EQUIPO'}

    while Not QrMoe.Eof do
    begin
      QrMoe.FieldbyName('sIdRecurso').asstring;
      QrMoe.Fieldbyname('sDescripcion').asstring;
      QrMoe.fieldbyname('sMedida').asstring;
      QrMoe.FieldByName('dCantidad').asfloat;
      QrMoe.Next;
    end;

    {$ENDREGION}

    {$REGION 'CONSULTAS - PARTIDAS'}
      QrColumnas := TZQuery.Create(nil);
      QrColumnas.Connection := Connection.zConnection;
      QrColumnas.SQL.Add( '' +
                          'SELECT ' +
                          '	ot.*, ' +
                          '	p.sDescripcion AS pernocta, ' +
                          '	pf.sDescripcion AS plataforma, ' +
                          '	bp.sIdpernocta AS PernoctaP ' +
                          'FROM ordenesdetrabajo AS ot ' +
                          ' INNER JOIN contratos AS c ' +
                          '	  ON (ot.sContrato=c.sContrato) ' +
                          ' INNER JOIN bitacoradeequipos AS bp ' +
                          '		ON (bp.scontrato=:OT AND bp.sNumeroOrden = ot.sNumeroOrden) ' +
                          ' INNER JOIN pernoctan AS p ' +
                          '	  ON (p.sidPernocta=bp.sIdpernocta) ' +
                          ' INNER JOIN plataformas AS pf ' +
                          '	  ON (pf.sidPlataforma=ot.sIdPlataforma) ' +
                          'WHERE ' +
                          '	(c.sContrato=:Contrato OR c.sCodigo=:Contrato) ' +
                          '	AND bp.dIdFecha= :Fecha ' +
                          'GROUP BY ' +
                          '	ot.sContrato, ' +
                          '	ot.sNumeroorden, ' +
                          '	bp.sidPernocta;');
      QrColumnas.ParamByName('OT').AsString:= ReporteDiario.FieldByName('sOrden').AsString;
      QrColumnas.ParamByName('Contrato').AsString:= ReporteDiario.FieldByName('sOrden').AsString;
      QrColumnas.ParamByName('Fecha').AsDate:=reportediario.FieldByName('dIdFecha').AsDateTime;
      QrColumnas.Open;
    {$ENDREGION}

    {$REGION 'INSERCION DE DATOS - PARTIDAS'}
      QrRecursos := TZQuery.Create(nil);
      QrRecursos.Connection := Connection.zConnection;
      while Not QrColumnas.Eof do
      begin
        {$REGION 'CABECERA'}
        qrColumnas.fieldbyname('snumeroorden').asstring;
        qrcolumnas.fieldbyname('pernocta').asstring;
        qrColumnas.FieldByName('plataforma').asstring;
        {$ENDREGION}

        {$REGION 'CONSULTA CUADRE'}
          QrRecursos.Active := False;
          QrRecursos.SQL.Clear;
          QrRecursos.SQL.Add( '' +
                              'SELECT ' +
                              '	bp.sIdEquipo, ' +
                              '	bp.sDescripcion, ' +
                              '	IF(SUM(bp.dCanthh) > 0, (SUM(ROUND(bp.dCanthh, 2)) + ' +
                              '  IFNULL((SELECT ' +
                              '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                              '          FROM ' +
                              '            bitacoradeajustes ba ' +
                              '          WHERE ' +
                              '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                              '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                              '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) , SUM(bp.dCantidad)) AS Total, ' +
                              ' IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste ' +
                              'FROM bitacoradeequipos bp ' +
                              'WHERE ' +
                              '	bp.scontrato = :Orden ' +
                              '	AND bp.snumeroorden = :Folio ' +
                              '	AND bp.didfecha = :Fecha ' +
                              '	AND bp.sidPernocta = :Pernocta ' +
                              'GROUP BY bp.sIdEquipo ');
          QrRecursos.ParamByName('Orden').AsString := QrColumnas.FieldByName('sContrato').AsString;
          QrRecursos.ParamByName('Folio').AsString := QrColumnas.FieldByName('sNumeroOrden').AsString;
          QrRecursos.ParamByName('Fecha').AsDate := reportediario.FieldByName('dIdFecha').AsDateTime;
          QrRecursos.ParamByName('Pernocta').AsString := QrColumnas.FieldByName('PernoctaP').AsString;
          QrRecursos.Open;
        {$ENDREGION}

        {$REGION 'INSERTA EL CUADRE'}
          QrMoe.First;
          while Not QrMoe.Eof do
          begin
            if QrRecursos.Locate('sIdEquipo', QrMoe.FieldByName('sIdRecurso').AsString, []) then
              //(QrRecursos.FieldByName('Total').AsFloat + QrRecursos.FieldByName('Ajuste').AsFloat, -2);
            QrMoe.Next;
          end;
        {$ENDREGION}
        QrColumnas.Next;
      end;      
    {$ENDREGION}



    {$REGION 'ITUTLO DE LA CABECERA'}
    with connection.qrybusca do
    begin
      active := false;
      sql.text := 'select sdescripcion from anexos where stipo = "EQUIPO"';
      open;

      fieldbyname('sdescripcion').asstring;
    end;
    {$ENDREGION}

  Finally
    QrMoe.Free;
    QrRecursos.Free;
  End;
end;

Procedure ReportePDF_ConcentradoDePernoctas2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
Var
  QrPernoctas,
  QrReportado,
  QrColumnas : TZReadOnlyQuery;

  iColumna : Integer;
  CuentaCol,iGrupo,iCiclo: Integer;
  DCiclo,dTotal:Double;

  mDatos:TJvMemoryData;
  mDatosDet:TJvMemoryData;
  mImprime:TJvMemoryData;
  Td_CImpDistribucionPer: TfrxDBDataset;
  Td_CDistribucionPer: TfrxDBDataset;
  Td_CDistribucionDetPer: TfrxDBDataset;
  QrAdicional:TzReadOnlyquery;
  ValTmp:Variant;
  CantTmp:Double;
  iPosTmp:Integer;
  sTdPuConstante : string;
begin
  {$REGION 'TZQUERY'}
  QrPernoctas := TZReadOnlyQuery.Create(nil);
  QrReportado := TZReadOnlyQuery.Create(nil);
  QrColumnas := TZReadOnlyQuery.Create(nil);
  QrAdicional := TZReadOnlyQuery.Create(nil);
  mDatos:=TJvMemoryData.Create(nil);
  mDatosDet:=TJvMemoryData.Create(nil);

  mImprime:=TJvMemoryData.Create(nil);
  Td_CDistribucionPer:=TfrxDBDataset.Create(nil);
  Td_CImpDistribucionPer:=TfrxDBDataset.Create(nil);
  Td_CDistribucionDetPer:=TfrxDBDataset.Create(nil);
  {$ENDREGION}

  try
    QrReportado.connection := connection.zconnection;
    QrPernoctas.connection := connection.zconnection;
    QrColumnas.connection := connection.zconnection;
    QrAdicional.Connection:=connection.zconnection;
    Td_CDistribucionPer.UserName:='Td_CDistribucionPer';
    Td_CImpDistribucionPer.UserName:='Td_CImpDistribucionPer';
    Td_CDistribucionDetPer.UserName:='Td_CDistribucionDetPer';

    with mDatos do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidrecurso', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 250, True);
      FieldDefs.Add('smedida', ftString, 100, false);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      FieldDefs.Add('dcantTotal3C', FtFloat, 0, True);
      for CuentaCol:=1 to TotalColPer do
      begin
        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPernocta' + Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPlataforma'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sTipo'+ Inttostr(CuentaCol), ftString, 10, false);
        FieldDefs.Add('dcantidad3C' + Inttostr(CuentaCol), FtFloat, 0, false);
      end;
      Active:=true;
    end;
    with mImprime do
    begin
      Active:=false;
      FieldDefs.Add('iCampo', ftInteger, 0, True);
      Active:=true;
    end;

    with mDatosDet do
    begin
      Active:=false;
      FieldDefs.Add('sContrato', ftString, 100, false);
      FieldDefs.Add('sIdPernocta', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 250, false);
      FieldDefs.Add('sMedida', ftString, 100, false);
      FieldDefs.Add('sPadre', ftString, 2, false);
      FieldDefs.Add('dCantidad' , FtFloat, 0, false);
      Active:=true;
    end;

    mImprime.EmptyTable;
    if (TImpresion=FtsCPernoctas) or (TImpresion=FtsAll) then
    begin
     {FOLIOS DEL DIA}
      {$REGION 'CONSULTA COLUMNAS - FOLIOS LABORADOS'}

      QrColumnas.active:=false;
      QrColumnas.SQL.Add( 'SELECT ot.sIdFolio, '+
                           'ot.sNumeroOrden, '+
                           'ot.sIdPernocta, '+
                           'ot.sIdPlataforma, '+
                           'p.sDescripcion as sPernocta, '+
                           'pf.sDescripcion as sPlataforma, '+
                           'bp.sTipoObra as sTipo '+
                          'FROM ordenesdetrabajo ot ' +
                          'INNER JOIN contratos AS c ON (ot.sContrato=c.sContrato) ' +
                          'INNER JOIN bitacoradepersonal AS bp ON (bp.scontrato=:OT AND bp.sNumeroOrden = ot.sNumeroOrden ) ' +
                          'INNER JOIN pernoctan AS p ON (p.sidPernocta=bp.sIdpernocta) ' +
                          'INNER JOIN plataformas AS pf ON (pf.sidPlataforma=bp.sIdPlataforma) ' +
                          'WHERE (c.sContrato=:OT OR c.sCodigo=:OT)	AND bp.dIdFecha= :Fecha ' +
                          'GROUP BY	ot.sContrato, ot.sNumeroorden, bp.sidPernocta, bp.sIdPlataforma, bp.sTipoObra '+
                          'order by ot.sContrato, ot.iOrden ASC, bp.sTipoObra  DESC');
      QrColumnas.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
      QrColumnas.ParamByName('Fecha').AsDate      := Reportediario.FieldByName('dIdFecha').AsDateTime;
      QrColumnas.Open;
      {$ENDREGION}

      {$region 'Pernocta c-3'}
      with qrPernoctas do
      begin
        active := false;
        // Cambios realizados por E. Rangel - 20/08/2016
        sql.text := 'select ' +
                    '  (sum(xround(bp.dCantHH, 2)) + ' +
                    '   IFNULL((SELECT ' +
                    '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                    '          FROM ' +
                    '            bitacoradeajustes ba ' +
                    '          WHERE ' +
                    '            ba.sContrato = bp.sContrato AND ' +
                    '            ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                    '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                    '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                    '            ba.dFecha = bp.dIdFecha), 0)) as dCantidad, ' +
                    '  bp.sNumeroOrden, bp.sIdPlataforma, ' +
                    '  ifnull(SUM(CAST(bp.dAjuste AS CHAR)),0) AS Ajuste ' +
                    'from ' +
                    '  bitacoradepersonal bp ' +
                    'where ' +
                    '  bp.sContrato =:Contrato and bp.dIdFecha =:Fecha and ' +
                    '  bp.sNumeroOrden=:Folio and bp.sTipoObra =:Tipo and ' +
                    '  bp.lImprime = "Si" and bp.lAplicaPernocta= "Si" ' +
                    'group by ' +
                    '  bp.dIdFecha, bp.sIdPersonal';
        {sql.text := 'select xround(sum( bp.dCantHH),2) as dCantidad, bp.sNumeroOrden, bp.sIdPlataforma, '+
                    'ifnull(SUM(bp.dAjuste),0) AS Ajuste ' +
                    'from bitacoradepersonal bp '+
                    'where bp.sContrato =:Contrato '+
                    'and bp.dIdFecha =:Fecha '+
                    'and bp.sNumeroOrden=:Folio and bp.sTipoObra =:Tipo and bp.lImprime = "Si" and bp.lAplicaPernocta= "Si" '+
                    'group by bp.dIdFecha, bp.sIdPersonal ';}
        parambyname('contrato').asstring      := ReporteDiario.FieldByName('sOrden').asstring;
        parambyname('fecha').asdate           := ReporteDiario.FieldByName('dIdFecha').asDatetime;

      end;
      {$endregion}

      //Primero Consultamos los tipos de pernocta
      QrReportado.Active:=false;
      QrReportado.SQL.Text:='select sIdCuenta, sDescripcion, sMedida, sIdPadre from cuentas';
      QrReportado.Open;

      while not QrReportado.Eof do
      begin
          mDatosDet.Append;
          mDatosDet.FieldByName('sContrato').AsString    := global_Contrato;
          mDatosDet.FieldByName('sIdPernocta').AsString  := QrReportado.FieldByName('sIdCuenta').AsString;
          mDatosDet.FieldByName('sdescripcion').AsString := QrReportado.FieldByName('sDescripcion').AsString;
          mDatosDet.FieldByName('sMedida').AsString      := QrReportado.FieldByName('sMedida').AsString;
          mDatosDet.FieldByName('sPadre').AsString       := QrReportado.FieldByName('sIdPadre').AsString;
          mDatosDet.FieldByName('dCantidad').AsFloat     := 0;
          QrReportado.Next;
      end;

      //Ahora consultamos las categorias del reporte..
      QrReportado.Active:=false;
      QrReportado.SQL.Text:='select * from categorias where lPersonalAnexo ="Si" order by iOrden';
      QrReportado.Open;

      QrAdicional.Active:=false;
      {QrAdicional.SQL.Text:='select ifnull(sum(dCantidad),0) as dCantidad from bitacoradepernocta where ' +
                            'sContrato=:Contrato and dIdFecha=:fecha and sNumeroOrden=:Folio and '+
                            'sIdCategoria =:Categoria  group by sContrato';}

      QrAdicional.SQL.Text:='select ' +
                            'ifnull(sum(bpr.dCantidad),0) AS XXX, ' +
                            '  (ifnull(sum(bpr.dCantidad),0) + ' +
                            '   IFNULL((SELECT ' +
                            '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                            '          FROM ' +
                            '            bitacoradeajustes ba ' +
                            '          WHERE ' +
                            '            ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                            '            ba.sIdCategoria = bpr.sIdCategoria AND ' +
                            '            ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0)) as dCantidad ' +
                            'from ' +
                            '  bitacoradepernocta bpr ' +
                            'where ' +
                            '  bpr.sContrato=:Contrato and ' +
                            '  bpr.dIdFecha=:fecha and ' +
                            '  bpr.sNumeroOrden=:Folio and ' +
                            '  bpr.sIdCategoria =:Categoria ' +
                            'group by ' +
                            '  bpr.sContrato';

      DCiclo:=QrColumnas.RecordCount/TotalColPer;
      iCiclo:=Trunc(DCiclo);
      if (dCiclo -iCiclo)>0 then
        Inc(iCiclo,1);


      iGrupo:=1;
      while iGrupo<=iCiclo do
      begin
        QrReportado.First;
        while not QrReportado.Eof do
        begin
          dTotal:=0;
          if iGrupo=1 then
            QrColumnas.First
          else
            QrColumnas.RecNo:=((iGrupo-1) * TotalColPer)+ 1;
          //Catgorias de Catalogo PEP
          mDatos.Append;
          mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
          mDatos.FieldByName('sidrecurso').AsString   :=QrReportado.FieldByName('sIdCategoria').asstring;
          mDatos.FieldByName('sdescripcion').AsString :=QrReportado.fieldbyname('sdescripcionpernocta').asstring;
          mDatos.FieldByName('smedida').AsString      := '';

          CuentaCol:=1;
          ValTmp   :=0;
          CantTmp  :=0;
          iPosTmp  :=0;
          sTdPuConstante := qrReportado.FieldByName('sTdPu').AsString;
          while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalColPer)) do
          begin
            //Asiganar folio y Plataforma
            mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString := qrColumnas.FieldbyName('sIdFolio').AsString;
            mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  := QrColumnas.Fieldbyname('splataforma').Asstring;
            mDatos.FieldByName('sPernocta'+ Inttostr(CuentaCol)).AsString    := QrColumnas.Fieldbyname('sPernocta').Asstring;
            mDatos.FieldByName('sTipo'+ Inttostr(CuentaCol)).AsString        := QrColumnas.Fieldbyname('sTipo').Asstring;

            //Consultar bitacoradepernoctas- Directas
            if ((sTdPuConstante <> 'PU') and (sTdPuConstante <> 'TD')) and (qrColumnas.FieldbyName('sTipo').AsString = 'TD') then
            begin
                QrAdicional.Active:=false;
                QrAdicional.ParamByName('Contrato').AsString  := ReporteDiario.FieldByName('sOrden').asstring;
                QrAdicional.ParamByName('Folio').AsString     := qrColumnas.FieldbyName('snumeroorden').AsString;
                QrAdicional.ParamByName('Fecha').Asdate       := ReporteDiario.FieldByName('dIdFecha').asDatetime;
                QrAdicional.ParamByName('Categoria').AsString := QrReportado.FieldByName('sIdCategoria').AsString;
                QrAdicional.Open;

                if QrAdicional.RecordCount = 1 then
                begin
                    mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:= xRound(QrAdicional.FieldByName('dCantidad').value,2);
                    dTotal:=dTotal+ xRound(QrAdicional.FieldByName('dCantidad').Value,2);

                    ValTmp:= ValTmp + (QrAdicional.FieldByName('dCantidad').Value);

                    if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                    begin
                        CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                        iPosTmp:= CuentaCol;
                    end;

                    if QrReportado.FieldByName('sTdPU').AsString = '3C' then
                       mDatos.FieldByName('dcantidad3C' + Inttostr(CuentaCol)).AsFloat:= mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;

                end
                else
                begin
                    mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat   := 0;
                    mDatos.FieldByName('dcantidad3C' + Inttostr(CuentaCol)).AsFloat := 0;
                    dTotal:=dTotal+ xRound(0,2);
                    ValTmp:= ValTmp + 0;
                    if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                    begin
                       CantTmp :=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                       iPosTmp := CuentaCol;
                    end;

                end;
            end
            else
            begin
                if sTdPuConstante = qrColumnas.FieldByName('sTipo').AsString then
                begin
                    QrPernoctas.Active:=false;
                    QrPernoctas.ParamByName('Folio').AsString := qrColumnas.FieldbyName('snumeroorden').AsString;
                    QrPernoctas.ParamByName('Tipo').AsString  := qrColumnas.FieldByName('sTipo').AsString;
                    QrPernoctas.Open;

                    if QrPernoctas.RecordCount > 0 then
                    begin
                        while not QrPernoctas.Eof do
                        begin
                            mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=  mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat + (QrPernoctas.FieldByName('dCantidad').value + QrPernoctas.FieldByName('Ajuste').Value);
                            dTotal:=dTotal+ (QrPernoctas.FieldByName('dCantidad').Value + QrPernoctas.FieldByName('Ajuste').Value);

                            ValTmp:= ValTmp + (QrPernoctas.FieldByName('dCantidad').Value + QrPernoctas.FieldByName('Ajuste').Value);

                            if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                            begin
                                CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                                iPosTmp:= CuentaCol;
                            end;

                            mDatos.FieldByName('dcantidad3C' + Inttostr(CuentaCol)).AsFloat:= mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                            QrPernoctas.Next;
                        end;
                    end
                    else
                    begin
                        mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat   := 0;
                        mDatos.FieldByName('dcantidad3C' + Inttostr(CuentaCol)).AsFloat := 0;
                        dTotal:=dTotal+ xRound(0,2);
                        ValTmp:= ValTmp + 0;
                        if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                        begin
                           CantTmp :=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                           iPosTmp := CuentaCol;
                        end;
                    end;
                end;
            end;

            Inc(CuentaCol);
            QrColumnas.Next;
          end;

          if dTotal<>xRound(ValTmp,2) then
          begin
              mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
              + (xRound(ValTmp,2)-dTotal);
              mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2);

              mDatos.FieldByName('dcantTotal3C').AsFloat := 0;
              if (QrReportado.FieldByName('sTdPU').AsString = '3C') or (QrReportado.FieldByName('sTdPU').AsString = 'PU') or (QrReportado.FieldByName('sTdPU').AsString = 'TD') then
              begin
                  mDatos.FieldByName('dcantidad3C' + Inttostr(iPosTmp)).AsFloat := mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat;
                  mDatos.FieldByName('dcantTotal3C').AsFloat:=xRound(ValTmp,2);
              end;
          end
          else
          begin
              mDatos.FieldByName('dcantTotal').AsFloat   := dTotal;

              mDatos.FieldByName('dcantTotal3C').AsFloat := 0;
              if (QrReportado.FieldByName('sTdPU').AsString = '3C') or (QrReportado.FieldByName('sTdPU').AsString = 'PU') or (QrReportado.FieldByName('sTdPU').AsString = 'TD') then
                 mDatos.FieldByName('dcantTotal3C').AsFloat := dTotal;
          end;
          mDatos.Post;

          //Aqui acumulamos las pernoctas
           if QrReportado.FieldByName('sTdPu').AsString='' then
            mDatosDet.Locate('sPadre',QrReportado.FieldByName('sTdPu').AsString,[])
          else
            mDatosDet.Locate('sPadre',QrReportado.FieldByName('sTdPu').AsString,[]);

          mDatosDet.Edit;
          mDatosDet.FieldByName('dCantidad').AsFloat:= mDatosDet.FieldByName('dCantidad').AsFloat + mDatos.FieldByName('dcantTotal').AsFloat;
          mDatosDet.Post;

          QrReportado.next;
        end;
        Inc(iGrupo,1);
      end;

      mImprime.Append;
      mImprime.FieldByName('iCampo').AsInteger:=1;
      mImprime.Post;

    end;

    Td_CDistribucionDetPer.DataSet :=mDatosDet;//QrReportado;
    Td_CDistribucionDetPer.FieldAliases.Clear;
    Td_CDistribucionPer.DataSet    :=MDatos;
    Td_CDistribucionPer.FieldAliases.Clear;
    Td_CImpDistribucionPer.DataSet :=MImprime;
    Td_CImpDistribucionPer.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_CDistribucionPer);
    Reporte.DataSets.Add(Td_CImpDistribucionPer);
    Reporte.DataSets.Add(Td_CDistribucionDetPer);
  finally
    QrPernoctas.Destroy;
    //QrReportado.Destroy;
    QrColumnas.Destroy;
  end;

end;

Procedure ReportePDF_TotalDePersonal2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
var
  qrPersonal, qryPersonalSol, qryPersonalGral, qryPersonalPep, qrNotas   : TZReadOnlyQuery;
  Td_ResumenPersonal,
  Td_NotasResumenP,
  Td_ResumenPersonalSol,
  Td_ResumenPersonalGral,
  Td_ResumenPersonalPep,
  Td_ImpResumenPersonal : TfrxDBDataset;
  rxPersonal : TClientDataSet;
begin
    qrPersonal := TZReadOnlyQuery.Create(nil);
    qrPersonal.connection := connection.zconnection;
    qryPersonalPep := TZReadOnlyQuery.Create(nil);
    qryPersonalPep.connection := connection.zconnection;
    qryPersonalSol := TZReadOnlyQuery.Create(nil);
    qryPersonalSol.connection := connection.zconnection;
    qryPersonalGral := TZReadOnlyQuery.Create(nil);
    qryPersonalGral.connection := connection.zconnection;
    qrNotas := TZReadOnlyQuery.Create(nil);
    qrNotas.connection := connection.zConnection;

    Td_ResumenPersonal   := TfrxDBDataset.Create(nil);
    Td_ResumenPersonal.UserName :='Td_ResumenPersonal';
    Td_ResumenPersonalSol   := TfrxDBDataset.Create(nil);
    Td_ResumenPersonalSol.UserName :='Td_ResumenPersonalSol';
    Td_ResumenPersonalGral   := TfrxDBDataset.Create(nil);
    Td_ResumenPersonalGral.UserName :='Td_ResumenPersonalGral';
    Td_NotasResumenP     := TfrxDBDataset.Create(nil);
    Td_NotasResumenP.UserName   :='Td_NotasResumenP';
    Td_ImpResumenPersonal:= TfrxDBDataset.Create(nil);
    Td_ImpResumenPersonal.UserName:='Td_ImpResumenPersonal';
    Td_ResumenPersonalPep   := TfrxDBDataset.Create(nil);
    Td_ResumenPersonalPep.UserName :='Td_ResumenPersonalPep';

    rxPersonal := TClientDataSet.Create(nil);
    rxPersonal.Active := False;
    rxPersonal.FieldDefs.Add('sDescripcionTipo', ftString, 150, False);
    rxPersonal.FieldDefs.Add('sContrato', ftString, 20, False);
    rxPersonal.FieldDefs.Add('dCantidad', ftFloat, 0, False);
    rxPersonal.CreateDataSet;
    rxPersonal.Active := True;

    rxPersonal.Append;
    rxPersonal.FieldByName('sDescripcionTipo').AsString := '';
    rxPersonal.FieldByName('sContrato').AsString := '';
    rxPersonal.FieldByName('dCantidad').AsFloat := 0;
    rxPersonal.Post;

    if (TImpresion=FtsTPersonal) or (TImpresion=FtsAll) then
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select dFechaVigencia from categorias where dFechaVigencia <=:fecha ');
        connection.zCommand.ParamByName('fecha').AsDate   := Date;
        connection.zCommand.Open;

        //Personal Normal
        qrPersonal.Active := False;
        qrPersonal.SQL.Add( 'select tp.sIdTipoPersonal, td.sIdTripulacion, t.sDescripcion, tp.sDescripcion as titulo, td.iNacionales, td.iExtranjeros, c.sLabelContrato, lPersonalEQ '+
                  'from tripulaciondiaria td '+
                  'inner join tiposdepersonal tp on (tp.sIdTipoPersonal = td.sIdCategoria and sLabelTotal = "NORMAL") '+
                  'inner join tripulacion t on (t.dFechaVigencia =:Vigencia and t.sIdCategoria = tp.sIdTipoPersonal and t.sIdTripulacion = td.sIdTripulacion) '+
                  'inner join contratos c on (c.sContrato =:Contrato) '+
                  'where td.sContrato =:ContratoBarco and td.dIdFecha =:Fecha '+
                  'order by tp.iOrden, t.iOrden ');
        qrPersonal.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        qrPersonal.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
        qrPersonal.ParamByName('ContratoBarco').AsString := global_Contrato_Barco;
        qrPersonal.ParamByName('Vigencia').AsDate        := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
        qrPersonal.Open;

        //personal de pemex
        qryPersonalPep.Active := False;

        qryPersonalPep.SQL.Add('select ' +
                               '  c.sDescripcion, ' +
                               '  (sum(b.dCantidad) + ' +
                               '   IFNULL((SELECT ' +
                               '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                               '           FROM ' +
                               '             bitacoradeajustes ba ' +
                               '           WHERE ' +
                               '             ba.sContrato = b.sContrato AND ba.sNumeroOrden = b.sNumeroOrden AND ' +
                               '             ba.sIdPartidaAnexo = b.sIdCuenta AND ba.dFecha = b.dIdFecha), 0)) as cantidad, ' +
                               '  c.sMiGrupoResumen ' +
                               'from ' +
                               '  bitacoradepernocta b ' +
                               'inner join ' +
                               '  categorias c ' +
                               '    on (c.sIdCategoria = b.sIdCategoria and ' +
                               '        c.dFechaVigencia = :Vigencia and ' +
                               '        c.sAgrupadorFinal <> "N/A") ' +
                               'where ' +
                               '  b.sContrato =:Contrato and ' +
                               '  b.dIdFecha =:Fecha ' +
                               'group by ' +
                               '  c.sIdCategoria ' +
                               'order by ' +
                               '  c.iOrden');

        {qryPersonalPep.SQL.Add('select c.sDescripcion, sum(b.dCantidad) as cantidad, c.sMiGrupoResumen from bitacoradepernocta b '+
                        'inner join categorias c on (c.sIdCategoria = b.sIdCategoria and c.dFechaVigencia = :Vigencia and c.sAgrupadorFinal <> "N/A") '+
                        'where b.sContrato =:Contrato and b.dIdFecha =:Fecha '+
                        'group by c.sIdCategoria '+
                        'order by c.iOrden');}
        qryPersonalPep.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        qryPersonalPep.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
        qryPersonalPep.ParamByName('Vigencia').AsDate        := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
        qryPersonalPep.Open;

        //personal solcitado
        qryPersonalSol.Active := False;
        qryPersonalSol.SQL.Add('select tp.sIdTipoPersonal, td.sIdTripulacion, t.sDescripcion, tp.sDescripcion as titulo, td.iNacionales, td.iExtranjeros, c.sLabelContrato, lPersonalEQ '+
                  'from tripulaciondiaria td '+
                  'inner join tiposdepersonal tp on (tp.sIdTipoPersonal = td.sIdCategoria and sLabelTotal = "SOLICITADO") '+
                  'inner join tripulacion t on (t.dFechaVigencia =:Vigencia and t.sIdCategoria = tp.sIdTipoPersonal and t.sIdTripulacion = td.sIdTripulacion) '+
                  'inner join contratos c on (c.sContrato =:Contrato) '+
                  'where td.sContrato =:ContratoBarco and td.dIdFecha =:Fecha '+
                  'order by tp.iOrden, t.iOrden');
        qryPersonalSol.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        qryPersonalSol.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
        qryPersonalSol.ParamByName('ContratoBarco').AsString := global_Contrato_Barco;
        qryPersonalSol.ParamByName('Vigencia').AsDate        := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
        qryPersonalSol.Open;

        //Personal de tripulacion y compañia
        qryPersonalGral.Active := False;
        qryPersonalGral.SQL.Clear;
        qryPersonalGral.SQL.Add('select td.sIdTripulacion, t.sDescripcion, tp.sDescripcion as titulo, SUM(td.iNacionales) as iNacionales, td.iExtranjeros, c.sLabelContrato, lPersonalEQ '+
                      'from tripulaciondiaria td '+
                      'inner join tiposdepersonal tp on (tp.sIdTipoPersonal = td.sIdCategoria and lPernocta = "No" and lPersonalEQ = "No" ) '+
                      'inner join tripulacion t on (t.dFechaVigencia =:Vigencia and t.sIdCategoria = tp.sIdTipoPersonal and t.sIdTripulacion = td.sIdTripulacion) '+
                      'inner join contratos c on (c.sContrato =:Contrato) '+
                      'where td.sContrato =:ContratoBarco and td.dIdFecha =:Fecha ' +
                      'group by tp.sIdTipoPersonal order by tp.iOrden');
        qryPersonalGral.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        qryPersonalGral.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
        qryPersonalGral.ParamByName('ContratoBarco').AsString := global_Contrato_Barco;
        qryPersonalGral.ParamByName('Vigencia').AsDate        := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
        qryPersonalGral.Open;

        qryPersonalGral.First;
        while not qryPersonalGral.Eof do
        begin
            rxPersonal.Append;
            rxPersonal.FieldByName('sDescripcionTipo').AsString := qryPersonalGral.FieldByName('titulo').AsString;
            rxPersonal.FieldByName('sContrato').AsString        := global_Contrato_Barco;
            rxPersonal.FieldByName('dCantidad').AsFloat         := qryPersonalGral.FieldByName('iNacionales').AsFloat;
            rxPersonal.Post;
            qryPersonalGral.Next;
        end;

        //Personal de PEMEX
        qryPersonalGral.Active := False;
        qryPersonalGral.SQL.Clear;

        qryPersonalGral.SQL.Add('select ' +
                                '  c.sDescripcion, ' +
                                '  (sum(b.dCantidad) + ' +
                                '   IFNULL((SELECT ' +
                                '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '           FROM ' +
                                '             bitacoradeajustes ba ' +
                                '           WHERE ' +
                                '             ba.sContrato = b.sContrato AND ba.sNumeroOrden = b.sNumeroOrden AND ' +
                                '             ba.sIdCategoria = b.sIdCategoria AND ' +
                                '             ba.sIdPartidaAnexo = b.sIdCuenta AND ba.dFecha = b.dIdFecha), 0)) as cantidad, ' +
                                '  c.sAgrupadorFinal ' +
                                'from ' +
                                '  bitacoradepernocta b ' +
                                'inner join ' +
                                '  categorias c ' +
                                '    on (c.sIdCategoria = b.sIdCategoria and ' +
                                '        c.dFechaVigencia = :Vigencia and c.sAgrupadorFinal <> "N/A") ' +
                                'where ' +
                                '  b.sContrato =:Contrato and ' +
                                '  b.dIdFecha =:Fecha ' +
                                'group by ' +
                                '  c.sAgrupadorFinal ' +
                                'order by ' +
                                '  c.iOrden');

          {qryPersonalGral.SQL.Add( 'select c.sDescripcion, sum(b.dCantidad) as cantidad, c.sAgrupadorFinal from bitacoradepernocta b '+
                        'inner join categorias c on (c.sIdCategoria = b.sIdCategoria and c.dFechaVigencia = :Vigencia and c.sAgrupadorFinal <> "N/A") '+
                        'where b.sContrato =:Contrato and b.dIdFecha =:Fecha '+
                        'group by c.sAgrupadorFinal '+
                        'order by c.iOrden');}
        qryPersonalGral.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        qryPersonalGral.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
        qryPersonalGral.ParamByName('Vigencia').AsDate        := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
        qryPersonalGral.Open;

        qryPersonalGral.First;
        while not qryPersonalGral.Eof do
        begin
            rxPersonal.Append;
            rxPersonal.FieldByName('sDescripcionTipo').AsString := qryPersonalGral.FieldByName('sAgrupadorFinal').AsString;
            rxPersonal.FieldByName('sContrato').AsString        := global_Contrato_Barco;
            rxPersonal.FieldByName('dCantidad').AsFloat         := qryPersonalGral.FieldByName('cantidad').AsFloat;
            rxPersonal.Post;
            qryPersonalGral.Next;
        end;

        //Personal de contruccion
        qryPersonalGral.Active := False;
        qryPersonalGral.SQL.Clear;
        qryPersonalGral.SQL.Add( 'select td.sIdTripulacion, t.sDescripcion, tp.sDescripcion as titulo, SUM(td.iNacionales) as iNacionales, td.iExtranjeros, c.sLabelContrato, lPersonalEQ '+
                      'from tripulaciondiaria td '+
                      'inner join tiposdepersonal tp on (tp.sIdTipoPersonal = td.sIdCategoria and lPernocta = "No" and lPersonalEQ = "Si" ) '+
                      'inner join tripulacion t on (t.dFechaVigencia =:Vigencia and t.sIdCategoria = tp.sIdTipoPersonal and t.sIdTripulacion = td.sIdTripulacion) '+
                      'inner join contratos c on (c.sContrato =:Contrato) '+
                      'where td.sContrato =:ContratoBarco and td.dIdFecha =:Fecha ' +
                      'group by tp.sIdTipoPersonal order by tp.iOrden');
        qryPersonalGral.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        qryPersonalGral.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
        qryPersonalGral.ParamByName('ContratoBarco').AsString := global_Contrato_Barco;
        qryPersonalGral.ParamByName('Vigencia').AsDate        := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
        qryPersonalGral.Open;

        qryPersonalGral.First;
        while not qryPersonalGral.Eof do
        begin
            rxPersonal.Append;
            rxPersonal.FieldByName('sDescripcionTipo').AsString := qryPersonalGral.FieldByName('titulo').AsString;
            rxPersonal.FieldByName('sContrato').AsString        := global_Contrato_Barco;
            rxPersonal.FieldByName('dCantidad').AsFloat         := qryPersonalGral.FieldByName('iNacionales').AsFloat;
            rxPersonal.Post;
            qryPersonalGral.Next;
        end;

        qrNotas.SQL.Text := 'select '+
                              'sNotaGeneral as mDescripcion '+
                            'from notas_generales '+
                            'where '+
                              'sContrato = :orden and '+
                              'dIdFecha = :fecha and '+
                              'eAplicaResumenPersonal = "Si"; ';
        qrNotas.ParamByName( 'orden' ).AsString := global_contrato;
        qrNotas.ParamByName( 'fecha' ).AsDateTime := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        qrNotas.Open;

        rxPersonal.First;
        rxPersonal.Delete;

        Td_ImpResumenPersonal.DataSet := qrPersonal;
        Td_NotasResumenP.DataSet      := qrNotas;

        {Esta secccion es para mostrar la hoja vacia sino existen datos.}
        if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
           if qrPersonal.RecordCount = 0 then
              Td_ImpResumenPersonal.DataSet := rxPersonal;
    end
    else
    begin
        rxPersonal.Delete;
        Td_ResumenPersonal.DataSet   := rxPersonal;
        Td_ImpResumenPersonal.DataSet:= rxPersonal;
    end;

    Td_NotasResumenP.DataSet:= qrNotas;
    Td_NotasResumenP.FieldAliases.Clear;

    Td_ResumenPersonal.DataSet:= qrPersonal;
    Td_ResumenPersonal.FieldAliases.Clear;

    Td_ResumenPersonalSol.DataSet:= QryPersonalSol;
    Td_ResumenPersonalSol.FieldAliases.Clear;

    Td_ResumenPersonalGral.DataSet:= rxPersonal;
    Td_ResumenPersonalGral.FieldAliases.Clear;

    Td_ImpResumenPersonal.DataSet:= rxPersonal;
    Td_ImpResumenPersonal.FieldAliases.Clear;

    Td_ResumenPersonalPep.DataSet:= qryPersonalPep;
    Td_ResumenPersonalPep.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_ResumenPersonal);
    Reporte.DataSets.Add(Td_ResumenPersonalSol);
    Reporte.DataSets.Add(Td_ResumenPersonalGral);
    Reporte.DataSets.Add(Td_NotasResumenP);
    Reporte.DataSets.Add(Td_ResumenPersonalPep);
end;

Procedure ReportePDF_ResumenMaterial2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
Var
   QryResumenMaterial: TZQuery;
   lInsertar : boolean;
   sFolio    : string;
   Td_resumenMaterial,
   Td_ImpResumenMaterial  : TfrxDBDataset;
   rxResumenMaterial: TClientDataSet;
begin
  Try
    QryResumenMaterial := TZQuery.Create(nil);
    QryResumenMaterial.Connection := Connection.zConnection;

    Td_resumenMaterial:= TfrxDBDataset.Create(nil);
    Td_resumenMaterial.UserName:='Td_resumenMaterial';

    Td_ImpResumenMaterial:= TfrxDBDataset.Create(nil);
    Td_ImpResumenMaterial.UserName:='Td_ImpResumenMaterial';

    rxResumenMaterial := TClientDataSet.Create(nil);
    rxResumenMaterial.Active := False;
    rxResumenMaterial.FieldDefs.Add('sNumeroOrden', ftString, 0, False);
    rxResumenMaterial.FieldDefs.Add('sNumeroActividad', ftString, 0, False);
    rxResumenMaterial.FieldDefs.Add('sTrazabilidad', ftString, 0, False);
    rxResumenMaterial.FieldDefs.Add('dCantidad', ftFloat, 0, False);
    rxResumenMaterial.FieldDefs.Add('sMedida', ftString, 0, False);
    rxResumenMaterial.FieldDefs.Add('sIdProveedor', ftString, 0, False);
    rxResumenMaterial.FieldDefs.Add('mDescripcion', ftMemo, 0, False);
    rxResumenMaterial.CreateDataSet;
    rxResumenMaterial.Active := True;

    if (TImpresion=FtsRMaterial) or (TImpresion=FtsAll) then
    begin
        {$REGION 'CONSULTA MATERIALES'}

          QryResumenMaterial.SQL.Add( 'select s.sNumeroOrden, b.sNumeroActividad, b.sTrazabilidad, i.mDescripcion, i.sMedida, sum(b.dCantidad) as dCantidad, i.sIdProveedor from almacen_salida s '+
                            'inner join bitacoradesalida b '+
                            'on(b.sContrato = s.sContrato and b.iFolioSalida = s.iFolioSalida and b.sext = s.sext) '+
                            'inner join insumos i '+
                            'on (i.sContrato =:Contrato and i.sIdInsumo = b.sIdInsumo and i.sTrazabilidad = b.sTrazabilidad ) '+
                            'where s.sContrato =:orden and b.dFechaSalida =:fecha '+
                            'group by b.sNumeroOrden, b.sNumeroActividad, b.sIdInsumo, b.sTrazabilidad Order by b.sNumeroOrden, b.sNumeroActividad, b.sTrazabilidad');
          QryResumenMaterial.ParamByName('Fecha').AsDateTime  := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
          QryResumenMaterial.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sOrden').AsString;
          QryResumenMaterial.ParamByName('Contrato').AsString := global_Contrato_Barco;
          QryResumenMaterial.Open;

        {$ENDREGION}
        Td_resumenMaterial.DataSet:= QryResumenMaterial;
        Td_ImpResumenMaterial.DataSet:= QryResumenMaterial;

        {Esta secccion es para mostrar la hoja vacia sino existen datos.}
        if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
        begin
            if QryResumenMaterial.RecordCount = 0 then
            begin
               rxResumenMaterial.Append;
               rxResumenMAterial.FieldValues['sNumeroOrden'] := '';
               rxResumenMaterial.Post;

               Td_resumenMaterial.DataSet   := rxResumenMaterial;
               Td_ImpResumenMaterial.DataSet:= rxResumenMaterial;
            end;
        end;
    end
    else
    begin
        {Solo se envian los campos del datset..}
        Td_resumenMaterial.DataSet   := rxResumenMaterial;
        Td_ImpResumenMaterial.DataSet:= rxResumenMaterial;
    end;

    Td_resumenMaterial.FieldAliases.Clear;
    Td_ImpResumenMaterial.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_resumenMaterial);
    Reporte.DataSets.Add(Td_ImpResumenMaterial);
  Finally

  End;
end;

Procedure ReportePDF_ConsumosDiesel(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
begin
  Try

  Finally

  End;
end;

Procedure ReportePDF_Listadepersonal(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
Var
  lista_personal: TZQuery;
  sCategoria : string;
  mImprime:TJvMemoryData;
  Td_CImpListadoPer: TfrxDBDataset;
  Td_CListadoPer: TfrxDBDataset;
begin
  lista_personal := TZQuery.Create(nil);
  mImprime:=TJvMemoryData.Create(nil);
  Td_CListadoPer:=TfrxDBDataset.Create(nil);
  Td_CImpListadoPer:=TfrxDBDataset.Create(nil);
  Try
    {$REGION 'CONSULTA'}

    Td_CListadoPer.UserName:='Td_CListadoPer';
    Td_CImpListadoPer.UserName:='Td_CImpListadoPer';

    lista_personal.Connection := Connection.zConnection;
    lista_personal.Active := False;
    lista_personal.Sql.Clear;
    lista_personal.Sql.Add('select td.*, cat.sDescripcion as categoria,cp.sdescripcion as compania, td.sIdPernocta, e.sDescripcion as barco '+
                  'from tripulaciondiaria_listado td '+
                  'inner join categorias cat on (cat.sIdCategoria = td.sIdCategoria and cat.dFechaVigencia <= td.dIdFecha) '+
                  'left join compersonal cp on(cp.sIdCompania=td.sIdCompania) '  +
                  'inner join embarcaciones e '+
                  'on (e.sIdEmbarcacion = td.sIdPernocta) '+
                  'where td.sContrato=:contrato and td.sOrden =:Orden and td.dIdFecha=:fecha and td.lImprimeListado="Si" and td.sIdTurno =:turno '+
                  'order by e.iOrden, cat.sIdCategoria, td.sOrden, td.sDescripcion,td.snumerocabina desc ');

    with mImprime do
    begin
      Active:=false;
      FieldDefs.Add('iCampo', ftInteger, 0, True);
      Active:=true;
    end;

    mImprime.EmptyTable;
    if (TImpresion=FtsLPersonal) or (TImpresion=FtsAll) then
    begin
      lista_personal.Active := False;
      lista_personal.ParamByName('Contrato').DataType := ftString;
      lista_personal.ParamByName('Contrato').Value    := global_Contrato_Barco;
      lista_personal.ParamByName('Orden').asString    := Reportediario.FieldByName('sOrden').AsString;
      lista_personal.ParamByName('Turno').DataType    := ftString;
      lista_personal.ParamByName('Turno').Value       := ReporteDiario.FieldByName('sIdTurno').AsString;
      lista_personal.ParamByName('Fecha').DataType    := ftDate;
      lista_personal.ParamByName('Fecha').Value       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;;
      lista_personal.Open;

      {Esta secccion es para mostrar la hoja vacia sino existen datos.}
      //if  then

          if (lista_personal.RecordCount > 0) or
          (connection.configuracion.FieldValues['eHojasBlanco'] = 'Si') then
          begin
            mImprime.Append;
            mImprime.FieldByName('iCampo').AsInteger:=1;
            mImprime.Post;
          end;
    end
    else
      lista_personal.Open;
    {$ENDREGION}
    Td_CListadoPer.DataSet:=lista_personal;
    Td_CListadoPer.FieldAliases.Clear;
    Td_CImpListadoPer.DataSet:=MImprime;
    Td_CImpListadoPer.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_CListadoPer);
    Reporte.DataSets.Add(Td_CImpListadoPer);


  Finally

  End;
end;

Procedure ReportePDF_DistribucionDeEquipos2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
var
  QrMoe,
  QrColumnas,
  QrRecursos,
  QryEmbarcacion: TZQuery;
  mDatos:TJvMemoryData;
  ZqEqBarcos : tzreadonlyquery;
  CuentaCol,iGrupo,iCiclo: Integer;
  DCiclo,dTotal:Double;
  Td_CDistribucionE:TfrxDBDataset;
  mImprime:TJvMemoryData;
  Td_CImpDistribucionE: TfrxDBDataset;
  ValTmp:Variant;
  CantTmp, dSubTotal, dSubTotalPU :Double;
  iPosTmp:Integer;
  lImprimeRecurso : string;
begin
  ZqEqBarcos := TZReadOnlyQuery.Create(nil);
  ZqEqBarcos.Connection := connection.zConnection;
  ZqEqBarcos.Active := False;
  ZqEqBarcos.SQL.Clear;
  ZqEqBarcos.SQL.Text := ' Select sidequipo,iunificador from equipos '+
                         ' where scontrato = :contrato and Length(iunificador) > 0  ';
  ZqEqBarcos.ParamByName('contrato').AsString := global_Contrato_Barco;
  ZqEqBarcos.Open;

  QrMoe := TZQuery.Create(nil);
  QrColumnas := TZQuery.Create(nil);
  QrRecursos := TZQuery.Create(nil);
  mDatos:=TJvMemoryData.Create(nil);
  mImprime:=TJvMemoryData.Create(nil);
  Td_CDistribucionE:=TfrxDBDataset.Create(nil);
  Td_CImpDistribucionE:=TfrxDBDataset.Create(nil);
  Try
    QrColumnas.Connection := Connection.zConnection;
    QrMoe.Connection := Connection.zConnection;
    QrRecursos.Connection := Connection.zConnection;
    Td_CDistribucionE.UserName:='Td_CDistribucionE';
    Td_CImpDistribucionE.UserName:='Td_CImpDistribucionE';
    with mDatos do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('iTotalGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidrecurso', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 350, True);
      FieldDefs.Add('smedida', ftString, 100, True);
      FieldDefs.Add('dcantSol', FtFloat, 0, True);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      FieldDefs.Add('dcantTotalPU', FtFloat, 0, True);
      for CuentaCol:=1 to TotalCol do
      begin
        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPernocta' + Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPlataforma'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sTipo'+ Inttostr(CuentaCol), ftString, 10, false);
      end;
      Active:=true;
    end;

    with mImprime do
    begin
      Active:=false;
      FieldDefs.Add('iCampo', ftInteger, 0, True);
      Active:=true;
    end;

    mImprime.EmptyTable;

    if (TImpresion=FtsCEquipos) or (TImpresion=FtsAll) then
    begin
      {$REGION 'Recursos'}
      QrRecursos.Active := False;
      QrRecursos.SQL.Clear;
      QrRecursos.SQL.Add( 'select be.sIdEquipo, (sum(xround(be.dCantHH, 2)) + ' +
                          '  IFNULL((SELECT ' +
                          '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                          '          FROM ' +
                          '            bitacoradeajustes ba ' +
                          '          WHERE ' +
                          '            ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                          '            ba.sNumeroActividad = be.sNumeroActividad AND ' +
                          '            ba.sIdPartidaAnexo = be.sIdEquipo AND ba.dFecha = be.dIdFecha), 0)) as Total, '+
                          'IFNULL(SUM(CAST(be.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                          'be.lImprime '+
                          'from bitacoradeequipos be '+
                          'inner join equipos e '+
                          'on ( e.sIdEquipo = be.sIdEquipo ) '+
                          'inner join bitacoradeactividades ba '+
                          'on ( ba.sContrato = :contrato and ba.dIdFecha = :fecha '+
                               'and ba.sNumeroOrden = :folio '+
                               'and ba.iIdDiario = be.iIdDiario '+
                          'and ba.iIdTarea = be.iIdTarea and ba.iIdActividad = be.iIdActividad) '+
                          'where e.sContrato = :contratoBarco '+
                          'and be.sContrato = :contrato '+
                          'and be.sNumeroOrden = :folio '+
                          'and be.sIdPernocta = :pernocta '+
                          'and be.dIdFecha = :fecha '+
                          'and be.sIdEquipo=:equipo and be.sTipoObra =:Tipo ' +
                          'group by e.sIdEquipo '+
                          'order by e.iItemOrden');
      {$endregion}

      {$REGION 'CONSULTAS - PARTIDAS'}
      QrColumnas.Active:=false;
      QrColumnas.SQL.Add( 'select ot.sContrato,ot.sIdFolio, '+
                          'ot.sNumeroOrden, be.sTipoObra as sTipo, '+
                          'be.sIdPernocta AS PernoctaP, '+
                          'ot.sIdPlataforma AS idPlataforma, '+
                          'p.sDescripcion as Pernocta, '+
                          'pt.sDescripcion as Plataforma, '+
                          'be.sTipoObra as sTipo '+
                          'from ordenesdetrabajo ot '+
                          'inner join contratos c '+
                          'on ( c.sContrato = ot.sContrato ) '+
                          'inner join bitacoradeequipos be '+
                          'on ( be.sContrato = ot.sContrato and be.sNumeroOrden = ot.sNumeroOrden ) '+
                          'inner join pernoctan p '+
                          'on ( ot.sIdPernocta = p.sIdPernocta) '+
                          'inner join plataformas pt '+
                          'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                          'where c.sContrato = :contrato '+
                          'and be.dIdFecha = :fecha '+
                          'group by ot.sNumeroOrden, be.sIdPernocta, be.sIdPlataforma, be.sTipoObra '+
                          'order by ot.sContrato, ot.iOrden ASC, be.sTipoObra  DESC' );
        QrColumnas.ParamByName('Contrato').AsString:= ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate:=reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;
      {$ENDREGION}


      {$REGION 'CONSULTAS - TODOS LOS EQUIPOS REGISTRADOS EN MOE'}

      QrMoe.Active := False;
      QrMoe.SQL.Clear;
      QrMoe.SQL.Add('select mr.sIdRecurso, e.sDescripcion, e.sMedida, mr.dCantidad '+
                    'from moe m '+
                    'inner join moerecursos mr '+
                    'on ( mr.iIdMoe = m.iIdMoe ) '+
                    'inner join equipos e '+
                    'on ( e.sContrato = :contratobarco and e.sIdEquipo = mr.sIdRecurso ) '+
                    'where m.dIdFecha = (select max(didfecha) from moe where didfecha <=:Fecha and sContrato = :contrato) '+
                    'and m.sContrato = :contrato '+
                    'and mr.eTipoRecurso = "Equipo" '+
                    'order by e.iItemOrden');
      QrMoe.ParamByName('contratobarco').AsString := Global_Contrato_Barco;
      QrMoe.ParamByName('contrato').AsString := ReporteDiario.FieldByName('sOrden').AsString;
      QrMoe.ParamByName('Fecha').AsDateTime := reportediario.FieldByName('dIdFecha').AsDateTime;
      QrMoe.Open;

      {$ENDREGION}

      {$REGION 'INSERCION DE DATOS - INFORMACIÓN DEL EQUIPO'}

      DCiclo:=QrColumnas.RecordCount/TotalCol;
      iCiclo:=Trunc(DCiclo);
      if (dCiclo -iCiclo)>0 then
        Inc(iCiclo,1);

      iGrupo:=1;
      dSubTotal := 0;
      dSubTotalPU := 0;
      while iGrupo<=iCiclo do
      begin
        with qrMoe do
        begin
          first;
          while not eof do
          begin
            if iGrupo > 1 then
            begin
               mDatos.Locate('sidrecurso',FieldByName('sidrecurso').asstring, []);
               dSubTotal   := mDatos.FieldByName('dcantTotal').AsFloat;
               dSubTotalPU := mDatos.FieldByName('dcantTotalPU').AsFloat;
            end;

            mDatos.Append;
            mDatos.FieldByName('iGrupo').AsInteger      :=Igrupo;
            mDatos.FieldByName('iTotalGrupo').AsInteger :=iCiclo;
            mDatos.FieldByName('sidrecurso').AsString   :=FieldByName('sidrecurso').asstring;
            mDatos.FieldByName('sdescripcion').AsString :=fieldbyname('sdescripcion').asstring;
            mDatos.FieldByName('smedida').AsString      :=fieldbyname('smedida').asstring;
            mDatos.FieldByName('dcantSol').AsFloat      :=fieldbyname('dcantidad').AsFloat;
            mDatos.FieldByName('dcantTotalPU').AsFloat  :=dSubTotalPU;
            mDatos.FieldByName('dCantTotal').AsFloat    :=dSubTotal;
            dTotal:=0;
            if iGrupo=1 then
              QrColumnas.First
            else
              QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

            CuentaCol:=1;
            ValTmp:=0;
            CantTmp:=0;
            iPosTmp:=0;
            lImprimeRecurso := 'Si';
            while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
            begin

              mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString := QrColumnas.FieldbyName('sIdFolio').AsString;
              mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString   := QrColumnas.Fieldbyname('Pernocta').asstring;
              mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  := QrColumnas.Fieldbyname('Plataforma').Asstring;
              mDatos.FieldByName('sTipo'+ Inttostr(CuentaCol)).AsString        := QrColumnas.Fieldbyname('sTipo').Asstring;

              QrRecursos.Active := False;
              QrRecursos.ParamByName('Equipo').AsString := FieldByName('sIdRecurso').AsString;

              QrRecursos.ParamByName('contratoBarco').AsString := global_contrato_barco;
              QrRecursos.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
              QrRecursos.ParamByName('folio').AsString         := QrColumnas.FieldByName('sNumeroOrden').AsString;
              QrRecursos.ParamByName('fecha').AsDate           := reportediario.FieldByName('dIdFecha').AsDateTime;
              QrRecursos.ParamByName('pernocta').AsString      := QrColumnas.FieldByName('PernoctaP').AsString;
              QrRecursos.ParamByName('Tipo').AsString          := QrColumnas.Fieldbyname('sTipo').Asstring;
              QrRecursos.Open;

              if QrRecursos.RecordCount>0 then
              begin
                mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat := (QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value);
                dTotal                                                        := dTotal+ (QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value);
                ValTmp                                                        := (ValTmp + QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value );
                if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                begin
                  CantTmp := mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                  iPosTmp := CuentaCol;
                end;

                if QrColumnas.Fieldbyname('sTipo').Asstring = 'TD' then
                   mDatos.FieldByName('dcantTotal').AsFloat   := mDatos.FieldByName('dcantTotal').AsFloat + mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat
                else
                   mDatos.FieldByName('dcantTotalPU').AsFloat := mDatos.FieldByName('dcantTotalPU').AsFloat + mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;

                lImprimeRecurso := QrRecursos.FieldByName('lImprime').AsString;
              end;
              QrColumnas.next;
              Inc(CuentaCol);
            end;

              if dTotal<>xRound(ValTmp,2) then
              begin
                 mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:= mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                 + (xRound(ValTmp,2)-dTotal);
              end;

              if lImprimeRecurso = 'Si' then
                 mDatos.Post
              else
                 mDatos.Cancel;

            next;
          end;
        end;
        Inc(iGrupo,1);
      end;


//      {Esta secccion es para mostrar la hoja vacia sino existen datos.}
//      if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
//         if MDatos.RecordCount >= 0 then
//      else
      if MDatos.RecordCount > 0 then
      begin
        mImprime.Append;
        mImprime.FieldByName('iCampo').AsInteger:=1;
        mImprime.Post;
      end;
      {$ENDREGION}
    end;
    Td_CDistribucionE.DataSet:=MDatos;
    Td_CDistribucionE.FieldAliases.Clear;
    Td_CImpDistribucionE.DataSet:=MImprime;
    Td_CImpDistribucionE.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_CDistribucionE);
    Reporte.DataSets.Add(Td_CImpDistribucionE);
  Finally
    QrMoe.Free;
    QrRecursos.Free;
    QrColumnas.Destroy;
  End;
end;

Procedure ReportePDF_DistribucionDeEquiposTierra(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
var
  QrMoe,
  QrColumnas,
  QrRecursos,
  QryEmbarcacion: TZQuery;
  mDatos:TJvMemoryData;
  ZqEqBarcos : tzreadonlyquery;
  CuentaCol,iGrupo,iCiclo: Integer;
  DCiclo,dTotal:Double;
  Td_CDistribucionE:TfrxDBDataset;
  mImprime:TJvMemoryData;
  Td_CImpDistribucionE: TfrxDBDataset;
  ValTmp:Variant;
  CantTmp:Double;
  iPosTmp:Integer;
begin
  ZqEqBarcos := TZReadOnlyQuery.Create(nil);
  ZqEqBarcos.Connection := connection.zConnection;
  ZqEqBarcos.Active := False;
  ZqEqBarcos.SQL.Clear;
  ZqEqBarcos.SQL.Text := ' Select sidequipo,iunificador from equipos '+
                         ' where scontrato = :contrato and Length(iunificador) > 0  ';
  ZqEqBarcos.ParamByName('contrato').AsString := global_Contrato_Barco;
  ZqEqBarcos.Open;

  QrMoe := TZQuery.Create(nil);
  QrColumnas := TZQuery.Create(nil);
  QrRecursos := TZQuery.Create(nil);
  mDatos:=TJvMemoryData.Create(nil);
  mImprime:=TJvMemoryData.Create(nil);
  Td_CDistribucionE:=TfrxDBDataset.Create(nil);
  Td_CImpDistribucionE:=TfrxDBDataset.Create(nil);
  Try
    QrColumnas.Connection := Connection.zConnection;
    QrMoe.Connection := Connection.zConnection;
    QrRecursos.Connection := Connection.zConnection;
    Td_CDistribucionE.UserName:='Td_CDistribucionE';
    Td_CImpDistribucionE.UserName:='Td_CImpDistribucionE';
    with mDatos do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidrecurso', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 250, True);
      FieldDefs.Add('smedida', ftString, 100, True);
      FieldDefs.Add('dcantSol', FtFloat, 0, True);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      FieldDefs.Add('dcantTotalPU', FtFloat, 0, True);
      for CuentaCol:=1 to TotalCol do
      begin
        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPernocta' + Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPlataforma'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sTipo'+ Inttostr(CuentaCol), ftString, 10, false);
      end;
      Active:=true;
    end;

    with mImprime do
    begin
      Active:=false;
      FieldDefs.Add('iCampo', ftInteger, 0, True);
      Active:=true;
    end;

    mImprime.EmptyTable;

    if (TImpresion=FtsCEquipos) or (TImpresion=FtsAll) then
    begin
      QrRecursos.Active := False;
      QrRecursos.SQL.Clear;
      QrRecursos.SQL.Add( 'select be.sIdEquipo, (sum(be.dCantHH) + ' +
                          '  IFNULL((SELECT ' +
                          '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                          '          FROM ' +
                          '            bitacoradeajustes ba ' +
                          '          WHERE ' +
                          '            ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                          '            ba.sNumeroActividad = be.sNumeroActividad AND ' +
                          '            ba.sIdPartidaAnexo = be.sIdEquipo AND ba.dFecha = be.dIdFecha), 0)) as dCantidad '+
                          'from bitacoradeequipos be '+
                          'inner join equipos e '+
                          'on ( e.sIdEquipo = be.sIdEquipo ) '+
                          'inner join bitacoradeactividades ba '+
                          'on ( ba.sContrato = :contrato and ba.dIdFecha = :fecha '+
                               'and ba.sNumeroOrden = :folio '+
                               'and ba.iIdDiario = be.iIdDiario '+
                          'and ba.iIdTarea = be.iIdTarea and ba.iIdActividad = be.iIdActividad) '+
                          'where e.sContrato = :contratoBarco '+
                          'and be.sContrato = :contrato '+
                          'and be.sNumeroOrden = :folio '+
                          'and be.sIdPernocta = :pernocta '+
                          'and be.dIdFecha = :fecha '+
                          'and be.sIdEquipo=:equipo ' +
                          'group by e.sIdEquipo '+
                          'order by e.iItemOrden');

          {$REGION 'CONSULTAS - PARTIDAS'}
      QrColumnas.Active:=false;
      QrColumnas.SQL.Add( 'select ot.sContrato,ot.sIdFolio, '+
                                 'ot.sNumeroOrden, ot.sTipo, '+
                                 'be.sIdPernocta AS PernoctaP, '+
                                 'ot.sIdPlataforma AS idPlataforma, '+
                                 'p.sDescripcion as Pernocta, '+
                                 'pt.sDescripcion as Plataforma '+
                          'from ordenesdetrabajo ot '+
                          'inner join contratos c '+
                          'on ( c.sContrato = ot.sContrato ) '+
                          'inner join bitacoradeequipos be '+
                          'on ( be.sContrato = ot.sContrato and be.sNumeroOrden = ot.sNumeroOrden ) '+
                          'inner join pernoctan p '+
                          'on ( ot.sIdPernocta = p.sIdPernocta) '+
                          'inner join plataformas pt '+
                          'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                          'where c.sContrato = :contrato '+
                          'and be.dIdFecha = :fecha '+
                          'group by ot.sIdFolio, p.sIdPernocta' );
        QrColumnas.ParamByName('Contrato').AsString:= ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate:=reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;
      {$ENDREGION}


      {$REGION 'CONSULTAS - TODOS LOS EQUIPOS REGISTRADOS EN MOE'}

      QrMoe.Active := False;
      QrMoe.SQL.Clear;
      QrMoe.SQL.Add('select mr.sIdRecurso, e.sDescripcion, e.sMedida, mr.dCantidad '+
                    'from moe m '+
                    'inner join moerecursos mr '+
                    'on ( mr.iIdMoe = m.iIdMoe ) '+
                    'inner join equipos e '+
                    'on ( e.sContrato = :contratobarco and e.sIdEquipo = mr.sIdRecurso ) '+
                    'where m.dIdFecha = (select max(didfecha) from moe where didfecha <=:Fecha and sContrato = :contrato) '+
                    'and m.sContrato = :contrato '+
                    'and mr.eTipoRecurso = "Equipo" '+
                    'order by e.iItemOrden');
      QrMoe.ParamByName('contratobarco').AsString := Global_Contrato_Barco;
      QrMoe.ParamByName('contrato').AsString := ReporteDiario.FieldByName('sOrden').AsString;
      QrMoe.ParamByName('Fecha').AsDateTime := reportediario.FieldByName('dIdFecha').AsDateTime;
      QrMoe.Open;

      {$ENDREGION}

      {$REGION 'INSERCION DE DATOS - INFORMACIÓN DEL EQUIPO'}

      DCiclo:=QrColumnas.RecordCount/TotalCol;
      iCiclo:=Trunc(DCiclo);
      if (dCiclo -iCiclo)>0 then
        Inc(iCiclo,1);
      iGrupo:=1;
      while iGrupo<=iCiclo do
      begin
        with qrMoe do
        begin
          first;
          while not eof do
          begin
            mDatos.Append;
            mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
            mDatos.FieldByName('sidrecurso').AsString:=FieldByName('sidrecurso').asstring;
            mDatos.FieldByName('sdescripcion').AsString:=fieldbyname('sdescripcion').asstring;
            mDatos.FieldByName('smedida').AsString:=fieldbyname('smedida').asstring;
            mDatos.FieldByName('dcantSol').AsFloat:=fieldbyname('dcantidad').AsFloat;
            mDatos.FieldByName('dcantTotalPU').AsFloat:=0;
            dTotal:=0;
            if iGrupo=1 then
              QrColumnas.First
            else
              QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

            CuentaCol:=1;
            ValTmp:=0;
            CantTmp:=0;
            iPosTmp:=0;
            while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
            begin

              mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString:=qrColumnas.FieldbyName('sIdFolio').AsString;
              mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('Pernocta').asstring;
              mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('Plataforma').Asstring;
              mDatos.FieldByName('sTipo'+ Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('sTipo').Asstring;

              QrRecursos.Active := False;
              QrRecursos.ParamByName('Equipo').AsString := FieldByName('sIdRecurso').AsString;

              QrRecursos.ParamByName('contratoBarco').AsString := global_contrato_barco;
              QrRecursos.ParamByName('contrato').AsString := ReporteDiario.FieldByName('sOrden').AsString;
              QrRecursos.ParamByName('folio').AsString := QrColumnas.FieldByName('sNumeroOrden').AsString;
              QrRecursos.ParamByName('fecha').AsDate := reportediario.FieldByName('dIdFecha').AsDateTime;
              QrRecursos.ParamByName('pernocta').AsString := QrColumnas.FieldByName('PernoctaP').AsString;               
              QrRecursos.Open;

              if QrRecursos.RecordCount>0 then
              begin
                mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(QrRecursos.FieldByName('dCantidad').Value,2);
                dTotal:=dTotal+ xRound(QrRecursos.FieldByName('dCantidad').Value,2);
                ValTmp:=ValTmp + QrRecursos.FieldByName('dCantidad').Value;
                if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                begin
                  CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                  iPosTmp:=CuentaCol;
                end;
              end;
              QrColumnas.next;
              Inc(CuentaCol);
            end;
            {if dTotal>0 then
            begin  }
              if dTotal<>xRound(ValTmp,2) then
              begin
                mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:= mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                + (xRound(ValTmp,2)-dTotal);
                if QrColumnas.Fieldbyname('sTipo').Asstring = 'TD' then
                   mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2)
                else
                   mDatos.FieldByName('dcantTotalPU').AsFloat:=xRound(ValTmp,2);
              end
              else
              begin
                 if QrColumnas.Fieldbyname('sTipo').Asstring = 'TD' then
                    mDatos.FieldByName('dcantTotal').AsFloat:=dTotal
                 else
                    mDatos.FieldByName('dcantTotalPU').AsFloat:=dTotal;
              end;
              mDatos.Post;
           { end
            else
              mDatos.cancel;}
            next;
          end;
        end;
        Inc(iGrupo,1);
      end;
          

//      {Esta secccion es para mostrar la hoja vacia sino existen datos.}
//      if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
//         if MDatos.RecordCount >= 0 then
//      else
      if MDatos.RecordCount > 0 then
      begin
        mImprime.Append;
        mImprime.FieldByName('iCampo').AsInteger:=1;
        mImprime.Post;
      end;
      {$ENDREGION}
    end;
    Td_CDistribucionE.DataSet:=MDatos;
    Td_CDistribucionE.FieldAliases.Clear;
    Td_CImpDistribucionE.DataSet:=MImprime;
    Td_CImpDistribucionE.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_CDistribucionE);
    Reporte.DataSets.Add(Td_CImpDistribucionE);
  Finally
    QrMoe.Free;
    QrRecursos.Free;
    QrColumnas.Destroy;
  End;
end;

procedure ReportePDF_ResumenDeRecursos(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
var
  QrMoeP,         //Moe Personal
  QrMoeE,         //Moe Equipo
  QrColumnas,     //Folios, Pernoctas, Plataformas
  QrMovBarco,     //Movimientos del barco
  QrPersonal,     //Personal
  QrEquipo,       //Equipo
  QrPernoctas     //Pernoctas
  : TZReadOnlyQuery;

  iFilaM,   //Moe
  iFilaP,   //Personal
  iFilaE,   //Equipo
  iFilaPr,  //Pernoctas
  iColumna,
  x : Integer;
begin
  //Sam
  try
    try
      {$REGION 'CREA TZREADQUERY'}
      QrMovBarco := TZReadOnlyQuery.Create(nil);
      QrMovBarco.connection := connection.zconnection;
      QrColumnas := TZReadOnlyQuery.Create(nil);
      QrColumnas.connection := connection.zconnection;
      QrMoeP := TZReadOnlyQuery.Create(nil);
      QrMoeP.connection := connection.zconnection;
      QrMoeE := TZReadOnlyQuery.Create(nil);
      QrMoeE.connection := connection.zconnection;
      QrPersonal := TZReadOnlyQuery.Create(nil);
      QrPersonal.connection := connection.zconnection;
      QrEquipo := TZReadOnLyQuery.Create(nil);
      QrEquipo.connection := connection.zconnection;
      QrPernoctas := TZReadOnlyQuery.Create(nil);
      QrPernoctas.Connection := connection.zconnection;
      {$ENDREGION}

      {FILIOS DEL DIA}
      {$REGION 'CONSULTA COLUMNAS - FOLIOS LABORADOS'}
      with QrColumnas do
      begin
        active := false;
        sql.text := 'select ot.sIdFolio, '+
                           'ot.sIdPernocta, '+
                           'ot.sIdPlataforma, '+
                           'p.sDescripcion as sPernocta, '+
                           'pt.sDescripcion as sPlataforma '+
                    'from ordenesdetrabajo ot '+
                    'inner join contratos c '+
                    'on ( c.sContrato = ot.sContrato ) '+
                    'inner join bitacoradeequipos be '+
                    'on ( be.sContrato = ot.sContrato and be.sNumeroOrden = ot.sNumeroOrden ) '+
                    'inner join pernoctan p '+
                    'on ( ot.sIdPernocta = p.sIdPernocta) '+
                    'inner join plataformas pt '+
                    'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                    'where c.sContrato = :contrato '+
                    'and be.dIdFecha = :fecha '+
                    'group by ot.sIdFolio, p.sIdPernocta';
        parambyname('contrato').asstring := ReporteDiario.fieldbyname('sOrden').asstring;
        parambyname('fecha').asdate      := reportediario.fieldbyname('didfecha').asdatetime;
        open;
      end;
      {$ENDREGION}

      if QrColumnas.RecordCount > 0 then
      begin
        {COLUMNAS TIENEN QUE ESTAR ALINEADAS CON EL LOGO
        EN CASO QUE NO EXISTAN SUFICIENTES DEBERAN CREARSE}
        {$REGION 'CREA COLUMNAS'}
        with QrColumnas do
        begin
          first;

        {$ENDREGION}

        {CONSULTA MOVIMIENTOS DEL BARCO}
        {$REGION 'CONSULTA MOVIMIENTOS DEL BARCO YA CON SU FACTOR POR FOLIO'}
        with QrMovBarco do
        begin
          active := false;
          sql.text := 'select tm.sIdTipoMovimiento, '+
                             'tm.sDescripcion, '+
                             'tm.sMedida, '+
                             ' ot.sIdFolio, '+
                             'ot.sIdPernocta, '+
                             'ot.sIdPlataforma, '+
                             'sum( me.sFactor ) as sFactor '+
                      'from movimientosdeembarcacion me '+
                      'inner join tiposdemovimiento tm '+
                        'on ( me.sContrato = tm.sContrato '+
                           'and me.sClasificacion = tm.sIdTipoMovimiento) '+
                      'inner join movimientosxfolios mf '+
                        'on ( mf.sContrato = me.sContrato '+
                           'and mf.dIdFecha = me.dIdFecha '+
                           'and mf.iIdDiario = me.iIdDiario) '+
                      'inner join ordenesdetrabajo ot '+
                      'on ( ot.sContrato = :contrato '+
                           'and ot.sIdFolio = mf.sFolio)'+
                      'where tm.sContrato = :contratoBarco '+
                      'and tm.sClasificacion = "Movimiento de Barco" '+
                      'and me.dIdFecha = :fecha '+
                      'group by tm.sIdTipoMovimiento';
          parambyname('contratobarco').asstring := global_contrato_barco;
          parambyname('contrato').asstring := reportediario.FieldByName('sOrden').asstring;
          parambyname('fecha').AsDate := reportediario.fieldbyname('didfecha').asdatetime;
          open;
        end;
        {$ENDREGION}

        {TRABAJA EN EXCEL LOS MOVIMIENTOS}
        {$REGION 'ESCRIBE LOS MOV BARCO'}
        if QrMovBarco.RecordCount > 0 then
        begin
          with QrMovBarco do
          begin
            while not Eof do
            begin
              fieldbyname('sidtipomovimiento').asstring;
              fieldbyname('sdescripcion').asstring;
              fieldbyname('smedida').asstring;
              next;
            end;
          end;
        end;
        {$ENDREGION}

        {PERSONAL REGISTRADO EN EL MOE}
        {$REGION 'CONSULTA Y ESCRIBE PERSONAL SOLICITADO EN EL ULTIMO MOE VIGENTE'}
        with QrMoeP do
        begin
          active := false;    
          sql.text := 'select p.sDescripcion, '+
                             'mr.*, p.lSumaSolicitado, '+
                             'p.sMedida, m.iidmoe '+
                      'from moe AS m '+
                      'inner join moerecursos as mr '+
                        'on (mr.iidMoe=m.iidMoe) '+
                      'inner join personal as p '+
                        'ON (p.scontrato=:contratoBarco '+
                            'and p.sidpersonal=mr.sidRecurso) '+
                      'where m.didfecha = (select max(didfecha) '+
                                          'from moe where didfecha <=:Fecha '+
                                          'and sContrato = :contrato) '+
                      'and m.sContrato = :contrato '+
                      'and mr.eTipoRecurso = "Personal" '+
                      'order by p.iItemOrden';
          parambyname('contratoBarco').asstring := global_contrato_barco;
          parambyname('contrato').asstring := reportediario.fieldbyname('sOrden').asstring;
          parambyname('fecha').asdate := reportediario.fieldbyname('didfecha').asdatetime;
          open;

          if recordcount > 0 then
          begin
            first;
            while not Eof do
            begin
              fieldbyname('sidrecurso').AsString;
              fieldbyname('sdescripcion').asstring;
              fieldbyname('smedida').asstring;
              next;
            end;
          end;
        end;
        {$ENDREGION}
        {EQUIPO REGISTRADO EN EL MOE}
        {$REGION 'CONSULTA Y ESCRIBE EL EQUIPO SOLICITADO EN EL ULTIMO MOE VIGENTE'}
        with QrMoeE do
        begin
          active := false;
          sql.Text := 'SELECT p.sidequipo, p.sDescripcion, p.sMedida ' +
                      'FROM moerecursos AS mr ' +
                      'INNER JOIN moe AS m ' +
                      'ON (m.iidMoe=mr.iidMoe) ' +
                      'INNER JOIN equipos AS p ' +
                      'ON (p.sContrato=:Contrato AND p.sIdEquipo=mr.sidRecurso) ' +
                      'LEFT JOIN gruposdeequipo AS ge ' +
                      'ON (ge.iId = p.iId_AgrupadorEquipo) ' +
                      'WHERE ' +
                      'dIdFecha = (SELECT max(didfecha) FROM moe WHERE didfecha <=:Fecha AND sContrato = :ContratoNormal) ' +
                      'AND m.sContrato = :ContratoNormal ' +
                      'AND mr.eTipoRecurso="Equipo" ' +
                      'ORDER BY p.iItemOrden';
          parambyname('contrato').asstring := global_contrato_barco;
          parambyname('contratoNormal').asstring := reportediario.fieldbyname('sOrden').asstring;
          parambyname('fecha').asdate := reportediario.fieldbyname('didfecha').asdatetime;
          open;

          if recordcount > 0 then
          begin
            first;
            while not eof do
            begin
              fieldbyname('sidequipo').AsString;
              fieldbyname('sdescripcion').asstring;
              fieldbyname('smedida').asstring;
              next;
            end;
          end;
        end;
        {$ENDREGION}

        qrColumnas.First;
        while not QrColumnas.Eof do
        begin
          {DATOS DE LA COLUMNA}
          {$REGION 'COLUMNAS'}
          QrColumnas.FieldByName('sidfolio').asstring;
          QrColumnas.FieldByName('splataforma').asstring;
          {$ENDREGION}

          {BARCO ANEXO C}
          {$REGION 'INSERTA BARCO'}
          try
            try
              QrMovBarco.Filtered := False;
              QrMovBarco.Filter := 'sIdFolio = '+QuotedStr(QrColumnas.FieldByName('sIdFolio').AsString);
              QrMovBarco.Filtered := true;
            except
              ;
            end;
             //Que hace excel Aqui??
            {while excel.cells[ifilam, 1].text <> '' do
            begin
              if ( QrMovBarco.Locate('sIdPlataforma', qrColumnas.FieldByName('sIdPlataforma').asstring, []) )
                  and ( QrMovBarco.fieldByName('sIdTipoMovimiento').asstring =  excel.cells[ifilam, 1].text) then
                excel.range[columnanombre(icolumna)+inttostr(ifilam)].value := QrMovBarco.FieldByName('sFactor').AsString;
              inc(ifilam);
            end;   }
          finally
            QrMovBarco.Filtered := False;
          end;
          {$ENDREGION}

          {PERSONAL ANEXO C-1}
          {$REGION 'CONSULTA PERSONAL REPORTADO'}
          iFilaP := iFilaM + 3;
          with QrPersonal do
          begin
            active  := false;
            sql.text := 'SELECT bp.sIdPersonal, bp.sDescripcion, ' +
                        'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                        'IF(SUM(bp.dCanthh) > 0, (SUM(ROUND(bp.dCanthh, 2)) +   IFNULL((SELECT ' +
                        '            SUM(round(CAST(ba.dAjuste AS CHAR), 2)) ' +
                        '          FROM ' +
                        '            bitacoradeajustes ba ' +
                        '          WHERE ' +
                        '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                        '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                        '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)), SUM(bp.dCantidad)) AS Total ' +
                        'FROM bitacoradepersonal bp ' +
                        'WHERE bp.scontrato = :contrato AND bp.sNumeroOrden = :Folio AND bp.didfecha = :fecha ' +
                        'AND bp.sidPernocta = :Pernocta AND bp.sidplataforma = :Plataforma GROUP BY bp.sIdPersonal';
            parambyname('contrato').asstring := reportediario.FieldByName('sOrden').asstring;
            parambyname('folio').asstring := QrColumnas.fieldbyname('sidfolio').asstring;
            parambyname('fecha').asdate := reportediario.fieldbyname('didfecha').asdatetime;
            parambyname('pernocta').asstring := QrColumnas.fieldbyname('sidpernocta').asstring;
            parambyname('plataforma').AsString := QrColumnas.FieldByName('sidplataforma').asstring;
            open;

            if recordcount > 0 then
            begin
              first;
              //Aqui tambien esta excel
              {while excel.cells[ifilap, 1].text <> '' do
              begin
                if QrMoeP.Locate( 'sIdRecurso', excel.cells[ifilap, 1].text ,[]) then
                begin
                  if Locate('sIdPersonal', QrMoeP.FieldByName('sidrecurso').AsString, []) then
                    excel.range[columnanombre(icolumna)+inttostr(ifilap)].value := fieldbyname('total').asfloat;
                end;
                inc(ifilap);
              end;}
            end;
          end;
          {$ENDREGION}

          {EQUIPO ANEXO C-2}
          {$REGION 'CONSULTA EQUIPO REPORTADO'}
          ifilaE := ifilap + 3;
          with QrEquipo do
          begin
            active := false;
            sql.text := 'select be.sIdEquipo, (sum(be.dCantHH) + ' +
                        '  IFNULL((SELECT ' +
                        '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                        '          FROM ' +
                        '            bitacoradeajustes ba ' +
                        '          WHERE ' +
                        '            ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                        '            ba.sNumeroActividad = be.sNumeroActividad AND ' +
                        '            ba.sIdPartidaAnexo = be.sIdEquipo AND ba.dFecha = be.dIdFecha), 0)) as dCantidad '+
                        'from bitacoradeequipos be '+
                        'inner join equipos e '+
                        'on ( e.sIdEquipo = be.sIdEquipo ) '+
                        'inner join bitacoradeactividades ba '+
                        'on ( ba.sContrato = :contrato and ba.dIdFecha = :fecha '+
                             'and ba.sNumeroOrden = :folio '+
                             'and ba.iIdDiario = be.iIdDiario ) '+
                        'where e.sContrato = :contratoBarco '+
                        'and be.sContrato = :contrato '+
                        'and be.sNumeroOrden = :folio '+
                        'and be.sIdPernocta = :pernocta '+
                        'and be.dIdFecha = :fecha '+
                        'group by e.sIdEquipo '+
                        'order by e.iItemOrden';
            parambyname('contratoBarco').asstring := global_contrato_barco;
            parambyname('contrato').asstring := reportediario.fieldbyname('sOrden').asstring;
            parambyname('folio').asstring := QrColumnas.fieldbyname('sidfolio').asstring;
            parambyname('pernocta').asstring := QrColumnas.fieldbyname('sidpernocta').asstring;
            parambyname('fecha').asdate := reportediario.fieldbyname('didfecha').asdatetime;
            open;

            if recordcount > 0 then
            begin
              first;
              //Aqui tambien esta excel
              {while excel.cells[ifilae, 1].text <> '' do
              begin
                if QrMoeE.Locate( 'sidequipo', excel.cells[ifilae, 1].text ,[]) then
                begin
                  if Locate('sidequipo', QrMoeE.FieldByName('sidequipo').AsString, []) then
                    excel.range[columnanombre(icolumna)+inttostr(ifilae)].value := fieldbyname('dCantidad').asfloat;
                end;
                inc(ifilae);
              end; }
            end;
          end;
          {$ENDREGION}
          QrColumnas.Next;
        end;

        {CONSULTA E INSERTA PERNOCTAS
        PD: LAS PERNOCTAS SON POR OT NO POR FOLIO}
        {$REGION 'CONSULTA PERNOCTAS'}
        with qrPernoctas do
        begin
          active := false;
          sql.text := 'select c.sIdPernocta, '+
                             'c.sDescripcion, '+
                             'c.sMedida, '+
                             '@dAjuste := ( select ifnull( bpr.dCantidad, 0) '+
                                                     'from bitacoradepernocta bpr '+
                                                     'where bpr.sContrato = :contrato '+
                                                     'and bpr.dIdFecha = :fecha '+
                                                     'and bpr.sNumeroOrden <> "@" ) as dAjuste, '+
                             '@dCantidad := (sum( bp.dCantHH ) + ' +
                             '  IFNULL((SELECT ' +
                             '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                             '          FROM ' +
                             '            bitacoradeajustes ba ' +
                             '          WHERE ' +
                             '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                             '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                             '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as dCantidadHH, '+
                             '(if( @dAjuste < @dCantidad, '+
                               '( @dCantidad - @dAjuste ), '+
                               '( if( @dAjuste = @dCantidad, '+
                                     '@dCantidad, '+
                                     '( @dCantidad + @dAjuste ) '+
                                 ') '+
                               ') '+
                             ') ) as dCantidadAjuste, '+

                       '((sum( bp.dCantHH ) + ' +
                       '  IFNULL((SELECT ' +
                       '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                       '          FROM ' +
                       '            bitacoradeajustes ba ' +
                       '          WHERE ' +
                       '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                       '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                       '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) - ( select ifnull( bpr.dCantidad, 0) '+
                                               'from bitacoradepernocta bpr '+
                                               'where bpr.sContrato = :contrato '+
                                               'and bpr.dIdFecha = :fecha '+
                                               'and bpr.sNumeroOrden <> "@" ) ) as dCantidad '+
                      'from cuentas c '+
                      'inner join bitacoradepersonal bp '+
                        'on ( bp.lAplicaPernocta = "Si" '+
                          'and bp.sTipoPernocta = c.sIdPernocta) '+

                      'inner join moerecursos mr '+
                        'on ( mr.sIdRecurso = bp.sIdPersonal '+
                          'and mr.eTipoRecurso = "Personal" '+
                          'and mr.iIdMoe = ( select m.iIdMoe from moe m where m.sContrato = :contrato '+
                                            'and m.dIdFecha <= :fecha order by m.dIdFecha desc limit 1) ) '+

                      'inner join personal p '+
                        'on ( p.sContrato = :contratoBarco '+
                          'and p.sIdPersonal = bp.sIdPersonal ) '+

                      'inner join tiposdepersonal tp '+
                        'on ( p.sIdTipoPersonal = tp.sIdTipoPersonal ) '+

                      'inner join bitacoradeactividades ba '+
                        'on ( ba.sNumeroActividad = bp.sNumeroActividad AND '+
                          'and ba.dIdFecha = bp.dIdFecha '+
                          'and ba.sIdTipoMovimiento = "E" ) '+

                      'inner join ordenesdetrabajo ot '+
                        'on ( ot.sContrato = :contrato and ot.sNumeroOrden = ba.sNumeroOrden ) '+

                      'inner join pernoctan pr '+
                        'on ( pr.sIdPernocta = ot.sIdPernocta ) '+

                      'inner join plataformas pl '+
                        'on ( pl.sIdPlataforma = ot.sIdPlataforma ) '+

                      'where bp.sContrato = :contrato '+
                      'and bp.dIdFecha = :fecha '+

                      'group by c.sIdPernocta '+
                      'order by c.sIdPernocta';
          parambyname('contrato').asstring := ReporteDiario.FieldByName('sOrden').asstring;
          parambyname('fecha').asdate := ReporteDiario.FieldByName('dIdFecha').asDatetime;
          parambyname('contratoBarco').asstring := global_contrato_barco;
          open;
        end;

        if qrPernoctas.recordcount > 0 then
        begin

          qrPernoctas.first;
          while not qrPernoctas.eof do
          begin

            qrPernoctas.fieldbyname('sIdPernocta').asstring;
            qrPernoctas.fieldbyname('sDescripcion').asstring;
            qrPernoctas.fieldbyname('smedida').asstring;
            qrPernoctas.fieldbyname('dCantidad').AsFloat;
            qrpernoctas.next;
          end;
        end;

        {$ENDREGION}


      end;
    end;
    except
      on e:exception do
      begin
        //Habra que determinar que hacer con los mensajes
        //MessageDlg('No se puede imprimir el reporte diario: '+e.Message, mtInformation, [mbOk], 0);
      end;
    end;
  finally
    QrMovBarco.Free;
    QrColumnas.Free;
    QrMoeP.Free;
    QrMoeE.Free;
    QrPersonal.Free;
    QrEquipo.Free;
    QrPernoctas.Free;
  end;
end;

procedure ReportePDF_PartidasAnexoC(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
var
  QryAnexo,
  QryRecursos : TZReadOnlyQuery;
  Td_PartidasAnexo,
  Td_ImpPartidasAnexo : TfrxDBDataset;
  rxAnexo: TClientDataSet;
begin
  try
      {$REGION 'CREA TZREADQUERY'}
      QryAnexo := TZReadOnlyQuery.Create(nil);
      QryAnexo.connection := connection.zconnection;

      QryRecursos := TZReadOnlyQuery.Create(nil);
      QryRecursos.connection := connection.zconnection;

      Td_PartidasAnexo:= TfrxDBDataset.Create(nil);
      Td_PartidasAnexo.UserName:='Td_PartidasAnexo';

      Td_ImpPartidasAnexo:= TfrxDBDataset.Create(nil);
      Td_ImpPartidasAnexo.UserName:='Td_ImpPartidasAnexo';

      rxAnexo := TClientDataSet.Create(nil);
      rxAnexo.Active := False;
      rxAnexo.FieldDefs.Add('sTipo', ftString, 25, False);
      rxAnexo.FieldDefs.Add('sOT', ftString, 15, False);
      rxAnexo.FieldDefs.Add('sTipoAnexo', ftString, 15, False);
      rxAnexo.FieldDefs.Add('sAnexoPaquete', ftString, 2, False);
      rxAnexo.FieldDefs.Add('sIdRecurso', ftString, 50, False);
      rxAnexo.FieldDefs.Add('sDescripcion', ftString, 1000, False);
      rxAnexo.FieldDefs.Add('sMedida', ftString, 20, False);
      rxAnexo.FieldDefs.Add('dCantidad', ftFloat, 0, False);
      rxAnexo.CreateDataSet;
      rxAnexo.Active := True;

      {$ENDREGION}
      if (TImpresion=FtsRRecursos) or (TImpresion=FtsAll) then
      begin
          QryAnexo.Active := False;
          QryAnexo.SQL.Clear;
          QryAnexo.SQL.Add('select * from anexos where (sTipo = "Barco" or (sTipo = "Personal" and sTierra="No") or sTipo = "Equipo" or sTipo = "Pernocta" '+
                           'or sTipo = "Anexo" or sTipo = "AnexoEXT" or sTipo = "AnexoC1" or sTipo = "AnexoC2" or sTipo = "AnexoC3" or sTipo ="AnexoC4") order by iOrden ');
          QryAnexo.Open;

          while not QryAnexo.Eof do
          begin
              rxAnexo.Append;
              rxAnexo.FieldValues['sTipo']        := 'Anexo';
              rxAnexo.FieldValues['sTipoAnexo']   := QryAnexo.FieldByName('sTipo').AsString;
              rxAnexo.FieldValues['sAnexoPaquete']:= 'Si';
              rxAnexo.FieldValues['sOT']          := reportediario.FieldByName('sOrden').AsString;
              rxAnexo.FieldValues['sIdRecurso']   := QryAnexo.FieldByName('sAnexo').AsString;
              rxAnexo.FieldValues['sDescripcion'] := QryAnexo.FieldByName('sDescripcion').AsString;
              rxAnexo.Post;

              {$REGION 'CONSULTA MOVIMIENTOS DEL BARCO YA CON SU FACTOR POR FOLIO'}
              if QryAnexo.FieldByName('sTipo').AsString = 'BARCO' then
              begin

                QryRecursos.Active := False;
                QryRecursos.SQL.Clear;
                QryRecursos.SQL.Add('SELECT ' +
                                    '  tm.sIdTipoMovimiento as sRecurso, ' +
                                    '  tm.sDescripcion, ' +
                                    '  tm.sTipo, ' +
                                     'tm.sMedida, '+
                                    '  ( ' +
                                    '    SELECT ' +
                                    '      ifnull(SUM( me.sFactor ),0) ' +
                                    '    FROM movimientosdeembarcacion AS me ' +
                                    '    WHERE me.sContrato = tm.sContrato ' +
                                    '    AND me.sOrden = :Orden ' +
                                    '    AND me.dIdFecha = :Fecha ' +
                                    '    AND me.sIdFase = "OPER" ' +
                                    '    AND me.sClasificacion = tm.sIdTipoMovimiento ' +
                                    '  ) AS dCantidad,tm.lImprimeCeros ' +
                                    'FROM tiposdemovimiento AS tm ' +
                                    'WHERE ' +
                                    '  tm.sContrato = :ContratoBarco ' +
                                    '  AND tm.sClasificacion = "Movimiento de Barco" '+
                                    'having ((dCantidad<>0) or ((tm.lImprimeCeros ="Si") and (dCantidad=0)))'
                                   );

                QryRecursos.Params.ParamByName('contratobarco').AsString := global_contrato_barco;
                QryRecursos.Params.ParamByName('Orden').AsString      := reportediario.FieldByName('sOrden').AsString;
                QryRecursos.Params.ParamByName('fecha').AsDate           := reportediario.FieldByName('dIdFecha').AsDateTime;
                QryRecursos.Open;
              end;

              if QryAnexo.FieldByName('sTipo').AsString = 'PERSONAL' then
              begin

                QryRecursos.Active := False;
                QryRecursos.SQL.Clear;
                // mODIFICACIONES Realizadas por E. Rangel - 20/08/2016
                QryRecursos.SQL.Add('SELECT ' +
                                    '  bp.sIdPersonal as sRecurso, bp.sDescripcion, p.sMedida, ' +
                                    '  IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                    '  IF(SUM(bp.dCanthh) > 0, (SUM(ROUND(bp.dCanthh, 2)) + ' +
                                    '   IFNULL((SELECT ' +
                                    '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                    '          FROM ' +
                                    '            bitacoradeajustes ba ' +
                                    '          WHERE ' +
                                    '            ba.sContrato = bp.sContrato AND ' +
                                    '            ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                    '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                    '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                    '            ba.dFecha = bp.dIdFecha), 0)), SUM(bp.dCantidad)) AS dCantidad ' +
                                    'FROM ' +
                                    '  bitacoradepersonal bp ' +
                                    'inner join ' +
                                    '  personal p ' +
                                    '    on (p.sContrato = bp.sContrato and p.sIdPersonal = bp.sIdPersonal) ' +
                                    'WHERE ' +
                                    '  bp.scontrato =:contrato AND bp.dIdfecha = :fecha ' +
                                    'group by ' +
                                    '  bp.sIdPersonal ' +
                                    'order by ' +
                                    '  p.iItemOrden');
                {QryRecursos.SQL.Add('SELECT bp.sIdPersonal as sRecurso, bp.sDescripcion, p.sMedida, SUM(bp.dAjuste) AS Ajuste, '+
                            'IF(SUM(bp.dCanthh) > 0, ROUND(SUM(bp.dCanthh), 2), SUM(bp.dCantidad)) AS dCantidad '+
                            'FROM bitacoradepersonal bp '+
                            'inner join personal p on (p.sContrato = bp.sContrato and p.sIdPersonal = bp.sIdPersonal) '+
                            'WHERE bp.scontrato =:contrato AND bp.dIdfecha = :fecha '+
                            'group by bp.sIdPersonal order by p.iItemOrden');}
                QryRecursos.Params.ParamByName('contrato').AsString := reportediario.FieldByName('sOrden').AsString;
                QryRecursos.Params.ParamByName('fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
                QryRecursos.Open;
              end;

              if QryAnexo.FieldByName('sTipo').AsString = 'EQUIPO' then
              begin
                QryRecursos.Active := False;
                QryRecursos.SQL.Clear;
                QryRecursos.SQL.Add('SELECT bp.sIdEquipo as sRecurso, bp.sDescripcion, p.sMedida, ' +
                            'IFNULL(SUM(CAST(be.dAjuste AS CHAR)), 0) AS Ajuste, '+
                            'IF(SUM(bp.dCanthh) > 0, (SUM(ROUND(bp.dCanthh, 2)) + ' +
                            '  IFNULL((SELECT ' +
                            '            SUM(ROUND(CAST(ba.dAjuste AS CHAR), 2)) ' +
                            '          FROM ' +
                            '            bitacoradeajustes ba ' +
                            '          WHERE ' +
                            '            ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                            '            ba.sNumeroActividad = be.sNumeroActividad AND ' +
                            '            ba.sIdPartidaAnexo = be.sIdEquipo AND be.dFecha = be.dIdFecha), 0)), SUM(bp.dCantidad)) AS dCantidad '+
                            'FROM bitacoradeequipos bp '+
                            'inner join equipos p on (p.sContrato = bp.sContrato and p.sIdEquipo = bp.sIdEquipo) '+
                            'WHERE bp.scontrato =:contrato AND bp.dIdfecha = :fecha '+
                            'group by bp.sIdEquipo order by p.iItemOrden');
                QryRecursos.Params.ParamByName('contrato').AsString := reportediario.FieldByName('sOrden').AsString;
                QryRecursos.Params.ParamByName('fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
                QryRecursos.Open;
              end;

              if QryAnexo.FieldByName('sTipo').AsString = 'PERNOCTA' then //ABBY
              begin
                if connection.contrato.FieldByName('eLugarOt').AsString = 'Abordo' then
                begin
                    QryRecursos.Active := False;
                    QryRecursos.SQL.Clear;
                    {QryRecursos.SQL.text:='select c.sIdCuenta,c.sIdPernocta as sRecurso,c.sDescripcion,c.sMedida,' + #10 +
                                          '(ifnull((select ifnull(sum(bpr.dCantidad),0) from bitacoradepernocta bpr' + #10 +
                                          'where bpr.sContrato =:Orden and bpr.didfecha=:fecha and (bpr.sNumeroOrden <> "@"' + #10 +
                                          ') and c.sidcuenta=bpr.sidcuenta group by bpr.sContrato),0)) as dCantidad from cuentas c' + #10 +
                                          'order by c.sIdCuenta';}

QryRecursos.SQL.text := 'select ' +
                        '  c.sIdCuenta, ' +
                        '  c.sIdPernocta as sRecurso, ' +
                        '  c.sDescripcion, ' +
                        '  c.sMedida, ' +
                        '  (ifnull((select ' +
                        '             (ifnull(sum(bpr.dCantidad),0) + ' +
                        '              IFNULL((SELECT ' +
                        '                        SUM(CAST(ba.dAjuste AS CHAR)) ' +
                        '                      FROM ' +
                        '                        bitacoradeajustes ba ' +
                        '                      WHERE ' +
                        '                        ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                        '                        ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0)) ' +
                        '           from ' +
                        '             bitacoradepernocta bpr ' +
                        '           where ' +
                        '             bpr.sContrato =:Orden and ' +
                        '             bpr.didfecha=:fecha and ' +
                        '             (bpr.sNumeroOrden <> "@") and ' +
                        '             c.sidcuenta=bpr.sidcuenta ' +
                        '           group by ' +
                        '             bpr.sContrato),0)) as dCantidad ' +
                        'from ' +
                        '  cuentas c ' +
                        'order by ' +
                        '  c.sIdCuenta';

                    QryRecursos.Params.ParamByName('Orden').AsString := reportediario.FieldByName('sOrden').AsString;
                    QryRecursos.Params.ParamByName('fecha').AsDate   := reportediario.FieldByName('dIdFecha').AsDateTime;
                    QryRecursos.Open;
                end;
              end;

              if (QryAnexo.FieldByName('sTipo').AsString = 'ANEXO') or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOEXT')
                 or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOC1') or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOC2')
                 or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOC3') or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOC4') then
              begin
                QryRecursos.Active := False;
                QryRecursos.SQL.Clear;
                QryRecursos.SQL.Add('SELECT ao.sNumeroActividad as sRecurso, ao.mDescripcion as sDescripcion, ao.sMedida, sum(ba.dCantidad) AS dCantidad '+
                                    'FROM bitacoradeactividades ba '+
                                    'inner join actividadesxorden ao on (ao.sContrato = ba.sContrato and ao.sIdConvenio = ba.sIdConvenio and ao.sWbs = ba.sWbs '+
                                    'and ao.sTipoActividad = "Actividad" and sTipoAnexo = "PU" and ao.sAnexo =:Anexo) '+
                                    'WHERE ba.scontrato =:contrato AND ba.dIdfecha =:fecha '+
                                    'group by ba.sNumeroActividad order by ao.iItemOrden');
                QryRecursos.Params.ParamByName('contrato').AsString := reportediario.FieldByName('sOrden').AsString;
                QryRecursos.Params.ParamByName('fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
                QryRecursos.Params.ParamByName('Anexo').AsString    := QryAnexo.FieldByName('sAnexo').AsString;
                QryRecursos.Open;
              end;

              if QryRecursos.RecordCount = 0 then
                 rxAnexo.Delete;

              QryRecursos.First;
              while not QryRecursos.Eof do
              begin
                  rxAnexo.Append;
                  rxAnexo.FieldValues['sTipo']        := 'Recurso';
                  rxAnexo.FieldValues['sOT']          := reportediario.FieldByName('sOrden').AsString;
                  rxAnexo.FieldValues['sIdRecurso']   := QryRecursos.FieldByName('sRecurso').AsString;
                  rxAnexo.FieldValues['sDescripcion'] := QryRecursos.FieldByName('sDescripcion').AsString;
                  rxAnexo.FieldValues['sMedida']      := QryRecursos.FieldByName('sMedida').AsString;
                  if QryAnexo.FieldByName('sTipo').AsString = 'PERNOCTA' then
                  begin
                    connection.QryBusca.Active:=false;
                    connection.QryBusca.SQL.Text:='select (ifnull(sum(bp.dCantHH),0) + ' +
                                                  '  IFNULL((SELECT ' +
                                                  '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                                  '          FROM ' +
                                                  '            bitacoradeajustes ba ' +
                                                  '          WHERE ' +
                                                  '            ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                                                  '            ba.sNumeroActividad = be.sNumeroActividad AND ' +
                                                  '            ba.sIdPartidaAnexo = be.sIdEquipo AND be.dFecha = bp.dIdFecha), 0)) as dCantidad from bitacoradepersonal bp' + #10 +
                                                  'inner join bitacoradeactividades ba on(bp.sContrato =ba.sContrato and' + #10 +
                                                  'ba.sNumeroOrden=bp.sNumeroOrden and  bp.dIdFecha = ba.dIdFecha and ba.iIdDiario=bp.iIdDiario and' + #10 +
                                                  'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad)' + #10 +
                                                  'inner join personal p on (p.sContrato =:ContratoBarco and p.sIdPersonal = bp.sIdPersonal)' + #10 +
                                                  'inner join cuentas c on (bp.sTipoPernocta = c.sIdPernocta)' + #10 +
                                                  'where bp.sContrato=:Orden and bp.dIdFecha=:Fecha and p.lpernocta="Si"' + #10 +
                                                  'and c.sIdCuenta=:Pernocta';

                    connection.QryBusca.ParamByName('ContratoBarco').AsString:= global_contrato_barco;
                    connection.QryBusca.ParamByName('Orden').AsString:=reportediario.FieldByName('sOrden').AsString;
                    connection.QryBusca.ParamByName('Fecha').AsDate:=reportediario.FieldByName('dIdFecha').AsDateTime;
                    connection.QryBusca.ParamByName('Pernocta').AsString:=QryRecursos.FieldByName('sIdcuenta').AsString;
                    connection.QryBusca.Open;
                    if connection.QryBusca.RecordCount=1 then
                      rxAnexo.FieldByName('dcantidad').AsFloat:=xRound(QryRecursos.FieldByName('dCantidad').value,2)+xRound(connection.QryBusca.FieldByName('dCantidad').value,2)
                    else
                      rxAnexo.FieldByName('dcantidad').AsFloat:=xRound(QryRecursos.FieldByName('dCantidad').value,2);


                  end
                  else
                  begin
                     if (QryAnexo.FieldByName('sTipo').AsString = 'ANEXO') or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOEXT')
                        or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOC1') or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOC2')
                        or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOC3') or (QryAnexo.FieldByName('sTipo').AsString = 'ANEXOC4') then
                     begin
                        rxAnexo.FieldValues['dCantidad']    := xRound(QryRecursos.FieldByName('dCantidad').value,4);
                        rxAnexo.FieldValues['sTipoAnexo']   := QryAnexo.FieldByName('sTipo').AsString;
                     end
                     else
                        rxAnexo.FieldValues['dCantidad']    := xRound(QryRecursos.FieldByName('dCantidad').value,2);
                  end;

                  if rxAnexo.FieldByName('dCantidad').asfloat >= 0 then
                    rxAnexo.Post
                  else
                    RxAnexo.Cancel;

                  QryRecursos.Next;
              end;

              {$ENDREGION}

              QryAnexo.Next;
          end;
      end;
      Td_PartidasAnexo.DataSet   := rxAnexo;
      Td_ImpPartidasAnexo.DataSet:= rxAnexo;

      Td_PartidasAnexo.FieldAliases.Clear;
      Td_ImpPartidasAnexo.FieldAliases.Clear;

      Reporte.DataSets.Add(Td_PartidasAnexo);
      Reporte.DataSets.Add(Td_ImpPartidasAnexo);

  finally

  end;
end;

procedure ReportePDF_ReporteFotografico(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
var
  QryReporteFotografico: TZQuery;
  Td_Fotografico,
  Td_ImpFotografico : TfrxDBDataset;
  rxFotografico: TClientDataSet;
begin

  {$REGION 'NOTAS'}
  Try

    QryReporteFotografico := TZQuery.Create(nil);
    QryReporteFotografico.Connection := Connection.zConnection;

    Td_Fotografico:= TfrxDBDataset.Create(nil);
    Td_Fotografico.UserName:='Td_Fotografico';

    Td_ImpFotografico:= TfrxDBDataset.Create(nil);
    Td_ImpFotografico.UserName:='Td_ImpFotografico';

    rxFotografico := TClientDataSet.Create(nil);
    rxFotografico.Active := False;
    rxFotografico.FieldDefs.Add('sContrato', ftMemo, 0, False);
    rxFotografico.CreateDataSet;
    rxFotografico.Active := True;
    rxFotografico.Append;
    rxFotografico.FieldValues['sContrato'] := '';
    rxFotografico.Post;

    if (TImpresion=FtsFotografico) or (TImpresion=FtsAll) then
    begin
        QryReporteFotografico.Active := False;
        QryReporteFotografico.SQL.Clear;
        QryReporteFotografico.SQL.Add('Select sContrato, sNumeroReporte, iImagen, bImagen, sDescripcion From reportefotografico ' +
          'Where sContrato = :Contrato And sNumeroReporte = :Reporte And lIncluye = "Si" Order By iImagen');
        QryReporteFotografico.Params.ParamByName('contrato').DataType := ftString;
        QryReporteFotografico.Params.ParamByName('contrato').Value    := reportediario.FieldByName('sOrden').AsString;;
        QryReporteFotografico.Params.ParamByName('Reporte').DataType  := ftString;
        QryReporteFotografico.Params.ParamByName('Reporte').Value     := reportediario.FieldByName('sNumeroReporte').AsString;;
        QryReporteFotografico.Open;

        Td_Fotografico.DataSet   := QryReporteFotografico;
        Td_ImpFotografico.DataSet:= QryReporteFotografico;

        {Esta secccion es para mostrar la hoja vacia sino existen datos.}
        if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
          if QryReporteFotografico.RecordCount = 0 then
             Td_ImpFotografico.DataSet := rxFotografico;
    end
    else
    begin
        rxFotografico.Delete;
        Td_Fotografico.DataSet   := rxFotografico;
        Td_ImpFotografico.DataSet:= rxFotografico;
    end;

    Td_Fotografico.FieldAliases.Clear;
    Td_ImpFotografico.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_Fotografico);
    Reporte.DataSets.Add(Td_ImpFotografico);

  finally

  end;
end;

Procedure EncabezadoPDF_Horizontal(Reportediario :TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
const
   Dias: array[1..7] of string = ('LUNES', 'MARTES', 'MIERCOLES', 'JUEVES', 'VIERNES', 'SABADO', 'DOMINGO');
Var
  zContrato, zConfiguracion,
  zEmbarcacion, zDuracion     : TZReadOnlyQuery;
  Td_contrato, Td_configuracion,
  Td_embarcacion, Td_duracion : TfrxDBDataset;
  iDia : integer;
  sDia : string;

begin
  Try

    Td_contrato:= TfrxDBDataset.Create(nil);
    Td_contrato.UserName:='td_contrato';

    Td_configuracion:= TfrxDBDataset.Create(nil);
    Td_configuracion.UserName:='dsConfiguracion';

    Td_embarcacion:= TfrxDBDataset.Create(nil);
    Td_embarcacion.UserName:='dsEmbarcacion';

    Td_duracion:= TfrxDBDataset.Create(nil);
    Td_duracion.UserName:='ds_Duracion';

    {Información del contrato}
    zContrato := TZReadOnlyQuery.Create(nil);
    zContrato.Connection := Connection.zConnection;
    zContrato.SQL.Add(' Select c.sLicitacion, c.sTitulo, c.sContrato, c.sTipoObra, c.mDescripcion, c.mCliente, c.bImagen, c.sUbicacion,  c.sCodigo, c.sLabelContrato, c.sCliente, '+
              'c.sIdResidencia, c.eLugarOT, m.sOficio, '+
              'a.bImagen as bImagenActivo, a.sDescripcion as sDescripcionActivo, c.mComentarios '+
              'FROM contratos c '+
              'inner join activos a on (c.sIdActivo = a.sIdActivo) '+
              'inner join residencias rs on (c.sIdResidencia = rs.sIdResidencia) '+
              'left join moe m on (m.sContrato =:Contrato and m.dIdFecha = (SELECT max(didfecha) FROM moe WHERE didfecha <=:Fecha AND sContrato = :contrato) ) '+
              'inner join configuracion c3 on (c.sContrato = c3.sContrato) '+
              'Where c.sContrato = :Contrato ');
    zContrato.ParamByName('fecha').AsDateTime  := reportediario.FieldByName('dIdfecha').AsDateTime;
    zContrato.ParamByName('contrato').AsString := reportediario.FieldByName('sOrden').AsString;
    zContrato.Open;

    {Información de la configuracion del sistema}
    zConfiguracion := TZReadOnlyQuery.Create(nil);
    zConfiguracion.Connection := Connection.zConnection;
    zConfiguracion.SQL.Add('select c.iFirmasReportes, c.sTipoPartida, c.sImprimePEP, ' +
              ' (select sContrato from contratos where sContrato =:contratobarco ) as sContratoBarco, ' +
              ' (select mDescripcion from contratos where sContrato =:contratobarco ) as mDescripcionBarco, ' +
              ' c.lLicencia, c.sReportesCIA, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3, c.iFirmasReportes, ' +
              ' if(c2.sContrato="BLP-OTt-002",if(:Fecha<DATE("2016-02-01"),c.bImagenAux1,c.bImagen),c.bImagen) as bImagen, c.sContrato, c.sNombre, c2.sCodigo, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, ' +
              ' c2.mDescripcion, c2.sTitulo, c2.mCliente, c2.bImagen as bImagenPEP ' +
              ' From contratos c2 '+
              ' INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
              ' Where c2.sContrato = :Contrato');
    zConfiguracion.Params.ParamByName('contrato').DataType      := ftString;
    zConfiguracion.Params.ParamByName('contrato').Value         := reportediario.FieldByName('sOrden').AsString;
    zConfiguracion.Params.ParamByName('contratobarco').DataType := ftString;
    zConfiguracion.Params.ParamByName('contratobarco').Value    := global_contrato_barco;
    zConfiguracion.ParamByName('fecha').AsDate  := reportediario.FieldByName('dIdfecha').AsDateTime;
    zConfiguracion.Open;

    {Busqueda de la embarcacion}
    zEmbarcacion := TZReadOnlyQuery.Create(nil);
    zEmbarcacion.Connection := Connection.zConnection;
    zEmbarcacion.SQL.Add('SELECT em.sDescripcion,em.sContrato, em.sIdEmbarcacion, cc.sCantidad as CantidadOlas,'+
                  'd.sDescripcion as DireccionOlas,cc.sLocalizacion, '+
                 // 'if(:orden=-1,cc.sLocalizacion,(select sLocalizacion from condicionesclimatologicas where sContrato =:orden and dIdFecha = :fecha and sIdTurno = :turno and iIdCondicion = 1 and sHorario = "24:00" group by iIdCondicion )) as sLocalizacion, '+

                  '(select sCantidad from condicionesclimatologicas where sContrato =:Orden and dIdFecha = :fecha and sIdTurno = :turno and iIdCondicion = 1 and sHorario = "24:00" group by iIdCondicion ) as CantidadViento, '+
                  '(select d2.sDescripcion from condicionesclimatologicas cc2 '+
                  'inner join direcciones d2 on (d2.iIdDireccion = cc2.iIdDireccion)'+
                  'where cc2.sContrato =:contrato and cc2.dIdFecha = :fecha and cc2.sIdTurno =:turno and cc2.iIdCondicion = 1  and cc2.sHorario = "24:00" group by iIdCondicion ) as DireccionViento, '+
                  '(select cd1.sMedida from condiciones cd1 where sDescripcion like "%viento%") as sMedidaCondicion1, '+
                  '(select cd1.sMedida from condiciones cd1 where sDescripcion like "%ola%") as sMedidaCondicion2 '+
                  'FROM embarcacion_vigencia AS ev '+
                  'INNER JOIN embarcaciones AS em ON (em.sIdEmbarcacion = ev.sIdEmbarcacion) '+
                  //'left join condicionesclimatologicas as cc on (cc.sContrato = ev.sContrato and cc.dIdFecha =:fecha and cc.sIdTurno =:turno and cc.iIdCondicion = 2 and cc.sHorario ="24:00") '+
                   'left join condicionesclimatologicas as cc on (cc.sContrato = :orden and cc.dIdFecha =:fecha and cc.sIdTurno =:turno and cc.iIdCondicion = 2 and cc.sHorario ="24:00") '+

                  'left join direcciones as d on (d.iIdDireccion = cc.iIdDireccion) '+
                  'WHERE ev.sContrato = :contrato and ev.dFechaInicio = (Select max(ev2.dfechainicio) '+
                  'from embarcacion_vigencia ev2 where ev.sContrato = em.sContrato and ev2.dfechainicio <= :fecha) order by sHorario DESC ');
    zEmbarcacion.ParamByName('Contrato').AsString := Global_Contrato_Barco;
    zEmbarcacion.ParamByName('fecha').AsDateTime  := reportediario.FieldByName('dIdfecha').AsDateTime;
    zEmbarcacion.ParamByName('turno').AsString    := reportediario.FieldByName('sIdTurno').AsString;
  //  if Tipo=FtAbordo then
  //    zEmbarcacion.ParamByName('Orden').AsInteger:=-1
  //  else
      zEmbarcacion.ParamByName('Orden').AsString :=ReporteDiario.FieldByName('sOrden').AsString;
    zEmbarcacion.Open;

    {Este bloque es para obtener el día}
     iDia := DayOfTheWeek(Reportediario.FieldByName('dIdFecha').AsDateTime);
     sDia := dias[iDia];

    {Consulto los días de contrato y de vigencia}
    zDuracion := TZReadOnlyQuery.Create(nil);
    zDuracion.Connection := Connection.zConnection;
    zDuracion.SQL.Add('SELECT ' +
                '	MIN(dFechaInicio) AS dInicioDeContrato, ' +
                '	MAX(dFechaFinal) AS dFinalDeContrato,   ' +
                '	DATEDIFF(MAX(dFechaFinal),MIN(dFechaInicio)) + 1 AS dDiasDeContrato, ' +
                '	DATEDIFF(MAX(dFechaFinal), :Hoy) AS dDiasRestantes,     ' +
                '	DATEDIFF(:Hoy, MIN(dFechaInicio)) + 1 AS dDiasTranscurridos, '+
                ' DATE_FORMAT(:hoy,"%d/%m/%Y") as dIdFecha, '+
                ' :Dia as DiaSemana '+
                ' FROM convenios WHERE sContrato = :Contrato');
    zDuracion.ParamByName('Hoy').AsDate        := Reportediario.FieldByName('dIdFecha').AsDateTime;
    zDuracion.ParamByName('Contrato').AsString := ReporteDiario.FieldByName('sContrato').AsString;
    zDuracion.ParamByName('Dia').AsString      := sDia;
    zDuracion.Open;

    Td_contrato.DataSet:=zContrato;
    Td_contrato.FieldAliases.Clear;
    Td_configuracion.DataSet:= zConfiguracion;
    Td_configuracion.FieldAliases.Clear;
    Td_embarcacion.DataSet:= zEmbarcacion;
    Td_embarcacion.FieldAliases.Clear;
    Td_duracion.DataSet:= zDuracion;
    Td_duracion.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_contrato);
    Reporte.DataSets.Add(Td_configuracion);
    Reporte.DataSets.Add(Td_embarcacion);
    Reporte.DataSets.Add(Td_duracion);

  Finally

  End;
end;

Procedure FirmasPDF_Generales(Reportediario :TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo);
Var
  zFirmas : TZReadOnlyQuery;
begin
  Try
    {Variables globales para Firmantes..}
    sCapitan              := 'SIN FIRMANTE';
    sSuperIntendente      := 'SIN FIRMANTE';
    sSuperIntendentePatio := 'SIN FIRMANTE';
    sRepresentanteTecnico := 'SIN FIRMANTE';
    sSupervisor           := 'SIN FIRMANTE';
    sSupervisorPatio      := 'SIN FIRMANTE';
    sSupervisorGenerador  := 'SIN FIRMANTE';
    sSupervisorEstimacion := 'SIN FIRMANTE';
    sSupervisorSubContratista   := 'SIN FIRMANTE';
    sResidente                  := 'SIN FIRMANTE';
    sPuestoSupervisorSubContratista := 'SIN PUESTO';
    sPuestoSuperintendente      := 'SIN PUESTO';
    sPuestoSupervisor           := 'SIN PUESTO';
    sPuestoSupervisorGenerador  := 'SIN PUESTO';
    sPuestoSupervisorEstimacion := 'SIN PUESTO';
    sSupervisorTierra           := 'SIN PUESTO';
    sPuestoSupervisorTierra     := 'SIN PUESTO';
    sPuestoRepresentanteTecnico := 'SIN PUESTO';
    sPuestoResidente            := 'SIN PUESTO';

    zFirmas := tzReadOnlyQuery.Create(nil);
    zFirmas.Connection := connection.zconnection;
    zFirmas.Active := False;
    zFirmas.SQL.Clear;
    if Reportediario.FieldByName('sNumeroOrden').AsString <> '' then
    begin
      zFirmas.SQL.Add('Select * from firmas where sContrato = :contrato and sIdTurno =:Turno and sNumeroOrden = :Orden And dIdFecha = :fecha');
      zFirmas.Params.ParamByName('Orden').DataType := ftString;
      zFirmas.Params.ParamByName('Orden').Value    := Reportediario.FieldByName('sNumeroOrden').AsString;
    end
    else
      zFirmas.SQL.Add('Select * from firmas where sContrato = :contrato and sIdTurno =:Turno And dIdFecha = :fecha');
    zFirmas.Params.ParamByName('Contrato').DataType := ftString;
    zFirmas.Params.ParamByName('Contrato').Value    := Reportediario.FieldByName('sOrden').AsString;
    zFirmas.Params.ParamByName('Turno').DataType    := ftString;
    zFirmas.Params.ParamByName('Turno').Value       := Reportediario.FieldByName('sIdTurno').AsString;
    zFirmas.Params.ParamByName('fecha').DataType    := ftDate;
    zFirmas.Params.ParamByName('fecha').Value       := Reportediario.FieldByName('dIdFecha').AsDateTime;
    zFirmas.Open;

    if zFirmas.RecordCount > 0 then
    begin
      sSuperintendente      := zFirmas.FieldValues['sFirmante1'];
      sSupervisor           := zFirmas.FieldValues['sFirmante2'];
      sSupervisorGenerador  := zFirmas.FieldValues['sFirmante3'];
      sSupervisorEstimacion := zFirmas.FieldValues['sFirmante4'];
      sSupervisorTierra     := zFirmas.FieldValues['sFirmante5'];
      sResidente            := zFirmas.FieldValues['sFirmante6'];
      sSuperintendentePatio := zFirmas.FieldValues['sFirmante7'];
      sSupervisorPatio      := zFirmas.FieldValues['sFirmante8'];
      sRepresentanteTecnico := zFirmas.FieldValues['sFirmante9'];
      sSupervisorSubContratista := zFirmas.FieldValues['sFirmante10'];
      sCapitan                  := zFirmas.FieldValues['sFirmante31'];

      sPuestoSuperintendente      := zFirmas.FieldValues['sPuesto1'];
      sPuestoSupervisor           := zFirmas.FieldValues['sPuesto2'];
      sPuestoSupervisorGenerador  := zFirmas.FieldValues['sPuesto3'];
      sPuestoSupervisorEstimacion := zFirmas.FieldValues['sPuesto4'];
      sPuestoSupervisorTierra     := zFirmas.FieldValues['sPuesto5'];
      sPuestoResidente            := zFirmas.FieldValues['sPuesto6'];
      sPuestoSuperintendentePatio := zFirmas.FieldValues['sPuesto7'];
      sPuestoSupervisorPatio      := zFirmas.FieldValues['sPuesto8'];
      sPuestoRepresentanteTecnico := zFirmas.FieldValues['sPuesto9'];
      sPuestoSupervisorSubContratista := zFirmas.FieldValues['sPuesto10'];
    end
    else
    begin
      zFirmas.Active := False;
      zFirmas.SQL.Clear;
      if Reportediario.FieldByName('sNumeroOrden').AsString <> '' then
      begin
        zFirmas.SQL.Add('Select * from firmas where sContrato = :contrato and sNumeroOrden = :Orden and sIdTurno =:Turno And dIdFecha <= :fecha Order By dIdFecha DESC');
        zFirmas.Params.ParamByName('Orden').DataType := ftString;
        zFirmas.Params.ParamByName('Orden').Value    := Reportediario.FieldByName('sNumeroOrden').AsString;
      end
      else
        zFirmas.SQL.Add('Select * from firmas where sContrato = :contrato and sIdTurno =:Turno And dIdFecha <= :fecha Order By dIdFecha DESC');

      zFirmas.Params.ParamByName('Contrato').DataType := ftString;
      zFirmas.Params.ParamByName('Contrato').Value    := Reportediario.FieldByName('sOrden').AsString;
      zFirmas.Params.ParamByName('Turno').DataType    := ftString;
      zFirmas.Params.ParamByName('Turno').Value       := Reportediario.FieldByName('sIdTurno').AsString;
      zFirmas.Params.ParamByName('fecha').DataType    := ftDate;
      zFirmas.Params.ParamByName('fecha').Value       := Reportediario.FieldByName('dIdFecha').AsDateTime;
      zFirmas.Open;

      if zFirmas.RecordCount > 0 then
      begin
        sSuperintendente      := zFirmas.FieldValues['sFirmante1'];
        sSupervisor           := zFirmas.FieldValues['sFirmante2'];
        sSupervisorGenerador  := zFirmas.FieldValues['sFirmante3'];
        sSupervisorEstimacion := zFirmas.FieldValues['sFirmante4'];
        sSupervisorTierra     := zFirmas.FieldValues['sFirmante5'];
        sResidente            := zFirmas.FieldValues['sFirmante6'];
        sSuperintendentePatio := zFirmas.FieldValues['sFirmante7'];
        sSupervisorPatio      := zFirmas.FieldValues['sFirmante8'];
        sRepresentanteTecnico := zFirmas.FieldValues['sFirmante9'];
        sSupervisorSubContratista := zFirmas.FieldValues['sFirmante10'];
        sCapitan                  := zFirmas.FieldValues['sFirmante31'];

        sPuestoSuperintendente      := zFirmas.FieldValues['sPuesto1'];
        sPuestoSupervisor           := zFirmas.FieldValues['sPuesto2'];
        sPuestoSupervisorGenerador  := zFirmas.FieldValues['sPuesto3'];
        sPuestoSupervisorEstimacion := zFirmas.FieldValues['sPuesto4'];
        sPuestoSupervisorTierra     := zFirmas.FieldValues['sPuesto5'];
        sPuestoResidente            := zFirmas.FieldValues['sPuesto6'];
        sPuestoSuperintendentePatio := zFirmas.FieldValues['sPuesto7'];
        sPuestoSupervisorPatio      := zFirmas.FieldValues['sPuesto8'];
        sPuestoRepresentanteTecnico := zFirmas.FieldValues['sPuesto9'];
        sPuestoSupervisorSubContratista := zFirmas.FieldValues['sPuesto10'];
      end
    end;
    zFirmas.Destroy;

  Finally

  End;
end;

Procedure ReportePDF_PartidasAnexoC_detalle(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);

Var
  i, x,y,CuentaCol,iGrupo,iCiclo: Integer;
  DCiclo,dTotal:Double;
  QryConsulta,
  QrMoe,
  QrColumnas,
  QrRecursos,
  QryCondiciones,
  QryEmbarcacion,
  QryOtroPersonal: TZQuery;
  QryAgrupador,
  qrReportado,
  qrPernoctas : TZQuery;

  dContrato_Inicio,
  dContrato_Final: TDateTime;
  mDatos:TJvMemoryData;
  mImprime:TJvMemoryData;
  Td_ImpDistribucion_detalle: TfrxDBDataset;
  Td_Distribucion_detalle: TfrxDBDataset;
  QrAdicional:TzReadOnlyquery;
  TmpAnexo:String;
  TmpDescAnexo:string;
  ValTmp:variant;
  CantTmp:Double;
  iPosTmp:Integer;
begin

  mDatos:=TJvMemoryData.Create(nil);
  mImprime:=TJvMemoryData.Create(nil);
  QrRecursos  := TZQuery.Create(nil);
  QrColumnas  := TZQuery.Create(nil);
  qrReportado := TZQuery.Create(nil);
  qrPernoctas := TZQuery.Create(nil);
  QrAdicional := TzReadOnlyquery.Create(nil);
  QrMoe       := TZQuery.Create(nil);
  Td_Distribucion_detalle:=TfrxDBDataset.Create(nil);
  Td_ImpDistribucion_detalle:=TfrxDBDataset.Create(nil);
  try
    QrMoe.Connection := Connection.zConnection;
    QrColumnas.Connection  := Connection.zConnection;
    QrRecursos.Connection  := Connection.zConnection;
    QrReportado.Connection := Connection.zConnection;
    QrPernoctas.Connection := Connection.zConnection;
    QrAdicional.Connection := Connection.zConnection;
    Td_Distribucion_detalle.UserName    :='Td_Distribucion_detalle';
    Td_ImpDistribucion_detalle.UserName :='Td_ImpDistribucion_detalle';
    with mDatos do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidAnexo', ftString, 10, false);
      FieldDefs.Add('sidrecurso', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 250, True);
      FieldDefs.Add('sAnexo', ftString, 250, false);
      FieldDefs.Add('sTitulo', ftString, 250, false);
      FieldDefs.Add('smedida', ftString, 100, True);
      FieldDefs.Add('sTipo', ftString, 10, false);
      FieldDefs.Add('dcantSol', FtFloat, 0, True);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      for CuentaCol:=1 to TotalCol do
      begin
        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPernocta' + Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPlataforma'+ Inttostr(CuentaCol), ftString, 100, false);
      end;
      Active:=true;
    end;
    with mImprime do
    begin
      Active:=false;
      FieldDefs.Add('iCampo', ftInteger, 0, True);
      Active:=true;
    end;

    mImprime.EmptyTable;
    if (TImpresion=FtsRRecursos_detalle) or (TImpresion=FtsAll) then
    begin

        {$REGION 'BARCO'}

        QrColumnas.active:=false;
        QrColumnas.SQL.Add( 'SELECT ot.sContrato, ot.sNumeroOrden, p.sDescripcion AS pernocta, p.sDescripcion AS pernocta, pf.sDescripcion AS plataforma, '+
                            'bp.sIdpernocta AS PernoctaP, bp.sIdPlataforma AS idPlataforma ' +
                            'FROM ordenesdetrabajo ot ' +
                            'INNER JOIN contratos AS c ON (ot.sContrato=c.sContrato) ' +
                            'INNER JOIN bitacoradepersonal AS bp ON (bp.scontrato=:OT AND bp.sNumeroOrden = ot.sNumeroOrden ) ' +
                            'INNER JOIN pernoctan AS p ON (p.sidPernocta=bp.sIdpernocta) ' +
                            'INNER JOIN plataformas AS pf ON (pf.sidPlataforma=bp.sIdPlataforma) ' +
                            'WHERE (c.sContrato=:OT OR c.sCodigo=:OT)	AND bp.dIdFecha= :Fecha ' +
                            'GROUP BY	ot.sContrato, ot.sNumeroorden, bp.sidPernocta, bp.sIdPlataforma');
        QrColumnas.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate      := Reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;

        qrMoe.Active:=false;
        QrMoe.SQL.Clear;
        qrMoe.SQL.Text:='select  t.sIdTipoMovimiento as sIdRecurso, t.*, a.sanexo,ifnull(sum(sFactor),0) as TotalFactor,a.sdescripcion as anexo, a.stitulo as tituloAnexo '+
                              'from tiposdemovimiento t '+
                              'inner join movimientosdeembarcacion m '+
                              'on (m.sContrato = t.sContrato and m.dIdFecha =:Fecha and m.sClasificacion = t.sIdTipoMovimiento and (m.sIdFase = "OPER" or m.sIdFase ="ESP")) '+
                              'left join anexos a on(a.sTipo= "BARCO") '+
                              'where t.sContrato =:contrato and t.sClasificacion = "Movimiento de Barco" group by m.sClasificacion order by t.iorden';
        qrMoe.ParamByName('Contrato').AsString := global_contrato_barco;
        qrMoe.ParamByName('Fecha').AsDate      := Reportediario.FieldByName('dIdFecha').AsDateTime;
        qrMoe.Open;

        qrRecursos.Active:=false;
        qrRecursos.SQL.Text:= 'select ifnull(sum(mf.sFactor),0) as dCantidad from movimientosxfolios mf' + #10 +
                              'inner join movimientosdeembarcacion me' + #10 +
                              'on(me.sContrato=mf.sContrato and me.dIdFecha=mf.dIdFecha and' + #10 +
                              'me.iIdDiario=mf.iIdDiario and me.sOrden=mf.sNumeroOrden)' + #10 +
                              'where mf.sContrato=:contratoBarco and mf.didfecha=:fecha and' + #10 +
                              'mf.sNumeroOrden=:contrato and mf.sFolio=:folio and me.sClasificacion=:Tipo';

        //movimientosxfolios where sContrato=:Contrato and dIdFecha=:fecha and iIdDiario=:Diario

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
            Inc(iCiclo,1);
        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
            with qrMoe do
            begin
                first;
                while not eof do
                begin
                    mDatos.Append;
                    mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
                    mDatos.FieldByName('sidAnexo').AsString     :=FieldByName('sanexo').asstring;
                    mDatos.FieldByName('sAnexo').AsString       :=FieldByName('anexo').asstring;
                    mDatos.FieldByName('sTitulo').AsString      :=FieldByName('tituloAnexo').asstring;
                    mDatos.FieldByName('sidrecurso').AsString   :=FieldByName('sidrecurso').asstring;
                    mDatos.FieldByName('sdescripcion').AsString :=fieldbyname('sdescripcion').asstring;
                    mDatos.FieldByName('smedida').AsString      :=fieldbyname('smedida').asstring;
                    mDatos.FieldByName('dcantSol').AsFloat      :=fieldbyname('dcantidad').AsFloat;
                    mDatos.FieldByName('sTipo').AsString        :='';
                    dTotal:=0;
                    if iGrupo=1 then
                       QrColumnas.First
                    else
                       QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

                    CuentaCol:=1;
                    while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
                    begin
                        mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString := qrColumnas.FieldbyName('sNumeroOrden').AsString;
                        mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString   := QrColumnas.Fieldbyname('Pernocta').asstring;
                        mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  := QrColumnas.Fieldbyname('Plataforma').Asstring;
                       // mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=QrRecursos.FieldByName('dCantidad').AsFloat;

                        QrRecursos.Active := False;
                        //QrRecursos.ParamByName('Equipo').AsString        := FieldByName('sIdRecurso').AsString;
                        QrRecursos.ParamByName('contratoBarco').AsString := global_contrato_barco;
                        QrRecursos.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
                        QrRecursos.ParamByName('folio').AsString         := QrColumnas.FieldByName('sNumeroOrden').AsString;
                        QrRecursos.ParamByName('fecha').AsDate           := reportediario.FieldByName('dIdFecha').AsDateTime;
                        QrRecursos.ParamByName('tipo').AsString          := FieldByName('sIdRecurso').AsString;
                        QrRecursos.Open;

                        if QrRecursos.RecordCount>0 then
                        begin
                          mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(QrRecursos.FieldByName('dCantidad').Value,4);
                          dTotal:=dTotal+ QrRecursos.FieldByName('dCantidad').AsFloat;
                        end;
                        QrColumnas.next;
                        Inc(CuentaCol);
                    end;

                    //if dTotal>0 then
                    //begin
                        mDatos.FieldByName('dcantTotal').AsFloat:=fieldbyname('TotalFactor').AsFloat;
                        mDatos.Post;
                    //end
                    //else
                    //   mDatos.cancel;
                    next;
                end;
            end;
            Inc(iGrupo,1);
        end;


        {$ENDREGION}

        {$REGION 'PERSONAL DE TERRA Y A BORDO'}
        QrRecursos.Active := False;
        QrRecursos.SQL.Clear;
        QrRecursos.SQL.Add('SELECT bp.sIdPersonal, bp.sDescripcion, ' +
                           'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                           'if(:Anexo=1,( ' +
                           'IF(SUM(bp.dCanthh) > 0, SUM(bp.dCanthh), SUM(bp.dCantidad))'+
                           '),SUM(bp.dCantidad)) AS Total ' +
                           'FROM bitacoradepersonal bp ' +
                           'WHERE bp.scontrato = :Orden AND bp.sNumeroOrden = :Folio AND bp.didfecha = :Fecha ' +
                           'AND bp.sidPernocta = :Pernocta AND bp.sidplataforma = :Plataforma ' +
                           'and bp.sIdPersonal=:Personal GROUP BY bp.sIdPersonal ');


        {$REGION 'CONSULTAS - PARTIDAS'}
        QrColumnas.active:=false;
        QrColumnas.SQL.Clear;
        QrColumnas.SQL.Add( 'SELECT ot.sContrato, ot.sNumeroOrden, p.sDescripcion AS pernocta, p.sDescripcion AS pernocta, pf.sDescripcion AS plataforma, '+
                            'bp.sIdpernocta AS PernoctaP, bp.sIdPlataforma AS idPlataforma ' +
                            'FROM ordenesdetrabajo ot ' +
                            'INNER JOIN contratos AS c ON (ot.sContrato=c.sContrato) ' +
                            'INNER JOIN bitacoradepersonal AS bp ON (bp.scontrato=:OT AND bp.sNumeroOrden = ot.sNumeroOrden ) ' +
                            'INNER JOIN pernoctan AS p ON (p.sidPernocta=bp.sIdpernocta) ' +
                            'INNER JOIN plataformas AS pf ON (pf.sidPlataforma=bp.sIdPlataforma) ' +
                            'WHERE (c.sContrato=:OT OR c.sCodigo=:OT)	AND bp.dIdFecha= :Fecha ' +
                            'GROUP BY	ot.sContrato, ot.sNumeroorden, bp.sidPernocta, bp.sIdPlataforma');
        QrColumnas.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;
        {$ENDREGION}


        {$REGION 'CONSULTAS - TODO EL PERSONAL SOLICITADO QUE SE REGISTRA EN EL MOE'}
        QrMoe.Active := False;
        QrMoe.SQL.Clear;   //ifnull(a.sdescripcion,"SIN ANEXO MAR/TIERRA") as anexo
        QrMoe.SQL.Add('SELECT p.sDescripcion,	mr.*, p.lSumaSolicitado, p.sMedida, ' +
                      'a.sanexo,ifnull(a.sdescripcion,"SIN ANEXO MAR/TIERRA") as anexo, a.stitulo as tituloAnexo,a.stierra '+
                      'FROM moe AS m ' +
                      'left JOIN moerecursos AS mr ON (mr.iidMoe=m.iidMoe) ' +
                      'INNER JOIN personal AS p ON (p.scontrato=:Contrato AND p.sidpersonal=mr.sidRecurso) ' +
                      'left join anexos a on(a.sAnexo=p.sAnexo) ' +
                      'WHERE m.didfecha = (SELECT max(didfecha) FROM moe WHERE didfecha <=:Fecha AND sContrato = :OT) ' +
                      'AND m.sContrato = :OT AND mr.eTipoRecurso = "Personal" ORDER BY a.iOrden,p.iItemOrden');
        QrMoe.ParamByName('Contrato').AsString := Global_Contrato_Barco;
        QrMoe.ParamByName('OT').AsString       := ReporteDiario.FieldByName('sOrden').AsString;
        QrMoe.ParamByName('Fecha').AsDateTime  := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrMoe.Open;
        {$ENDREGION}

        {$REGION 'INSERTA EL MOE'}


        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
          Inc(iCiclo,1);

        iGrupo:=1;
        TmpAnexo:='';
        TmpDescAnexo:='';
        while iGrupo<=iCiclo do
        begin
            with qrMoe do
            begin

                first;

                while not eof do
                begin
                    if (TmpAnexo='') and (FieldByName('sanexo').asstring<>'') then
                    begin
                      TmpAnexo:=FieldByName('sanexo').asstring;
                      TmpDescAnexo:=FieldByName('anexo').asstring;
                    end;

                    mDatos.Append;
                    mDatos.FieldByName('iGrupo').AsInteger      :=Igrupo;
                    mDatos.FieldByName('sidAnexo').AsString     :=TmpAnexo;//FieldByName('sanexo').asstring;
                    mDatos.FieldByName('sAnexo').AsString       :=TmpDescAnexo;//FieldByName('anexo').asstring;
                    mDatos.FieldByName('sTitulo').AsString      :=FieldByName('tituloAnexo').asstring;
                    mDatos.FieldByName('sidrecurso').AsString   :=FieldByName('sidrecurso').asstring;
                    mDatos.FieldByName('sdescripcion').AsString :=fieldbyname('sdescripcion').asstring;
                    mDatos.FieldByName('smedida').AsString      :=fieldbyname('smedida').asstring;
                    mDatos.FieldByName('dcantSol').AsFloat      :=fieldbyname('dcantidad').AsFloat;
                    mDatos.FieldByName('sTipo').AsString        :='';
                    dTotal:=0;
                    if iGrupo=1 then
                      QrColumnas.First
                    else
                      QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

                    CuentaCol:=1;
                    ValTmp:=0;
                    CantTmp:=0;
                    iPosTmp:=0;
                    while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
                    begin
                        mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString :=qrColumnas.FieldbyName('sNumeroOrden').AsString;
                        mDatos.FieldByName('sPernocta'   + Inttostr(CuentaCol)).AsString :=QrColumnas.Fieldbyname('Pernocta').asstring;
                        mDatos.FieldByName('sPlataforma' + Inttostr(CuentaCol)).AsString :=QrColumnas.Fieldbyname('Plataforma').Asstring;

                        QrRecursos.Active := False;
                        {if fieldbyname('sTierra').asstring='Si' then
                          QrRecursos.ParamByName('Anexo').AsINteger :=0
                        else }
                        QrRecursos.ParamByName('Anexo').AsINteger :=1;
                        QrRecursos.ParamByName('Orden').AsString      := QrColumnas.FieldByName('sContrato').AsString;
                        QrRecursos.ParamByName('Folio').AsString      := QrColumnas.FieldByName('sNumeroOrden').AsString;
                        QrRecursos.ParamByName('Fecha').AsDate        := reportediario.FieldByName('dIdFecha').AsDateTime;
                        QrRecursos.ParamByName('Pernocta').AsString   := QrColumnas.FieldByName('PernoctaP').AsString;
                        QrRecursos.ParamByName('Plataforma').AsString := QrColumnas.FieldByName('idPlataforma').AsString;
                        QrRecursos.ParamByName('Personal').AsString   := FieldByName('sIdRecurso').AsString;
                        QrRecursos.Open;

                        if QrRecursos.RecordCount>0 then
                        begin
                          mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=(xRound(QrRecursos.FieldByName('Total').Value,2) + QrRecursos.FieldByName('Ajuste').Value);
                          dTotal:=dTotal+ (xRound(QrRecursos.FieldByName('Total').Value,2) + QrRecursos.FieldByName('Ajuste').Value);
                          ValTmp:= ValTmp + (QrRecursos.FieldByName('Total').Value + QrRecursos.FieldByName('Ajuste').Value);
                          if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                          begin
                            CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                            iPosTmp:=CuentaCol;
                          end;
                        end;
                        QrColumnas.next;
                        Inc(CuentaCol);
                    end;

                    if dTotal>=0 then
                    begin
                      if dTotal<>xRound(ValTmp,2) then
                      begin
                        mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                          + (xRound(ValTmp,2)-dTotal);
                        mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2);
                      end
                      else
                        mDatos.FieldByName('dcantTotal').AsFloat:=dTotal;
                      mDatos.Post;
                    end
                    else
                      mDatos.Cancel;
                    next;
                end;
            end;
            Inc(iGrupo,1);
        end;

        {$ENDREGION}
        {$ENDREGION}

        {$REGION 'EQUIPOS..'}
        {Equipos...}
        QrRecursos.Active := False;
        QrRecursos.SQL.Clear;
        QrRecursos.SQL.Add( 'select be.sIdEquipo, (sum(be.dCantHH) + ' +
                            '  IFNULL((SELECT ' +
                            '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                            '          FROM ' +
                            '            bitacoradeajustes ba ' +
                            '          WHERE ' +
                            '            ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                            '            ba.sNumeroActividad = be.sNumeroActividad AND ' +
                            '            ba.sIdPartidaAnexo = be.sIdEquipo AND be.dFecha = be.dIdFecha), 0)) as dCantidad '+
                            'from bitacoradeequipos be '+
                            'inner join equipos e '+
                            'on ( e.sIdEquipo = be.sIdEquipo ) '+
                            'inner join bitacoradeactividades ba '+
                            'on ( ba.sContrato = :contrato and ba.dIdFecha = :fecha '+
                                 'and ba.sNumeroOrden = :folio '+
                                 'and ba.iIdDiario = be.iIdDiario '+
                            'and ba.iIdTarea = be.iIdTarea and ba.iIdActividad = be.iIdActividad) '+
                            'where e.sContrato = :contratoBarco '+
                            'and be.sContrato = :contrato '+
                            'and be.sNumeroOrden = :folio '+
                            'and be.sIdPernocta = :pernocta '+
                            'and be.dIdFecha = :fecha '+
                            'and be.sIdEquipo=:equipo ' +
                            'group by e.sIdEquipo '+
                            'order by e.iItemOrden');

        {$REGION 'CONSULTAS - PARTIDAS'}
        QrColumnas.Active:=false;
        QrColumnas.SQL.Clear;
        QrColumnas.SQL.Add( 'select ot.sContrato,ot.sIdFolio, '+
                                     'ot.sNumeroOrden, '+
                                     'be.sIdPernocta AS PernoctaP, '+
                                     'ot.sIdPlataforma AS idPlataforma, '+
                                     'p.sDescripcion as Pernocta, '+
                                     'pt.sDescripcion as Plataforma '+
                              'from ordenesdetrabajo ot '+
                              'inner join contratos c '+
                              'on ( c.sContrato = ot.sContrato ) '+
                              'inner join bitacoradeequipos be '+
                              'on ( be.sContrato = ot.sContrato and be.sNumeroOrden = ot.sNumeroOrden ) '+
                              'inner join pernoctan p '+
                              'on ( ot.sIdPernocta = p.sIdPernocta) '+
                              'inner join plataformas pt '+
                              'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                              'where c.sContrato = :contrato '+
                              'and be.dIdFecha = :fecha '+
                              'group by ot.sIdFolio, p.sIdPernocta' );
        QrColumnas.ParamByName('Contrato').AsString:= ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate:=reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;
        {$ENDREGION}


        {$REGION 'CONSULTAS - TODOS LOS EQUIPOS REGISTRADOS EN MOE'}

        QrMoe.Active := False;
        QrMoe.SQL.Clear;
        QrMoe.SQL.Add('select mr.sIdRecurso, e.sDescripcion, e.sMedida, mr.dCantidad, '+
                      'a.sanexo,a.sdescripcion as anexo, a.stitulo as tituloAnexo '+
                        'from moe m '+
                        'inner join moerecursos mr '+
                        'on ( mr.iIdMoe = m.iIdMoe ) '+
                        'inner join equipos e '+
                        'on ( e.sContrato = :contratobarco and e.sIdEquipo = mr.sIdRecurso ) '+
                        'left join anexos a on(a.sTipo= "EQUIPO") ' +
                        'where m.dIdFecha = (select max(didfecha) from moe where didfecha <=:Fecha and sContrato = :contrato) '+
                        'and m.sContrato = :contrato '+
                        'and mr.eTipoRecurso = "Equipo" '+
                        'order by e.iItemOrden');
        QrMoe.ParamByName('contratobarco').AsString := Global_Contrato_Barco;
        QrMoe.ParamByName('contrato').AsString := ReporteDiario.FieldByName('sOrden').AsString;
        QrMoe.ParamByName('Fecha').AsDateTime := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrMoe.Open;

        {$ENDREGION}

        {$REGION 'INSERCION DE DATOS - INFORMACIÓN DEL EQUIPO'}

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
            Inc(iCiclo,1);
        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
            with qrMoe do
            begin
                first;
                while not eof do
                begin
                    mDatos.Append;
                    mDatos.FieldByName('iGrupo').AsInteger       :=Igrupo;
                    mDatos.FieldByName('sidAnexo').AsString      :=FieldByName('sanexo').asstring;
                    mDatos.FieldByName('sAnexo').AsString        :=FieldByName('anexo').asstring;
                    mDatos.FieldByName('sTitulo').AsString       :=FieldByName('tituloAnexo').asstring;
                    mDatos.FieldByName('sidrecurso').AsString    :=FieldByName('sidrecurso').asstring;
                    mDatos.FieldByName('sdescripcion').AsString  :=fieldbyname('sdescripcion').asstring;
                    mDatos.FieldByName('smedida').AsString       :=fieldbyname('smedida').asstring;
                    mDatos.FieldByName('dcantSol').AsFloat       :=fieldbyname('dcantidad').AsFloat;
                     mDatos.FieldByName('sTipo').AsString        :='';
                    dTotal:=0;
                    if iGrupo=1 then
                       QrColumnas.First
                    else
                       QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

                    CuentaCol:=1;
                    ValTmp:=0;
                    CantTmp:=0;
                    iPosTmp:=0;
                    while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
                    begin
                        mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString := qrColumnas.FieldbyName('sNumeroOrden').AsString;
                        mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString   := QrColumnas.Fieldbyname('Pernocta').asstring;
                        mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  := QrColumnas.Fieldbyname('Plataforma').Asstring;

                        QrRecursos.Active := False;
                        QrRecursos.ParamByName('Equipo').AsString        := FieldByName('sIdRecurso').AsString;
                        QrRecursos.ParamByName('contratoBarco').AsString := global_contrato_barco;
                        QrRecursos.ParamByName('contrato').AsString      := ReporteDiario.FieldByName('sOrden').AsString;
                        QrRecursos.ParamByName('folio').AsString         := QrColumnas.FieldByName('sNumeroOrden').AsString;
                        QrRecursos.ParamByName('fecha').AsDate           := reportediario.FieldByName('dIdFecha').AsDateTime;
                        QrRecursos.ParamByName('pernocta').AsString      := QrColumnas.FieldByName('PernoctaP').AsString;
                        QrRecursos.Open;
                        //ABBY
                        if QrRecursos.RecordCount>0 then
                        begin
                          mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(QrRecursos.FieldByName('dCantidad').Value,2);
                          dTotal:=dTotal+ xRound(QrRecursos.FieldByName('dCantidad').Value,2);
                          ValTmp:=ValTmp + QrRecursos.FieldByName('dCantidad').Value;
                          if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                          begin
                            CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                            iPosTmp:=CuentaCol;
                          end;
                        end;
                        QrColumnas.next;
                        Inc(CuentaCol);
                    end;

                    if dTotal>0 then
                    begin
                      if dTotal<>xRound(ValTmp,2) then
                      begin
                        mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                        + (xRound(ValTmp,2)-dTotal);
                        mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2);
                      end
                      else
                        mDatos.FieldByName('dcantTotal').AsFloat:=dTotal;
                      mDatos.Post;
                    end
                    else
                       mDatos.cancel;
                    next;
                end;
            end;
            Inc(iGrupo,1);
        end;


     {Esta secccion es para mostrar la hoja vacia sino existen datos.}
//     if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
//        if MDatos.RecordCount>=0 then
//     else
//        if MDatos.RecordCount>0 then
//     begin
//
//     end;

        mImprime.Append;
        mImprime.FieldByName('iCampo').AsInteger:=1;
        mImprime.Post;

        {$ENDREGION}
        {$ENDREGION}

        {$REGION 'PERNOCTAS..'}
        with QrColumnas do
        begin
          active := false;
          sql.text := 'select ot.sIdFolio, '+
                             'ot.sNumeroOrden, '+
                             'ot.sIdPernocta, '+
                             'ot.sIdPlataforma, '+
                             'p.sDescripcion as sPernocta, '+
                             'pt.sDescripcion as sPlataforma '+
                      'from ordenesdetrabajo ot '+
                      'inner join contratos c '+
                      'on ( c.sContrato = ot.sContrato ) '+
                      'inner join bitacoradeequipos be '+
                      'on ( be.sContrato = ot.sContrato and be.sNumeroOrden = ot.sNumeroOrden ) '+
                      'inner join pernoctan p '+
                      'on ( ot.sIdPernocta = p.sIdPernocta) '+
                      'inner join plataformas pt '+
                      'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                      'where c.sContrato = :contrato '+
                      'and be.dIdFecha = :fecha '+
                      'group by ot.sIdFolio, p.sIdPernocta';
          parambyname('contrato').asstring := ReporteDiario.fieldbyname('sOrden').asstring;
          parambyname('fecha').asdate      := reportediario.fieldbyname('didfecha').asdatetime;
          open;
        end;

        with qrPernoctas do
        begin
          active := false;
          sql.text := 'select c.sIdPernocta, '+
                             'c.sDescripcion, '+
                             'c.sMedida, '+

                       '( (sum( bp.dCantHH ) + ' +
                       ' IFNULL((SELECT ' +
                       '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                       '          FROM ' +
                       '            bitacoradeajustes ba ' +
                       '          WHERE ' +
                       '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                       '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                       '            ba.sIdPartidaAnexo = c.sIdCuenta AND ba.dFecha = bp.dIdFecha), 0)) ' +
                       ' - ifnull(( select ifnull( sum(bpr.dCantidad), 0) '+
                                               'from bitacoradepernocta bpr '+
                                               'where bpr.sContrato = :contrato '+
                                               'and bpr.dIdFecha = :fecha '+
                                               'and (bpr.sNumeroOrden <> "@" and bpr.sNumeroOrden=ba.snumeroorden) group by bpr.dIdFecha),0) ) as dCantidad '+

                      'from cuentas c '+
                      'left join bitacoradepersonal bp '+
                        'on ( '+ //bp.lAplicaPernocta = "Si"
                          'bp.sContrato = :contrato '+
                          'and bp.dIdFecha = :fecha '+
                          'and bp.sTipoPernocta = c.sIdPernocta ) '+

                      'left join moerecursos mr '+
                        'on ( mr.sIdRecurso = bp.sIdPersonal '+
                          'and mr.eTipoRecurso = "Personal" '+
                          'and mr.iIdMoe = ( select m.iIdMoe from moe m where m.sContrato = :contrato '+
                                            'and m.dIdFecha <= :fecha order by m.dIdFecha desc limit 1) ) '+

                      'left join personal p '+
                        'on ( p.sContrato = :contratoBarco '+
                          'and p.sIdPersonal = bp.sIdPersonal ) '+

                      'left join tiposdepersonal tp '+
                        'on ( p.sIdTipoPersonal = tp.sIdTipoPersonal ) '+

                      'left join bitacoradeactividades ba '+
                      'on (bp.sContrato =ba.sContrato and ba.sNumeroOrden=bp.sNumeroOrden and '+
                      'bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                      'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad ' +
                      'and ba.sIdTipoMovimiento = "E" ) '+

                      'left join ordenesdetrabajo ot '+
                        'on ( ot.sContrato = :contrato and ot.sNumeroOrden = ba.sNumeroOrden ) '+

                      'left join pernoctan pr '+
                        'on ( pr.sIdPernocta = ot.sIdPernocta ) '+

                      'left join plataformas pl '+
                        'on ( pl.sIdPlataforma = ot.sIdPlataforma ) '+
                      'where bp.snumeroorden=:Folio and c.sidpernocta=:Pernocta ' +
                      'and p.lpernocta="Si" ' +
                      //'where bp.sContrato = :contrato '+
                      //'and bp.dIdFecha = :fecha '+

                      'group by c.sIdPernocta '+
                      'order by c.sIdPernocta';
          parambyname('contrato').asstring := ReporteDiario.FieldByName('sOrden').asstring;
          parambyname('fecha').asdate := ReporteDiario.FieldByName('dIdFecha').asDatetime;
          parambyname('contratoBarco').asstring := global_contrato_barco;

        end;

        QrReportado.Active:=false;
        QrReportado.SQL.Text:='select c.*, a.sanexo,a.sdescripcion as anexo, a.stitulo as tituloAnexo '+
                              'from cuentas c '+
                              'left join anexos a on(a.sTipo= "PERNOCTA")';
        QrREportado.Open;


        QrAdicional.Active:=false;

        QrAdicional.SQL.Text:='select ' +
                              '  (ifnull(sum(bpr.dCantidad),0) + ' +
                              '   IFNULL((SELECT ' +
                              '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                              '           FROM ' +
                              '             bitacoradeajustes ba ' +
                              '           WHERE ' +
                              '             ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                              '             ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0)) as dCantidad ' +
                              'from ' +
                              '  bitacoradepernocta bpr ' +
                              'where ' +
                              '  bpr.sContrato=:Contrato and ' +
                              '  bpr.dIdFecha=:fecha and ' +
                              '  bpr.sNumeroOrden=:Folio and ' +
                              '  bpr.sIdCuenta=:Pernocta ';

        {QrAdicional.SQL.Text:='select ifnull(sum(dCantidad),0) as dCantidad from bitacoradepernocta where ' +
                      'sContrato=:Contrato and dIdFecha=:fecha and sNumeroOrden=:Folio and '+
                      'sIdCuenta=:Pernocta ';}

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
          Inc(iCiclo,1);

        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
          QrReportado.First;
          while not QrReportado.Eof do
          begin
            dTotal:=0;
            if iGrupo=1 then
              QrColumnas.First
            else
              QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

            mDatos.Append;
            mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
            mDatos.FieldByName('sidAnexo').AsString     := QrReportado.FieldByName('sanexo').asstring;
            mDatos.FieldByName('sAnexo').AsString       := QrReportado.FieldByName('anexo').asstring;
            mDatos.FieldByName('sTitulo').AsString      := QrReportado.FieldByName('tituloAnexo').asstring;
            mDatos.FieldByName('sidrecurso').AsString   := QrReportado.FieldByName('sidpernocta').asstring;
            mDatos.FieldByName('sdescripcion').AsString := QrReportado.fieldbyname('sdescripcion').asstring;
            mDatos.FieldByName('smedida').AsString      := QrReportado.fieldbyname('smedida').asstring;
            mDatos.FieldByName('dcantSol').AsFloat      :=0;
             mDatos.FieldByName('sTipo').AsString       :='';

            CuentaCol:=1;
            ValTmp:=0;
            CantTmp:=0;
            iPosTmp:=0;
            while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
            begin
              mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString:=qrColumnas.FieldbyName('snumeroorden').AsString;
              mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString:=QrColumnas.Fieldbyname('splataforma').Asstring;

              with QrPernoctas do
              begin
                Active:=false;
                parambyname('Folio').AsString:=qrColumnas.FieldbyName('snumeroorden').AsString;
                parambyname('Pernocta').AsString:=QrReportado.FieldByName('sIdPernocta').AsString;
                Open;

                if Recordcount=0 then
                begin
                  ///Aqui va
                  QrAdicional.Active:=false;
                  QrAdicional.ParamByName('Contrato').AsString:=ReporteDiario.FieldByName('sOrden').asstring;
                  QrAdicional.ParamByName('Folio').AsString:=qrColumnas.FieldbyName('snumeroorden').AsString;
                  QrAdicional.ParamByName('Fecha').Asdate:=ReporteDiario.FieldByName('dIdFecha').asDatetime;
                  QrAdicional.ParamByName('Pernocta').AsString:= QrReportado.FieldByName('sIdCuenta').AsString;
                  QrAdicional.Open;
                  if QrAdicional.RecordCount=1 then
                  begin
                    mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=
                    xRound(QrAdicional.FieldByName('dCantidad').value,2);
                    dTotal:=dTotal+ xRound(QrAdicional.FieldByName('dCantidad').Value,2);
                    ValTmp:= ValTmp + (QrAdicional.FieldByName('dCantidad').Value);
                    if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                    begin
                      CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                      iPosTmp:= CuentaCol;
                    end;
                  end

                end;


                first;
                while not eof do
                begin

                  QrAdicional.Active:=false;
                  QrAdicional.ParamByName('Contrato').AsString:=ReporteDiario.FieldByName('sOrden').asstring;
                  QrAdicional.ParamByName('Folio').AsString:=qrColumnas.FieldbyName('snumeroorden').AsString;
                  QrAdicional.ParamByName('Fecha').Asdate:=ReporteDiario.FieldByName('dIdFecha').asDatetime;
                  QrAdicional.ParamByName('Pernocta').AsString:= QrReportado.FieldByName('sIdCuenta').AsString;
                  QrAdicional.Open;

                  if QrAdicional.RecordCount=1 then
                  begin
                    mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(FieldByName('dCantidad').Value,2)+
                    QrAdicional.FieldByName('dCantidad').AsFloat;
                    dTotal:=dTotal+ xRound(FieldByName('dCantidad').Value,2) +QrAdicional.FieldByName('dCantidad').AsFloat;
                    ValTmp:= ValTmp + (FieldByName('dCantidad').Value + QrAdicional.FieldByName('dCantidad').AsFloat);
                    if CantTmp < mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                    begin
                      CantTmp := mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                      iPosTmp:=CuentaCol;
                    end;
                  end
                  else
                  begin
                    mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(FieldByName('dCantidad').Value,2);
                    dTotal:=dTotal+ xRound(FieldByName('dCantidad').Value,2);
                    ValTmp:= ValTmp + FieldByName('dCantidad').Value;
                    if CantTmp < mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                    begin
                      CantTmp := mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                      iPosTmp:=CuentaCol;
                    end;
                  end;
                  next;
                end;
                //next;
              end;
              Inc(CuentaCol);
              QrColumnas.Next;
            end;
            if dTotal<>xRound(ValTmp,2) then
            begin
              mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
              + (xRound(ValTmp,2)-dTotal);
              mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,2);
            end
            else
              mDatos.FieldByName('dcantTotal').AsFloat:=dTotal;
            mDatos.Post;
            QrReportado.next;
          end;
          Inc(iGrupo,1);
        end;
        {$ENDREGION}

        {$REGION 'PARTIDAS DE PU'}

         {$REGION 'PARTIDAS PU.'}
        {Equipos...}
        QrRecursos.Active := False;
        QrRecursos.SQL.Clear;
        QrRecursos.SQL.Add( 'select b.sNumeroActividad, sum(b.dCantidad) as dCantidad '+
                            'from bitacoradeactividades b '+
                            'where b.sContrato = :Contrato and b.sNumeroOrden = :Orden and b.dIdFecha = :Fecha '+
                            'and b.sIdTipoMovimiento = "E" and b.sWbs = :Wbs group by b.sNumeroActividad');

        {$REGION 'CONSULTAS - PARTIDAS'}
        QrColumnas.Active:=false;
        QrColumnas.SQL.Clear;
        QrColumnas.SQL.Add( 'select ot.sContrato,ot.sIdFolio, '+
                                     'ot.sNumeroOrden, '+
                                     'ot.sIdPlataforma AS idPlataforma, '+
                                     'p.sDescripcion as Pernocta, '+
                                     'pt.sDescripcion as Plataforma '+
                              'from ordenesdetrabajo ot '+
                              'inner join contratos c '+
                              'on ( c.sContrato = ot.sContrato ) '+
                              'inner join bitacoradeactividades ba '+
                              'on ( ba.sContrato = ot.sContrato and ba.sNumeroOrden = ot.sNumeroOrden ) '+
                              'inner join pernoctan p '+
                              'on ( ot.sIdPernocta = p.sIdPernocta) '+
                              'inner join plataformas pt '+
                              'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                              'where c.sContrato = :contrato '+
                              'and ba.dIdFecha = :fecha '+
                              'group by ot.sIdFolio, p.sIdPernocta' );
        QrColumnas.ParamByName('Contrato').AsString:= ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate:=reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;
        {$ENDREGION}

        QrMoe.Active := False;
        QrMoe.SQL.Clear;
        QrMoe.SQL.Add('select b.sNumeroActividad as sIdRecurso, b.mDescripcion, a.sMedida, a.dCantidad, a.sAnexo, b.sWbs, '+
                      'n.sdescripcion as anexo, n.stitulo as tituloAnexo, n.sTipo '+
                      'from bitacoradeactividades b '+
                      'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sIdConvenio = b.sIdConvenio and a.sNumeroOrden  = a.sNumeroOrden '+
                      'and a.sTipoAnexo= "PU" and a.sNumeroActividad = b.sNumeroActividad and a.sWbs = b.sWbs and a.sTipoActividad = "Actividad") '+
                      'inner join anexos n on (a.sAnexo = n.sAnexo) '+
                      'where b.sContrato = :Contrato and b.dIdFecha = :Fecha '+
                      'and b.sIdTipoMovimiento = "E" group by n.iOrden, b.sContrato, b.sNumeroOrden, a.sNumeroActividad order by a.iItemOrden ');
        QrMoe.ParamByName('contrato').AsString := ReporteDiario.FieldByName('sOrden').AsString;
        QrMoe.ParamByName('Fecha').AsDateTime  := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrMoe.Open;


        {$REGION 'INSERCION DE DATOS - INFORMACIÓN DEL PU'}

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
            Inc(iCiclo,1);
        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
            with qrMoe do
            begin
                first;
                while not eof do
                begin
                    mDatos.Append;
                    mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
                    mDatos.FieldByName('sidAnexo').AsString     :=FieldByName('sanexo').asstring;
                    mDatos.FieldByName('sAnexo').AsString       :=FieldByName('anexo').asstring;
                    mDatos.FieldByName('sTitulo').AsString      :=FieldByName('tituloAnexo').asstring;
                    mDatos.FieldByName('sidrecurso').AsString   :=FieldByName('sidrecurso').asstring;
                    mDatos.FieldByName('sdescripcion').AsString :=fieldbyname('mdescripcion').asstring;
                    mDatos.FieldByName('smedida').AsString      :=fieldbyname('smedida').asstring;
                    mDatos.FieldByName('dcantSol').AsFloat      :=fieldbyname('dcantidad').AsFloat;
                    mDatos.FieldByName('sTipo').AsString        :=FieldByName('sTipo').asstring;
                    dTotal:=0;
                    if iGrupo=1 then
                       QrColumnas.First
                    else
                       QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

                    CuentaCol:=1;
                    ValTmp:=0;
                    CantTmp:=0;
                    iPosTmp:=0;
                    while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
                    begin
                        mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString := qrColumnas.FieldbyName('sNumeroOrden').AsString;
                        mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString   := QrColumnas.Fieldbyname('Pernocta').asstring;
                        mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  := QrColumnas.Fieldbyname('Plataforma').Asstring;

                        QrRecursos.Active := False;
                        QrRecursos.ParamByName('Wbs').AsString      := FieldByName('sWbs').AsString;
                        QrRecursos.ParamByName('contrato').AsString := ReporteDiario.FieldByName('sOrden').AsString;
                        QrRecursos.ParamByName('Orden').AsString    := QrColumnas.FieldByName('sNumeroOrden').AsString;
                        QrRecursos.ParamByName('fecha').AsDate      := reportediario.FieldByName('dIdFecha').AsDateTime;
                        QrRecursos.Open;
                        //ABBY
                        if QrRecursos.RecordCount>0 then
                        begin
                          mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(QrRecursos.FieldByName('dCantidad').Value,4);
                          dTotal:=dTotal+ xRound(QrRecursos.FieldByName('dCantidad').Value,4);
                          ValTmp:=ValTmp + QrRecursos.FieldByName('dCantidad').Value;
                          if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                          begin
                            CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                            iPosTmp:=CuentaCol;
                          end;
                        end;
                        QrColumnas.next;
                        Inc(CuentaCol);
                    end;

                    if dTotal>=0 then
                    begin
                      if dTotal<>xRound(ValTmp,4) then
                      begin
                        mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                        + (xRound(ValTmp,4)-dTotal);
                        mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,4);
                      end
                      else
                        mDatos.FieldByName('dcantTotal').AsFloat:=dTotal;
                      mDatos.Post;
                    end
                    else
                       mDatos.cancel;
                    next;
                end;
            end;
            Inc(iGrupo,1);
        end;
        {$ENDREGION}

    end;
    Td_Distribucion_detalle.DataSet:=MDatos;
    Td_Distribucion_detalle.FieldAliases.Clear;
    Td_ImpDistribucion_detalle.DataSet:=MImprime;
    Td_ImpDistribucion_detalle.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_Distribucion_detalle);
    Reporte.DataSets.Add(Td_ImpDistribucion_detalle);
  finally
    QrRecursos.destroy;
    QrColumnas.destroy;
    QrMoe.destroy;
  end;
end;


Procedure ReportePDF_PartidasAnexoC_detalleV2(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);

Var
  i, x,y,CuentaCol,iGrupo,iCiclo: Integer;
  DCiclo,dTotal:Double;
  QryConsulta,
  QrMoe,
  QrColumnas,
  QrRecursos,
  QryCondiciones,
  QryEmbarcacion,
  QryOtroPersonal: TZQuery;
  QryAgrupador,
  qrReportado,
  qrPernoctas : TZQuery;

  ZFolios : TZReadOnlyQuery;

  dContrato_Inicio,
  dContrato_Final: TDateTime;
  mDatos:TJvMemoryData;
  mImprime:TJvMemoryData;
  Td_ImpDistribucion_detalle2: TfrxDBDataset;
  Td_Distribucion_detalle2: TfrxDBDataset;
  QrAdicional:TzReadOnlyquery;
  TmpAnexo:String;
  TmpDescAnexo:string;
  ValTmp:variant;
  CantTmp:Double;
  iPosTmp:Integer;
  sTituloAnexo : string;
begin

  mDatos:=TJvMemoryData.Create(nil);
  mImprime:=TJvMemoryData.Create(nil);
  QrRecursos  := TZQuery.Create(nil);
  QrColumnas  := TZQuery.Create(nil);
  qrReportado := TZQuery.Create(nil);
  qrPernoctas := TZQuery.Create(nil);
  QrAdicional := TzReadOnlyquery.Create(nil);
  QrMoe       := TZQuery.Create(nil);
  ZFolios := TZReadOnlyQuery.Create( nil );
  Td_Distribucion_detalle2:=TfrxDBDataset.Create(nil);
  Td_ImpDistribucion_detalle2:=TfrxDBDataset.Create(nil);
  try
    QrMoe.Connection := Connection.zConnection;
    QrColumnas.Connection  := Connection.zConnection;
    QrRecursos.Connection  := Connection.zConnection;
    QrReportado.Connection := Connection.zConnection;
    QrPernoctas.Connection := Connection.zConnection;
    QrAdicional.Connection := Connection.zConnection;
    ZFolios.Connection := connection.zConnection;
    Td_Distribucion_detalle2.UserName    :='Td_Distribucion_detalle2';
    Td_ImpDistribucion_detalle2.UserName :='Td_ImpDistribucion_detalle2';
    with mDatos do
    begin
      Active:=false;
      FieldDefs.Add('iGrupo', ftInteger, 0, True);
      FieldDefs.Add('sidAnexo', ftString, 10, false);
      FieldDefs.Add('sidrecurso', ftString, 100, True);
      FieldDefs.Add('sdescripcion', ftString, 500, True);
      FieldDefs.Add('sAnexo', ftString, 250, false);
      FieldDefs.Add('sTitulo', ftString, 250, false);
      FieldDefs.Add('smedida', ftString, 100, True);
      FieldDefs.Add('sTipo', ftString, 10, false);
      FieldDefs.Add('dcantSol', ftString, 255, False);
      FieldDefs.Add('dcantTotal', FtFloat, 0, True);
      FieldDefs.Add('dcantidad0', FtFloat, 0, false);
      for CuentaCol:=1 to TotalCol do
      begin
        FieldDefs.Add('sPda' + Inttostr(CuentaCol), ftString, 50, false);

        FieldDefs.Add('dcantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
        FieldDefs.Add('sNumeroOrden'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPernocta' + Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sPlataforma'+ Inttostr(CuentaCol), ftString, 100, false);
        FieldDefs.Add('sTipo'+ Inttostr(CuentaCol), ftString, 10, false);
      end;
      Active:=true;
    end;
    with mImprime do
    begin
      Active:=false;
      FieldDefs.Add('iCampo', ftInteger, 0, True);
      Active:=true;
    end;

    mImprime.EmptyTable;
    if (TImpresion=FtsRRecursos_detalleV2) or (TImpresion=FtsAll) then
    begin
      {$REGION 'PARTIDAS DE PU'}

        mImprime.Append;
        mImprime.FieldByName('iCampo').AsInteger:=1;
        mImprime.Post;

        {Equipos...}
        QrRecursos.Active := False;
        QrRecursos.SQL.Clear;
        QrRecursos.SQL.Add( 'select b.sIdMaterial, sum(b.dCantidad) as dCantidad '+
                            'from bitacorademateriales b '+
                            'where b.sContrato = :Contrato and b.sNumeroOrden = :Orden and b.dIdFecha = :Fecha and b.sIdMaterial =:Actividad '+
                            'group by b.sIdMaterial');

        {$REGION 'CONSULTAS - PARTIDAS'}
        QrColumnas.Active:=false;
        QrColumnas.SQL.Clear;
        QrColumnas.SQL.Add( 'select ot.sContrato,ot.sIdFolio, ot.sTipo, '+
                                     'ot.sNumeroOrden, '+
                                     'ot.sIdPlataforma AS idPlataforma, '+
                                     'p.sDescripcion as Pernocta, '+
                                     'pt.sDescripcion as Plataforma '+
                              'from ordenesdetrabajo ot '+
                              'inner join contratos c '+
                              'on ( c.sContrato = ot.sContrato ) '+
                              'inner join bitacorademateriales ba '+
                              'on ( ba.sContrato = ot.sContrato and ba.sNumeroOrden = ot.sNumeroOrden ) '+
                              'left join pernoctan p '+
                              'on ( p.sIdPernocta = ba.sIdPernocta) '+
                              'inner join plataformas pt '+
                              'on ( ot.sIdPlataforma = pt.sIdPlataforma ) '+
                              'where c.sContrato = :contrato '+
                              'and ba.dIdFecha = :fecha '+
                              'group by ot.sIdFolio, p.sIdPernocta order by ot.iOrden ' );
        QrColumnas.ParamByName('Contrato').AsString:= ReporteDiario.FieldByName('sOrden').AsString;
        QrColumnas.ParamByName('Fecha').AsDate:=reportediario.FieldByName('dIdFecha').AsDateTime;
        QrColumnas.Open;

        QrColumnas.First;
        {$ENDREGION}

        QrMoe.Active := False;
        QrMoe.SQL.Clear;
        QrMoe.SQL.Add('select b.sIdMaterial as sIdRecurso, a.mDescripcion, a.sMedida, a.dCantidadAnexo as dCantidad, a.sAnexo, b.sWbs, '+
                  'n.sdescripcion as anexo, n.stitulo as tituloAnexo, n.sTipo '+
                  'from bitacorademateriales b '+
                  'inner join actividadesxanexo a on (a.sContrato =:Contrato and a.sNumeroActividad = b.sIdMaterial and a.sTipoAnexo= "PU" and a.sTipoActividad = "Actividad") '+
                  'left join anexos n on (a.sAnexo = n.sAnexo) '+
                  'where b.sContrato =:Orden and b.dIdFecha =:Fecha '+
                  'group by n.iOrden, b.sContrato, a.sNumeroActividad order by n.iOrden, a.iItemOrden ');
        QrMoe.ParamByName('contrato').AsString := global_contrato_barco;
        QrMoe.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sOrden').AsString;
        QrMoe.ParamByName('Fecha').AsDateTime  := reportediario.FieldByName('dIdFecha').AsDateTime;
        QrMoe.Open;


        {$REGION 'INSERCION DE DATOS - INFORMACIÓN DEL PU'}

        DCiclo:=QrColumnas.RecordCount/TotalCol;
        iCiclo:=Trunc(DCiclo);
        if (dCiclo -iCiclo)>0 then
            Inc(iCiclo,1);
        iGrupo:=1;
        while iGrupo<=iCiclo do
        begin
            with qrMoe do
            begin
                first;
                while not eof do
                begin
                    mDatos.Append;
                    mDatos.FieldByName('iGrupo').AsInteger:=Igrupo;
                    mDatos.FieldByName('sidAnexo').AsString     :=FieldByName('sanexo').asstring;
                    mDatos.FieldByName('sAnexo').AsString       :=FieldByName('anexo').asstring;
                    mDatos.FieldByName('sTitulo').AsString      :=FieldByName('tituloAnexo').asstring;
                    mDatos.FieldByName('sidrecurso').AsString   :=FieldByName('sidrecurso').asstring;
                    mDatos.FieldByName('sdescripcion').AsString :=fieldbyname('mdescripcion').asstring;
                    mDatos.FieldByName('smedida').AsString      :=fieldbyname('smedida').asstring;
                    mDatos.FieldByName('dcantSol').AsFloat      :=fieldbyname('dcantidad').AsFloat;
                    mDatos.FieldByName('sTipo').AsString        :=FieldByName('sTipo').asstring;
                    dTotal:=0;
                    if iGrupo=1 then
                       QrColumnas.First
                    else
                       QrColumnas.RecNo:=((iGrupo-1) * TotalCol)+ 1;

                    CuentaCol:=1;
                    ValTmp:=0;
                    CantTmp:=0;
                    iPosTmp:=0;
                    while not (QrColumnas.Eof) and (QrColumnas.RecNo<=((iGrupo) * TotalCol)) do
                    begin
                        mDatos.FieldByName('sNumeroOrden'+ Inttostr(CuentaCol)).AsString := qrColumnas.FieldbyName('sIdFolio').AsString;
                        mDatos.FieldByName('sPernocta' + Inttostr(CuentaCol)).AsString   := QrColumnas.Fieldbyname('Pernocta').asstring;
                        mDatos.FieldByName('sPlataforma'+ Inttostr(CuentaCol)).AsString  := QrColumnas.Fieldbyname('Plataforma').Asstring;
                        mDatos.FieldByName('sTipo'+ Inttostr(CuentaCol)).AsString        := QrColumnas.Fieldbyname('sTipo').Asstring;

                        QrRecursos.Active := False;
                        QrRecursos.ParamByName('Actividad').AsString := FieldByName('sIdRecurso').AsString;
                        QrRecursos.ParamByName('contrato').AsString  := ReporteDiario.FieldByName('sOrden').AsString;
                        QrRecursos.ParamByName('Orden').AsString     := QrColumnas.FieldByName('sNumeroOrden').AsString;
                        QrRecursos.ParamByName('fecha').AsDate       := reportediario.FieldByName('dIdFecha').AsDateTime;
                        QrRecursos.Open;
                        //ABBY
                        if QrRecursos.RecordCount>0 then
                        begin
                          mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat:=xRound(QrRecursos.FieldByName('dCantidad').Value,4);
                          dTotal:=dTotal+ xRound(QrRecursos.FieldByName('dCantidad').Value,4);
                          ValTmp:=ValTmp + QrRecursos.FieldByName('dCantidad').Value;
                          if CantTmp<mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat then
                          begin
                            CantTmp:=mDatos.FieldByName('dcantidad' + Inttostr(CuentaCol)).AsFloat;
                            iPosTmp:=CuentaCol;
                          end;
                        end;
                        QrColumnas.next;
                        Inc(CuentaCol);
                    end;

                    if dTotal>=0 then
                    begin
                      if dTotal<>xRound(ValTmp,4) then
                      begin
                        mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat:=mDatos.FieldByName('dcantidad' + Inttostr(iPosTmp)).AsFloat
                        + (xRound(ValTmp,4)-dTotal);
                        mDatos.FieldByName('dcantTotal').AsFloat:=xRound(ValTmp,4);
                      end
                      else
                        mDatos.FieldByName('dcantTotal').AsFloat:=dTotal;
                      mDatos.Post;
                    end
                    else
                       mDatos.cancel;
                    next;
                end;
            end;
            Inc(iGrupo,1);
        end;
        {$ENDREGION}

        {$endregion}

    end;
    mDatos.first;
    Td_Distribucion_detalle2.DataSet:=MDatos;
    Td_Distribucion_detalle2.FieldAliases.Clear;
    Td_ImpDistribucion_detalle2.DataSet:=MDatos;
    Td_ImpDistribucion_detalle2.FieldAliases.Clear;

    Reporte.DataSets.Add(Td_Distribucion_detalle2);
    Reporte.DataSets.Add(Td_ImpDistribucion_detalle2);
  finally
    QrRecursos.destroy;
    QrColumnas.destroy;
    QrMoe.destroy;
  end;
end;

Procedure NotaCampo(ParamContrato,ParamFolio:String;Var Reporte: TfrxReport);
const
  Asql:Array[1..4] of String =('SELECT ' +
                               '  bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion, ' +
                               '  p.smedida as medida, a.sDescripcion as Titulo, ' +
                               '  IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                               '  (IF(SUM(bp.dCanthh) > 0, (SUM(bp.dCanthh)), SUM(bp.dCantidad))) AS Total ' +
                               'FROM ' +
                               '  bitacoradepersonal bp ' +
                               'inner join ' +
                               '  personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                               'inner join ' +
                               '  anexos a ' +
                               '    on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="No") ' +
                               'WHERE ' +
                               '  bp.scontrato = :Orden AND bp.sNumeroOrden = :Folio ' +
                               'GROUP BY ' +
                               '  bp.sIdPersonal'
{  Asql:Array[1..4] of String =( 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,SUM(bp.dAjuste) AS Ajuste,(IF(SUM(bp.dCanthh) > 0, '+
                                '(SUM(bp.dCanthh)), SUM(bp.dCantidad))) AS Total FROM bitacoradepersonal bp '+
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="No") '+
                                'WHERE bp.scontrato = :Orden AND bp.sNumeroOrden = :Folio GROUP BY bp.sIdPersonal '
}
                              , 'select be.sIdEquipo as sIdRecurso,be.sDescripcion as Descripcion,e.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="EQUIPO" limit 1),"No existe") as Titulo,'+
                                '(sum(be.dCantHH) ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = bp.sContrato AND ' +
                                '           ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                '           ba.dFecha = bp.dIdFecha), 0)) as total ' +
                                'from bitacoradeequipos be inner join equipos e on '+
                                '(e.sContrato=:contratoBarco and e.sIdEquipo=be.sIdEquipo) inner join bitacoradeactividades '+
                                'ba on(ba.sContrato = be.sContrato and ba.dIdFecha = be.didfecha and ba.sNumeroOrden = '+
                                'be.sNumeroOrden and ba.iIdDiario = be.iIdDiario and ba.iIdTarea = be.iIdTarea and '+
                                'ba.iIdActividad=be.iIdActividad) where  be.sContrato=:Orden and be.sNumeroOrden=:folio '+
                                'group by e.sIdEquipo order by e.iItemOrden'

                              , 'select c.sIdCuenta,c.sIdPernocta as sIdRecurso,c.sDescripcion as Descripcion,c.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="PERNOCTA" limit 1),"No existe") as Titulo,'+
                                '(sum(bp.dCantHH) ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = bpr.sContrato AND ' +
                                '           ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                                '           ba.sIdPartidaAnexo = bpr.sIdCuenta AND ' +
                                '           ba.sNumeroActividad = bpr.sNumeroActividad AND ' +
                                '           ba.dFecha = bpr.dIdFecha), 0)) ' +
                                '- ifnull((select ifnull(sum(bpr.dCantidad),0) from bitacoradepernocta bpr '+
                                'where bpr.sContrato = ba.sContrato and ba.didfecha=bpr.didfecha and (bpr.sNumeroOrden <> "@" '+
                                'and bpr.sNumeroOrden=ba.snumeroorden) group by bpr.sContrato),0)) as Total from cuentas c '+
                                'left join bitacoradeactividades ba on(ba.sContrato=:Orden and ba.sNumeroOrden=:Folio'+
                                ') left join bitacoradepersonal bp on (bp.sContrato =ba.sContrato and '+
                                'ba.sNumeroOrden=bp.sNumeroOrden and  bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                                'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad and bp.sTipoPernocta = c.sIdPernocta '+
                                'and bp.lAplicaPernocta = "Si") left join personal p on (p.sContrato = :contratoBarco '+
                                'and p.sIdPersonal = bp.sIdPersonal) left join tiposdepersonal tp on ( p.sIdTipoPersonal = '+
                                'tp.sIdTipoPersonal) left join ordenesdetrabajo ot '+
                                'on (ot.sContrato = :Orden and ot.sNumeroOrden = ba.sNumeroOrden) left join pernoctan pr '+
                                'on (pr.sIdPernocta = ot.sIdPernocta) left join plataformas pl on(pl.sIdPlataforma=ot.sIdPlataforma) '+
                                'where bp.snumeroorden=:Folio group by c.sIdPernocta order by c.sIdPernocta '
                              
                              , 'SELECT ' +
                                '  bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida, a.sDescripcion as Titulo, ' +
                                '  IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                '  SUM(bp.dCanthh) AS Total ' +
                                'FROM ' +
                                '  bitacoradepersonal bp ' +
                                'inner join ' +
                                '  personal p ' +
                                '    on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                                'inner join ' +
                                '  anexos a ' +
                                '    on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="Si") ' +
                                'WHERE ' +
                                '  bp.scontrato = :Orden AND bp.sNumeroOrden = :Folio ' +
                                'GROUP BY ' +
                                '  bp.sIdPersonal');
                            { , 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,SUM(bp.dAjuste) AS Ajuste,SUM(bp.dCanthh) AS Total '+
                                'FROM bitacoradepersonal bp inner join personal p on(p.sContrato=:ContratoBarco and '+
                                'p.sIdPersonal=bp.sIdPersonal) inner join anexos a on(a.sAnexo=p.sAnexo and '+
                                'a.sTipo="PERSONAL" and a.sTierra="Si") WHERE bp.scontrato = :Orden AND '+
                                'bp.sNumeroOrden = :Folio GROUP BY bp.sIdPersonal ');}
var
  QrActividades:TzReadOnlyQuery;
  QrProrrateo: TzReadOnlyQuery;
  MryRecursos:TjvMemoryData;
  iGrupo:Integer;
  QrConsulta,
  QrResumenMaterial,
  QrObservaciones:TzReadOnlyQuery;
  Td_Partidas,
  Td_ProrrateoBarco,
  Td_CDistribucionP,
  Td_resumenMaterial,
  Td_NotasGrales: TfrxDBDataset;

begin
  QrActividades:=TzReadOnlyQuery.Create(nil);
  QrProrrateo:=TzREadOnlyQuery.Create(nil);
  QrConsulta:=TzREadOnlyQuery.Create(nil);
  MryRecursos:=TJvMemoryData.Create(nil);
  QrObservaciones:=TzREadOnlyQuery.Create(nil);
  QrResumenMaterial:=TzREadOnlyQuery.Create(nil);

  Td_Partidas:= TfrxDBDataset.Create(nil);
  Td_ProrrateoBarco:= TfrxDBDataset.Create(nil);
  Td_CDistribucionP:= TfrxDBDataset.Create(nil);
  Td_resumenMaterial:= TfrxDBDataset.Create(nil);
  Td_NotasGrales:=TfrxDBDataset.Create(nil);

  try
    try
      Td_Partidas.UserName:='Td_Partidas';
      Td_ProrrateoBarco.UserName:='Td_ProrrateoBarco';
      Td_CDistribucionP.UserName:='Td_CDistribucionP';
      Td_resumenMaterial.UserName:='Td_resumenMaterial';
      Td_NotasGrales.UserName:='Td_NotasGrales';

      QrConsulta.Connection:=Connection.zConnection;
      QrACtividades.Connection:=Connection.zConnection;
      QrACtividades.SQL.Text:='select ao.*,' + #10 +
                              'ifnull((SELECT  SUM(dAvance) FROM bitacoradeactividades' + #10 +
                              'WHERE sContrato = ao.sContrato AND sNumeroOrden = ao.sNumeroOrden' + #10 +
                              'and sIdTipoMovimiento = "E" AND sWbs = ao.sWbs),0) AS dAvanceGeneralPartida,' + #10 +
                              'if(ao.lExtraordinario="No",0,1) as lOrdenExt,' + #10 +
                              'if(ao.lExtraordinario="No","ACTIVIDADES DEL PROGRAMA","ACTIVIDADES ADICIONALES") as lOrdenTitulo' + #10 +
                              'from actividadesxorden ao' + #10 +
                              'where ao.scontrato=:Contrato and ao.sNumeroOrden=:Orden' + #10 +
                              'and ao.sTipoActividad="Actividad"' + #10 +
                              'order by lOrdenExt,ao.iItemOrden' ;
      QrACtividades.ParamByName('Contrato').AsString:=ParamContrato;
      QrActividades.ParamByName('Orden').AsString:=ParamFolio;
      try
        QrACtividades.Open;
      except
        Raise;
      end;

      QrProrrateo.Connection:=Connection.zConnection;
      QrProrrateo.SQL.Text:='SELECT tm.sContrato,tm.sIdTipoMovimiento,tm.sDescripcion,tm.sTipo,tm.smedida,' + #10 +
                            '(SELECT ifnull(SUM( me.sFactor ),0) FROM' + #10 +
                            'movimientosdeembarcacion AS me inner join' + #10 +
                            'movimientosxfolios mf on (me.sContrato = mf.sContrato' + #10 +
                            'and me.iIdDiario = mf.iIdDiario) WHERE' + #10 +
                            'me.sContrato = tm.sContrato AND me.sOrden =:Orden' + #10 +
                            'and mf.sFolio=:Folio AND me.sIdFase = "OPER"' + #10 +
                            'AND me.sClasificacion = tm.sIdTipoMovimiento) AS sFactor,' + #10 +
                            'ifnull((select sdescripcion from anexos where sTipo="BARCO" limit 1)'+
                            ',"No existe") as Titulo ' +
                            'FROM tiposdemovimiento AS tm WHERE tm.sContrato =:ContratoBarco' + #10 +
                            'AND tm.sClasificacion = "Movimiento de Barco" and' + #10 +
                            'tm.lImprimeCeros ="Si"';
      QrProrrateo.ParamByName('ContratoBarco').AsString:=Global_Contrato_Barco;
      QrProrrateo.ParamByName('Orden').AsString:=ParamContrato;
      QrProrrateo.ParamByName('Folio').AsString:=ParamFolio;
      QrProrrateo.Open;



      with MryRecursos do
      begin
        Active:=False;
        FieldDefs.Add('iGrupo', ftInteger, 0, True);
        //FieldDefs.Add('sidAnexo', ftString, 10, false);
        FieldDefs.Add('sidrecurso', ftString, 100, True);
        FieldDefs.Add('sdescripcion', ftString, 250, True);
        //FieldDefs.Add('sAnexo', ftString, 250, false);
        FieldDefs.Add('sTitulo', ftString, 250, false);
        FieldDefs.Add('smedida', ftString, 100, True);
        FieldDefs.Add('dcantidad', FtFloat, 0, True);
        Active:=true;
      end;

      for iGrupo := 1 to length(Asql) do
      begin
        QrConsulta.Active:=false;
        QrConsulta.SQL.Text:=ASql[iGrupo];
        QrConsulta.ParamByName('ContratoBarco').AsString:=Global_Contrato_Barco;
        QrConsulta.ParamByName('Orden').AsString:=ParamContrato;
        QrConsulta.ParamByName('Folio').AsString:=ParamFolio;
        QrConsulta.Open;
        while not QrConsulta.Eof do
        begin
          MryRecursos.Append;
          MryRecursos.FieldByName('iGrupo').AsInteger:=iGrupo;
          MryRecursos.FieldByName('sIdRecurso').asString:=QrConsulta.FieldByName('sIdRecurso').AsString;
          MryRecursos.FieldByName('sDescripcion').asString:=QrConsulta.FieldByName('descripcion').AsString;
          MryRecursos.FieldByName('sMedida').asString:=QrConsulta.FieldByName('medida').AsString;
          MryRecursos.FieldByName('sTitulo').asString:=QrConsulta.FieldByName('Titulo').AsString;
          if (iGrupo=1) or (iGrupo=4) then
            MryRecursos.FieldByName('dcantidad').AsFloat:=(xRound(QrConsulta.FieldByName('Total').AsFloat,2) + QrConsulta.FieldByName('Ajuste').AsFloat)
          else
            if iGrupo=3 then
            begin

                connection.QryBusca.Active:=false;
                {connection.QryBusca.SQL.Text:='select ifnull(sum(dCantidad),0) as dCantidad from bitacoradepernocta where ' +
                            'sContrato=:Orden and sNumeroOrden=:Folio and sIdCuenta=:Pernocta ';}

                connection.QryBusca.SQL.Text:='select ' +
                                              '  (ifnull(sum(bpr.dCantidad),0) + ' +
                                              '   IFNULL((SELECT ' +
                                              '             SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                              '           FROM ' +
                                              '             bitacoradeajustes ba ' +
                                              '           WHERE ' +
                                              '             ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                                              '             ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0)) as dCantidad ' +
                                              'from ' +
                                              '  bitacoradepernocta bpr ' +
                                              'where ' +
                                              '  bpr.sContrato=:Orden and ' +
                                              '  bpr.sNumeroOrden=:Folio and ' +
                                              '  bpr.sIdCuenta=:Pernocta';

                  {connection.QryBusca.SQL.Text:='select sum(bpc.dCantidad) as dCantidad' + #10 +
                                              'from bitacoradepersonal_cuadre bpc' + #10 +
                                              'inner join personal p' + #10 +
                                              'on ( p.sContrato = :ContratoBarco and bpc.sIdPersonal = p.sIdPersonal )' + #10 +
                                              'inner join tiposdepersonal tp' + #10 +
                                              'on (tp.sIdTipoPersonal = p.sIdTipoPersonal )' + #10 +
                                              'where bpc.sContrato =:Orden and bpc.snumeroorden=:Folio' + #10 +
                                              'and tp.lPernocta="Si"' ; }

                connection.QryBusca.ParamByName('Pernocta').AsString:=QrConsulta.FieldByName('sIdCuenta').AsString ;
                connection.QryBusca.ParamByName('Orden').AsString:=ParamContrato;
                connection.QryBusca.ParamByName('Folio').AsString:=ParamFolio;
                connection.QryBusca.Open;
                if connection.QryBusca.RecordCount=1 then
                  MryRecursos.FieldByName('dcantidad').AsFloat:=xRound(QrConsulta.FieldByName('total').AsFloat,2)+connection.QryBusca.FieldByName('dCantidad').AsFloat
                else
                  MryRecursos.FieldByName('dcantidad').AsFloat:=xRound(QrConsulta.FieldByName('total').AsFloat,2);
            end
            else
              MryRecursos.FieldByName('dcantidad').AsFloat:=xRound(QrConsulta.FieldByName('total').AsFloat,2);

          MryRecursos.Post;
          QrConsulta.Next;
        end;
      end;

      QrResumenMaterial.Connection:=Connection.zConnection;
      QrResumenMaterial.SQL.text:=  'select s.sNumeroOrden, b.sNumeroActividad, b.sTrazabilidad,i.sIdInsumo,i.mDescripcion, i.sMedida, sum(b.dCantidad) as dCantidad, i.sIdProveedor from almacen_salida s '+
                                    'inner join bitacoradesalida b '+
                                    'on(b.sContrato = s.sContrato and b.iFolioSalida = s.iFolioSalida and b.sext = s.sext) '+
                                    'inner join insumos i '+
                                    'on (i.sContrato =:Contrato and i.sIdInsumo = b.sIdInsumo and i.sTrazabilidad = b.sTrazabilidad ) '+
                                    'INNER JOIN actividadesxorden ao ON ( ' +
                                    '	ao.sContrato = b.sContrato and ao.sIdConvenio =:Convenio and ao.sNumeroOrden = b.sNumeroOrden and ' +
                                    '	ao.sNumeroActividad = b.sNumeroActividad and ao.sTipoActividad = "Actividad" ) ' +
                                    'where s.sContrato =:orden and s.sNumeroOrden=:Folio '+
                                    'group by b.sNumeroOrden,b.sIdInsumo Order by b.sNumeroOrden, ao.iItemOrden, b.sTrazabilidad';
      QrResumenMaterial.ParamByName('Orden').AsString    := ParamContrato;
      QrResumenMaterial.ParamByName('Folio').AsString    := ParamFolio;
      QrResumenMaterial.ParamByName('Contrato').AsString := global_Contrato_Barco;
      QrResumenMaterial.ParamByName('Convenio').AsString := global_Convenio;
      QrResumenMaterial.Open;

      QrObservaciones.Connection:=Connection.zConnection;
      QrObservaciones.SQL.Text:='select ncg.*,nco.dfecha as Fecha,nco.sObservacion from notacampo_general ncg' + #10 +
                                'left join notacampo_observaciones nco on(ncg.iIdNota=nco.iIdNota) ' +
                                'where ncg.sContrato=:Contrato and ncg.sNumeroOrden=:Folio';
      QrObservaciones.ParamByName('Contrato').AsString:=ParamContrato;
      QrObservaciones.ParamByName('Folio').AsString:=ParamFolio;
      QrObservaciones.Open;





      Td_Partidas.DataSet:= QrACtividades;
      Td_Partidas.FieldAliases.Clear;
      Td_ProrrateoBarco.DataSet:= QrProrrateo;
      Td_ProrrateoBarco.FieldAliases.Clear;
      Td_CDistribucionP.DataSet:= MryRecursos;
      Td_CDistribucionP.FieldAliases.Clear;
      Td_resumenMaterial.DataSet:= QrResumenMaterial;
      Td_resumenMaterial.FieldAliases.Clear;
      Td_NotasGrales.DataSet:=QrObservaciones;
      Td_NotasGrales.FieldAliases.Clear;

      Reporte.DataSets.Add(Td_Partidas);
      Reporte.DataSets.Add(Td_ProrrateoBarco);
      Reporte.DataSets.Add(Td_CDistribucionP);
      Reporte.DataSets.Add(Td_resumenMaterial);
      Reporte.DataSets.Add(Td_NotasGrales);

    except

    end;
  finally
   { ReportePDF_ClearDataset(Reporte);
    QrActividades.Destroy;
    QrProrrateo.Destroy;
    QrConsulta.Destroy;
    MryRecursos.Destroy;  }
  end;

end;

Procedure TdConfiguracion(ParamContrato,ParamFolio:String;Var Reporte: TfrxReport);
var
  Td_contrato,
  Td_configuracion,
  Td_ConfiguracionOrden,
  TD_ConfigOTBarco: TfrxDBDataset;
  zContrato,
  zConfiguracion,
  QrConfigFolio,
  QrConfigBarco:TZReadOnlyQuery;
begin
  try
    try
      Td_contrato:= TfrxDBDataset.Create(nil);
      Td_contrato.UserName:='contrato';

      Td_configuracion:= TfrxDBDataset.Create(nil);
      Td_configuracion.UserName:='dsConfiguracion';

      Td_ConfiguracionOrden:= TfrxDBDataset.Create(nil);
      Td_ConfiguracionOrden.UserName:='DsConfiguracionOrden';

      TD_ConfigOTBarco:= TfrxDBDataset.Create(nil);
      TD_ConfigOTBarco.UserName:='TD_ConfigOTBarco';

      {Información del contrato}
      zContrato := TZReadOnlyQuery.Create(nil);
      zContrato.Connection := Connection.zConnection;
      zContrato.SQL.Add(' Select c.sLicitacion, c.sTitulo, c.sContrato, c.sTipoObra, c.mDescripcion, c.mCliente, c.bImagen, c.sUbicacion,  c.sCodigo, c.sLabelContrato, c.sCliente, '+
                'c.sIdResidencia, c.eLugarOT, '+
                'c2.sDescripcion as sConvenio, c2.dFechaInicio, c2.dFechaFinal, c2.dMontoMN, c2.dMontoDLL, '+
                'c2.dFecha, a.bImagen as bImagenActivo, a.sDescripcion as sDescripcionActivo, c.mComentarios '+
                'FROM contratos c '+
                'inner join activos a on (c.sIdActivo = a.sIdActivo) '+
                'inner join residencias rs on (c.sIdResidencia = rs.sIdResidencia) '+
                'inner join configuracion c3 on (c.sContrato = c3.sContrato) '+
                'inner join convenios c2 on (c3.sContrato = c2.sContrato And c3.sIdConvenio = c2.sIdConvenio) '+
                'Where c.sContrato = :Contrato ');
      zContrato.ParamByName('contrato').AsString := ParamContrato;
      zContrato.Open;

      {Información de la configuracion del sistema}
      zConfiguracion := TZReadOnlyQuery.Create(nil);
      zConfiguracion.Connection := Connection.zConnection;
      zConfiguracion.SQL.Add('select c.iFirmasReportes, c.sTipoPartida, c.sImprimePEP, ' +
                ' (select sContrato from contratos where sContrato =:contratobarco ) as sContratoBarco, ' +
                ' (select mDescripcion from contratos where sContrato =:contratobarco ) as mDescripcionBarco, ' +
                ' (select mcliente from contratos where sContrato =:contratobarco ) as mClienteBarco, ' +
                ' c.lLicencia, c.sReportesCIA, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3, c.iFirmasGeneradores, ' +
                ' c.bImagen, c.sContrato, c.sNombre, c2.sCodigo, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, ' +
                ' c2.mDescripcion, c2.sTitulo, c2.mCliente, c2.bImagen as bImagenPEP, cv.dFechaInicio, cv.dfechaFinal ' +
                ',concat(c.sDireccion1," ",c.sDireccion2) as direccion,c.sCiudad,c.sTelefono,c.sFax'   +
                ' From contratos c2 '+
                ' INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
                ' INNER JOIN convenios cv on (cv.sContrato = c2.sContrato and cv.sIdConvenio =:convenio) '+
                ' Where c2.sContrato = :Contrato');
      zConfiguracion.Params.ParamByName('contrato').DataType := ftString;
      zConfiguracion.Params.ParamByName('contrato').Value    := ParamContrato;
      zConfiguracion.Params.ParamByName('contratobarco').DataType := ftString;
      zConfiguracion.Params.ParamByName('contratobarco').Value    := global_contrato_barco;
      zConfiguracion.Params.ParamByName('convenio').DataType := ftString;
      zConfiguracion.Params.ParamByName('convenio').Value    := global_convenio;
      zConfiguracion.Open;

      QrConfigFolio := TZReadOnlyQuery.Create(nil);
      QrConfigFolio.Connection:=Connection.zConnection;
      QrConfigFolio.sql.Text:='select ot.*,p.sDescripcion as Plataforma from ordenesdetrabajo ot inner join plataformas p ' +
                              'on(p.sIdPlataforma=ot.sIdPlataforma) '+
                              'where ot.sContrato=:Contrato and ot.sNumeroOrden=:Orden';
      QrConfigFolio.ParamByName('Contrato').AsString:=ParamContrato;
      QrConfigFolio.ParamByName('Orden').AsString:=ParamFolio;
      try
        QrConfigFolio.Open;
      except
        raise;
      end;

      QrConfigBarco:=TzREadOnlyQuery.Create(nil);
      QrConfigBarco.Connection:=Connection.zConnection;
      QrConfigBarco.SQL.Text:='select * from contratos c inner join convenios cv '+
                              'on(c.sContrato and cv.sContrato) '+
                              'where c.sContrato=:ContratoBarco and cv.sIdConvenio=:Convenio';
      QrConfigBarco.ParamByName('ContratoBarco').AsString:= global_contrato_barco;
      QrConfigBarco.ParamByName('Convenio').AsString:= global_convenio;
      QrConfigBarco.Open;

      TD_ConfigOTBarco.DataSet:=QrConfigBarco;
      TD_ConfigOTBarco.FieldAliases.Clear;
      Td_contrato.DataSet:= zContrato;
      Td_contrato.FieldAliases.Clear;
      Td_configuracion.DataSet:= zConfiguracion;
      Td_configuracion.FieldAliases.Clear;
      Td_ConfiguracionOrden.DataSet:= QrConfigFolio;
      Td_ConfiguracionOrden.FieldAliases.Clear;



      Reporte.DataSets.Add(Td_contrato);
      Reporte.DataSets.Add(Td_configuracion);
      Reporte.DataSets.Add(Td_ConfiguracionOrden);
      Reporte.DataSets.Add(TD_ConfigOTBarco);

    except

    end;

  finally

  end;


end;

Procedure TdConfiguracionGenerador(ParamContrato,ParamFolio:String;Var Reporte: TfrxReport;FechaI,FechaT:TDateTime);
var
  Td_contrato,
  Td_configuracion,
  Td_ConfiguracionOrden,
  TD_ConfigOTBarco: TfrxDBDataset;
  zContrato,
  zConfiguracion,
  QrConfigFolio,
  QrConfigBarco:TZReadOnlyQuery;
begin
  try
    try
      Td_contrato:= TfrxDBDataset.Create(nil);
      Td_contrato.UserName:='Td_contrato';

      Td_configuracion:= TfrxDBDataset.Create(nil);
      Td_configuracion.UserName:='DsConfiguracion';

      Td_ConfiguracionOrden:= TfrxDBDataset.Create(nil);
      Td_ConfiguracionOrden.UserName:='DsConfiguracionOrden';

      TD_ConfigOTBarco:= TfrxDBDataset.Create(nil);
      TD_ConfigOTBarco.UserName:='TD_ConfigOTBarco';

      {Información del contrato}
      zContrato := TZReadOnlyQuery.Create(nil);
      zContrato.Connection := Connection.zConnection;
      zContrato.SQL.Add(' Select c.sLicitacion, c.sTitulo, c.sContrato, c.sTipoObra, c.mDescripcion, c.mCliente, c.bImagen, c.sUbicacion,  c.sCodigo, c.sLabelContrato, c.sCliente, '+
                'c.sIdResidencia, c.eLugarOT, '+
                'c2.sDescripcion as sConvenio, c2.dFechaInicio, c2.dFechaFinal,'+
                'c2.dFecha, a.bImagen as bImagenActivo, a.sDescripcion as sDescripcionActivo, c.mComentarios '+
                'FROM contratos c '+
                'inner join activos a on (c.sIdActivo = a.sIdActivo) '+
                'inner join residencias rs on (c.sIdResidencia = rs.sIdResidencia) '+
                'inner join configuracion c3 on (c.sContrato = c3.sContrato) '+
                'inner join convenios c2 on (c3.sContrato = c2.sContrato And c3.sIdConvenio = c2.sIdConvenio) '+
                'Where c.sContrato = :Contrato ');
      zContrato.ParamByName('contrato').AsString := ParamContrato;
      zContrato.Open;

      {Información de la configuracion del sistema}
      zConfiguracion := TZReadOnlyQuery.Create(nil);
      zConfiguracion.Connection := Connection.zConnection;
      zConfiguracion.SQL.Add('select c.iFirmasReportes, c.sTipoPartida, c.sImprimePEP, ' +
                ' (select sContrato from contratos where sContrato =:contratobarco ) as sContratoBarco, ' +
                ' (select mDescripcion from contratos where sContrato =:contratobarco ) as mDescripcionBarco, ' +
                ' (select mcliente from contratos where sContrato =:contratobarco ) as mClienteBarco, ' +
                ' c.lLicencia, c.sReportesCIA, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3, c.iFirmasGeneradores, ' +
                ' c.bImagen, c.sContrato, c.sNombre, c2.sCodigo, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, ' +
                ' c2.mDescripcion, c2.sTitulo, c2.mCliente, c2.bImagen as bImagenPEP, cv.dFechaInicio, cv.dfechaFinal ' +
                ',concat(c.sDireccion1," ",c.sDireccion2) as direccion,c.sCiudad,c.sTelefono,c.sFax'   +
                ' From contratos c2 '+
                ' INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
                ' INNER JOIN convenios cv on (cv.sContrato = c2.sContrato and cv.sIdConvenio =:convenio) '+
                ' Where c2.sContrato = :Contrato');
      zConfiguracion.Params.ParamByName('contrato').DataType := ftString;
      zConfiguracion.Params.ParamByName('contrato').Value    := ParamContrato;
      zConfiguracion.Params.ParamByName('contratobarco').DataType := ftString;
      zConfiguracion.Params.ParamByName('contratobarco').Value    := global_contrato_barco;
      zConfiguracion.Params.ParamByName('convenio').DataType := ftString;
      zConfiguracion.Params.ParamByName('convenio').Value    := global_convenio;
      zConfiguracion.Open;

      QrConfigFolio := TZReadOnlyQuery.Create(nil);
      QrConfigFolio.Connection:=Connection.zConnection;
      QrConfigFolio.sql.Text:='select ot.*,:FechaI as FechaInicio,:FechaT as FechaTermino,p.sDescripcion as Plataforma from ordenesdetrabajo ot inner join plataformas p ' +
                              'on(p.sIdPlataforma=ot.sIdPlataforma) '+
                              'where ot.sContrato=:Contrato and ot.sNumeroOrden=:Orden';
      QrConfigFolio.ParamByName('Contrato').AsString:=ParamContrato;
      QrConfigFolio.ParamByName('Orden').AsString:=ParamFolio;
      QrConfigFolio.ParamByName('FechaI').AsString:=formatdatetime('dd/mm/yyyy',FechaI);
      QrConfigFolio.ParamByName('FechaT').AsString:= formatdatetime('dd/mm/yyyy',FechaT);
      try
        QrConfigFolio.Open;
      except
        raise;
      end;

      QrConfigBarco:=TzREadOnlyQuery.Create(nil);
      QrConfigBarco.Connection:=Connection.zConnection;
      QrConfigBarco.SQL.Text:='select * from contratos c inner join convenios cv '+
                              'on(c.sContrato and cv.sContrato) '+
                              'where c.sContrato=:ContratoBarco and cv.sIdConvenio=:Convenio';
      QrConfigBarco.ParamByName('ContratoBarco').AsString:= global_contrato_barco;
      QrConfigBarco.ParamByName('Convenio').AsString:= global_convenio;
      QrConfigBarco.Open;

      TD_ConfigOTBarco.DataSet:=QrConfigBarco;
      TD_ConfigOTBarco.FieldAliases.Clear;
      Td_contrato.DataSet:= zContrato;
      Td_contrato.FieldAliases.Clear;
      Td_configuracion.DataSet:= zConfiguracion;
      Td_configuracion.FieldAliases.Clear;
      Td_ConfiguracionOrden.DataSet:= QrConfigFolio;
      Td_ConfiguracionOrden.FieldAliases.Clear;



      Reporte.DataSets.Add(Td_contrato);
      Reporte.DataSets.Add(Td_configuracion);
      Reporte.DataSets.Add(Td_ConfiguracionOrden);
      Reporte.DataSets.Add(TD_ConfigOTBarco);

    except
      on e:Exception do
        ShowMessage(e.Message +   ' : ' + e.ClassName);
    end;

  finally

  end;


end;

Procedure TdConfiguracionGenerador(ParamContrato,ParamFolio:String;Var Reporte: TfrxReport);
var
  Td_contrato,
  Td_configuracion,
  Td_ConfiguracionOrden,
  TD_ConfigOTBarco: TfrxDBDataset;
  zContrato,
  zConfiguracion,
  QrConfigFolio,
  QrConfigBarco:TZReadOnlyQuery;
begin
  try
    try
      Td_contrato:= TfrxDBDataset.Create(nil);
      Td_contrato.UserName:='Td_contrato';

      Td_configuracion:= TfrxDBDataset.Create(nil);
      Td_configuracion.UserName:='DsConfiguracion';

      Td_ConfiguracionOrden:= TfrxDBDataset.Create(nil);
      Td_ConfiguracionOrden.UserName:='DsConfiguracionOrden';

      TD_ConfigOTBarco:= TfrxDBDataset.Create(nil);
      TD_ConfigOTBarco.UserName:='TD_ConfigOTBarco';

      {Información del contrato}
      zContrato := TZReadOnlyQuery.Create(nil);
      zContrato.Connection := Connection.zConnection;
      zContrato.SQL.Add(' Select c.sLicitacion, c.sTitulo, c.sContrato, c.sTipoObra, c.mDescripcion, c.mCliente, c.bImagen, c.sUbicacion,  c.sCodigo, c.sLabelContrato, c.sCliente, '+
                'c.sIdResidencia, c.eLugarOT, '+
                'c2.sDescripcion as sConvenio, c2.dFechaInicio, c2.dFechaFinal, c2.dMontoMN, c2.dMontoDLL, '+
                'c2.dFecha, a.bImagen as bImagenActivo, a.sDescripcion as sDescripcionActivo, c.mComentarios '+
                'FROM contratos c '+
                'inner join activos a on (c.sIdActivo = a.sIdActivo) '+
                'inner join residencias rs on (c.sIdResidencia = rs.sIdResidencia) '+
                'inner join configuracion c3 on (c.sContrato = c3.sContrato) '+
                'inner join convenios c2 on (c3.sContrato = c2.sContrato And c3.sIdConvenio = c2.sIdConvenio) '+
                'Where c.sContrato = :Contrato ');
      zContrato.ParamByName('contrato').AsString := ParamContrato;
      zContrato.Open;

      {Información de la configuracion del sistema}
      zConfiguracion := TZReadOnlyQuery.Create(nil);
      zConfiguracion.Connection := Connection.zConnection;
      zConfiguracion.SQL.Add('select c.iFirmasReportes, c.sTipoPartida, c.sImprimePEP, ' +
                ' (select sContrato from contratos where sContrato =:contratobarco ) as sContratoBarco, ' +
                ' (select mDescripcion from contratos where sContrato =:contratobarco ) as mDescripcionBarco, ' +
                ' (select mcliente from contratos where sContrato =:contratobarco ) as mClienteBarco, ' +
                ' c.lLicencia, c.sReportesCIA, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3, c.iFirmasGeneradores, ' +
                ' c.bImagen, c.sContrato, c.sNombre, c2.sCodigo, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, ' +
                ' c2.mDescripcion, c2.sTitulo, c2.mCliente, c2.bImagen as bImagenPEP, cv.dFechaInicio, cv.dfechaFinal ' +
                ',concat(c.sDireccion1," ",c.sDireccion2) as direccion,c.sCiudad,c.sTelefono,c.sFax'   +
                ' From contratos c2 '+
                ' INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
                ' INNER JOIN convenios cv on (cv.sContrato = c2.sContrato and cv.sIdConvenio =:convenio) '+
                ' Where c2.sContrato = :Contrato');
      zConfiguracion.Params.ParamByName('contrato').DataType := ftString;
      zConfiguracion.Params.ParamByName('contrato').Value    := ParamContrato;
      zConfiguracion.Params.ParamByName('contratobarco').DataType := ftString;
      zConfiguracion.Params.ParamByName('contratobarco').Value    := global_contrato_barco;
      zConfiguracion.Params.ParamByName('convenio').DataType := ftString;
      zConfiguracion.Params.ParamByName('convenio').Value    := global_convenio;
      zConfiguracion.Open;

      QrConfigFolio := TZReadOnlyQuery.Create(nil);
      QrConfigFolio.Connection:=Connection.zConnection;
      QrConfigFolio.sql.Text:='select ot.*,p.sDescripcion as Plataforma from ordenesdetrabajo ot inner join plataformas p ' +
                              'on(p.sIdPlataforma=ot.sIdPlataforma) '+
                              'where ot.sContrato=:Contrato and ot.sNumeroOrden=:Orden';
      QrConfigFolio.ParamByName('Contrato').AsString:=ParamContrato;
      QrConfigFolio.ParamByName('Orden').AsString:=ParamFolio;
      try
        QrConfigFolio.Open;
      except
        raise;
      end;

      QrConfigBarco:=TzREadOnlyQuery.Create(nil);
      QrConfigBarco.Connection:=Connection.zConnection;
      QrConfigBarco.SQL.Text:='select * from contratos c inner join convenios cv '+
                              'on(c.sContrato and cv.sContrato) '+
                              'where c.sContrato=:ContratoBarco and cv.sIdConvenio=:Convenio';
      QrConfigBarco.ParamByName('ContratoBarco').AsString:= global_contrato_barco;
      QrConfigBarco.ParamByName('Convenio').AsString:= global_convenio;
      QrConfigBarco.Open;

      TD_ConfigOTBarco.DataSet:=QrConfigBarco;
      TD_ConfigOTBarco.FieldAliases.Clear;
      Td_contrato.DataSet:= zContrato;
      Td_contrato.FieldAliases.Clear;
      Td_configuracion.DataSet:= zConfiguracion;
      Td_configuracion.FieldAliases.Clear;
      Td_ConfiguracionOrden.DataSet:= QrConfigFolio;
      Td_ConfiguracionOrden.FieldAliases.Clear;



      Reporte.DataSets.Add(Td_contrato);
      Reporte.DataSets.Add(Td_configuracion);
      Reporte.DataSets.Add(Td_ConfiguracionOrden);
      Reporte.DataSets.Add(TD_ConfigOTBarco);

    except
      on e:Exception do
        ShowMessage(e.Message +   ' : ' + e.ClassName);
    end;

  finally

  end;


end;
// aQUI TE QUEDASTE RANGEL
Procedure TdGenerador(ParamContrato:String;ParamFechaI,ParamFechaT:TDateTime;Var Reporte: TfrxReport;IsGeneral:Boolean=True);
const
  Asql:Array[1..4] of String =('SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida, ' +
                               'a.sDescripcion as Titulo,a.sAnexo, ' +
                               'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                               '(IF(SUM(bp.dCanthh) > 0, ' +
                               '(SUM(bp.dCanthh) + ' +
                               ' IFNULL((SELECT ' +
                               '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                               '          FROM ' +
                               '            bitacoradeajustes ba ' +
                               '          WHERE ' +
                               '            ba.sContrato = bp.sContrato AND ' +
                               '            ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                               '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                               '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                               '            ba.dFecha = bp.dIdFecha), 0)), SUM(bp.dCantidad))) AS Total,bp.dIdFecha FROM bitacoradepersonal bp ' +
                               'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                               'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="No") ' +
                               'WHERE bp.scontrato = :Orden AND ((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) and bp.dIdFecha between :Inicio and :termino ' +
                               'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'

{  Asql:Array[1..4] of String =( 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,a.sAnexo,SUM(bp.dAjuste) AS Ajuste,(IF(SUM(bp.dCanthh) > 0, '+
                                'SUM(bp.dCanthh), SUM(bp.dCantidad))) AS Total,bp.dIdFecha FROM bitacoradepersonal bp '+
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="No") '+
                                'WHERE bp.scontrato = :Orden AND ((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'
}
                              , 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,a.sDescripcion as Titulo, ' +
                                'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                'SUM(bp.dCantidad) AS Total,bp.dIdFecha, a.sAnexo ' +
                                'FROM bitacoradepersonal bp ' +
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="Si") ' +
                                'WHERE bp.scontrato = :Orden AND ((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) and bp.dIdFecha between :Inicio and :termino ' +
                                'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'

{                              , 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,SUM(bp.dAjuste) AS Ajuste,SUM(bp.dCantidad) AS Total,bp.dIdFecha, '+
                                'a.sAnexo FROM bitacoradepersonal bp inner join personal p on(p.sContrato=:ContratoBarco and '+
                                'p.sIdPersonal=bp.sIdPersonal) inner join anexos a on(a.sAnexo=p.sAnexo and '+
                                'a.sTipo="PERSONAL" and a.sTierra="Si") WHERE bp.scontrato = :Orden AND '+
                                '((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) ' +
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'
}                             
                              , 'select be.sIdEquipo as sIdRecurso,be.sDescripcion as Descripcion,e.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="EQUIPO" limit 1),"No existe") as Titulo,'+
                                'ifnull((select sAnexo from anexos where sTipo="EQUIPO" limit 1),"N/E") as sAnexo,'+
                                '(IFNULL(sum(cast(be.dCantHH as char)), 0) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = be.sContrato AND ' +
                                '           ba.sNumeroOrden = be.sNumeroOrden AND ' +
                                '           ba.sIdPartidaAnexo = be.sIdPersonal AND ' +
                                '           ba.sNumeroActividad = be.sNumeroActividad AND ' +
                                '           ba.dFecha = be.dIdFecha), 0)) as total,be.didfecha from bitacoradeequipos be inner join equipos e on '+
                                '(e.sContrato=:contratoBarco and e.sIdEquipo=be.sIdEquipo) inner join bitacoradeactividades '+
                                'ba on(ba.sContrato = be.sContrato and ba.dIdFecha = be.didfecha and ba.sNumeroOrden = '+
                                'be.sNumeroOrden and ba.iIdDiario = be.iIdDiario and ba.iIdTarea = be.iIdTarea and '+
                                'ba.iIdActividad=be.iIdActividad) where  be.sContrato=:Orden and ((:folio=-1) or (:folio<>-1 and '+
                                'be.sNumeroOrden=:folio)) and be.dIdFecha between :Inicio and :termino ' +
                                'group by e.sIdEquipo,be.didfecha order by e.iItemOrden,e.sIdEquipo,be.didfecha'

                              , 'select c.sIdCuenta,c.sIdPernocta as sIdRecurso,c.sDescripcion as Descripcion,c.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="PERNOCTA" limit 1),"No existe") as Titulo,'+
                                'ifnull((select sAnexo from anexos where sTipo="PERNOCTA" limit 1),"N/E") as sAnexo,'+
                                '((sum(bp.dCantHH) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = bp.sContrato AND ' +
                                '           ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                '           ba.dFecha = bp.dIdFecha), 0)) - ifnull((select ifnull(sum(bpr.dCantidad),0) from bitacoradepernocta bpr '+
                                'where bpr.sContrato = ba.sContrato and ba.didfecha=bpr.didfecha and (bpr.sNumeroOrden <> "@" '+
                                'and bpr.sNumeroOrden=ba.snumeroorden) group by bpr.sContrato),0)) as Total,bp.didfecha from cuentas c '+
                                'left join bitacoradeactividades ba on(ba.sContrato=:Orden and ((:Folio=-1) or (:Folio<>-1 and ba.sNumeroOrden=:Folio)) '+
                                ') left join bitacoradepersonal bp on (bp.sContrato =ba.sContrato and '+
                                'ba.sNumeroOrden=bp.sNumeroOrden and  bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                                'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad and bp.sTipoPernocta = c.sIdPernocta '+
                                ') left join personal p on (p.sContrato = :contratoBarco '+
                                'and p.sIdPersonal = bp.sIdPersonal) left join tiposdepersonal tp on ( p.sIdTipoPersonal = '+
                                'tp.sIdTipoPersonal) left join ordenesdetrabajo ot '+
                                'on (ot.sContrato = :Orden and ot.sNumeroOrden = ba.sNumeroOrden) left join pernoctan pr '+
                                'on (pr.sIdPernocta = ot.sIdPernocta) left join plataformas pl on(pl.sIdPlataforma=ot.sIdPlataforma) '+
                                'where ((:Folio=-1) or (:Folio<>-1 and bp.snumeroorden=:Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'group by c.sIdPernocta,bp.didfecha order by c.sIdPernocta,bp.didfecha '

                              );



var
  MryRecursos:TjvMemoryData;
  iGrupo:Integer;
  QrConsulta,QrFolios:TzReadOnlyQuery;
  sIdRecurso:String;
  ValTmp:Variant;
  CantTmp,dTotal:Double;
  iPosTmp,iPos:Integer;
  sDia:String;
  dsGeneradorGeneral:TfrxDBDataset;
begin
  QrFolios:=TzREadOnlyQUery.Create(nil);
  QrConsulta:=TzREadOnlyQUery.Create(nil);
  try
    try
      QrConsulta.Connection:=Connection.zConnection;
      QrFolios.Connection:=Connection.zConnection;
      dsGeneradorGeneral:=TfrxDBDataset.Create(nil);
      MryRecursos:=TjvMemoryData.Create(nil);

      dsGeneradorGeneral.UserName:='dsGeneradorGeneral';

      with MryRecursos do
      begin
        Active:=False;
        FieldDefs.Add('iGrupo', ftInteger, 0, True);
        FieldDefs.Add('sFolio', ftString, 100, True);
        FieldDefs.Add('sDescFolio', ftString, 500, True);
        FieldDefs.Add('sidrecurso', ftString, 100, True);
        FieldDefs.Add('sdescripcion', ftString, 500, True);
        FieldDefs.Add('sAnexo', ftString, 50, false);
        FieldDefs.Add('sTitulo', ftString, 500, false);
        FieldDefs.Add('smedida', ftString, 100, True);
        FieldDefs.Add('dTotal', FtFloat, 0, True);
        for IPos := 1 to 31 do
          FieldDefs.Add('dia' + InttoStr(IPos), FtFloat, 0, True);
        Active:=true;
      end;

      QrFolios.Active:=False;
      QrFolios.SQL.Text:= 'select ba.*,ot.mdescripcion as descFolio from bitacoradeactividades ba '+
                          'inner join ordenesdetrabajo ot '+
                          'on(ot.sContrato=ba.sContrato and ot.sNumeroOrden=ba.sNumeroorden) '+
                          'where ba.sContrato=:Contrato and ba.dIdFecha ' +
                          'between :Inicio and :termino';

      if IsGeneral then
        QrFolios.SQL.Add('group by ba.sContrato')
      else
        QrFolios.SQL.Add('group by ba.sContrato,ba.sNumeroOrden');

      QrFolios.ParamByName('Contrato').AsString:=ParamContrato;
      QrFolios.ParamByName('Inicio').AsDate:=ParamFechaI;
      QrFolios.ParamByName('Termino').AsDate:=ParamFechaT;
      QrFolios.Open;

      while not QrFolios.Eof do
      begin
        QrConsulta.Active:=False;              //Descripcion,e.sMedida as medida
        QrConsulta.SQL.Text:='SELECT tm.sContrato,tm.sIdTipoMovimiento,tm.sIdTipoMovimiento as sIdRecurso,tm.sDescripcion as Descripcion,' + #10 +
                              'tm.sTipo,tm.smedida as medida,me.didfecha,mf.sFolio,ifnull(SUM( me.sFactor ),0) AS Total,' + #10 +
                              'ifnull((select sdescripcion from anexos where sTipo="BARCO" limit 1)'+
                              ',"No existe") as Titulo, ' +
                               'ifnull((select sAnexo from anexos where sTipo="BARCO" limit 1)'+
                              ',"No existe") as sAnexo ' +
                              'FROM tiposdemovimiento AS tm '  +
                              'inner join movimientosdeembarcacion AS me '   +
                              'on(me.sContrato = tm.sContrato AND me.sClasificacion = tm.sIdTipoMovimiento) ' + #10 +
                              'inner join movimientosxfolios mf on (me.sContrato = mf.sContrato and me.sOrden=mf.sNumeroOrden' + #10 +
                              'and me.didfecha=mf.dIdFecha and me.iIdDiario = mf.iIdDiario) ' +
                              'WHERE tm.sContrato =:ContratoBarco AND me.sOrden =:Orden' + #10 +
                              'and ((:Folio=-1) or (:Folio<>-1 and mf.sFolio=:Folio)) ' + #10 +
                              'AND tm.sClasificacion = "Movimiento de Barco" and' + #10 +
                              'tm.lGenera ="Si" and me.didfecha between :Inicio and :termino ' +
                              'group by me.sContrato,tm.sIdTipoMovimiento,me.didfecha '+
                              'order by me.sContrato,tm.sIdTipoMovimiento,me.didfecha';

        {if IsGeneral then
          QrConsulta.SQL.Add('group by me.sContrato,tm.sIdTipoMovimiento,me.didfecha '+
                             'order by me.sContrato,tm.sIdTipoMovimiento,me.didfecha')
        else
          QrConsulta.SQL.Add('group by me.sContrato,tm.sIdTipoMovimiento,me.didfecha '+
                             'order by me.sContrato,tm.sIdTipoMovimiento,me.didfecha');}

        QrConsulta.ParamByName('ContratoBarco').AsString:=Global_Contrato_Barco;
        QrConsulta.ParamByName('Orden').AsString:=ParamContrato;

        if IsGeneral then
          QrConsulta.ParamByName('Folio').AsString:='-1'
        else
          QrConsulta.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
        QrConsulta.ParamByName('Inicio').AsDate:=ParamFechaI;
        QrConsulta.ParamByName('Termino').AsDate:=ParamFechaT;
        QrConsulta.Open;

        sIdRecurso:='';

        while not QrConsulta.Eof do
        begin
          if sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString then
          begin
            if MryRecursos.State=dsInsert then
            begin
              if dTotal<>xRound(ValTmp,Mantisa) then
              begin
                MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat := MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat +
                  + (xRound(ValTmp,Mantisa)-dTotal);
                MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
              end
              else
                MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
              MryRecursos.Post;
            end;

            ValTmp:=0;
            CantTmp:=0;
            iPosTmp:=0;
            dTotal:=0;

            sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.Append;
            MryRecursos.FieldByName('iGrupo').AsInteger:=QrFolios.RecNo;
            if IsGeneral then
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sContrato').AsString
            else
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
            MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
            MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
            MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
            MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
            MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
            MryRecursos.FieldByName('dTotal').AsFloat:=0;
            for iPos := 1 to 31 do
              MryRecursos.FieldValues['dia' + Trim(IntToStr(iPos)) ] := 0.000000 ;
          end;

          sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;
          MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
          ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
          dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;
          if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
          begin
            CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
            iPosTmp:=DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime);
          end;

          QrConsulta.Next;
        end;

        if MryRecursos.State=dsInsert then
        begin
          if dTotal<>xRound(ValTmp,Mantisa) then
          begin
            MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat := MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat +
              + (xRound(ValTmp,Mantisa)-dTotal);
            MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
          end
          else
            MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
          MryRecursos.Post;
        end;


        QrConsulta.Active:=fALSE;
        QrConsulta.SQL.text:= 'select s.sNumeroOrden, b.sNumeroActividad, b.sTrazabilidad,i.sIdInsumo as sIdRecurso,'+
                              'i.mDescripcion as descripcion,i.sMedida as medida, sum(b.dCantidad) as total,i.sIdProveedor,'+
                              'b.dFechaSalida as dIdFecha,ifnull((select sdescripcion from anexos where sTipo="MATERIAL" '+
                              'and sTierra="No" order by iorden limit 1),"No existe") as Titulo, ' +
                              'ifnull((select sAnexo from anexos where sTipo="MATERIAL" and sTierra="No" '+
                              'order by iorden limit 1),"No existe") as sAnexo ' +
                              'from almacen_salida s '+
                              'inner join bitacoradesalida b '+
                              'on(b.sContrato = s.sContrato and b.iFolioSalida = s.iFolioSalida and b.sext = s.sext) '+
                              'inner join insumos i '+
                              'on (i.sContrato =:Contrato and i.sIdInsumo = b.sIdInsumo and i.sTrazabilidad = b.sTrazabilidad ) '+
                              //'INNER JOIN actividadesxorden ao ON ( ' +
                              //'ao.sContrato = b.sContrato and ao.sIdConvenio =:Convenio and ao.sNumeroOrden = b.sNumeroOrden and ' +
                              //'ao.sNumeroActividad = b.sNumeroActividad and ao.sTipoActividad = "Actividad" ) ' +
                              'where s.sContrato =:orden and ((:Folio=-1) or (:Folio<>-1 and s.sNumeroOrden=:Folio)) '+
                              'and b.dFechaSalida between :Inicio and :termino '+
                              'group by b.sContrato,b.sIdInsumo,didfecha Order by b.sNumeroOrden,b.sIdInsumo,didfecha';
        QrConsulta.ParamByName('Orden').AsString    := ParamContrato;
        if IsGeneral then
          QrConsulta.ParamByName('Folio').AsString    := '-1'
        else
          QrConsulta.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
        QrConsulta.ParamByName('Inicio').AsDate:=ParamFechaI;
        QrConsulta.ParamByName('Termino').AsDate:=ParamFechaT;
        QrConsulta.ParamByName('Contrato').AsString := global_Contrato_Barco;
        QrConsulta.Open;

        sIdRecurso:='';

        while not QrConsulta.Eof do
        begin
          if sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString then
          begin
            if MryRecursos.State=dsInsert then
            begin
              if dTotal<>xRound(ValTmp,Mantisa) then
              begin
                MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat := MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat +
                  + (xRound(ValTmp,Mantisa)-dTotal);
                MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
              end
              else
                MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
              MryRecursos.Post;
            end;

            ValTmp:=0;
            CantTmp:=0;
            iPosTmp:=0;
            dTotal:=0;

            sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.Append;
            MryRecursos.FieldByName('iGrupo').AsInteger:=QrFolios.RecNo;
            if IsGeneral then
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sContrato').AsString
            else
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
            MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
            MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
            MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
            MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
            MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
            MryRecursos.FieldByName('dTotal').AsFloat:=0;
            for iPos := 1 to 31 do
              MryRecursos.FieldValues['dia' + Trim(IntToStr(iPos)) ] := 0.000000 ;
          end;

          sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;
          MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
          ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
          dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;
          if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
          begin
            CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
            iPosTmp:=DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime);
          end;

          QrConsulta.Next;
        end;

        if MryRecursos.State=dsInsert then
        begin
          if dTotal<>xRound(ValTmp,Mantisa) then
          begin
            MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat := MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat +
              + (xRound(ValTmp,Mantisa)-dTotal);
            MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
          end
          else
            MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
          MryRecursos.Post;
        end;

        for iGrupo := 1 to length(Asql) do
        begin
          QrConsulta.Active:=false;
          QrConsulta.SQL.Text:=ASql[iGrupo];
          if IsGeneral then
            QrConsulta.ParamByName('Folio').AsString    := '-1'
          else
            QrConsulta.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
          QrConsulta.ParamByName('Inicio').AsDate:=ParamFechaI;
          QrConsulta.ParamByName('Termino').AsDate:=ParamFechaT;
          QrConsulta.ParamByName('ContratoBarco').AsString:=Global_Contrato_Barco;
          QrConsulta.ParamByName('Orden').AsString:=ParamContrato;
          QrConsulta.Open;

          sIdRecurso:='';

          while not QrConsulta.Eof do
          begin
            if sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString then
            begin
              if MryRecursos.State=dsInsert then
              begin
                if dTotal<>xRound(ValTmp,Mantisa) then
                begin
                  MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat := MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat +
                    + (xRound(ValTmp,Mantisa)-dTotal);
                  MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
                end
                else
                  MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
                MryRecursos.Post;
              end;

              ValTmp:=0;
              CantTmp:=0;
              iPosTmp:=0;
              dTotal:=0;

              sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
              MryRecursos.Append;
              MryRecursos.FieldByName('iGrupo').AsInteger:=QrFolios.RecNo;
              if IsGeneral then
                MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sContrato').AsString
              else
                MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
              MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
              MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
              MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
              MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
              MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
              MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
              MryRecursos.FieldByName('dTotal').AsFloat:=0;
              for iPos := 1 to 31 do
                MryRecursos.FieldValues['dia' + Trim(IntToStr(iPos)) ] := 0.000000 ;
            end;

            sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;

            if iGrupo<3 then
            begin
              MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat;
              ValTmp:=ValTmp + (QrConsulta.FieldByName('total').Value + + QrConsulta.FieldByName('Ajuste').AsFloat);
              if CantTmp<(xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat) then
              begin
                CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat;
                iPosTmp:=DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime);
              end;
            end
            else
              if (iGrupo=4) and (QrConsulta.FieldByName('sIdcuenta').AsString='1') then
              begin
                connection.QryBusca.Active:=false;
                connection.QryBusca.SQL.Text:='select sum(bpc.dCantidad) as dCantidad' + #10 +
                                              'from bitacoradepersonal_cuadre bpc' + #10 +
                                              'inner join personal p' + #10 +
                                              'on ( p.sContrato = :ContratoBarco and bpc.sIdPersonal = p.sIdPersonal )' + #10 +
                                              'inner join tiposdepersonal tp' + #10 +
                                              'on (tp.sIdTipoPersonal = p.sIdTipoPersonal )' + #10 +
                                              'where bpc.sContrato =:Orden and bpc.snumeroorden=:Folio' + #10 +
                                              'and tp.lPernocta="Si" and bpc.dIdFecha=:Fecha' ;

                connection.QryBusca.ParamByName('ContratoBarco').AsString:= global_contrato_barco;
                connection.QryBusca.ParamByName('Orden').AsString:=ParamContrato;
                connection.QryBusca.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
                connection.QryBusca.ParamByName('Fecha').AsDate:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
                connection.QryBusca.Open;
                if connection.QryBusca.RecordCount=1 then
                begin
                  MryRecursos.FieldByName(sDia).AsFloat:=xRound(QrConsulta.FieldByName('total').AsFloat,Mantisa)+connection.QryBusca.FieldByName('dCantidad').AsFloat;
                  ValTmp:=ValTmp + (QrConsulta.FieldByName('total').Value + +connection.QryBusca.FieldByName('dCantidad').AsFloat);
                  if CantTmp<(xRound(QrConsulta.FieldByName('total').Value,Mantisa)++connection.QryBusca.FieldByName('dCantidad').AsFloat) then
                  begin
                    CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + +connection.QryBusca.FieldByName('dCantidad').AsFloat;
                    iPosTmp:=DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime);
                  end;
                end
                else
                begin
                  MryRecursos.FieldByName(sDia).AsFloat:=xRound(QrConsulta.FieldByName('total').AsFloat,Mantisa);
                  ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
                  if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
                  begin
                    CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
                    iPosTmp:=DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime);
                  end;


                end;
              end
              else
              begin
                MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
                ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
                if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
                begin
                  CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
                  iPosTmp:=DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime);
                end;
              end;

            dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;

            QrConsulta.Next;
          end;

          if MryRecursos.State=dsInsert then
          begin
            if dTotal<>xRound(ValTmp,Mantisa) then
            begin
              MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat := MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat +
                + (xRound(ValTmp,Mantisa)-dTotal);
              MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
            end
            else
              MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
            MryRecursos.Post;
          end;


        end;
        QrFolios.Next;
      end;

      dsGeneradorGeneral.DataSet:=MryRecursos;
      dsGeneradorGeneral.FieldAliases.Clear;

      Reporte.DataSets.Add(dsGeneradorGeneral);


    except
      on e:Exception do
        Showmessage(e.Message + ', ' + e.ClassName);
    end;

  finally
    QrConsulta.Destroy;
  end;
end;

Procedure TdGenerador(ParamContrato:String;ParamFechaI,ParamFechaT:TDateTime;Var MryRecursos:TjvMemoryData;IsGeneral:Boolean=True;Progress:TAdvProgress=Nil);
const
  Asql:Array[1..4] of String =('SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,a.sDescripcion as Titulo,a.sAnexo, ' +
                               'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                               '(IF(SUM(bp.dCanthh) > 0, SUM(bp.dCanthh), SUM(bp.dCantidad))) AS Total,bp.dIdFecha ' +
                               'FROM bitacoradepersonal bp ' +
                               'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                               'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="No") ' +
                               'WHERE bp.scontrato = :Orden AND ((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) and bp.dIdFecha between :Inicio and :termino ' +
                               'GROUP BY bp.sIdPersonal,bp.didfecha ' +
                               'order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'

  {Asql:Array[1..4] of String =( 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,a.sAnexo,SUM(bp.dAjuste) AS Ajuste,(IF(SUM(bp.dCanthh) > 0, '+
                                'SUM(bp.dCanthh), SUM(bp.dCantidad))) AS Total,bp.dIdFecha FROM bitacoradepersonal bp '+
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="No") '+
                                'WHERE bp.scontrato = :Orden AND ((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'
}
                              , 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,a.sDescripcion as Titulo, ' +
                                'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                'SUM(bp.dCantidad) AS Total,bp.dIdFecha, a.sAnexo ' +
                                'FROM bitacoradepersonal bp ' +
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="Si") ' +
                                'WHERE bp.scontrato = :Orden AND ((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) and bp.dIdFecha between :Inicio and :termino ' +
                                'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'

{                              , 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,SUM(bp.dAjuste) AS Ajuste,SUM(bp.dCantidad) AS Total,bp.dIdFecha, '+
                                'a.sAnexo FROM bitacoradepersonal bp inner join personal p on(p.sContrato=:ContratoBarco and '+
                                'p.sIdPersonal=bp.sIdPersonal) inner join anexos a on(a.sAnexo=p.sAnexo and '+
                                'a.sTipo="PERSONAL" and a.sTierra="Si") WHERE bp.scontrato = :Orden AND '+
                                '((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) ' +
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'}

                              , 'select be.sIdEquipo as sIdRecurso,be.sDescripcion as Descripcion,e.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="EQUIPO" limit 1),"No existe") as Titulo,'+
                                'ifnull((select sAnexo from anexos where sTipo="EQUIPO" limit 1),"N/E") as sAnexo,'+
                                '(IFNULL(sum(be.dCantHH), 0) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = be.sContrato AND ' +
                                '           ba.sNumeroOrden = be.sNumeroOrden AND ' +
                                '           ba.sIdPartidaAnexo = be.sIdPersonal AND ' +
                                '           ba.sNumeroActividad = be.sNumeroActividad AND ' +
                                '           ba.dFecha = be.dIdFecha), 0)) as total,be.didfecha from bitacoradeequipos be inner join equipos e on '+
                                '(e.sContrato=:contratoBarco and e.sIdEquipo=be.sIdEquipo) inner join bitacoradeactividades '+
                                'ba on(ba.sContrato = be.sContrato and ba.dIdFecha = be.didfecha and ba.sNumeroOrden = '+
                                'be.sNumeroOrden and ba.iIdDiario = be.iIdDiario and ba.iIdTarea = be.iIdTarea and '+
                                'ba.iIdActividad=be.iIdActividad) where  be.sContrato=:Orden and ((:folio=-1) or (:folio<>-1 and '+
                                'be.sNumeroOrden=:folio)) and be.dIdFecha between :Inicio and :termino ' +
                                'group by e.sIdEquipo,be.didfecha order by e.iItemOrden,e.sIdEquipo,be.didfecha'

                              , 'select c.sIdCuenta,c.sIdPernocta as sIdRecurso,c.sDescripcion as Descripcion,c.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="PERNOCTA" limit 1),"No existe") as Titulo,'+
                                'ifnull((select sAnexo from anexos where sTipo="PERNOCTA" limit 1),"N/E") as sAnexo,'+
                                '((sum(bp.dCantHH) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = bp.sContrato AND ' +
                                '           ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                '           ba.dFecha = bp.dIdFecha), 0)) - ifnull((select ifnull(sum(bpr.dCantidad),0) from bitacoradepernocta bpr '+
                                'where bpr.sContrato = ba.sContrato and ba.didfecha=bpr.didfecha and (bpr.sNumeroOrden <> "@" '+
                                'and bpr.sNumeroOrden=ba.snumeroorden) group by bpr.sContrato),0)) as Total,bp.didfecha from cuentas c '+
                                'left join bitacoradeactividades ba on(ba.sContrato=:Orden and ((:Folio=-1) or (:Folio<>-1 and ba.sNumeroOrden=:Folio)) '+
                                ') left join bitacoradepersonal bp on (bp.sContrato =ba.sContrato and '+
                                'ba.sNumeroOrden=bp.sNumeroOrden and  bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                                'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad and bp.sTipoPernocta = c.sIdPernocta '+
                                ') left join personal p on (p.sContrato = :contratoBarco '+
                                'and p.sIdPersonal = bp.sIdPersonal) left join tiposdepersonal tp on ( p.sIdTipoPersonal = '+
                                'tp.sIdTipoPersonal) left join ordenesdetrabajo ot '+
                                'on (ot.sContrato = :Orden and ot.sNumeroOrden = ba.sNumeroOrden) left join pernoctan pr '+
                                'on (pr.sIdPernocta = ot.sIdPernocta) left join plataformas pl on(pl.sIdPlataforma=ot.sIdPlataforma) '+
                                'where ((:Folio=-1) or (:Folio<>-1 and bp.snumeroorden=:Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'group by c.sIdPernocta,bp.didfecha order by c.sIdPernocta,bp.didfecha '

                              );


var
  iGrupo:Integer;
  QrConsulta,QrFolios:TzReadOnlyQuery;
  sIdRecurso:String;
  ValTmp:Variant;
  CantTmp,dTotal:Double;
  iPos,TotalDias,TmpAv:Integer;
  sDia:String;
  dDia,dPosTmp:TDateTime;
  sTipo:String;

begin
  QrFolios:=TzREadOnlyQUery.Create(nil);
  QrConsulta:=TzREadOnlyQUery.Create(nil);
  try
    try
      QrConsulta.Connection:=Connection.zConnection;
      QrFolios.Connection:=Connection.zConnection;
      if MryRecursos.Active then
      begin
        MryRecursos.EmptyTable;
      end;

      TotalDias:=DaysBetween(ParamFechaT,ParamFechaI) + 1;
      dDia:= ParamFechaI;
      with MryRecursos do
      begin
        Active:=False;
        MryRecursos.FieldDefs.Clear;
        FieldDefs.Add('iGrupo', ftInteger, 0, True);
        FieldDefs.Add('sTipo', ftString, 20, True);
        FieldDefs.Add('sFolio', ftString, 100, True);
        FieldDefs.Add('sDescFolio', ftString, 500, True);
        FieldDefs.Add('sidrecurso', ftString, 100, True);
        FieldDefs.Add('sdescripcion', ftString, 500, True);
        FieldDefs.Add('sAnexo', ftString, 50, false);
        FieldDefs.Add('sTitulo', ftString, 500, false);
        FieldDefs.Add('smedida', ftString, 100, True);

        for IPos := 1 to TotalDias do
        begin
          sDia:=FormatDateTime('ddddd',ddia);
          FieldDefs.Add(sDia, FtFloat, 0, True);
          FieldDefs.Add('Limit'+sDia, FtFloat, 0, True);
          dDia:=IncDay(dDia);
          if Progress<>nil then
          begin
            Progress.Position:=Trunc((IPos*10)/TotalDias);
            Application.ProcessMessages;
          end;
        end;
        FieldDefs.Add('dTotal', FtFloat, 0, True);
        Active:=true;
        FieldList.Update;
        if Progress<>nil then
        begin
          Progress.Position:=15;
          Application.ProcessMessages;
        end;
      end;

      QrFolios.Active:=False;
      QrFolios.SQL.Text:= 'select ba.*,ot.mdescripcion as descFolio from bitacoradeactividades ba '+
                          'inner join ordenesdetrabajo ot '+
                          'on(ot.sContrato=ba.sContrato and ot.sNumeroOrden=ba.sNumeroorden) '+
                          'where ba.sContrato=:Contrato and ba.dIdFecha ' +
                          'between :Inicio and :termino';

      if IsGeneral then
        QrFolios.SQL.Add('group by ba.sContrato')
      else
        QrFolios.SQL.Add('group by ba.sContrato,ba.sNumeroOrden');

      QrFolios.ParamByName('Contrato').AsString:=ParamContrato;
      QrFolios.ParamByName('Inicio').AsDate:=ParamFechaI;
      QrFolios.ParamByName('Termino').AsDate:=ParamFechaT;
      QrFolios.Open;

      while not QrFolios.Eof do
      begin
        sTipo:='BARCO';
        QrConsulta.Active:=False;              //Descripcion,e.sMedida as medida
        QrConsulta.SQL.Text:='SELECT tm.sContrato,tm.sIdTipoMovimiento,tm.sIdTipoMovimiento as sIdRecurso,tm.sDescripcion as Descripcion,' + #10 +
                              'tm.sTipo,tm.smedida as medida,me.didfecha,mf.sFolio,ifnull(SUM( me.sFactor ),0) AS Total,' + #10 +
                              'ifnull((select sdescripcion from anexos where sTipo="BARCO" limit 1)'+
                              ',"No existe") as Titulo, ' +
                               'ifnull((select sAnexo from anexos where sTipo="BARCO" limit 1)'+
                              ',"No existe") as sAnexo ' +
                              'FROM tiposdemovimiento AS tm '  +
                              'inner join movimientosdeembarcacion AS me '   +
                              'on(me.sContrato = tm.sContrato AND me.sClasificacion = tm.sIdTipoMovimiento) ' + #10 +
                              'inner join movimientosxfolios mf on (me.sContrato = mf.sContrato and me.sOrden=mf.sNumeroOrden' + #10 +
                              'and me.didfecha=mf.dIdFecha and me.iIdDiario = mf.iIdDiario) ' +
                              'WHERE tm.sContrato =:ContratoBarco AND me.sOrden =:Orden' + #10 +
                              'and ((:Folio=-1) or (:Folio<>-1 and mf.sFolio=:Folio)) ' + #10 +
                              'AND tm.sClasificacion = "Movimiento de Barco" and' + #10 +
                              'tm.lGenera ="Si" and me.didfecha between :Inicio and :termino ' +
                              'group by me.sContrato,tm.sIdTipoMovimiento,me.didfecha '+
                              'order by me.sContrato,tm.sIdTipoMovimiento,me.didfecha';

        {if IsGeneral then
          QrConsulta.SQL.Add('group by me.sContrato,tm.sIdTipoMovimiento,me.didfecha '+
                             'order by me.sContrato,tm.sIdTipoMovimiento,me.didfecha')
        else
          QrConsulta.SQL.Add('group by me.sContrato,tm.sIdTipoMovimiento,me.didfecha '+
                             'order by me.sContrato,tm.sIdTipoMovimiento,me.didfecha');}

        QrConsulta.ParamByName('ContratoBarco').AsString:=Global_Contrato_Barco;
        QrConsulta.ParamByName('Orden').AsString:=ParamContrato;

        if IsGeneral then
          QrConsulta.ParamByName('Folio').AsString:='-1'
        else
          QrConsulta.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
        QrConsulta.ParamByName('Inicio').AsDate:=ParamFechaI;
        QrConsulta.ParamByName('Termino').AsDate:=ParamFechaT;
        QrConsulta.Open;

        sIdRecurso:='';
        TmpAv:=0;
        while not QrConsulta.Eof do
        begin
          if sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString then
          begin
            if MryRecursos.State=dsInsert then
            begin
              if dTotal<>xRound(ValTmp,Mantisa) then
              begin
                sDia:=FormatDateTime('ddddd',dPosTmp);
                MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
                  + (xRound(ValTmp,Mantisa)-dTotal);
                MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;
                MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
              end
              else
                MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
              MryRecursos.Post;
            end;

            ValTmp:=0;
            CantTmp:=0;
            dTotal:=0;

            sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.Append;
            MryRecursos.FieldByName('iGrupo').AsInteger:=QrFolios.RecNo;
            MryRecursos.FieldByName('sTipo').AsString:= sTipo;
            if IsGeneral then
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sContrato').AsString
            else
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
            MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
            MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
            MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
            MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
            MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
            MryRecursos.FieldByName('dTotal').AsFloat:=0;
            dDia:= ParamFechaI;
            for iPos := 1 to TotalDias do
            begin
              sDia:=FormatDateTime('ddddd',ddia);
              MryRecursos.FieldValues[sDia] := 0.000000 ;
              MryRecursos.FieldValues['Limit'+sDia] := 0.000000 ;
              dDia:=IncDay(dDia);
            end;

          end;


          //sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;

          sDia:=FormatDateTime('ddddd',QrConsulta.FieldByName('dIdFecha').AsDateTime);
          MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
          MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;
          ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
          dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;
          if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
          begin
            CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
            dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
          end;

          if Progress<>nil then
          begin
            if TmpAv<>Trunc((QrConsulta.RecNo*10)/QrConsulta.RecordCount) then
            begin
              Progress.Position:=Progress.Position +(Trunc((QrConsulta.RecNo*10)/QrConsulta.RecordCount) - TmpAv) ;
              Application.ProcessMessages;
              TmpAv:= Trunc((QrConsulta.RecNo*10)/QrConsulta.RecordCount);
            end;
          end;

          QrConsulta.Next;
        end;

        if MryRecursos.State=dsInsert then
        begin
          if dTotal<>xRound(ValTmp,Mantisa) then
          begin
            sDia:=FormatDateTime('ddddd',dPosTmp);
            MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
              + (xRound(ValTmp,Mantisa)-dTotal);

            MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;
            MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
          end
          else
            MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
          MryRecursos.Post;
        end;
        sTipo:='MATERIAL';
        QrConsulta.Active:=fALSE;
        QrConsulta.SQL.text:= 'select s.sNumeroOrden, b.sNumeroActividad, b.sTrazabilidad,i.sIdInsumo as sIdRecurso,'+
                              'i.mDescripcion as descripcion,i.sMedida as medida, sum(b.dCantidad) as total,i.sIdProveedor,'+
                              'b.dFechaSalida as dIdFecha,ifnull((select sdescripcion from anexos where sTipo="MATERIAL" '+
                              'and sTierra="No" order by iorden limit 1),"No existe") as Titulo, ' +
                              'ifnull((select sAnexo from anexos where sTipo="MATERIAL" and sTierra="No" '+
                              'order by iorden limit 1),"No existe") as sAnexo ' +
                              'from almacen_salida s '+
                              'inner join bitacoradesalida b '+
                              'on(b.sContrato = s.sContrato and b.iFolioSalida = s.iFolioSalida and b.sext = s.sext) '+
                              'inner join insumos i '+
                              'on (i.sContrato =:Contrato and i.sIdInsumo = b.sIdInsumo and i.sTrazabilidad = b.sTrazabilidad ) '+
                              //'INNER JOIN actividadesxorden ao ON ( ' +
                              //'ao.sContrato = b.sContrato and ao.sIdConvenio =:Convenio and ao.sNumeroOrden = b.sNumeroOrden and ' +
                              //'ao.sNumeroActividad = b.sNumeroActividad and ao.sTipoActividad = "Actividad" ) ' +
                              'where s.sContrato =:orden and ((:Folio=-1) or (:Folio<>-1 and s.sNumeroOrden=:Folio)) '+
                              'and b.dFechaSalida between :Inicio and :termino '+
                              'group by b.sContrato,b.sIdInsumo,didfecha Order by b.sNumeroOrden,b.sIdInsumo,didfecha';
        QrConsulta.ParamByName('Orden').AsString    := ParamContrato;
        if IsGeneral then
          QrConsulta.ParamByName('Folio').AsString    := '-1'
        else
          QrConsulta.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
        QrConsulta.ParamByName('Inicio').AsDate:=ParamFechaI;
        QrConsulta.ParamByName('Termino').AsDate:=ParamFechaT;
        QrConsulta.ParamByName('Contrato').AsString := global_Contrato_Barco;
        QrConsulta.Open;

        sIdRecurso:='';
        TmpAv:=0;
        while not QrConsulta.Eof do
        begin
          if sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString then
          begin
            if MryRecursos.State=dsInsert then
            begin
              if dTotal<>xRound(ValTmp,Mantisa) then
              begin
                sDia:=FormatDateTime('ddddd',dPosTmp);
                MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
                  + (xRound(ValTmp,Mantisa)-dTotal);
                MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

                MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
              end
              else
                MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
              MryRecursos.Post;
            end;





            ValTmp:=0;
            CantTmp:=0;
            dTotal:=0;

            sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.Append;
            MryRecursos.FieldByName('iGrupo').AsInteger:=QrFolios.RecNo;
            MryRecursos.FieldByName('sTipo').AsString:= sTipo;
            if IsGeneral then
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sContrato').AsString
            else
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
            MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
            MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
            MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
            MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
            MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
            MryRecursos.FieldByName('dTotal').AsFloat:=0;
            dDia:= ParamFechaI;
            for iPos := 1 to TotalDias do
            begin
              sDia:=FormatDateTime('ddddd',ddia);
              MryRecursos.FieldValues[sDia] := 0.000000 ;
              MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

              dDia:=IncDay(dDia);
            end;
          end;


          //sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;
          sDia:=FormatDateTime('ddddd',QrConsulta.FieldByName('dIdFecha').AsDateTime);
          MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
          MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

          ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
          dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;
          if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
          begin
            CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
            dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
          end;

          if Progress<>nil then
          begin
            if TmpAv<>Trunc((QrConsulta.RecNo*10)/QrConsulta.RecordCount) then
            begin
              Progress.Position:=Progress.Position +(Trunc((QrConsulta.RecNo*10)/QrConsulta.RecordCount) - TmpAv) ;
              Application.ProcessMessages;
              TmpAv:= Trunc((QrConsulta.RecNo*10)/QrConsulta.RecordCount);
            end;
          end;

          QrConsulta.Next;
        end;

        if MryRecursos.State=dsInsert then
        begin
          if dTotal<>xRound(ValTmp,Mantisa) then
          begin
            sDia:=FormatDateTime('ddddd',dPosTmp);
            MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
              + (xRound(ValTmp,Mantisa)-dTotal);
            MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

            MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
          end
          else
            MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
          MryRecursos.Post;
        end;

        for iGrupo := 1 to length(Asql) do
        begin
          if iGrupo<3 then
            sTipo:='PERSONAL';
          if iGrupo=3 then
            sTipo:='EQUIPO';
          if iGrupo=4 then
            sTipo:='PERNOCTA';



          QrConsulta.Active:=false;
          QrConsulta.SQL.Text:=ASql[iGrupo];
          if IsGeneral then
            QrConsulta.ParamByName('Folio').AsString    := '-1'
          else
            QrConsulta.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
          QrConsulta.ParamByName('Inicio').AsDate:=ParamFechaI;
          QrConsulta.ParamByName('Termino').AsDate:=ParamFechaT;
          QrConsulta.ParamByName('ContratoBarco').AsString:=Global_Contrato_Barco;
          QrConsulta.ParamByName('Orden').AsString:=ParamContrato;
          QrConsulta.Open;

          sIdRecurso:='';
          TmpAv:=0;
          while not QrConsulta.Eof do
          begin
            if sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString then
            begin
              if MryRecursos.State=dsInsert then
              begin
                if dTotal<>xRound(ValTmp,Mantisa) then
                begin
                  sDia:=FormatDateTime('ddddd',dPosTmp);
                  MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
                    + (xRound(ValTmp,Mantisa)-dTotal);
                  MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

                  MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
                end
                else
                  MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
                MryRecursos.Post;
              end;

              ValTmp:=0;
              CantTmp:=0;
              dTotal:=0;

              sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
              MryRecursos.Append;
              MryRecursos.FieldByName('iGrupo').AsInteger:=QrFolios.RecNo;
              MryRecursos.FieldByName('sTipo').AsString:= sTipo;
              if IsGeneral then
                MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sContrato').AsString
              else
                MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
              MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
              MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
              MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
              MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
              MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
              MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
              MryRecursos.FieldByName('dTotal').AsFloat:=0;
              dDia:= ParamFechaI;
              for iPos := 1 to TotalDias do
              begin
                sDia:=FormatDateTime('ddddd',ddia);
                MryRecursos.FieldValues[sDia] := 0.000000 ;
                MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

                dDia:=IncDay(dDia);
              end;
            end;

            //sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;
            sDia:=FormatDateTime('ddddd',QrConsulta.FieldByName('dIdFecha').AsDateTime);
            if iGrupo<3 then
            begin
              MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat;

              ValTmp:=ValTmp + (QrConsulta.FieldByName('total').Value + + QrConsulta.FieldByName('Ajuste').AsFloat);
              if CantTmp<(xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat) then
              begin
                CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat;
                dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
              end;
            end
            else
              if (iGrupo=4) and (QrConsulta.FieldByName('sIdcuenta').AsString='1') then
              begin
                connection.QryBusca.Active:=false;
                connection.QryBusca.SQL.Text:='select sum(bpc.dCantidad) as dCantidad' + #10 +
                                              'from bitacoradepersonal_cuadre bpc' + #10 +
                                              'inner join personal p' + #10 +
                                              'on ( p.sContrato = :ContratoBarco and bpc.sIdPersonal = p.sIdPersonal )' + #10 +
                                              'inner join tiposdepersonal tp' + #10 +
                                              'on (tp.sIdTipoPersonal = p.sIdTipoPersonal )' + #10 +
                                              'where bpc.sContrato =:Orden and bpc.snumeroorden=:Folio' + #10 +
                                              'and tp.lPernocta="Si" and bpc.dIdFecha=:Fecha' ;

                connection.QryBusca.ParamByName('ContratoBarco').AsString:= global_contrato_barco;
                connection.QryBusca.ParamByName('Orden').AsString:=ParamContrato;
                connection.QryBusca.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
                connection.QryBusca.ParamByName('Fecha').AsDate:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
                connection.QryBusca.Open;
                if connection.QryBusca.RecordCount=1 then
                begin
                  MryRecursos.FieldByName(sDia).AsFloat:=xRound(QrConsulta.FieldByName('total').AsFloat,Mantisa)+connection.QryBusca.FieldByName('dCantidad').AsFloat;
                  ValTmp:=ValTmp + (QrConsulta.FieldByName('total').Value + +connection.QryBusca.FieldByName('dCantidad').AsFloat);
                  if CantTmp<(xRound(QrConsulta.FieldByName('total').Value,Mantisa)++connection.QryBusca.FieldByName('dCantidad').AsFloat) then
                  begin
                    CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + +connection.QryBusca.FieldByName('dCantidad').AsFloat;
                    dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
                  end;
                end
                else
                begin
                  MryRecursos.FieldByName(sDia).AsFloat:=xRound(QrConsulta.FieldByName('total').AsFloat,Mantisa);
                  ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
                  if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
                  begin
                    CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
                    dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
                  end;


                end;
              end
              else
              begin
                MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
                ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
                if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
                begin
                  CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
                  dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
                end;
              end;
            MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

            dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;
            if Progress<>nil then
            begin
              if TmpAv<>Trunc((QrConsulta.RecNo*10)/QrConsulta.RecordCount) then
              begin
                Progress.Position:=Progress.Position +(Trunc((QrConsulta.RecNo*10)/QrConsulta.RecordCount) - TmpAv) ;
                Application.ProcessMessages;
                TmpAv:= Trunc((QrConsulta.RecNo*10)/QrConsulta.RecordCount);
              end;
            end;
            QrConsulta.Next;
          end;

          {
            dDia:= ParamFechaI;
            for IPos := 1 to TotalDias do
        begin
          sDia:=FormatDateTime('ddddd',ddia);
          FieldDefs.Add(sDia, FtFloat, 0, True);
          dDia:=IncDay(dDia);
        end;
        sDia:=FormatDateTime('ddddd',dPosTmp);
        }
          if MryRecursos.State=dsInsert then
          begin
            if dTotal<>xRound(ValTmp,Mantisa) then
            begin
              sDia:=FormatDateTime('ddddd',dPosTmp);
              MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
                + (xRound(ValTmp,Mantisa)-dTotal);
              MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

              MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
            end
            else
              MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
            MryRecursos.Post;
          end;


        end;
        QrFolios.Next;
      end;

    except
      on e:Exception do
        Showmessage(e.Message + ', ' + e.ClassName);
    end;

  finally
    QrConsulta.Destroy;
  end;
end;

Procedure TdGenerador(ParamContrato:String;ParamFechaI,ParamFechaT:TDateTime;Var MryRecursos:TClientDataSet;IsGeneral:Boolean=True); Overload;
const
  Asql:Array[1..4] of String =('SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,a.sDescripcion as Titulo,a.sAnexo, ' +
                               'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                               '(IF(SUM(bp.dCanthh) > 0, (SUM(bp.dCanthh) + ' +
                               ' IFNULL((SELECT ' +
                               '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                               '          FROM ' +
                               '            bitacoradeajustes ba ' +
                               '          WHERE ' +
                               '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                               '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                               '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)), SUM(bp.dCantidad))) AS Total,bp.dIdFecha ' +
                               'FROM bitacoradepersonal bp ' +
                               'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                               'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="No") ' +
                               'WHERE bp.scontrato = :Orden AND ((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) and bp.dIdFecha between :Inicio and :termino ' +
                               'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'

  {Asql:Array[1..4] of String =( 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,a.sAnexo,SUM(bp.dAjuste) AS Ajuste,(IF(SUM(bp.dCanthh) > 0, '+
                                'SUM(bp.dCanthh), SUM(bp.dCantidad))) AS Total,bp.dIdFecha FROM bitacoradepersonal bp '+
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="No") '+
                                'WHERE bp.scontrato = :Orden AND ((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'}

                              , 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,a.sDescripcion as Titulo, ' +
                                'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                'SUM(bp.dCantidad) AS Total,bp.dIdFecha, a.sAnexo ' +
                                'FROM bitacoradepersonal bp ' +
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal) ' +
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="Si") ' +
                                'WHERE bp.scontrato = :Orden AND ((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) and bp.dIdFecha between :Inicio and :termino ' +
                                'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'

                              {, 'SELECT bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,SUM(bp.dAjuste) AS Ajuste,SUM(bp.dCantidad) AS Total,bp.dIdFecha, '+
                                'a.sAnexo FROM bitacoradepersonal bp inner join personal p on(p.sContrato=:ContratoBarco and '+
                                'p.sIdPersonal=bp.sIdPersonal) inner join anexos a on(a.sAnexo=p.sAnexo and '+
                                'a.sTipo="PERSONAL" and a.sTierra="Si") WHERE bp.scontrato = :Orden AND '+
                                '((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) ' +
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'GROUP BY bp.sIdPersonal,bp.didfecha order by p.iItemOrden,bp.sIdPersonal,bp.didfecha'}

                              , 'select be.sIdEquipo as sIdRecurso,be.sDescripcion as Descripcion,e.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="EQUIPO" limit 1),"No existe") as Titulo,'+
                                'ifnull((select sAnexo from anexos where sTipo="EQUIPO" limit 1),"N/E") as sAnexo,'+
                                '(sum(be.dCantHH) + ' +
                                ' IFNULL((SELECT ' +
                               '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                               '          FROM ' +
                               '            bitacoradeajustes ba ' +
                               '          WHERE ' +
                               '            ba.sContrato = be.sContrato AND ba.sNumeroOrden = be.sNumeroOrden AND ' +
                               '            ba.sNumeroActividad = be.sNumeroActividad AND ' +
                               '            ba.sIdPartidaAnexo = be.sIdEquipo AND ba.dFecha = be.dIdFecha), 0)) as total,be.didfecha from bitacoradeequipos be inner join equipos e on '+
                                '(e.sContrato=:contratoBarco and e.sIdEquipo=be.sIdEquipo) inner join bitacoradeactividades '+
                                'ba on(ba.sContrato = be.sContrato and ba.dIdFecha = be.didfecha and ba.sNumeroOrden = '+
                                'be.sNumeroOrden and ba.iIdDiario = be.iIdDiario and ba.iIdTarea = be.iIdTarea and '+
                                'ba.iIdActividad=be.iIdActividad) where  be.sContrato=:Orden and ((:folio=-1) or (:folio<>-1 and '+
                                'be.sNumeroOrden=:folio)) and be.dIdFecha between :Inicio and :termino ' +
                                'group by e.sIdEquipo,be.didfecha order by e.iItemOrden,e.sIdEquipo,be.didfecha'

                              , 'select c.sIdCuenta,c.sIdPernocta as sIdRecurso,c.sDescripcion as Descripcion,c.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="PERNOCTA" limit 1),"No existe") as Titulo,'+
                                'ifnull((select sAnexo from anexos where sTipo="PERNOCTA" limit 1),"N/E") as sAnexo,'+
                                '((sum(bp.dCantHH) + ' +
                                ' IFNULL((SELECT ' +
                                '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '          FROM ' +
                                '            bitacoradeajustes ba ' +
                                '          WHERE ' +
                                '            ba.sContrato = bpr.sContrato AND ba.sNumeroOrden = bpr.sNumeroOrden AND ' +
                                '            ba.sNumeroActividad = bpr.sNumeroActividad AND ' +
                                '            ba.sIdPartidaAnexo = bpr.sIdCuenta AND ba.dFecha = bpr.dIdFecha), 0)) - ifnull((select ifnull(sum(bpr.dCantidad),0) from bitacoradepernocta bpr '+
                                'where bpr.sContrato = ba.sContrato and ba.didfecha=bpr.didfecha and (bpr.sNumeroOrden <> "@" '+
                                'and bpr.sNumeroOrden=ba.snumeroorden) group by bpr.sContrato),0)) as Total,bp.didfecha from cuentas c '+
                                'left join bitacoradeactividades ba on(ba.sContrato=:Orden and ((:Folio=-1) or (:Folio<>-1 and ba.sNumeroOrden=:Folio)) '+
                                ') left join bitacoradepersonal bp on (bp.sContrato =ba.sContrato and '+
                                'ba.sNumeroOrden=bp.sNumeroOrden and  bp.dIdFecha = ba.didfecha and ba.iIdDiario=bp.iIdDiario and '+
                                'ba.iIdTarea=bp.iIdTarea and ba.iIdActividad=bp.iIdActividad and bp.sTipoPernocta = c.sIdPernocta '+
                                ') left join personal p on (p.sContrato = :contratoBarco '+
                                'and p.sIdPersonal = bp.sIdPersonal) left join tiposdepersonal tp on ( p.sIdTipoPersonal = '+
                                'tp.sIdTipoPersonal) left join ordenesdetrabajo ot '+
                                'on (ot.sContrato = :Orden and ot.sNumeroOrden = ba.sNumeroOrden) left join pernoctan pr '+
                                'on (pr.sIdPernocta = ot.sIdPernocta) left join plataformas pl on(pl.sIdPlataforma=ot.sIdPlataforma) '+
                                'where ((:Folio=-1) or (:Folio<>-1 and bp.snumeroorden=:Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'group by c.sIdPernocta,bp.didfecha order by c.sIdPernocta,bp.didfecha '

                              );


var
  iGrupo:Integer;
  QrConsulta,QrFolios:TzReadOnlyQuery;
  sIdRecurso:String;
  ValTmp:Variant;
  CantTmp,dTotal:Double;
  iPos,TotalDias:Integer;
  sDia:String;
  dDia,dPosTmp:TDateTime;
  sTipo:String;
begin
  QrFolios:=TzREadOnlyQUery.Create(nil);
  QrConsulta:=TzREadOnlyQUery.Create(nil);
  try
    try
      QrConsulta.Connection:=Connection.zConnection;
      QrFolios.Connection:=Connection.zConnection;
      if MryRecursos.Active then
      begin
        MryRecursos.EmptyDataSet;  //.EmptyTable;
      end;

      TotalDias:=DaysBetween(ParamFechaT,ParamFechaI) + 1;
      dDia:= ParamFechaI;
      with MryRecursos do
      begin
        Active:=False;
        MryRecursos.FieldDefs.Clear;
        FieldDefs.Add('iGrupo', ftInteger, 0, True);
        FieldDefs.Add('sTipo', ftString, 20, True);
        FieldDefs.Add('sFolio', ftString, 100, True);
        FieldDefs.Add('sDescFolio', ftString, 500, True);
        FieldDefs.Add('sidrecurso', ftString, 100, True);
        FieldDefs.Add('sdescripcion', ftString, 500, True);
        FieldDefs.Add('sAnexo', ftString, 50, false);
        FieldDefs.Add('sTitulo', ftString, 500, false);
        FieldDefs.Add('smedida', ftString, 100, True);

        for IPos := 1 to TotalDias do
        begin
          sDia:=FormatDateTime('ddddd',ddia);
          FieldDefs.Add(sDia, FtFloat, 0, True);
          FieldDefs.Add('Limit'+sDia, FtFloat, 0, True);
          dDia:=IncDay(dDia);
        end;
        FieldDefs.Add('dTotal', FtFloat, 0, True);
        CreateDataSet;

        Open;
        FieldList.Update;
      end;


      QrFolios.Active:=False;
      QrFolios.SQL.Text:= 'select ba.*,ot.mdescripcion as descFolio from bitacoradeactividades ba '+
                          'inner join ordenesdetrabajo ot '+
                          'on(ot.sContrato=ba.sContrato and ot.sNumeroOrden=ba.sNumeroorden) '+
                          'where ba.sContrato=:Contrato and ba.dIdFecha ' +
                          'between :Inicio and :termino';

      if IsGeneral then
        QrFolios.SQL.Add('group by ba.sContrato')
      else
        QrFolios.SQL.Add('group by ba.sContrato,ba.sNumeroOrden');

      QrFolios.ParamByName('Contrato').AsString:=ParamContrato;
      QrFolios.ParamByName('Inicio').AsDate:=ParamFechaI;
      QrFolios.ParamByName('Termino').AsDate:=ParamFechaT;
      QrFolios.Open;

      while not QrFolios.Eof do
      begin
        sTipo:='BARCO';
        QrConsulta.Active:=False;              //Descripcion,e.sMedida as medida
        QrConsulta.SQL.Text:='SELECT tm.sContrato,tm.sIdTipoMovimiento,tm.sIdTipoMovimiento as sIdRecurso,tm.sDescripcion as Descripcion,' + #10 +
                              'tm.sTipo,tm.smedida as medida,me.didfecha,mf.sFolio,ifnull(SUM( me.sFactor ),0) AS Total,' + #10 +
                              'ifnull((select sdescripcion from anexos where sTipo="BARCO" limit 1)'+
                              ',"No existe") as Titulo, ' +
                               'ifnull((select sAnexo from anexos where sTipo="BARCO" limit 1)'+
                              ',"No existe") as sAnexo ' +
                              'FROM tiposdemovimiento AS tm '  +
                              'inner join movimientosdeembarcacion AS me '   +
                              'on(me.sContrato = tm.sContrato AND me.sClasificacion = tm.sIdTipoMovimiento) ' + #10 +
                              'inner join movimientosxfolios mf on (me.sContrato = mf.sContrato and me.sOrden=mf.sNumeroOrden' + #10 +
                              'and me.didfecha=mf.dIdFecha and me.iIdDiario = mf.iIdDiario) ' +
                              'WHERE tm.sContrato =:ContratoBarco AND me.sOrden =:Orden' + #10 +
                              'and ((:Folio=-1) or (:Folio<>-1 and mf.sFolio=:Folio)) ' + #10 +
                              'AND tm.sClasificacion = "Movimiento de Barco" and' + #10 +
                              'tm.lGenera ="Si" and me.didfecha between :Inicio and :termino ' +
                              'group by me.sContrato,tm.sIdTipoMovimiento,me.didfecha '+
                              'order by me.sContrato,tm.sIdTipoMovimiento,me.didfecha';

        {if IsGeneral then
          QrConsulta.SQL.Add('group by me.sContrato,tm.sIdTipoMovimiento,me.didfecha '+
                             'order by me.sContrato,tm.sIdTipoMovimiento,me.didfecha')
        else
          QrConsulta.SQL.Add('group by me.sContrato,tm.sIdTipoMovimiento,me.didfecha '+
                             'order by me.sContrato,tm.sIdTipoMovimiento,me.didfecha');}

        QrConsulta.ParamByName('ContratoBarco').AsString:=Global_Contrato_Barco;
        QrConsulta.ParamByName('Orden').AsString:=ParamContrato;

        if IsGeneral then
          QrConsulta.ParamByName('Folio').AsString:='-1'
        else
          QrConsulta.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
        QrConsulta.ParamByName('Inicio').AsDate:=ParamFechaI;
        QrConsulta.ParamByName('Termino').AsDate:=ParamFechaT;
        QrConsulta.Open;

        sIdRecurso:='';

        while not QrConsulta.Eof do
        begin
          if sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString then
          begin
            if MryRecursos.State=dsInsert then
            begin
              if dTotal<>xRound(ValTmp,Mantisa) then
              begin
                sDia:=FormatDateTime('ddddd',dPosTmp);
                MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
                  + (xRound(ValTmp,Mantisa)-dTotal);
                MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;
                MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
              end
              else
                MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
              MryRecursos.Post;
            end;

            ValTmp:=0;
            CantTmp:=0;
            dTotal:=0;

            sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.Append;
            MryRecursos.FieldByName('iGrupo').AsInteger:=QrFolios.RecNo;
            MryRecursos.FieldByName('sTipo').AsString:= sTipo;
            if IsGeneral then
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sContrato').AsString
            else
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
            MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
            MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
            MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
            MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
            MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
            MryRecursos.FieldByName('dTotal').AsFloat:=0;
            dDia:= ParamFechaI;
            for iPos := 1 to TotalDias do
            begin
              sDia:=FormatDateTime('ddddd',ddia);
              MryRecursos.FieldValues[sDia] := 0.000000 ;
              MryRecursos.FieldValues['Limit'+sDia] := 0.000000 ;
              dDia:=IncDay(dDia);
            end;

          end;


          //sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;

          sDia:=FormatDateTime('ddddd',QrConsulta.FieldByName('dIdFecha').AsDateTime);
          MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
          MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;
          ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
          dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;
          if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
          begin
            CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
            dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
          end;

          QrConsulta.Next;
        end;

        if MryRecursos.State=dsInsert then
        begin
          if dTotal<>xRound(ValTmp,Mantisa) then
          begin
            sDia:=FormatDateTime('ddddd',dPosTmp);
            MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
              + (xRound(ValTmp,Mantisa)-dTotal);

            MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;
            MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
          end
          else
            MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
          MryRecursos.Post;
        end;
        sTipo:='MATERIAL';
        QrConsulta.Active:=fALSE;
        QrConsulta.SQL.text:= 'select s.sNumeroOrden, b.sNumeroActividad, b.sTrazabilidad,i.sIdInsumo as sIdRecurso,'+
                              'i.mDescripcion as descripcion,i.sMedida as medida, sum(b.dCantidad) as total,i.sIdProveedor,'+
                              'b.dFechaSalida as dIdFecha,ifnull((select sdescripcion from anexos where sTipo="MATERIAL" '+
                              'and sTierra="No" order by iorden limit 1),"No existe") as Titulo, ' +
                              'ifnull((select sAnexo from anexos where sTipo="MATERIAL" and sTierra="No" '+
                              'order by iorden limit 1),"No existe") as sAnexo ' +
                              'from almacen_salida s '+
                              'inner join bitacoradesalida b '+
                              'on(b.sContrato = s.sContrato and b.iFolioSalida = s.iFolioSalida and b.sext = s.sext) '+
                              'inner join insumos i '+
                              'on (i.sContrato =:Contrato and i.sIdInsumo = b.sIdInsumo and i.sTrazabilidad = b.sTrazabilidad ) '+
                              //'INNER JOIN actividadesxorden ao ON ( ' +
                              //'ao.sContrato = b.sContrato and ao.sIdConvenio =:Convenio and ao.sNumeroOrden = b.sNumeroOrden and ' +
                              //'ao.sNumeroActividad = b.sNumeroActividad and ao.sTipoActividad = "Actividad" ) ' +
                              'where s.sContrato =:orden and ((:Folio=-1) or (:Folio<>-1 and s.sNumeroOrden=:Folio)) '+
                              'and b.dFechaSalida between :Inicio and :termino '+
                              'group by b.sContrato,b.sIdInsumo,didfecha Order by b.sNumeroOrden,b.sIdInsumo,didfecha';
        QrConsulta.ParamByName('Orden').AsString    := ParamContrato;
        if IsGeneral then
          QrConsulta.ParamByName('Folio').AsString    := '-1'
        else
          QrConsulta.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
        QrConsulta.ParamByName('Inicio').AsDate:=ParamFechaI;
        QrConsulta.ParamByName('Termino').AsDate:=ParamFechaT;
        QrConsulta.ParamByName('Contrato').AsString := global_Contrato_Barco;
        QrConsulta.Open;

        sIdRecurso:='';

        while not QrConsulta.Eof do
        begin
          if sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString then
          begin
            if MryRecursos.State=dsInsert then
            begin
              if dTotal<>xRound(ValTmp,Mantisa) then
              begin
                sDia:=FormatDateTime('ddddd',dPosTmp);
                MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
                  + (xRound(ValTmp,Mantisa)-dTotal);
                MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

                MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
              end
              else
                MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
              MryRecursos.Post;
            end;





            ValTmp:=0;
            CantTmp:=0;
            dTotal:=0;

            sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.Append;
            MryRecursos.FieldByName('iGrupo').AsInteger:=QrFolios.RecNo;
            MryRecursos.FieldByName('sTipo').AsString:= sTipo;
            if IsGeneral then
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sContrato').AsString
            else
              MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
            MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
            MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
            MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
            MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
            MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
            MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
            MryRecursos.FieldByName('dTotal').AsFloat:=0;
            dDia:= ParamFechaI;
            for iPos := 1 to TotalDias do
            begin
              sDia:=FormatDateTime('ddddd',ddia);
              MryRecursos.FieldValues[sDia] := 0.000000 ;
              MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

              dDia:=IncDay(dDia);
            end;
          end;


          //sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;
          sDia:=FormatDateTime('ddddd',QrConsulta.FieldByName('dIdFecha').AsDateTime);
          MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
          MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

          ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
          dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;
          if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
          begin
            CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
            dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
          end;

          QrConsulta.Next;
        end;

        if MryRecursos.State=dsInsert then
        begin
          if dTotal<>xRound(ValTmp,Mantisa) then
          begin
            sDia:=FormatDateTime('ddddd',dPosTmp);
            MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
              + (xRound(ValTmp,Mantisa)-dTotal);
            MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

            MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
          end
          else
            MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
          MryRecursos.Post;
        end;

        for iGrupo := 1 to length(Asql) do
        begin
          if iGrupo<3 then
            sTipo:='PERSONAL';
          if iGrupo=3 then
            sTipo:='EQUIPO';
          if iGrupo=4 then
            sTipo:='PERNOCTA';



          QrConsulta.Active:=false;
          QrConsulta.SQL.Text:=ASql[iGrupo];
          if IsGeneral then
            QrConsulta.ParamByName('Folio').AsString    := '-1'
          else
            QrConsulta.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
          QrConsulta.ParamByName('Inicio').AsDate:=ParamFechaI;
          QrConsulta.ParamByName('Termino').AsDate:=ParamFechaT;
          QrConsulta.ParamByName('ContratoBarco').AsString:=Global_Contrato_Barco;
          QrConsulta.ParamByName('Orden').AsString:=ParamContrato;
          QrConsulta.Open;

          sIdRecurso:='';

          while not QrConsulta.Eof do
          begin
            if sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString then
            begin
              if MryRecursos.State=dsInsert then
              begin
                if dTotal<>xRound(ValTmp,Mantisa) then
                begin
                  sDia:=FormatDateTime('ddddd',dPosTmp);
                  MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
                    + (xRound(ValTmp,Mantisa)-dTotal);
                  MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

                  MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
                end
                else
                  MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
                MryRecursos.Post;
              end;

              ValTmp:=0;
              CantTmp:=0;
              dTotal:=0;

              sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
              MryRecursos.Append;
              MryRecursos.FieldByName('iGrupo').AsInteger:=QrFolios.RecNo;
              MryRecursos.FieldByName('sTipo').AsString:= sTipo;
              if IsGeneral then
                MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sContrato').AsString
              else
                MryRecursos.FieldByName('sFolio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
              MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
              MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
              MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
              MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
              MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
              MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
              MryRecursos.FieldByName('dTotal').AsFloat:=0;
              dDia:= ParamFechaI;
              for iPos := 1 to TotalDias do
              begin
                sDia:=FormatDateTime('ddddd',ddia);
                MryRecursos.FieldValues[sDia] := 0.000000 ;
                MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

                dDia:=IncDay(dDia);
              end;
            end;

            //sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;
            sDia:=FormatDateTime('ddddd',QrConsulta.FieldByName('dIdFecha').AsDateTime);
            if iGrupo<3 then
            begin
              MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat;

              ValTmp:=ValTmp + (QrConsulta.FieldByName('total').Value + + QrConsulta.FieldByName('Ajuste').AsFloat);
              if CantTmp<(xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat) then
              begin
                CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat;
                dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
              end;
            end
            else
              if (iGrupo=4) and (QrConsulta.FieldByName('sIdcuenta').AsString='1') then
              begin
                connection.QryBusca.Active:=false;
                connection.QryBusca.SQL.Text:='select sum(bpc.dCantidad) as dCantidad' + #10 +
                                              'from bitacoradepersonal_cuadre bpc' + #10 +
                                              'inner join personal p' + #10 +
                                              'on ( p.sContrato = :ContratoBarco and bpc.sIdPersonal = p.sIdPersonal )' + #10 +
                                              'inner join tiposdepersonal tp' + #10 +
                                              'on (tp.sIdTipoPersonal = p.sIdTipoPersonal )' + #10 +
                                              'where bpc.sContrato =:Orden and bpc.snumeroorden=:Folio' + #10 +
                                              'and tp.lPernocta="Si" and bpc.dIdFecha=:Fecha' ;

                connection.QryBusca.ParamByName('ContratoBarco').AsString:= global_contrato_barco;
                connection.QryBusca.ParamByName('Orden').AsString:=ParamContrato;
                connection.QryBusca.ParamByName('Folio').AsString:=QrFolios.FieldByName('sNumeroOrden').AsString;
                connection.QryBusca.ParamByName('Fecha').AsDate:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
                connection.QryBusca.Open;
                if connection.QryBusca.RecordCount=1 then
                begin
                  MryRecursos.FieldByName(sDia).AsFloat:=xRound(QrConsulta.FieldByName('total').AsFloat,Mantisa)+connection.QryBusca.FieldByName('dCantidad').AsFloat;
                  ValTmp:=ValTmp + (QrConsulta.FieldByName('total').Value + +connection.QryBusca.FieldByName('dCantidad').AsFloat);
                  if CantTmp<(xRound(QrConsulta.FieldByName('total').Value,Mantisa)++connection.QryBusca.FieldByName('dCantidad').AsFloat) then
                  begin
                    CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + +connection.QryBusca.FieldByName('dCantidad').AsFloat;
                    dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
                  end;
                end
                else
                begin
                  MryRecursos.FieldByName(sDia).AsFloat:=xRound(QrConsulta.FieldByName('total').AsFloat,Mantisa);
                  ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
                  if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
                  begin
                    CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
                    dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
                  end;


                end;
              end
              else
              begin
                MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
                ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
                if CantTmp<xRound(QrConsulta.FieldByName('total').Value,Mantisa) then
                begin
                  CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa);
                  dPosTmp:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
                end;
              end;
            MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

            dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;

            QrConsulta.Next;
          end;

          {
            dDia:= ParamFechaI;
            for IPos := 1 to TotalDias do
        begin
          sDia:=FormatDateTime('ddddd',ddia);
          FieldDefs.Add(sDia, FtFloat, 0, True);
          dDia:=IncDay(dDia);
        end;
        sDia:=FormatDateTime('ddddd',dPosTmp);
        }
          if MryRecursos.State=dsInsert then
          begin
            if dTotal<>xRound(ValTmp,Mantisa) then
            begin
              sDia:=FormatDateTime('ddddd',dPosTmp);
              MryRecursos.FieldByName(sDia).AsFloat := MryRecursos.FieldByName(sDia).AsFloat +
                + (xRound(ValTmp,Mantisa)-dTotal);
              MryRecursos.FieldByName('Limit'+sDia).AsFloat :=MryRecursos.FieldByName(sDia).AsFloat;

              MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,Mantisa);
            end
            else
              MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
            MryRecursos.Post;
          end;


        end;
        QrFolios.Next;
      end;

    except
      on e:Exception do
        Showmessage(e.Message + ', ' + e.ClassName);
    end;

  finally
    QrConsulta.Destroy;
  end;
end;

Procedure ExportOTProject(ParamContrato,ParamConvenio,ParamFolio:String;PgbAvance: TcxProgressBar=nil);
var
  MsProject:Variant;
  Error:Boolean;
  QrActividades,QrAvances,
  QrFechas:TzReadOnlyQuery;
  ListaParams:TStringList;
  Ren,ItotalReg : integer;
  Task:Variant;
  iCol:Integer;
begin
  try
    Error:=False;
    try
      MsProject:=GetActiveOLEObject ('MSProject.Application');
    except
      MsProject:=CreateOleObject ('MSProject.Application');
    end;
  except
    Error:=true;
  end;

  if NOt error then
  begin
    QrActividades:=TzReadOnlyQuery.Create(nil);
    QrAvances:=TzReadOnlyQuery.Create(nil);
    QrFechas:=TzReadOnlyQuery.Create(nil);
    ListaParams:=TStringList.Create;
    try
      try
        QrFechas.Connection:=Connection.zConnection;

        MsProject.visible:=false;
        MsProject.DisplayAlerts:=false;
        MsProject.projects.add(False);
        MsProject.ActiveProject.Title:=ParamFolio;
        QrActividades.Connection:=Connection.zConnection;
        QrActividades.SQL.Text:='Select * from actividadesxorden where scontrato=:Contrato ' +
                                'and sIdConvenio=:Convenio and sNumeroOrden=:Folio '+
                                'and inivel=(select min(inivel) from actividadesxorden '+
                                'where sContrato =:Contrato and sIdConvenio =:Convenio ' +
                                'and sNumeroOrden =:Folio) order by iItemOrden';

        QrAvances.Connection:=Connection.zConnection;
        QrAvances.SQL.Text:='SELECT  SUM(dAvance) FROM bitacoradeactividades '+
                            'WHERE sContrato =:Contrato AND sNumeroOrden =:Folio ' +
                            'AND sWbs=:Wbs and sNumeroActividad=:Actividad ' +
                            'AND sIdTipoMovimiento = "E" ';

        QrActividades.ParamByName('Contrato').AsString:=ParamContrato;
        QrActividades.ParamByName('Convenio').AsString:=ParamConvenio;
        QrActividades.ParamByName('Folio').AsString:=ParamFolio;
        QrActividades.Open;

        MsProject.SelectTaskColumn('Nombre');
//        MsProject.Font32Ex('Arial','8',False);
        MsProject.TableEditEx('&Entrada',true,emptyParam,EmptyParam,emptyparam,emptyparam,'Texto1',
        emptyparam,emptyparam,emptyparam,true,true,255,1,2,emptyParam, emptyParam, emptyParam, emptyparam, emptyParam);
        //emptyparam,emptyparam,emptyparam,true,true,255,1,2,emptyParam, emptyParam, emptyParam, emptyParam, emptyParam);

        MsProject.TableApply('&Entrada');

        //MsProject.SelectTaskColumn('Nombre');
        //MsProject.Font32Ex('Arial','8',False);

      MsProject.TableEditEx('&Entrada',True,emptyParam,EmptyParam,emptyparam,'Texto1','Texto1',
        'Act.',emptyParam,emptyParam,emptyParam,emptyParam,emptyParam,emptyParam,emptyParam,2,
        emptyParam,emptyParam,emptyParam,emptyParam);
        MsProject.TableApply('&Entrada');

        //electTaskColumn Column:="Texto1"
    //Font32Ex Name:="Arial", Size:="8", Bold:=False

        MsProject.SelectTaskColumn('Duración');
        MsProject.ColumnDelete;
        MsProject.SelectTaskColumn('Predecesoras');
        MsProject.ColumnDelete;
        MsProject.SelectTaskColumn('Indicadores');
        MsProject.ColumnDelete;
        {MsProject.SelectTaskColumn('Modo de Tarea');
        MsProject.ColumnDelete;
         }
        MsProject.SelectTaskColumn('Nombres de los recursos');
        MsProject.ColumnDelete;

        iCol:=5;
        MsProject.TableEditEx('&Entrada',true,emptyParam,EmptyParam,emptyparam,emptyparam,'Duración',
        emptyparam,emptyparam,emptyparam,true,true,255,1,iCol,emptyParam, emptyParam, emptyParam, emptyParam, emptyParam);
        MsProject.TableApply('&Entrada');

        inc(iCol);
        MsProject.TableEditEx('&Entrada',true,emptyParam,EmptyParam,emptyparam,emptyparam,'Comienzo real',
        emptyparam,emptyparam,emptyparam,true,true,255,1,iCol,emptyParam, emptyParam, emptyParam, emptyParam, emptyParam);
        MsProject.TableApply('&Entrada');

        inc(iCol);
        MsProject.TableEditEx('&Entrada',true,emptyParam,EmptyParam,emptyparam,emptyparam,'Fin real',
        emptyparam,emptyparam,emptyparam,true,true,255,1,iCol,emptyParam, emptyParam, emptyParam, emptyParam, emptyParam);
        MsProject.TableApply('&Entrada');

        inc(iCol);
        MsProject.TableEditEx('&Entrada',true,emptyParam,EmptyParam,emptyparam,emptyparam,'Duración real',
        emptyparam,emptyparam,emptyparam,true,true,255,1,iCol,emptyParam, emptyParam, emptyParam, emptyParam, emptyParam);
        MsProject.TableApply('&Entrada');

        iCol:=2;
        MsProject.TableEditEx('&Entrada',True,emptyParam,EmptyParam,emptyparam,'Nombre','Nombre',
        emptyParam,50,emptyParam,emptyParam,emptyParam,emptyParam,emptyParam,iCol,emptyParam,
        emptyParam,emptyParam,true,emptyParam);
        //TableEditEx(Name, TaskTable, Create, OverwriteExisting, NewName, FieldName, NewFieldName,
      //Title, Width, Align, ShowInMenu, LockFirstColumn, DateFormat, RowHeight, ColumnPosition, AlignTitle, HeaderAutoRowHeightAdjustment, HeaderTextWrap, WrapText, ShowAddNewColumn)

        MsProject.TableApply('&Entrada');

        inc(iCol);

        MsProject.TableEditEx('&Entrada',True,emptyParam,EmptyParam,emptyparam,'Comienzo','Comienzo',
        'Comienzo Previsto',emptyParam,emptyParam,emptyParam,emptyParam,emptyParam,emptyParam,iCol,emptyParam,
        emptyParam,emptyParam,emptyParam,emptyParam);
        MsProject.TableApply('&Entrada');

         inc(iCol);
        MsProject.TableEditEx('&Entrada',True,emptyParam,EmptyParam,emptyparam,'Fin','Fin',
        'Fin Previsto',10,0,true,true,255,1,iCol,0,
        emptyParam,emptyParam,emptyParam,emptyParam);
        MsProject.TableApply('&Entrada');

         inc(iCol);
        MsProject.TableEditEx('&Entrada',True,emptyParam,EmptyParam,emptyparam,'Duración','Duración',
        'Duración Prevista',10,0,true,true,255,1,iCol,0,
        emptyParam,emptyParam,emptyParam,emptyParam);
        MsProject.TableApply('&Entrada');

      //  TableEditEx Name:="&Entrada", TaskTable:=True, NewName:="", NewFieldName:="Duración", Width:=20, ShowInMenu:=True, LockFirstColumn:=True, DateFormat:=255, RowHeight:=1, ColumnPosition:=6


      //TableEditEx(Name, TaskTable, Create, OverwriteExisting, NewName, FieldName, NewFieldName,
      //Title, Width, Align, ShowInMenu, LockFirstColumn, DateFormat, RowHeight, ColumnPosition, AlignTitle, HeaderAutoRowHeightAdjustment, HeaderTextWrap, WrapText, ShowAddNewColumn)

        while not QrActividades.Eof do
        begin
          MsProject.ActiveProject.Tasks.add(QrActividades.FieldByName('mDescripcion').AsString);
          //showmessage(IntTostr(MsProject.ActiveProject.Tasks.Count));
          Task:=MsProject.ActiveProject.Tasks.item[MsProject.ActiveProject.Tasks.Count];
          Task.text1:=QrActividades.FieldByName('sNumeroActividad').AsString;
          Task.start:=QrActividades.FieldByName('dFechaInicio').AsDateTime;
          Task.Finish:=QrActividades.FieldByName('dFechaFinal').AsDateTime;

          if QrActividades.FieldByName('sTipoActividad').AsString='Paquete' then
          begin


          end
          else
          begin
            QrFechas.Active:=False;
            QrFechas.SQL.Text:= 'select cast((concat(min(didfecha)," ",min(sHoraInicio))) as datetime) as fecha '+
                                'from bitacoradeactividades where sContrato=:Contrato and '+
                                'sNumeroOrden=:Folio and swbs=:wbs and sNumeroActividad=:Actividad and '+
                                'sIdTipoMovimiento="ED" ';

            QrFechas.ParamByName('Contrato').AsString:=QrActividades.FieldByName('sContrato').AsString;
            QrFechas.ParamByName('Folio').AsString:=QrActividades.FieldByName('sNumeroOrden').AsString;
            QrFechas.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            QrFechas.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            QrFechas.Open;

            if QrFechas.RecordCount>0 then
              if not QrFechas.FieldByName('fecha').IsNull then
                Task.Actualstart:=QrFechas.FieldByName('fecha').AsDateTime;

            QrFechas.Active:=False;
            QrFechas.SQL.Text:= 'select cast((concat(Max(didfecha)," ",max(sHoraFinal))) as datetime) as fecha '+
                                'from bitacoradeactividades where sContrato=:Contrato and '+
                                'sNumeroOrden=:Folio and swbs=:wbs and sNumeroActividad=:Actividad and '+
                                'sIdTipoMovimiento="ED" ';

            QrFechas.ParamByName('Contrato').AsString:=QrActividades.FieldByName('sContrato').AsString;
            QrFechas.ParamByName('Folio').AsString:=QrActividades.FieldByName('sNumeroOrden').AsString;
            QrFechas.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            QrFechas.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            QrFechas.Open;
            if QrFechas.RecordCount>0 then
              if not QrFechas.FieldByName('fecha').IsNull then
                Task.ActualFinish:=QrFechas.FieldByName('fecha').value;




            QrAvances.Active:=False;
            QrAvances.ParamByName('Contrato').AsString:=QrActividades.FieldByName('sContrato').AsString;
            QrAvances.ParamByName('Folio').AsString:=QrActividades.FieldByName('sNumeroOrden').AsString;
            QrAvances.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            QrAvances.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            QrAvances.Open;


          end;

          with ListaParams do
          begin
            Clear;
            Add('Contrato='+QrActividades.FieldByName('sContrato').AsString);
            Add('Convenio='+ParamConvenio);
            Add('Folio='+QrActividades.FieldByName('sNumeroOrden').AsString);
            Add('Wbs='+QrActividades.FieldByName('swbs').AsString);
          end;

          //SelectTaskField Row:=1, Column:="Modo de tarea"
          //SetTaskField Field:="Modo de tarea", Value:="No"
          //MsProject.SelectTaskField(1,'Modo de tarea');
          //MsProject.SetTaskField('Modo de tarea','No');
          MsProject.SelectTaskField(MsProject.ActiveProject.Tasks.Count,'Modo de tarea',false);
          MsProject.SetTaskField('Modo de tarea','No',MsProject.ActiveProject.Tasks.Count,emptyparam);
          VolumenActividades(MsProject,ListaParams,Ren,ItotalReg,PgbAvance);
          QrActividades.Next;
        end;
      except
        on e:Exception do
          showmessage(e.Message + ', ' + e.ClassName);
      end;
    finally
      MsProject.visible:=true;
      MsProject.DisplayAlerts:=true;

      if PgbAvance<>nil then
      begin
        PgbAvance.Visible:=False;
        Application.ProcessMessages;
      end;
      QrActividades.Destroy;
      QrAvances.Destroy;
      ListaParams.Destroy;
      QrFechas.Destroy;
    end;




  end;
///Exportando
end;

function VolumenActividades(var MsProject:Variant;ListaParams:TstringList;var Ren:Integer;const ItotalReg:Integer;PgbAvance: TcxProgressBar=nil):Boolean;
var
  QrActividades,QrAvances:TzReadOnlyQuery;
  NewListaParams:TstringList;
  Task:variant;
  QrFechas:TzReadOnlyQuery;
begin
  QrActividades:=TzReadOnlyQuery.Create(nil);
  QrAvances:=TzReadOnlyQuery.Create(nil);
  QrFechas:=TzReadOnlyQuery.Create(nil);
  NewListaParams:=TstringList.Create;
  try
    try
      QrFechas.Connection:=connection.zConnection;
      QrAvances.Connection:=Connection.zConnection;
      QrAvances.SQL.Text:='SELECT  SUM(dAvance) FROM bitacoradeactividades '+
                          'WHERE sContrato =:Contrato AND sNumeroOrden =:Folio ' +
                          'AND sWbs=:Wbs and sNumeroActividad=:Actividad ' +
                          'AND sIdTipoMovimiento = "E" ';



      QrActividades.Connection:=Connection.zConnection;
      QrActividades.SQL.Text:='select * from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio ' +
                              'and sNumeroOrden =:Folio and swbsanterior=:wbsAnt order by iItemOrden';
      QrActividades.Params.ParamByName('Contrato').Value    := ListaParams.Values['Contrato'];
      QrActividades.Params.ParamByName('Convenio').Value    := ListaParams.Values['Convenio'];
      QrActividades.Params.ParamByName('Folio').Value       := ListaParams.Values['Folio'];
      QrActividades.ParamByName('wbsAnt').AsString:=ListaParams.Values['Wbs'];
      QrActividades.Open ;



      while not QrActividades.Eof do
      begin
        MsProject.ActiveProject.Tasks.add(QrActividades.FieldByName('mDescripcion').AsString);
        if QrActividades.RecNo=1 then
        begin
          MsProject.SelectTaskField(1,'Nombre');
          MsProject.OutlineIndent;
        end;
        //Task:=MsProject.ActiveProject.Tasks.item[MsProject.ActiveProject.Tasks.Count-1];
        Task:=MsProject.ActiveProject.Tasks.item[MsProject.ActiveProject.Tasks.Count];
//        showmessage(IntTostr(MsProject.ActiveProject.Tasks.Count));
        Task.text1:=QrActividades.FieldByName('sNumeroActividad').AsString;
        Task.start:=QrActividades.FieldByName('dFechaInicio').AsDateTime;
        Task.Finish:=QrActividades.FieldByName('dFechaFinal').AsDateTime;
        //MsProject.SelectTaskField Row:=0, Column:="Modo de tarea"
        MsProject.SelectTaskField(MsProject.ActiveProject.Tasks.Count,'Modo de tarea',false);
        MsProject.SetTaskField('Modo de tarea','No',MsProject.ActiveProject.Tasks.Count,emptyparam);
        if QrActividades.FieldByName('sTipoActividad').AsString='Paquete' then
        begin


        end
        else
        begin
          QrFechas.Active:=False;
          QrFechas.SQL.Text:= 'select cast((concat(min(didfecha)," ",min(sHoraInicio))) as datetime) as fecha '+
                              'from bitacoradeactividades where sContrato=:Contrato and '+
                              'sNumeroOrden=:Folio and swbs=:wbs and sNumeroActividad=:Actividad and '+
                              'sIdTipoMovimiento="ED" ';

          QrFechas.ParamByName('Contrato').AsString:=QrActividades.FieldByName('sContrato').AsString;
          QrFechas.ParamByName('Folio').AsString:=QrActividades.FieldByName('sNumeroOrden').AsString;
          QrFechas.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QrFechas.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QrFechas.Open;

          if QrFechas.RecordCount>0 then
            if not QrFechas.FieldByName('fecha').IsNull then
              Task.Actualstart:=QrFechas.FieldByName('fecha').AsDateTime;

          QrFechas.Active:=False;
          QrFechas.SQL.Text:= 'select cast((concat(Max(didfecha)," ",max(sHoraFinal))) as datetime) as fecha '+
                              'from bitacoradeactividades where sContrato=:Contrato and '+
                              'sNumeroOrden=:Folio and swbs=:wbs and sNumeroActividad=:Actividad and '+
                              'sIdTipoMovimiento="ED" ';

          QrFechas.ParamByName('Contrato').AsString:=QrActividades.FieldByName('sContrato').AsString;
          QrFechas.ParamByName('Folio').AsString:=QrActividades.FieldByName('sNumeroOrden').AsString;
          QrFechas.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QrFechas.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QrFechas.Open;
          if QrFechas.RecordCount>0 then
            if not QrFechas.FieldByName('fecha').IsNull then
              Task.ActualFinish:=QrFechas.FieldByName('fecha').value;

          QrAvances.Active:=False;
          QrAvances.ParamByName('Contrato').AsString:=QrActividades.FieldByName('sContrato').AsString;
          QrAvances.ParamByName('Folio').AsString:=QrActividades.FieldByName('sNumeroOrden').AsString;
          QrAvances.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QrAvances.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QrAvances.Open;


        end;
        with NewListaParams do
        begin
          Clear;
          Add('Contrato='+QrActividades.FieldByName('sContrato').AsString);
          Add('Convenio='+ListaParams.Values['Convenio']);
          Add('Folio='+QrActividades.FieldByName('sNumeroOrden').AsString);
          Add('Wbs='+QrActividades.FieldByName('swbs').AsString);
        end;
        VolumenActividades(MsProject,NewListaParams,Ren,ItotalReg,PgbAvance);
        QrActividades.Next;
      end;
    except

    end;

  finally
    NewListaParams.Destroy;
    QrActividades.Destroy;
    QrAvances.Destroy;
    QrFechas.destroy;
  end;

end;

Procedure TdGenerador(ListaParams:TstringList;TipoGenerador:FtGenerador;Var Reporte: TfrxReport;IsMonto:Boolean=False);Overload;
////  FtGenerador=(FtGBarco=1,FtGPersonal=2,FtGEquipo=3,FtGMaterial=4,FtGPernocta=5,FtGAnexoPU=6,FtTiempoExtra=7);

const
{$Region 'SQL'}
  Asql:Array[1..7,1..3] of String =( ('SELECT me.sOrden as Contrato,mf.sFolio as Folio,tm.sIdTipoMovimiento as sIdRecurso,tm.sDescripcion as Descripcion,' + #10 +
                                'tm.sTipo,tm.smedida as medida,me.didfecha,mf.sFolio,ifnull(SUM( mf.sFactor ),0) AS Total,' + #10 +
                                'ifnull((select sdescripcion from anexos where sTipo="BARCO" limit 1)'+
                                ',"No existe") as Titulo, ' +
                                 'ifnull((select sAnexo from anexos where sTipo="BARCO" limit 1)'+
                                ',"No existe") as sAnexo, ' +
                                 'ifnull((select sMedida from anexos where sTipo="BARCO" limit 1)'+
                                ',"No existe") as sMedidaAnexo, ' +
                                'ifnull((select stitulo from anexos where sTipo="BARCO" limit 1)'+
                                ',"No existe") as sleyenda, ' +
                                'tm.dVentaMn,tm.dVentaDll ' +
                                'FROM tiposdemovimiento AS tm '  +
                                'inner join movimientosdeembarcacion AS me '   +
                                'on(me.sContrato = tm.sContrato AND me.sClasificacion = tm.sIdTipoMovimiento) ' + #10 +
                                'inner join movimientosxfolios mf on (me.sContrato = mf.sContrato and me.sOrden=mf.sNumeroOrden' + #10 +
                                'and me.didfecha=mf.dIdFecha and me.iIdDiario = mf.iIdDiario) ' +
                                'WHERE tm.sContrato =:ContratoBarco and ' + #10 +
                                '((:Orden=-1) or (:Orden<>-1 and me.sOrden =:Orden)) '+
                                'and ((:Folio=-1) or (:Folio<>-1 and mf.sFolio=:Folio)) ' + #10 +
                                'AND tm.sClasificacion = "Movimiento de Barco" and' + #10 +
                                'tm.lGenera ="Si" and me.didfecha between :Inicio and :termino ' +
                                'group by tm.sIdTipoMovimiento,me.didfecha '+
                                'order by tm.sIdTipoMovimiento,me.didfecha',
                                'group by me.sOrden,mf.sFolio,',
                                'order by me.sOrden,mf.sFolio,')
                                //'group by me.sOrden,mf.sFolio,tm.sIdTipoMovimiento,me.didfecha '+
                                //'order by me.sOrden,mf.sFolio,tm.sIdTipoMovimiento,me.didfecha'
                              ,
                                ('(SELECT bp.sContrato as Contrato,bp.sNumeroOrden as Folio,bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida, '+
                                'a.sDescripcion as Titulo,a.sAnexo, ' +
                                'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                '(IF(SUM(bp.dCanthh) > 0, '+
                                '(SUM(bp.dCanthh) + ' +
                                ' IFNULL((SELECT ' +
                                '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '          FROM ' +
                                '            bitacoradeajustes ba ' +
                                '          WHERE ' +
                                '            ba.sContrato = bp.sContrato AND ' +
                                '            ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                '            ba.dFecha = bp.dIdFecha), 0)), ' +
                                'SUM(bp.dCantidad))) AS Total,bp.dIdFecha,p.iItemOrden,a.iOrden,'+
                                'p.dVentaMn,p.dVentaDll,a.sMedida as sMedidaAnexo,a.stitulo as sleyenda FROM bitacoradepersonal bp '+
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal ) ' +     //and bp.sAgrupaPersonal = p.sIdTipoPersonal
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="PERSONAL" and a.sTierra="No") '+
                                'WHERE  ((:Orden=-1) or (:Orden<>-1 and bp.scontrato =:Orden)) AND ' +
                                '((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino ' +
                               // '((:Tipo="TE" and p.sIdTipoPersonal =:Tipo) or (:Tipo<>"TE" and p.sIdTipoPersonal<>"TE" )) '  +
                                'and p.lCobro=:Cobro '  +
                                'group by bp.sIdPersonal,bp.didfecha) '+
                                'union all '+
                                '(SELECT bp.sContrato as Contrato,bp.sNumeroOrden as Folio,bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,a.sAnexo, ' +
                                'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                '(SUM(bp.dCanthh) + ' +
                                ' IFNULL((SELECT ' +
                                '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '          FROM ' +
                                '            bitacoradeajustes ba ' +
                                '          WHERE ' +
                                '            ba.sContrato = bp.sContrato AND ' +
                                '            ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ' +
                                '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                '            ba.dFecha = bp.dIdFecha), 0)) AS Total,bp.dIdFecha, '+
                                'p.iItemOrden,a.iOrden,p.dVentaMn,p.dVentaDll,a.sMedida as sMedidaAnexo,a.stitulo as sleyenda  FROM bitacoradepersonal bp inner join personal p on(p.sContrato=:ContratoBarco and '+
                                'p.sIdPersonal=bp.sIdPersonal ) inner join anexos a on(a.sAnexo=p.sAnexo and '+    //and bp.sAgrupaPersonal = p.sIdTipoPersonal
                                'a.sTipo="PERSONAL" and a.sTierra="Si") '+
                                'WHERE  ((:Orden=-1) or (:Orden<>-1 and bp.scontrato =:Orden)) AND ' +
                                '((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) ' +
                                'and bp.dIdFecha between :Inicio and :termino ' +
                              //  '((:Tipo="TE" and p.sIdTipoPersonal =:Tipo) or (:Tipo<>"TE" and p.sIdTipoPersonal<>"TE" )) '  +
                                'and p.lCobro=:Cobro '  +
                                'group by bp.sIdPersonal,bp.didfecha) '+
                                'order by iOrden,iItemOrden,sIdRecurso,didfecha',
                                'group by bp.sContrato,bp.sNumeroOrden,',
                                'order by Contrato,Folio,'
                                )

                              , ('select be.sContrato as Contrato,be.sNumeroOrden as Folio,be.sIdEquipo as sIdRecurso,be.sDescripcion as Descripcion,e.sMedida as medida,'+
                                'ifnull((select sdescripcion from anexos where sTipo="EQUIPO" and sTierra=:Tierra limit 1),"No existe") as Titulo,'+
                                'ifnull((select sAnexo from anexos where sTipo="EQUIPO" and sTierra=:Tierra limit 1),"N/E") as sAnexo,'+
                                'ifnull((select sMedida from anexos where sTipo="EQUIPO" and sTierra=:Tierra limit 1),"N/E") as sMedidaAnexo,'+
                                 'ifnull((select sTitulo from anexos where sTipo="EQUIPO" and sTierra=:Tierra limit 1),"N/E") as sLeyenda,'+
                                'IFNULL(SUM(CAST(be.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                '(IFNULL(sum(be.dCantHH), 0) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = be.sContrato AND ' +
                                '           ba.sNumeroOrden = be.sNumeroOrden AND ' +
                                '           ba.sIdPartidaAnexo = be.sIdEquipo AND ' +
                                '           ba.sNumeroActividad = be.sNumeroActividad AND ' +
                                '           ba.dFecha = be.dIdFecha), 0)) as total,be.didfecha,e.dVentaMn,e.dVentaDll from bitacoradeequipos be inner join equipos e on '+
                                '(e.sContrato=:contratoBarco and e.sIdEquipo=be.sIdEquipo) inner join bitacoradeactividades '+
                                'ba on(ba.sContrato = be.sContrato and ba.dIdFecha = be.didfecha and ba.sNumeroOrden = '+
                                'be.sNumeroOrden and ba.iIdDiario = be.iIdDiario and ba.iIdTarea = be.iIdTarea and '+
                                'ba.iIdActividad=be.iIdActividad) '+
                                'where ((:Orden=-1) or (:Orden<>-1 and be.sContrato =:Orden)) and '+
                                '((:folio=-1) or (:folio<>-1 and be.sNumeroOrden=:folio)) ' +
                                'and be.dIdFecha between :Inicio and :termino and ' +
                              //  '((:Tipo="TE" and e.sIdTipoEquipo =:Tipo) or (:Tipo<>"TE" and e.sIdTipoEquipo<>"TE" )) '  +
                                'e.lCobro=:Cobro ' +
                                'group by e.sIdEquipo,be.didfecha '+
                                'order by e.iItemOrden,e.sIdEquipo,be.didfecha',
                                'group by be.sContrato,be.sNumeroOrden,',
                                'order by be.sContrato,be.sNumeroOrden,'

                                )
                              ,
                                ('select b.sContrato as Contrato,b.sNumeroOrden as Folio,i.sIdInsumo as sIdRecurso,'+
                                'i.mDescripcion as descripcion,i.sMedida as medida, sum(b.dCantidad) as total,i.sIdProveedor,'+
                                'b.dFechaSalida as dIdFecha,ifnull((select sdescripcion from anexos where sTipo="MATERIAL" '+
                                'and sTierra="No" order by iorden limit 1),"No existe") as Titulo, ' +
                                'ifnull((select sAnexo from anexos where sTipo="MATERIAL" and sTierra="No" '+
                                'order by iorden limit 1),"No existe") as sAnexo,i.dVentaMn,i.dVentaDll, ' +
                                 'ifnull((select sMedida from anexos where sTipo="MATERIAL" and sTierra="No" '+
                                'order by iorden limit 1),"No existe") as sMedidaAnexo '+
                                'ifnull((select sTitulo from anexos where sTipo="MATERIAL" and sTierra="No" '+
                                'order by iorden limit 1),"No existe") as sLeyenda '+
                                'from almacen_salida s '+
                                'inner join bitacoradesalida b '+
                                'on(b.sContrato = s.sContrato and b.iFolioSalida = s.iFolioSalida and b.sext = s.sext) '+
                                'inner join insumos i '+
                                'on (i.sContrato =:ContratoBarco and i.sIdInsumo = b.sIdInsumo and i.sTrazabilidad = b.sTrazabilidad ) '+
                                'where ((:orden=-1) or (:orden<>-1 and s.sContrato =:orden)) and '+
                                '((:Folio=-1) or (:Folio<>-1 and s.sNumeroOrden=:Folio)) '+
                                'and b.dFechaSalida between :Inicio and :termino '+
                                'group by b.sIdInsumo,didfecha '+
                                'Order by b.sIdInsumo,didfecha',
                                'group by b.sContrato,b.sNumeroOrden,',
                                'Order by b.sContrato,b.sNumeroOrden,'
                                )
                              , ( 'select '+
                                  'bp.sContrato as contrato, '+
                                  'bp.sNumeroOrden as Folio, '+
                                  'c.sIdPernocta as sIdRecurso, '+
                                  'c.sDescripcion as Descripcion, '+
                                  'c.sMedida as medida, '+
                                  'ifnull( ( select sdescripcion '+
                                            'from anexos '+
                                            'where sTipo = "PERNOCTA" limit 1 '+
                                          '), "No existe" '+
                                        ') as Titulo, '+#13+

                                  'ifnull( ( select sAnexo '+
                                            'from anexos '+
                                            'where sTipo = "PERNOCTA" '+
                                            'limit 1 '+
                                          '), "N/E" '+
                                        ') as sAnexo, '+

                                  'ifnull( ( select sMedida '+
                                            'from anexos '+
                                            'where sTipo = "PERNOCTA" '+
                                            'limit 1 '+
                                          '), "N/E" '+
                                        ') as sMedidaAnexo, '+#13+
                                       'ifnull( ( select sTitulo '+
                                            'from anexos '+
                                            'where sTipo = "PERNOCTA" '+
                                            'limit 1 '+
                                          '), "N/E" '+
                                        ') as sLeyenda, '+#13+

                                  '( (sum( bp.dCantHH ) + ' +
                                  '  IFNULL((SELECT ' +
                                  '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                  '          FROM ' +
                                  '            bitacoradeajustes ba ' +
                                  '          WHERE ' +
                                  '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                  '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                  '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) - ifnull( ( select ifnull( sum( bpr.dCantidad ), 0 ) '+
                                                                  'from bitacoradepernocta bpr  '+
                                                                  'where bpr.sContrato = ba.sContrato and '+
                                                                        'ba.didfecha=bpr.didfecha and '+
                                                                        '( bpr.sNumeroOrden <> "@" and '+
                                                                          'bpr.sNumeroOrden = ba.snumeroorden ) and '+
                                                                        'bpr.sIdCuenta = "" '+
                                                                  'group by bpr.sContrato '+
                                                                '), 0 '+
                                                              ') '+#13+
                                  ') as Total, '+

                                  'ifnull( ( select ifnull( sum( bpr.dCantidad), 0 ) '+
                                            'from bitacoradepernocta bpr '+
                                            'where bpr.sContrato = ba.sContrato and '+
                                                  'ba.didfecha = bpr.didfecha and '+
                                                  '( bpr.sNumeroOrden <> "@" and '+
                                                    'bpr.sNumeroOrden = ba.snumeroorden ) and '+
                                                  'bpr.sIdCuenta = c.sIdCuenta '+
                                            'group by '+
                                              'bpr.sContrato, '+
                                              'bpr.sNumeroOrden '+
                                          '), 0 '+
                                        ') as ajuste, '+#13+

                                  'bp.didfecha, '+
                                  'c.dVentaMn, '+
                                  'c.dVentaDll '+

                                'from cuentas c  '+

                                'left join bitacoradeactividades ba '+
                                  'on( ( ( :Orden =-1 ) or ( :Orden <> -1 and ba.sContrato=:Orden ) ) and '+
                                      '( ( :Folio = -1 ) or ( :Folio <> -1 and ba.sNumeroOrden=:Folio ) ) '+
                                    ') '+

                                'left join bitacoradepersonal bp '+
                                  'on ( bp.sContrato = ba.sContrato and '+
                                       'ba.sNumeroOrden = bp.sNumeroOrden and '+
                                       'bp.dIdFecha = ba.didfecha and '+
                                       'ba.iIdDiario=bp.iIdDiario and  '+
                                       'ba.iIdTarea = bp.iIdTarea and '+
                                       'ba.iIdActividad = bp.iIdActividad and '+
                                       'bp.sTipoPernocta = c.sIdPernocta  '+
                                      ') '+
                                'inner join anexos a '+
                                   'on (a.stipo = "PERSONAL" and a.sTierra = "No") '+

                                'inner join personal p '+
                                  'on ( p.sContrato = :contratoBarco and '+
                                       'p.sIdPersonal = bp.sIdPersonal and '+
                                       'p.sAnexo = a.sAnexo '+
                                       'and p.lPernocta = "Si"  ) '+

                                'left join tiposdepersonal tp '+
                                  'on ( p.sIdTipoPersonal = tp.sIdTipoPersonal) '+

                                'left join ordenesdetrabajo ot '+
                                  'on ( ot.sContrato = ba.sContrato and '+
                                       'ot.sNumeroOrden = ba.sNumeroOrden '+
                                     ') '+

                                'left join pernoctan pr  '+
                                  'on ( pr.sIdPernocta = ot.sIdPernocta ) '+
                                '   '+
                                'left join plataformas pl '+
                                  'on ( pl.sIdPlataforma = ot.sIdPlataforma ) '+#13+

                                'where '+
                                  '( ( :Orden = -1 ) or ( :Orden <> -1 and bp.sContrato=:Orden) ) and '+
                                  '( ( :Folio = -1) or ( :Folio <> -1 and bp.snumeroorden = :Folio ) ) and '+
                                  'bp.dIdFecha between :Inicio and :termino '+

                                'group by '+
                                  'c.sIdPernocta, '+
                                  'bp.didfecha '+#13+

                                'order by '+
                                  'c.sIdPernocta, '+
                                  'bp.didfecha ',
                                'group by bp.sContrato,bp.sNumeroOrden,',
                                'order by bp.sContrato,bp.sNumeroOrden,'
                                )
                              ,
                                ('select ba.sContrato as contrato,ba.sNumeroOrden as Folio,ba.sNumeroActividad as sIdRecurso,' + #10 +
                                'ao.mDescripcion as Descripcion,ao.sMedida as medida,sum(ba.dCantidad) as total,' + #10 + 
                                'a.sDescripcion as Titulo,a.sAnexo,a.sMedida as sMedidaAnexo,a.sTitulo as sLeyenda,ba.didfecha,ao.dVentaMn,ao.dVentaDll from bitacoradeactividades ba' + #10 +
                                'inner join actividadesxorden ao on(ao.sContrato=ba.sContrato and ao.sNumeroOrden=ba.sNumeroOrden' + #10 + 
                                'and ao.swbs=ba.swbs and ao.sNumeroACtividad=ba.sNumeroActividad) left join anexos a' + #10 + 
                                'on(a.sAnexo=ao.sAnexo) where ((:Orden =-1) or (:Orden<>-1 and ba.sContrato=:Orden)) and' + #10 + 
                                '((:Folio=-1) or (:Folio<>-1 and ba.snumeroorden=:Folio)) and ba.dIdFecha between :Inicio and' + #10 + 
                                ':termino and ba.sIdTipoMovimiento = "E" and ao.sTipoAnexo= "PU"' + #10 + 
                                'group by ba.swbs,ba.didfecha' + #10 + 
                                'order by ao.iItemOrden,ba.swbs,ba.didfecha',
                                'group by ba.sContrato,ba.sNumeroOrden,',
                                'order by ba.sContrato,ba.sNumeroOrden,'
                                )
                              ,
                              ('(SELECT bp.sContrato as Contrato,bp.sNumeroOrden as Folio,bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,a.sAnexo, ' +
                                'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                                '(IF(SUM(bp.dCanthh) > 0, '+
                                
                                '(SUM(bp.dCanthh) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)), SUM(bp.dCantidad))) AS Total,bp.dIdFecha,p.iItemOrden,a.iOrden,'+
                                'p.dVentaMn,p.dVentaDll,a.sMedida as sMedidaAnexo,a.sTitulo as sLeyenda FROM bitacoradepersonal bp '+
                                'inner join personal p on(p.sContrato=:ContratoBarco and p.sIdPersonal=bp.sIdPersonal ) ' +     //and bp.sAgrupaPersonal = p.sIdTipoPersonal
                                'inner join anexos a on(a.sAnexo=p.sAnexo and a.sTipo="TIEMPO_EXTRA" and a.sTierra="No") '+
                                'WHERE  ((:Orden=-1) or (:Orden<>-1 and bp.scontrato =:Orden)) AND ' +
                                '((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) '+
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'and p.lCobro=:Cobro  '  +
                                'group by bp.sIdPersonal,bp.didfecha) '+
                                'union all '+
                                '(SELECT bp.sContrato as Contrato,bp.sNumeroOrden as Folio,bp.sIdPersonal as sIdRecurso, bp.sDescripcion as Descripcion,p.smedida as medida,'+
                                'a.sDescripcion as Titulo,a.sAnexo, ' +
                                'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +

                                '(SUM(bp.dCanthh) + ' +
                                ' IFNULL((SELECT ' +
                                '           SUM(CAST(ba.dAjuste AS CHAR)) ' +
                                '         FROM ' +
                                '           bitacoradeajustes ba ' +
                                '         WHERE ' +
                                '           ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                                '           ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                                '           ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) AS Total,bp.dIdFecha, '+
                                'p.iItemOrden,a.iOrden,p.dVentaMn,p.dVentaDll,a.sMedida as sMedidaAnexo,a.sTitulo as sLeyenda FROM bitacoradepersonal bp inner join personal p on(p.sContrato=:ContratoBarco and '+
                                'p.sIdPersonal=bp.sIdPersonal ) inner join anexos a on(a.sAnexo=p.sAnexo and '+    //and bp.sAgrupaPersonal = p.sIdTipoPersonal
                                'a.sTipo="TIEMPO_EXTRA" and a.sTierra="Si") '+
                                'WHERE  ((:Orden=-1) or (:Orden<>-1 and bp.scontrato =:Orden)) AND ' +
                                '((:Folio=-1) or(:Folio<>-1 and bp.sNumeroOrden = :Folio)) ' +
                                'and bp.dIdFecha between :Inicio and :termino ' +
                                'and p.lCobro=:Cobro  '  +
                                'group by bp.sIdPersonal,bp.didfecha) '+
                                'order by iOrden,iItemOrden,sIdRecurso,didfecha',
                                'group by bp.sContrato,bp.sNumeroOrden,',
                                'order by Contrato,Folio,'
                                )
                              );

  //FtGenerador=(FtGBarco,FtGPersonal,FtGEquipo,FtGMaterial,FtGPernocta);
{$EndRegion}
var
  QrConsulta,QrFolios:TzREadOnlyQUery;
  MryRecursos:TjvMemoryData;
  dsGeneradorGeneral:TfrxDBDataset;
  IPos,iPosTmp,
  iGrupo:Integer;
  sIdRecurso:String;
  ValTmp:Variant;
  CantTmp,dTotal:Double;
  sDia,sConFol:String;
  TmpMantisa,iCiclo,
  MaxCiclo:Integer;
begin
  QrFolios:=TzREadOnlyQUery.Create(nil);
  QrConsulta:=TzREadOnlyQUery.Create(nil);
  try
    try
      QrConsulta.Connection:=Connection.zConnection;
      QrFolios.Connection:=Connection.zConnection;
      dsGeneradorGeneral:=TfrxDBDataset.Create(nil);
      MryRecursos:=TjvMemoryData.Create(nil);

      dsGeneradorGeneral.UserName:='dsGeneradorGeneral';

      QrFolios.Active:=False;
      QrFolios.SQL.Text:= 'select ot.sContrato,ot.sIdFolio as descFolio from ordenesdetrabajo ot '+
                          'where ot.sContrato=:Contrato and ot.sNumeroOrden=:Folio ';


      with MryRecursos do
      begin
        Active:=False;
        FieldDefs.Add('sContrato', ftString, 100, True);
        FieldDefs.Add('sFolio', ftString, 100, True);
        FieldDefs.Add('sDescFolio', ftString, 500, True);
        FieldDefs.Add('sidrecurso', ftString, 100, True);
        FieldDefs.Add('sdescripcion', ftString, 500, True);
        FieldDefs.Add('sAnexo', ftString, 50, false);
        FieldDefs.Add('sLeyenda', ftString, 50, false);
        FieldDefs.Add('sTitulo', ftString, 500, false);
        FieldDefs.Add('smedida', ftString, 100, True);
        FieldDefs.Add('smedidaanexo', ftString, 100, True);
        FieldDefs.Add('dTotal', FtFloat, 0, True);
        for IPos := 1 to 31 do
          FieldDefs.Add('dia' + InttoStr(IPos), FtFloat, 0, True);

        if isMonto then
        begin
          FieldDefs.Add('dVentaMn', FtCurrency, 0, True);
          FieldDefs.Add('dVentaDll', FtCurrency, 0, True);
          FieldDefs.Add('dTotalMn', FtCurrency, 0, True);
          FieldDefs.Add('dTotalDll', FtCurrency, 0, True);
        end;
        Active:=true;
      end;



      //FtGenerador=(FtGBarco,FtGPersonal,FtGEquipo,FtGMaterial,FtGPernocta);
      iGrupo:=ord(TipoGenerador);
      if (iGrupo=1) or (iGrupo=6) then
        TmpMantisa:=4
      else
        TmpMantisa:=2;

      iCiclo := iGrupo;


      QrConsulta.Active:=fALSE;
      
      if  ListaParams.Values['TIPO']='GENERAL' then
        QrConsulta.SQL.text:=aSql[iCiclo,1]
      else
      begin
        QrConsulta.SQL.text:=AnsiReplaceText(aSql[iCiclo,1],'group by',aSql[iCiclo,2]);
        QrConsulta.SQL.text:=AnsiReplaceText(QrConsulta.SQL.text,'order by',aSql[iCiclo,3]);
      end;

      if (TipoGenerador=FtGEquipo)  then
        if (connection.contrato.FieldByName('eLugarOt').AsString='Tierra') then
          QrConsulta.ParamByNAme('Tierra').AsString:='Si'
        else
          QrConsulta.ParamByNAme('Tierra').AsString:='No';

      QrConsulta.ParamByName('Orden').AsString    :=ListaParams.Values['CONTRATO'];
      QrConsulta.ParamByName('Folio').AsString    := ListaParams.Values['FOLIO'];
      QrConsulta.ParamByName('Inicio').AsDate:=strToDate(ListaParams.Values['INICIO']);
      QrConsulta.ParamByName('Termino').AsDate:=strToDate(ListaParams.Values['TERMINO']);
      if (iCiclo<>ord(FtGAnexoPu)) then
        QrConsulta.ParamByName('contratoBarco').AsString := ListaParams.Values['CONTRATO_BARCO'];
      if (iCiclo=ord(FtGPersonal)) or (iCiclo=Ord(FtGEquipo)) or (iCiclo=Ord(FtTiempoExtra)) then
      begin
     //   QrConsulta.ParamByName('Tipo').AsString    :=ListaParams.Values['TIEMPO'];
        QrConsulta.ParamByName('Cobro').AsString    := ListaParams.Values['COBRO'];
      end;

      QrConsulta.Open;

      sIdRecurso:='';
      sConFol:='';
      while not QrConsulta.Eof do
      begin
        if sConFol<>(QrConsulta.FieldByName('Contrato').AsString + QrConsulta.FieldByName('Folio').AsString) then
        begin
          QrFolios.Active:=false;
          QrFolios.ParamByName('Contrato').AsString:=QrConsulta.FieldByName('Contrato').AsString;
          QrFolios.ParamByName('Folio').AsString:=QrConsulta.FieldByName('Folio').AsString;
          QrFolios.open;
          sConFol:=(QrConsulta.FieldByName('Contrato').AsString + QrConsulta.FieldByName('Folio').AsString);
        end;


        if ((iGrupo<>5) and (sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString)) or
          ((iGrupo=5) and (ListaParams.Values['TIPO']='GENERAL') and (sIdRecurso<>QrConsulta.FieldByName('sIdRecurso').AsString))
           or ((iGrupo=5) and (ListaParams.Values['TIPO']<>'GENERAL') and
           (sIdRecurso<>QrConsulta.FieldByName('Folio').AsString + QrConsulta.FieldByName('sIdRecurso').AsString)) then
        begin
          if MryRecursos.State=dsInsert then
          begin
            if dTotal<>xRound(ValTmp,TmpMantisa) then
            begin
              MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat := MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat +
                + (xRound(ValTmp,TmpMantisa)-dTotal);
              MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,TmpMantisa);
            end
            else
              MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;

            if IsMonto then
            begin
              MryRecursos.FieldByName('dTotalMn').AsCurrency:=(MryRecursos.FieldByName('dTotal').AsFloat * MryRecursos.FieldByName('dventaMn').AsCurrency);
              MryRecursos.FieldByName('dTotalDll').AsCurrency:=(MryRecursos.FieldByName('dTotal').AsFloat * MryRecursos.FieldByName('dventaDll').AsCurrency);
            end;

            MryRecursos.Post;
          end;

          ValTmp:=0;
          CantTmp:=0;
          iPosTmp:=0;
          dTotal:=0;

          if iGrupo=5 then
          begin
            if  ListaParams.Values['TIPO']<>'GENERAL' then
              sIdRecurso:=QrConsulta.FieldByName('Folio').AsString + QrConsulta.FieldByName('sIdRecurso').AsString
            else
              sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
          end
          else
            sIdRecurso:=QrConsulta.FieldByName('sIdRecurso').AsString;
          MryRecursos.Append;
          MryRecursos.FieldByName('sContrato').AsString:=QrConsulta.FieldByName('Contrato').AsString;

          if  ListaParams.Values['TIPO']='GENERAL' then
            MryRecursos.FieldByName('sFolio').AsString:=QrConsulta.FieldByName('Contrato').AsString
          else
            MryRecursos.FieldByName('sFolio').AsString:=QrConsulta.FieldByName('Folio').AsString;

          MryRecursos.FieldByName('sDescFolio').AsString:=QrFolios.FieldByName('descFolio').AsString;
          MryRecursos.FieldByName('sidrecurso').AsString:=QrConsulta.FieldByName('sIdRecurso').AsString;
          MryRecursos.FieldByName('sdescripcion').AsString:=QrConsulta.FieldByName('descripcion').AsString;
          MryRecursos.FieldByName('sAnexo').AsString:=QrConsulta.FieldByName('sAnexo').AsString;
          MryRecursos.FieldByName('sTitulo').AsString:=QrConsulta.FieldByName('Titulo').AsString;
          MryRecursos.FieldByName('smedida').AsString:=QrConsulta.FieldByName('Medida').AsString;
          MryRecursos.FieldByName('sLeyenda').AsString:=QrConsulta.FieldByName('sLeyenda').AsString;
          MryRecursos.FieldByName('smedidaanexo').AsString:=QrConsulta.FieldByName('smedidaanexo').AsString;
          MryRecursos.FieldByName('dTotal').AsFloat:=0;

          if isMonto then
          begin
            MryRecursos.FieldByName('dVentaMn').AsCurrency:=QrConsulta.FieldByName('dventaMn').AsCurrency;
            MryRecursos.FieldByName('dVentaDll').AsCurrency:=QrConsulta.FieldByName('dventaDll').AsCurrency;
            MryRecursos.FieldByName('dTotalMn').AsCurrency:=0;
            MryRecursos.FieldByName('dTotalDll').AsCurrency:=0;
          end;


          for iPos := 1 to 31 do
            MryRecursos.FieldValues['dia' + Trim(IntToStr(iPos)) ] := 0.000000 ;
        end;

        sDia := 'dia' + Trim (InttoStr(DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime))) ;

        if (iGrupo=2)or (iGrupo=3) or (iGrupo=5) or (iGrupo=7) then
        begin
            MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,TmpMantisa) + QrConsulta.FieldByName('Ajuste').AsFloat;
            ValTmp:=ValTmp + (QrConsulta.FieldByName('total').Value + QrConsulta.FieldByName('Ajuste').AsFloat);
            if CantTmp<(xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat) then
            begin
              CantTmp:=xRound(QrConsulta.FieldByName('total').Value,Mantisa) + QrConsulta.FieldByName('Ajuste').AsFloat;
              iPosTmp:=DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime);
            end;
        end
        else
        begin
          MryRecursos.FieldByName(sDia).AsFloat :=xRound(QrConsulta.FieldByName('total').Value,TmpMantisa);
          ValTmp:=ValTmp + QrConsulta.FieldByName('total').Value;
          if CantTmp<xRound(QrConsulta.FieldByName('total').Value,TmpMantisa) then
          begin
            CantTmp:=xRound(QrConsulta.FieldByName('total').Value,TmpMantisa);
            iPosTmp:=DayOfTheMonth(QrConsulta.FieldByName('dIdFecha').AsDateTime);
          end;
        end;

        dTotal:=dTotal + MryRecursos.FieldByName(sDia).AsFloat;

        QrConsulta.Next;
      end;

      if MryRecursos.State=dsInsert then
      begin
        if dTotal<>xRound(ValTmp,TmpMantisa) then
        begin
          MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat := MryRecursos.FieldByName('dia' + Inttostr(iPosTmp)).AsFloat +
            + (xRound(ValTmp,TmpMantisa)-dTotal);
          MryRecursos.FieldByName('dTotal').AsFloat:=xRound(ValTmp,TmpMantisa);
        end
        else
        begin
          MryRecursos.FieldByName('dTotal').AsFloat:= dTotal;
        end;

        if IsMonto then
        begin
          MryRecursos.FieldByName('dTotalMn').AsCurrency:=(MryRecursos.FieldByName('dTotal').AsFloat * MryRecursos.FieldByName('dventaMn').AsCurrency);
          MryRecursos.FieldByName('dTotalDll').AsCurrency:=(MryRecursos.FieldByName('dTotal').AsFloat * MryRecursos.FieldByName('dventaDll').AsCurrency);
        end;

        MryRecursos.Post;
      end;
      dsGeneradorGeneral.DataSet:=MryRecursos;
      dsGeneradorGeneral.FieldAliases.Clear;
      Reporte.DataSets.Add(dsGeneradorGeneral);
    except
      on e:Exception do
        Showmessage(e.Message + ', ' + e.ClassName);
    end;

  finally
    QrConsulta.Destroy;
  end;                                                

end;


Procedure BitacoraPDF_ActividadesTareas(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccionBitacora; sParamEstatus, sParamNivelImprime : string);
Var
  QryInfoFolio,
  QryPartidas,
  QryDetallePartidas,
  QryFolios : TZQuery;

  Td_partidas,
  Td_ImpPartidas : TfrxDBDataset;
  dDiaSiguiente  : TDateTime;
  rxPartidas     : TClientDataSet;
  iTotalPersonal, indice : integer;
  lContinua      : boolean;
  QrAux:TzReadOnlyQuery;
  sIidDiario,
  siIdTarea,
  siIdActividad:String;
begin
    QrAux:=TzReadOnlyQuery.create(nil);
    Try
      {Creacion de objetos de consulta}
      QryFolios := TZQuery.Create(nil);
      QryFolios.Connection := Connection.zConnection;
      QryPartidas := TZQuery.Create(nil);
      QryPartidas.Connection := Connection.zConnection;
      QryDetallePartidas := TZQuery.Create(nil);
      QryDetallePartidas.Connection := Connection.zConnection;

      Td_partidas:= TfrxDBDataset.Create(nil);
      Td_partidas.UserName:='td_partidas';

      Td_ImpPartidas:= TfrxDBDataset.Create(nil);
      Td_ImpPartidas.UserName:='Td_ImpPartidas';

      QrAux.connection:=Connection.zconnection;
      rxPartidas := TClientDataSet.Create(nil);
      rxPartidas.Active := False;
      rxPartidas.FieldDefs.Add('sTipo', ftString, 25, False);
      rxPartidas.FieldDefs.Add('sOT', ftString, 15, False);
      rxPartidas.FieldDefs.Add('sFolio', ftString, 35, False);
      rxPartidas.FieldDefs.Add('sPlataforma', ftString, 50, False);
      rxPartidas.FieldDefs.Add('mDescripcionFolio', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('sNumeroActividad', ftString, 20, False);
      rxPartidas.FieldDefs.Add('mDescripcion', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('sHoraInicio', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sHoraFin', ftString, 5, False);
      rxPartidas.FieldDefs.Add('mDescripcionTarea', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('sEstatus', ftString, 15, False);
      rxPartidas.FieldDefs.Add('dFechaInicio', ftDate, 0, False);
      rxPartidas.FieldDefs.Add('dFechaFinal', ftDate, 0, False);
      rxPartidas.CreateDataSet;
      rxPartidas.Active := True;

      if (TImpresion=FtsTareas) then
      begin
          {$REGION 'FOLIOS REPORTADOS'}
          QryFolios.SQL.Clear;
          QryFolios.SQL.Add('SELECT ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.mDescripcion, p.sDescripcion as plataforma '+
                            'FROM ordenesdetrabajo AS ot ' +
                            '	INNER JOIN contratos AS c ' +
                            '		ON (ot.sContrato=c.sContrato) ' +
                            '	INNER JOIN bitacoradeactividades AS ba ' +
                            '		ON (ba.sContrato=c.sContrato AND ba.sNumeroOrden=ot.sNumeroOrden) ' +
                            '	LEFT JOIN tiposdemovimiento AS tm ' +
                            '		ON (tm.sContrato= :ContratoBarco AND tm.sIdTipoMovimiento=ba.sIdTipoMovimiento AND tm.sClasificacion="Tarifa Diaria") ' +
                            'inner join plataformas  p on (p.sIdPlataforma = ot.sIdPlataforma) '+
                            'WHERE c.sContrato= :Orden AND ba.dIdFecha= :Fecha and ot.sNumeroOrden =:Folio ' +
                            'GROUP BY ot.sContrato,	ot.sNumeroorden');
          QryFolios.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sContrato').AsString;
          QryFolios.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          QryFolios.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
          QryFolios.ParamByName('Folio').AsString         := ReporteDiario.FieldByName('sNumeroOrden').AsString;
          QryFolios.Open;

          while Not QryFolios.Eof do
          begin
              {$REGION 'CONSULTA - PARTIDAS REPORTADAS'}
              QryPartidas.SQL.Clear;
              QryPartidas.SQL.Add('SELECT  a.sContrato, a.sNumeroActividad, a.sWbs, a.sNumeroOrden, '+
                                  '	o.mdescripcion as sDescripcionActividad, a.iIdTarea_act ' +
                                  'FROM bitacoradeactividades a ' +
                                  'INNER JOIN actividadesxorden o ' +
                                  '		ON (o.sContrato = a.sContrato AND o.sIdConvenio = :Convenio AND o.sNumeroOrden = a.sNumeroOrden AND a.sWbs = o.sWbs AND o.sNumeroActividad = a.sNumeroActividad AND o.sTipoActividad = "Actividad" and o.sTipoAnexo = "ADM") ');
              if sParamNivelImprime = 'General' then
                 QryPartidas.SQL.Add('WHERE a.sContrato = :Orden AND a.sNumeroOrden = :Folio AND a.sIdTurno = :Turno AND a.sIdTipoMovimiento = "E" and a.lImprime ="Si" ')
              else
                 QryPartidas.SQL.Add('WHERE a.sContrato = :Orden AND a.dIdFecha = :Fecha AND a.sNumeroOrden = :Folio AND a.sIdTurno = :Turno AND a.sIdTipoMovimiento = "E" and a.lImprime ="Si" '+
                                  'GROUP BY	a.sWbs ' +
                                  'ORDER BY a.sContrato, o.iItemOrden, a.iIdTarea_act ');
              QryPartidas.ParamByName('Convenio').AsString := global_convenio;
              QryPartidas.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sContrato').AsString;
              if sParamNivelImprime = 'Seleccionado' then
                 QryPartidas.ParamByName('Fecha').AsDateTime  := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryPartidas.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryPartidas.ParamByName('Turno').AsString    := ReporteDiario.FieldByName('sIdTurno').AsString;
              QryPartidas.Open;
              {$ENDREGION}

              while Not QryPartidas.Eof do
              begin
                  {$REGION 'CONSULTA - DETALLE DE PARTIDAS'}
                  QryDetallePartidas.SQL.Clear;
                  QryDetallePartidas.SQL.Add(' select * from bitacoradeactividades_tareas where sContrato =:Contrato and sNumeroOrden =:Orden and sWbs =:Wbs '+ sParamEstatus);
                  QryDetallePartidas.ParamByName('Contrato').AsString := QryPartidas.FieldByName('sContrato').AsString;
                  QryDetallePartidas.ParamByName('Orden').AsString    := QryPartidas.FieldByName('sNumeroOrden').AsString;
                  QryDetallePartidas.ParamByName('Wbs').AsString      := QryPartidas.FieldByName('sWbs').AsString;
                  QryDetallePartidas.Open;
                  {$ENDREGION}

                  {$REGION 'INSERCIÓN DE DATOS'}
                  while Not QryDetallePartidas.Eof do
                  begin
                        rxPartidas.Append;
                        rxPartidas.FieldValues['sTipo']            := 'Tareas';
                        rxPartidas.FieldValues['sOT']              := QryFolios.FieldByName('sContrato').AsString;
                        rxPartidas.FieldValues['sFolio']           := QryFolios.FieldByName('sIdFolio').AsString;
                        rxPartidas.FieldValues['sPlataforma']      := QryFolios.FieldByName('plataforma').AsString;
                        rxPartidas.FieldValues['mDescripcionFolio']:= QryFolios.FieldByName('mDescripcion').AsString;
                        rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        rxPartidas.FieldValues['mDescripcion']     := QryPartidas.FieldByName('sDescripcionActividad').AsString;
                        rxPartidas.FieldValues['sHoraInicio']      := QryDetallePartidas.FieldByName('sHoraInicio').AsString;
                        rxPartidas.FieldValues['sHoraFin']         := QryDetallePartidas.FieldByName('sHoraFinal').AsString;
                        rxPartidas.FieldValues['mDescripcionTarea']:= QryDetallePartidas.FieldByName('mDescripcion').AsString;
                        rxPartidas.FieldValues['dFechaInicio']     := QryDetallePartidas.FieldByName('dFechaInicio').AsDateTime;
                        rxPartidas.FieldValues['dFechaFinal']      := QryDetallePartidas.FieldByName('dFechaFinal').AsDateTime;
                        rxPartidas.FieldValues['sEstatus']         := QryDetallePartidas.FieldByName('eEstatus').AsString;
                        rxPartidas.Post;
                      QryDetallePartidas.Next;
                  end;

                  if sParamNivelImprime = 'General' then
                     if QryDetallePartidas.RecordCount = 0 then
                     begin
                        rxPartidas.Append;
                        rxPartidas.FieldValues['sTipo']            := 'Tareas';
                        rxPartidas.FieldValues['sOT']              := QryFolios.FieldByName('sContrato').AsString;
                        rxPartidas.FieldValues['sFolio']           := QryFolios.FieldByName('sIdFolio').AsString;
                        rxPartidas.FieldValues['sPlataforma']      := QryFolios.FieldByName('plataforma').AsString;
                        rxPartidas.FieldValues['mDescripcionFolio']:= QryFolios.FieldByName('mDescripcion').AsString;
                        rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        rxPartidas.FieldValues['mDescripcion']     := QryPartidas.FieldByName('sDescripcionActividad').AsString;
                        rxPartidas.Append;
                     end;

                  {$ENDREGION}
                 QryPartidas.Next;
              end;

              QryFolios.Next;
          end;
          {Esta secccion es para mostrar la hoja vacia sino existen datos.}
          if rxPartidas.RecordCount = 0 then
          begin
              rxPartidas.Append;
              rxPartidas.FieldValues['sOT'] := '';
              rxPartidas.Post;
          end;
      end;

      {$ENDREGION}
      Td_Partidas.DataSet:= rxPartidas;
      Td_Partidas.FieldAliases.Clear;

      Td_ImpPartidas.DataSet:= rxPartidas;
      Td_ImpPartidas.FieldAliases.Clear;

      Reporte.DataSets.Add(Td_partidas);
      Reporte.DataSets.Add(Td_ImpPartidas);

    Finally
      QrAux.destroy;
    End;

end;


Procedure ReportePDF_AvancesCortes(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccion);
Var
  QryInfoFolio,
  QryPartidas,
  QryDetallePartidas,
  QryConsultaAvanceAnterior,
  QryConsultaAvancesAcumulados,
  QryFolios,
  QryPonderadoPartida  : TZQuery;

  Td_AvancesPartidas,
  Td_ImpAvPartidas : TfrxDBDataset;
  rxPartidas: TClientDataSet;
  indice : integer;
  lContinua  : boolean;
  sHora, sConvenio  : String;

  dAvanceRealAnt_folio,
  dAvanceRealAct_folio,
  dAvanceRealAcum_folio,

  dAvanceProgAnt_folio,
  dAvanceProgAct_folio,
  dAvanceProgAcum_folio :double;
begin
    Try
      {Creacion de objetos de consulta}
      QryConsultaAvancesAcumulados := TZQuery.Create(nil);
      QryConsultaAvancesAcumulados.Connection := Connection.zConnection;
      QryConsultaAvanceAnterior := TZQuery.Create(nil);
      QryConsultaAvanceAnterior.Connection := connection.zConnection;
      QryFolios := TZQuery.Create(nil);
      QryFolios.Connection := Connection.zConnection;
      QryPartidas := TZQuery.Create(nil);
      QryPartidas.Connection := Connection.zConnection;
      QryDetallePartidas := TZQuery.Create(nil);
      QryDetallePartidas.Connection := Connection.zConnection;
      QryInfoFolio := TZQuery.Create(nil);
      QryInfoFolio.Connection := Connection.zConnection;
      QryPonderadoPartida := TZQuery.Create(nil);
      QryPonderadoPartida.Connection := Connection.zConnection;

      Td_AvancesPartidas:= TfrxDBDataset.Create(nil);
      Td_AvancesPartidas.UserName:='td_AvancesPartidas';

      Td_ImpAvPartidas:= TfrxDBDataset.Create(nil);
      Td_ImpAvPartidas.UserName:='Td_ImpAvPartidas';


      rxPartidas := TClientDataSet.Create(nil);
      rxPartidas.Active := False;
      rxPartidas.FieldDefs.Add('sTipo', ftString, 25, False);
      rxPartidas.FieldDefs.Add('sOT', ftString, 15, False);
      rxPartidas.FieldDefs.Add('sFolio', ftString, 35, False);
      rxPartidas.FieldDefs.Add('sPlataforma', ftString, 50, False);
      rxPartidas.FieldDefs.Add('mDescripcionFolio', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntFolio_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActFolio_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumFolio_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntFolio_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActFolio_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumFolio_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sNumeroActividad', ftString, 20, False);
      rxPartidas.FieldDefs.Add('mDescripcion', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntPart_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActPart_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumPart_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntPart_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActPart_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumPart_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sHorario', ftString, 5, False);
      rxPartidas.CreateDataSet;
      rxPartidas.Active := True;

      if (TImpresion=FtsAvancesCortes) or (TImpresion=FtsAll) then
      begin
          {$REGION 'FOLIOS REPORTADOS'}
          QryFolios.SQL.Clear;
          QryFolios.SQL.Add('SELECT ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.mDescripcion, p.sDescripcion as plataforma, '+
                            ' ba.sIdConvenio as Convenio '+
                            'FROM ordenesdetrabajo AS ot ' +
                            '	INNER JOIN contratos AS c ' +
                            '		ON (ot.sContrato=c.sContrato) ' +
                            '	INNER JOIN bitacoradeactividades AS ba ' +
                            '		ON (ba.sContrato=c.sContrato AND ba.sNumeroOrden=ot.sNumeroOrden) ' +
                            '	LEFT JOIN tiposdemovimiento AS tm ' +
                            '		ON (tm.sContrato= :ContratoBarco AND tm.sIdTipoMovimiento=ba.sIdTipoMovimiento AND tm.sClasificacion="Tarifa Diaria") ' +
                            'inner join plataformas  p on (p.sIdPlataforma = ot.sIdPlataforma) '+
                            'WHERE c.sContrato= :Orden AND ba.dIdFecha= :Fecha ' +
                            'GROUP BY ot.sContrato,	ot.sNumeroorden');
          QryFolios.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
          QryFolios.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          QryFolios.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
          QryFolios.Open;

          while Not QryFolios.Eof do
          begin
               {$REGION 'AVANCE PROGRAMADO ANT. FOLIO'}
              dAvanceProgAnt_Folio  := 0;
              dAvanceProgAct_folio  := 0;
              dAvanceProgAcum_folio := 0;
              dAvanceRealAnt_folio  := 0;
              dAvanceRealAct_folio  := 0;
              dAvanceRealAcum_folio := 0;

              {Avances Programados del Folio..}
              QryConsultaAvancesAcumulados.SQL.Clear;
              QryConsultaAvancesAcumulados.SQL.Add('Select a.sIdConvenio, c.iGerencialInicio, c.dFecha,  a.dAvancePonderadoDia, a.dAvancePonderadoGlobal, a.iNumeroGerencial '+
                                           ' From avancesglobales a '+
                                           ' inner join convenios c '+
                                           ' on (c.sContrato = a.sContrato and  '+
                                           '     c.sIdConvenio = a.sIdConvenio) '+
                                           ' Where a.sContrato = :Orden And     '+
                                           '       a.sNumeroOrden = :Folio and  '+
                                           '       a.dIdFecha =:Fecha '+
                                           ' group by a.dIdFecha, a.iNumeroGerencial '+
                                           ' order by a.dIdFecha, a.iNumeroGerencial');
              QryConsultaAvancesAcumulados.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sOrden').AsString;
              QryConsultaAvancesAcumulados.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvancesAcumulados.ParamByName('Fecha').AsDate      := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryConsultaAvancesAcumulados.Open;


              {Avances Fisicos del Folio}
              QryConsultaAvanceAnterior.SQL.Clear;
              QryConsultaAvanceAnterior.SQL.Text := ObtenerSentencia( 'bitacoradeactividades', 'sql_avances_reales_frente', ftConsulta );
              QryConsultaAvanceAnterior.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sOrden').AsString;
              QryConsultaAvanceAnterior.ParamByName('Convenio').AsString := QryFolios.FieldByName('Convenio').AsString;
              QryConsultaAvanceAnterior.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvanceAnterior.ParamByName('Fecha').AsDate      := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryConsultaAvanceAnterior.Open;

              {$ENDREGION}

              for indice := 1 to 3 do
              begin
                  {$REGION 'AVANCES DEL FOLIO'}
                  QryConsultaAvanceAnterior.First;
                  while not QryConsultaAvanceAnterior.Eof do
                  begin
                      if indice = QryConsultaAvanceAnterior.FieldByName('iNumeroGerencial').AsInteger then
                      begin
                          dAvanceRealAnt_folio  := QryConsultaAvanceAnterior.FieldByName('AvanceAnterior').AsFloat;
                          dAvanceRealAct_folio  := QryConsultaAvanceAnterior.FieldByName('AvanceActual').AsFloat;
                          dAvanceRealAcum_folio := QryConsultaAvanceAnterior.FieldByName('AvanceAcumulado').AsFloat;
                      end;
                      QryConsultaAvanceAnterior.Next;
                  end;

                  QryConsultaAvancesAcumulados.First;
                  while not QryConsultaAvancesAcumulados.Eof do
                  begin
                      if indice = QryConsultaAvancesAcumulados.FieldByName('iNumeroGerencial').AsInteger then
                      begin
                          dAvanceProgAnt_folio  := xRound(QryConsultaAvancesAcumulados.FieldByName('dAvancePonderadoGlobal').AsFloat - QryConsultaAvancesAcumulados.FieldByName('dAvancePonderadoDia').AsFloat,2);
                          dAvanceProgAct_folio  := xRound(QryConsultaAvancesAcumulados.FieldByName('dAvancePonderadoDia').AsFloat,2);
                          dAvanceProgAcum_folio := xRound(QryConsultaAvancesAcumulados.FieldByName('dAvancePonderadoGlobal').AsFloat,2) ;
                          sConvenio             := QryConsultaAvancesAcumulados.FieldByName('sIdConvenio').AsString;
                      end;
                      QryConsultaAvancesAcumulados.Next;
                  end;

                  {$ENDREGION}

                  {$REGION 'CONSULTA AV. REAL - PARTIDAS REPORTADAS'}
                  QryPartidas.SQL.Clear;
                  QryPartidas.SQL.Add('SELECT  a.sContrato, a.iIdDiario,	a.sNumeroActividad, o.dPonderado, a.sWbs, '+
                                      ' (SELECT   SUM(dAvance) ' +
                                      ' FROM bitacoradeactividades ' +
                                      ' WHERE '+
                                      'dIdFecha < a.dIdFecha '+
                                      '   AND sIdTipoMovimiento = "E" ' +
                                      '   AND sWbs = a.sWbs ' +
                                      '   AND sContrato = a.sContrato ' +
                                      '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceAnterior_GeneralPartida_1, '+

                                      ' (SELECT   SUM(dAvance) ' +
                                      ' FROM bitacoradeactividades ' +
                                      ' WHERE '+
                                      'dIdFecha = a.dIdFecha and iNumeroGerencial <= :numero '+
                                      '   AND sIdTipoMovimiento = "E" ' +
                                      '   AND sWbs = a.sWbs ' +
                                      '   AND sContrato = a.sContrato ' +
                                      '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceAnterior_GeneralPartida, '+

                                      ' (SELECT  SUM(dAvance) ' +
                                      ' FROM bitacoradeactividades ' +
                                      ' WHERE ' +
                                      '   dIdFecha = a.dIdFecha and iNumeroGerencial =:Numero ' +
                                      '   AND sIdTipoMovimiento = "E" ' +
                                      '   AND sWbs = a.sWbs ' +
                                      '   AND sContrato = a.sContrato ' +
                                      '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceActual_GeneralPartida, ' +

                                      '	o.mdescripcion as sDescripcionActividad ' +
                                      'FROM bitacoradeactividades a ' +
                                      'INNER JOIN actividadesxorden o ' +
                                      '		ON (o.sContrato = a.sContrato AND o.sIdConvenio = :Convenio AND o.sNumeroOrden = a.sNumeroOrden AND a.sWbs = o.sWbs AND o.sNumeroActividad = a.sNumeroActividad AND o.sTipoActividad = "Actividad" and o.sTipoAnexo = "ADM") ' +

                                      'WHERE a.sContrato = :Orden AND a.dIdFecha = :Fecha AND a.sNumeroOrden = :Folio AND a.sIdTurno = :Turno AND a.sIdTipoMovimiento = "E" and a.lImprime ="Si" ' +
                                      'GROUP BY	a.sNumeroActividad ' +
                                      'ORDER BY a.sContrato, o.iItemOrden, a.sNumeroActividad, a.sHoraInicio ');
                  QryPartidas.ParamByName('Convenio').AsString := QryFolios.FieldByName('Convenio').AsString;
                  QryPartidas.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sOrden').AsString;
                  QryPartidas.ParamByName('Fecha').AsDateTime  := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                  QryPartidas.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
                  QryPartidas.ParamByName('Turno').AsString    := ReporteDiario.FieldByName('sIdTurno').AsString;
                  QryPartidas.ParamByName('Numero').AsInteger  := indice;
                  QryPartidas.Open;
                  {$ENDREGION}

                  while Not QryPartidas.Eof do
                  begin
                      {$REGION 'AVANCES PROGRAMADOS PARTIDAS'}
                      QryPonderadoPartida.SQL.Clear;
                      QryPonderadoPartida.SQL.Add('select b.sNumeroActividad, '+
                                            '(select sum(dCantidad) '+
                                            'from '+
                                            'distribuciondeactividades where sContrato = b.sContrato and sIdConvenio = b.sIdConvenio '+
                                            'and sNumeroOrden = b.sNumeroOrden and swbs = b.sWbs '+
                                            'and dIdFecha <:fecha '+
                                            ') as PonderadoAnt, '+

                                            '(select sum(dCantidad) '+
                                            'from '+
                                            'distribuciondeactividades where sContrato = b.sContrato and sIdConvenio = b.sIdConvenio '+
                                            'and sNumeroOrden = b.sNumeroOrden and swbs = b.sWbs '+
                                            'and dIdFecha = :Fecha and iNumeroGerencial < :Numero '+
                                            ') as PonderadoAcum, '+

                                            '(select dCantidad '+
                                            'from '+
                                            'distribuciondeactividades where sContrato = b.sContrato and sIdConvenio = b.sIdConvenio '+
                                            'and sNumeroOrden = b.sNumeroOrden and swbs = b.sWbs '+
                                            'and dIdFecha = :Fecha and iNumeroGerencial = :Numero '+
                                            ') as PonderadoActual '+

                                            'from distribuciondeactividades b where b.sContrato = :Orden and b.sIdConvenio = :Convenio '+
                                            'and b.sNumeroOrden = :Folio and b.swbs = :Wbs '+
                                            'group by b.sWbs ');
                      QryPonderadoPartida.ParamByName('Convenio').AsString := sConvenio;
                      QryPonderadoPartida.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sOrden').AsString;
                      QryPonderadoPartida.ParamByName('Fecha').AsDate      := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryPonderadoPartida.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
                      QryPonderadoPartida.ParamByName('Wbs').AsString      := QryPartidas.FieldByName('sWbs').AsString;
                      QryPonderadoPartida.ParamByName('Numero').AsInteger  := indice;
                      QryPonderadoPartida.Open;
                      {$ENDREGION}

                      {$REGION 'Insertar Datos en RxPartidas'}
                      rxPartidas.Append;
                      rxPartidas.FieldValues['sTipo']       := 'General';
                      rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                      rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                      rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                      rxPartidas.FieldValues['mDescripcionFolio'] := QryFolios.FieldValues['mDescripcion'];

                      rxPartidas.FieldValues['dAvanceAntFolio_Real']   := dAvanceRealAnt_folio;
                      rxPartidas.FieldValues['dAvanceActFolio_Real']   := dAvanceRealAct_folio ;
                      rxPartidas.FieldValues['dAvanceAcumFolio_Real']  := dAvanceRealAcum_folio;

                      rxPartidas.FieldValues['dAvanceAntFolio_Prog']   := dAvanceProgAnt_folio;
                      rxPartidas.FieldValues['dAvanceActFolio_Prog']   := dAvanceProgAct_folio;
                      rxPartidas.FieldValues['dAvanceAcumFolio_Prog']  := dAvanceProgAcum_folio;

                      rxPartidas.FieldValues['sNumeroActividad']       := QryPartidas.FieldByName('sNumeroActividad').AsString;
                      rxPartidas.FieldValues['mDescripcion']           := QryPartidas.FieldByName('sDescripcionActividad').AsString;

                      if indice = 1 then
                      begin
                          rxPartidas.FieldValues['dAvanceAntPart_Real']    := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida_1').AsFloat ;
                          rxPartidas.FieldValues['dAvanceActPart_Real']    := QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                          rxPartidas.FieldValues['dAvanceAcumPart_real']   := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida_1').AsFloat + QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;;
                      end
                      else
                      begin
                          rxPartidas.FieldValues['dAvanceAntPart_Real']    := (QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida_1').AsFloat + QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat) - QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                          rxPartidas.FieldValues['dAvanceActPart_Real']    := QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                          rxPartidas.FieldValues['dAvanceAcumPart_real']   := rxPartidas.FieldValues['dAvanceAntPart_Real'] + QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat ;
                      end;

                      rxPartidas.FieldValues['dAvanceAntPart_Prog']    := QryPonderadoPartida.FieldByName('PonderadoAnt').AsFloat + QryPonderadoPartida.FieldByName('PonderadoAcum').AsFloat;
                      rxPartidas.FieldValues['dAvanceAcumPart_Prog']   := QryPonderadoPartida.FieldByName('PonderadoAnt').AsFloat + QryPonderadoPartida.FieldByName('PonderadoAcum').AsFloat + QryPonderadoPartida.FieldByName('PonderadoActual').AsFloat;
                      rxPartidas.FieldValues['dAvanceActPart_Prog']    := QryPonderadoPartida.FieldByName('PonderadoActual').AsFloat;
                      if indice = 1 then
                         rxPartidas.FieldValues['sHorario']            := '05:00';
                      if indice = 2 then
                         rxPartidas.FieldValues['sHorario']            := '17:00';
                      if indice = 3 then
                         rxPartidas.FieldValues['sHorario']            := '24:00';
                      rxPartidas.Post;
                      {$ENDREGION}

                      QryPartidas.Next;
                  end;
              end;

              QryFolios.Next;
          end;
          {Esta secccion es para mostrar la hoja vacia sino existen datos.}
          if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
          begin
             if QryFolios.RecordCount = 0 then
             begin
                 rxPartidas.Append;
                 rxPartidas.FieldValues['sOT'] := '';
                 rxPartidas.Post;
             end;
          end;
      end;

      {$ENDREGION}
      Td_AvancesPartidas.DataSet:= rxPartidas;
      Td_AvancesPartidas.FieldAliases.Clear;

      Td_ImpAvPartidas.DataSet:= rxPartidas;
      Td_ImpAvPartidas.FieldAliases.Clear;

      Reporte.DataSets.Add(Td_AvancesPartidas);
      Reporte.DataSets.Add(Td_ImpAvPartidas);

    Finally

    End;

end;

Procedure BitacoraPDF_AvancesCortes(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccionBitacora);
Var
  QryInfoFolio,
  QryPartidas,
  QryDetallePartidas,
  QryConsultaAvanceAnterior,
  QryConsultaAvancesAcumulados,
  QryFolios,
  QryPonderadoPartida  : TZQuery;

  Td_AvancesPartidas,
  Td_ImpAvPartidas : TfrxDBDataset;
  dDiaSiguiente: TDateTime;
  rxPartidas: TClientDataSet;
  iTotalPersonal, indice : integer;
  lContinua      : boolean;
  QrAux:TzReadOnlyQuery;
  sIidDiario,
  siIdTarea,
  siIdActividad:String;

  dAvanceRealAnt_folio,
  dAvanceRealAct_folio,
  dAvanceRealAcum_folio,

  dAvanceProgAnt_folio,
  dAvanceProgAct_folio,
  dAvanceProgAcum_folio  :double;

  dFactor, dFactorAnt : double;
begin
    QrAux:=TzReadOnlyQuery.create(nil);
    Try
      {Creacion de objetos de consulta}
      QryConsultaAvancesAcumulados := TZQuery.Create(nil);
      QryConsultaAvancesAcumulados.Connection := Connection.zConnection;
      QryConsultaAvanceAnterior := TZQuery.Create(nil);
      QryConsultaAvanceAnterior.Connection := connection.zConnection;
      QryFolios := TZQuery.Create(nil);
      QryFolios.Connection := Connection.zConnection;
      QryPartidas := TZQuery.Create(nil);
      QryPartidas.Connection := Connection.zConnection;
      QryDetallePartidas := TZQuery.Create(nil);
      QryDetallePartidas.Connection := Connection.zConnection;
      QryInfoFolio := TZQuery.Create(nil);
      QryInfoFolio.Connection := Connection.zConnection;
      QryPonderadoPartida := TZQuery.Create(nil);
      QryPonderadoPartida.Connection := Connection.zConnection;

      Td_AvancesPartidas:= TfrxDBDataset.Create(nil);
      Td_AvancesPartidas.UserName:='td_AvancesPartidas';

      Td_ImpAvPartidas:= TfrxDBDataset.Create(nil);
      Td_ImpAvPartidas.UserName:='Td_ImpAvPartidas';

      QrAux.connection:=Connection.zconnection;
      rxPartidas := TClientDataSet.Create(nil);
      rxPartidas.Active := False;
      rxPartidas.FieldDefs.Add('sTipo', ftString, 25, False);
      rxPartidas.FieldDefs.Add('sOT', ftString, 15, False);
      rxPartidas.FieldDefs.Add('sFolio', ftString, 35, False);
      rxPartidas.FieldDefs.Add('sPlataforma', ftString, 50, False);
      rxPartidas.FieldDefs.Add('mDescripcionFolio', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntFolio_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActFolio_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumFolio_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntFolio_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActFolio_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumFolio_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sNumeroActividad', ftString, 20, False);
      rxPartidas.FieldDefs.Add('mDescripcion', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntPart_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActPart_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumPart_Real', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntPart_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActPart_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumPart_Prog', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sHorario', ftString, 5, False);
      rxPartidas.CreateDataSet;
      rxPartidas.Active := True;

      if (TImpresion=FtsAvancesParciales) then
      begin
          {$REGION 'FOLIOS REPORTADOS'}
          QryFolios.SQL.Clear;
          QryFolios.SQL.Add('SELECT ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.mDescripcion, p.sDescripcion as plataforma, '+
                            ' (select reprogramacion(:Orden, ot.sNumeroOrden, :Fecha)) as Convenio '+
                            'FROM ordenesdetrabajo AS ot ' +
                            '	INNER JOIN contratos AS c ' +
                            '		ON (ot.sContrato=c.sContrato) ' +
                            '	INNER JOIN bitacoradeactividades AS ba ' +
                            '		ON (ba.sContrato=c.sContrato AND ba.sNumeroOrden=ot.sNumeroOrden) ' +
                            '	LEFT JOIN tiposdemovimiento AS tm ' +
                            '		ON (tm.sContrato= :ContratoBarco AND tm.sIdTipoMovimiento=ba.sIdTipoMovimiento AND tm.sClasificacion="Tarifa Diaria") ' +
                            'inner join plataformas  p on (p.sIdPlataforma = ot.sIdPlataforma) '+
                            'WHERE c.sContrato= :Orden AND ba.dIdFecha= :Fecha ' +
                            'GROUP BY ot.sContrato,	ot.sNumeroorden');
          QryFolios.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sContrato').AsString;
          QryFolios.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          QryFolios.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
          QryFolios.Open;

          while Not QryFolios.Eof do
          begin
               {$REGION 'AVANCE PROGRAMADO ANT. FOLIO'}
              {Avances Programados del Folio..}
              QryConsultaAvanceAnterior.SQL.Clear;
              QryConsultaAvanceAnterior.SQL.Add('Select a.dAvancePonderadoDia, a.dAvancePonderadoGlobal ' +
                                                'From avancesglobales a ' +
                                                'Where a.sContrato = :Orden And a.sIdConvenio = :Convenio And a.sNumeroOrden = :Folio and a.dIdFecha =:Fecha ');
              QryConsultaAvanceAnterior.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sContrato').AsString;
              QryConsultaAvanceAnterior.ParamByName('Convenio').AsString := QryFolios.FieldByName('Convenio').AsString;
              QryConsultaAvanceAnterior.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvanceAnterior.ParamByName('Fecha').AsDate      := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryConsultaAvanceAnterior.Open;

              dAvanceProgAnt_Folio  := 0;
              dAvanceProgAct_folio  := 0;
              dAvanceProgAcum_folio := 0;
              dAvanceRealAnt_folio  := 0;
              dAvanceRealAct_folio  := 0;
              dAvanceRealAcum_folio := 0;
              if QryConsultaAvanceAnterior.RecordCount > 0 then
              begin
                  dAvanceProgAnt_folio  := QryConsultaAvanceAnterior.FieldByName('dAvancePonderadoGlobal').AsFloat - QryConsultaAvanceAnterior.FieldByName('dAvancePonderadoDia').AsFloat;
                  dAvanceProgAct_folio  := QryConsultaAvanceAnterior.FieldByName('dAvancePonderadoDia').AsFloat;
                  dAvanceProgAcum_folio := QryConsultaAvanceAnterior.FieldByName('dAvancePonderadoGlobal').AsFloat ;
              end;
              {$ENDREGION}

              for indice := 1 to 3 do
              begin
                  {$REGION 'AVANCES DEL FOLIO'}
                  {FISICOS - FOLIO}
                  {Avances anteriores}
                  if indice = 1 then
                  begin
                      QryConsultaAvanceAnterior.SQL.Clear;
                      QryConsultaAvanceAnterior.SQL.Add('Select sum((b.dPonderado / 100)* a.dAvance) as dAvanceAnterior '+
                                                'From actividadesxorden b '+
                                                'inner JOIN bitacoradeactividades a '+
                                                'ON (a.sContrato=b.sContrato And a.sWbs=b.sWbs and b.sNumeroOrden=a.sNumeroOrden and dIdFecha <:fecha ) '+
                                                'left JOIN tiposdemovimiento t '+
                                                'ON (b.sContrato=t.sContrato And a.sIdTipoMovimiento=t.sIdTipoMovimiento And t.sClasificacion="Tiempo en Operacion") '+
                                                'Where b.sContrato=:Orden And b.sIdConvenio=:Convenio And b.sNumeroOrden =:Folio Group By a.sContrato');
                      QryConsultaAvanceAnterior.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sContrato').AsString;
                      QryConsultaAvanceAnterior.ParamByName('Convenio').AsString := QryFolios.FieldByName('Convenio').AsString;
                      QryConsultaAvanceAnterior.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
                      QryConsultaAvanceAnterior.ParamByName('Fecha').AsDate      := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryConsultaAvanceAnterior.Open;
                      dAvanceRealAnt_folio := QryConsultaAvanceAnterior.FieldByName('dAvanceAnterior').AsFloat;
                  end;

                  //Avances Acumulados
                  QryConsultaAvancesAcumulados.SQL.Clear;
                  QryConsultaAvancesAcumulados.SQL.Add('Select sum((b.dPonderado / 100)* a.dAvance) as dAvanceAcumulado '+
                                              'From actividadesxorden b '+
                                              'inner JOIN bitacoradeactividades a '+
                                              'ON (a.sContrato=b.sContrato And a.sWbs=b.sWbs and b.sNumeroOrden=a.sNumeroOrden and dIdFecha =:Fecha and iNumeroGerencial =:numero ) '+
                                              'left JOIN tiposdemovimiento t '+
                                              'ON (b.sContrato=t.sContrato And a.sIdTipoMovimiento=t.sIdTipoMovimiento And t.sClasificacion="Tiempo en Operacion") '+
                                              'Where b.sContrato=:Orden And b.sIdConvenio=:Convenio And b.sNumeroOrden =:Folio Group By a.sContrato ');
                  QryConsultaAvancesAcumulados.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sContrato').AsString;
                  QryConsultaAvancesAcumulados.ParamByName('Convenio').AsString := QryFolios.FieldByName('Convenio').AsString;
                  QryConsultaAvancesAcumulados.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
                  QryConsultaAvancesAcumulados.ParamByName('Fecha').AsDate      := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                  QryConsultaAvancesAcumulados.ParamByName('numero').AsInteger  := indice;
                  QryConsultaAvancesAcumulados.Open;
                  dAvanceRealAct_folio := QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').AsFloat;

                  if indice = 1 then
                     dAvanceProgAcum_folio := (dAvanceProgAct_folio * 0.2083);
                  if indice = 2 then
                     dAvanceProgAcum_folio := (dAvanceProgAct_folio * 0.7083);
                  if indice = 3 then
                     dAvanceProgAcum_folio := (dAvanceProgAct_folio * 0.0834);

                  {$ENDREGION}

                  {$REGION 'CONSULTA AV. REAL - PARTIDAS REPORTADAS'}
                  QryPartidas.SQL.Clear;
                  QryPartidas.SQL.Add('SELECT  a.sContrato, a.iIdDiario,	a.sNumeroActividad, o.dPonderado, a.sWbs, '+
                                      ' (SELECT   SUM(dAvance) ' +
                                      ' FROM bitacoradeactividades ' +
                                      ' WHERE '+
                                      'dIdFecha < a.dIdFecha '+
                                      '   AND sIdTipoMovimiento = "E" ' +
                                      '   AND sWbs = a.sWbs ' +
                                      '   AND sContrato = a.sContrato ' +
                                      '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceAnterior_GeneralPartida_1, '+

                                      ' (SELECT   SUM(dAvance) ' +
                                      ' FROM bitacoradeactividades ' +
                                      ' WHERE '+
                                      'dIdFecha = a.dIdFecha and iNumeroGerencial <= :numero '+
                                      '   AND sIdTipoMovimiento = "E" ' +
                                      '   AND sWbs = a.sWbs ' +
                                      '   AND sContrato = a.sContrato ' +
                                      '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceAnterior_GeneralPartida, '+

                                      ' (SELECT  SUM(dAvance) ' +
                                      ' FROM bitacoradeactividades ' +
                                      ' WHERE ' +
                                      '   dIdFecha = a.dIdFecha and iNumeroGerencial =:Numero ' +
                                      '   AND sIdTipoMovimiento = "E" ' +
                                      '   AND sWbs = a.sWbs ' +
                                      '   AND sContrato = a.sContrato ' +
                                      '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceActual_GeneralPartida, ' +

                                      '	o.mdescripcion as sDescripcionActividad ' +
                                      'FROM bitacoradeactividades a ' +
                                      'INNER JOIN actividadesxorden o ' +
                                      '		ON (o.sContrato = a.sContrato AND o.sIdConvenio = :Convenio AND o.sNumeroOrden = a.sNumeroOrden AND a.sWbs = o.sWbs AND o.sNumeroActividad = a.sNumeroActividad AND o.sTipoActividad = "Actividad" and o.sTipoAnexo = "ADM") ' +

                                      'WHERE a.sContrato = :Orden AND a.dIdFecha = :Fecha AND a.sNumeroOrden = :Folio AND a.sIdTurno = :Turno AND a.sIdTipoMovimiento = "E" and a.lImprime ="Si" ' +
                                      'GROUP BY	a.sNumeroActividad ' +
                                      'ORDER BY a.sContrato, o.iItemOrden, a.sNumeroActividad, a.sHoraInicio ');
                  QryPartidas.ParamByName('Convenio').AsString := QryFolios.FieldByName('Convenio').AsString;
                  QryPartidas.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sContrato').AsString;
                  QryPartidas.ParamByName('Fecha').AsDateTime  := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                  QryPartidas.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
                  QryPartidas.ParamByName('Turno').AsString    := ReporteDiario.FieldByName('sIdTurno').AsString;
                  QryPartidas.ParamByName('Numero').AsInteger  := indice;
                  QryPartidas.Open;
                  {$ENDREGION}

                  if indice = 1 then
                  begin
                     dFactorAnt := 0.0;
                     dFactor    := 0.2083;
                  end;
                  if indice = 2 then
                  begin
                     dFactorAnt := 0.2083;
                     dFactor    := 0.9166;
                  end;
                  if indice = 3 then
                  begin
                     dFactorAnt := 0.9166;
                     dFactor    := 1;
                  end;


                  while Not QryPartidas.Eof do
                  begin
                      {$REGION 'AVANCES PROGRAMADOS PARTIDAS'}
                      QryPonderadoPartida.SQL.Clear;
                      QryPonderadoPartida.SQL.Add('select b.sNumeroActividad, '+
                                          '(select sum(dCantidad) '+
                                          ' from '+
                                          ' distribuciondeactividades where sContrato = b.sContrato and sIdConvenio = b.sIdConvenio '+
                                          'and sNumeroOrden = b.sNumeroOrden and swbs = b.sWbs '+
                                          'and dIdFecha < :fecha '+
                                          ') as PonderadoAnt, '+

                                          '(select (sum(dCantidad) * :Factor) '+
                                          ' from '+
                                          ' distribuciondeactividades where sContrato = b.sContrato and sIdConvenio = b.sIdConvenio '+
                                          'and sNumeroOrden = b.sNumeroOrden and swbs = b.sWbs '+
                                          'and dIdFecha = :Fecha '+
                                          ') as PonderadoActualFactor, '+

                                          '(select (sum(dCantidad) * :FactorAnt)  '+
                                          ' from '+
                                          ' distribuciondeactividades where sContrato = b.sContrato and sIdConvenio = b.sIdConvenio '+
                                          'and sNumeroOrden = b.sNumeroOrden and swbs = b.sWbs '+
                                          'and dIdFecha = :Fecha '+
                                          ') as PonderadoActualAcumFactor '+

                                          'from distribuciondeactividades b where b.sContrato = :Orden and b.sIdConvenio = :Convenio '+
                                          'and b.sNumeroOrden = :Folio and b.swbs = :Wbs '+
                                          'and b.dIdFecha <= :Fecha group by b.sNumeroActividad');
                      QryPonderadoPartida.ParamByName('Convenio').AsString := QryFolios.FieldByName('Convenio').AsString;
                      QryPonderadoPartida.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sContrato').AsString;
                      QryPonderadoPartida.ParamByName('Fecha').AsDate      := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                      QryPonderadoPartida.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
                      QryPonderadoPartida.ParamByName('Wbs').AsString      := QryPartidas.FieldByName('sWbs').AsString;
                      QryPonderadoPartida.ParamByName('Factor').AsFloat    := dFactor;
                      QryPonderadoPartida.ParamByName('FactorAnt').AsFloat := dFactorAnt;
                      QryPonderadoPartida.Open;
                      {$ENDREGION}

                      {$REGION 'Insertar Datos en RxPartidas'}
                      rxPartidas.Append;
                      rxPartidas.FieldValues['sTipo']       := 'General';
                      rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                      rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                      rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                      rxPartidas.FieldValues['mDescripcionFolio'] := QryFolios.FieldValues['mDescripcion'];

                      rxPartidas.FieldValues['dAvanceAntFolio_Real']   := dAvanceRealAnt_folio + dAvanceRealAcum_folio;
                      rxPartidas.FieldValues['dAvanceActFolio_Real']   := dAvanceRealAct_folio ;
                      rxPartidas.FieldValues['dAvanceAcumFolio_Real']  := dAvanceRealAnt_folio +  dAvanceRealAct_folio + dAvanceRealAcum_folio;

                      rxPartidas.FieldValues['dAvanceAntFolio_Prog']   := dAvanceProgAnt_folio;
                      rxPartidas.FieldValues['dAvanceActFolio_Prog']   := dAvanceProgAcum_folio;
                      rxPartidas.FieldValues['dAvanceAcumFolio_Prog']  := dAvanceProgAnt_folio + dAvanceProgAcum_folio;

                      rxPartidas.FieldValues['sNumeroActividad']       := QryPartidas.FieldByName('sNumeroActividad').AsString;
                      rxPartidas.FieldValues['mDescripcion']           := QryPartidas.FieldByName('sDescripcionActividad').AsString;

                      if indice = 1 then
                      begin
                          rxPartidas.FieldValues['dAvanceAntPart_Real']    := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida_1').AsFloat ;
                          rxPartidas.FieldValues['dAvanceActPart_Real']    := QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                          rxPartidas.FieldValues['dAvanceAcumPart_real']   := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida_1').AsFloat + QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;;
                      end
                      else
                      begin
                          rxPartidas.FieldValues['dAvanceAntPart_Real']    := (QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida_1').AsFloat + QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat) - QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                          rxPartidas.FieldValues['dAvanceActPart_Real']    := QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                          rxPartidas.FieldValues['dAvanceAcumPart_real']   := rxPartidas.FieldValues['dAvanceAntPart_Real'] + QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat ;
                      end;

                      rxPartidas.FieldValues['dAvanceAntPart_Prog']    := QryPonderadoPartida.FieldByName('PonderadoAnt').AsFloat + QryPonderadoPartida.FieldByName('PonderadoActualAcumFactor').AsFloat;
                      rxPartidas.FieldValues['dAvanceAcumPart_Prog']   := QryPonderadoPartida.FieldByName('PonderadoAnt').AsFloat + QryPonderadoPartida.FieldByName('PonderadoActualFactor').AsFloat;
                      rxPartidas.FieldValues['dAvanceActPart_Prog']    := rxPartidas.FieldValues['dAvanceAcumPart_Prog'] - rxPartidas.FieldValues['dAvanceAntPart_Prog'];
                      if indice = 1 then
                         rxPartidas.FieldValues['sHorario']            := '05:00';
                      if indice = 2 then
                         rxPartidas.FieldValues['sHorario']            := '17:00';
                      if indice = 3 then
                         rxPartidas.FieldValues['sHorario']            := '24:00';
                      rxPartidas.Post;
                      {$ENDREGION}

                      QryPartidas.Next;
                  end;
                  dAvanceProgAnt_folio  := dAvanceProgAnt_folio + dAvanceProgAcum_folio;
                  dAvanceRealAcum_folio := dAvanceRealAcum_Folio + dAvanceRealAct_folio;
              end;
              QryFolios.Next;
          end;
          {Esta secccion es para mostrar la hoja vacia sino existen datos.}
          if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
          begin
             if QryFolios.RecordCount = 0 then
             begin
                 rxPartidas.Append;
                 rxPartidas.FieldValues['sOT'] := '';
                 rxPartidas.Post;
             end;
          end;
      end;

      {$ENDREGION}
      Td_AvancesPartidas.DataSet:= rxPartidas;
      Td_AvancesPartidas.FieldAliases.Clear;

      Td_ImpAvPartidas.DataSet:= rxPartidas;
      Td_ImpAvPartidas.FieldAliases.Clear;

      Reporte.DataSets.Add(Td_AvancesPartidas);
      Reporte.DataSets.Add(Td_ImpAvPartidas);

    Finally
      QrAux.destroy;
    End;

end;

Procedure BitacoraPDF_ActividadesDetalle(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport;Tipo:FtTipo;TImpresion:FtSeccionBitacora);
Var
  QryInfoFolio,
  QryPartidas,
  QryDetallePartidas,
  QryConsultaAvanceAnterior,
  QryConsultaAvancesAcumulados,
  QryFolios : TZQuery;

  Td_partidas,
  Td_ImpPartidas: TfrxDBDataset;
  dDiaSiguiente: TDateTime;
  rxPartidas: TClientDataSet;
  iTotalPersonal, indice : integer;
  lContinua      : boolean;
  QrAux:TzReadOnlyQuery;
  sIidDiario,
  siIdTarea,
  siIdActividad:String;
begin
    QrAux:=TzReadOnlyQuery.create(nil);
    Try
      {Creacion de objetos de consulta}
      QryConsultaAvancesAcumulados := TZQuery.Create(nil);
      QryConsultaAvancesAcumulados.Connection := Connection.zConnection;
      QryConsultaAvanceAnterior := TZQuery.Create(nil);
      QryConsultaAvanceAnterior.Connection := connection.zConnection;
      QryFolios := TZQuery.Create(nil);
      QryFolios.Connection := Connection.zConnection;
      QryPartidas := TZQuery.Create(nil);
      QryPartidas.Connection := Connection.zConnection;
      QryDetallePartidas := TZQuery.Create(nil);
      QryDetallePartidas.Connection := Connection.zConnection;
      QryInfoFolio := TZQuery.Create(nil);
      QryInfoFolio.Connection := Connection.zConnection;

      Td_partidas:= TfrxDBDataset.Create(nil);
      Td_partidas.UserName:='td_partidas';

      Td_ImpPartidas:= TfrxDBDataset.Create(nil);
      Td_ImpPartidas.UserName:='Td_ImpPartidas';

      QrAux.connection:=Connection.zconnection;
      rxPartidas := TClientDataSet.Create(nil);
      rxPartidas.Active := False;
      rxPartidas.FieldDefs.Add('sTipo', ftString, 25, False);
      rxPartidas.FieldDefs.Add('sOT', ftString, 15, False);
      rxPartidas.FieldDefs.Add('sFolio', ftString, 35, False);
      rxPartidas.FieldDefs.Add('sPlataforma', ftString, 50, False);
      rxPartidas.FieldDefs.Add('mDescripcionFolio', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sNumeroActividad', ftString, 20, False);
      rxPartidas.FieldDefs.Add('mDescripcion', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('iIdDiario', ftInteger, 0, False);
      rxPartidas.FieldDefs.Add('sHoraInicio', ftString, 5, False);
      rxPartidas.FieldDefs.Add('sHoraFin', ftString, 5, False);
      rxPartidas.FieldDefs.Add('mDescripcionHorario', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('sAbierto', ftString, 2, False);
      rxPartidas.CreateDataSet;
      rxPartidas.Active := True;

      if (TImpresion=FtsTareas) then
      begin
          {$REGION 'FOLIOS REPORTADOS'}
          QryFolios.SQL.Clear;
          QryFolios.SQL.Add('SELECT ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.mDescripcion, p.sDescripcion as plataforma '+
                            'FROM ordenesdetrabajo AS ot ' +
                            '	INNER JOIN contratos AS c ' +
                            '		ON (ot.sContrato=c.sContrato) ' +
                            '	INNER JOIN bitacoradeactividades AS ba ' +
                            '		ON (ba.sContrato=c.sContrato AND ba.sNumeroOrden=ot.sNumeroOrden) ' +
                            '	LEFT JOIN tiposdemovimiento AS tm ' +
                            '		ON (tm.sContrato= :ContratoBarco AND tm.sIdTipoMovimiento=ba.sIdTipoMovimiento AND tm.sClasificacion="Tarifa Diaria") ' +
                            'inner join plataformas  p on (p.sIdPlataforma = ot.sIdPlataforma) '+
                            'WHERE c.sContrato= :Orden AND ba.dIdFecha= :Fecha ' +
                            'GROUP BY ot.sContrato,	ot.sNumeroorden');
          QryFolios.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sContrato').AsString;
          QryFolios.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          QryFolios.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
          QryFolios.Open;

          while Not QryFolios.Eof do
          begin
              {$REGION 'AVANCES DEL FOLIO'}
              {Avances anteriores}
              QryConsultaAvanceAnterior.SQL.Clear;
              QryConsultaAvanceAnterior.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', ReporteDiario.FieldByName('dIdFecha').AsDateTime)+'", :Orden, :Folio), 4) AS dAvanceAnterior;';
              QryConsultaAvanceAnterior.ParamByName('Orden').AsString := ReporteDiario.FieldByName('sContrato').AsString;
              QryConsultaAvanceAnterior.ParamByName('Folio').AsString := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvanceAnterior.Open;

              //Avances Acumulados
              dDiaSiguiente := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              dDiaSiguiente := IncDay(dDiaSiguiente);

              QryConsultaAvancesAcumulados.SQL.Clear;
              QryConsultaAvancesAcumulados.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 4) AS dAvanceAcumulado;';
              QryConsultaAvancesAcumulados.ParamByName('Orden').AsString := ReporteDiario.FieldByName('sContrato').AsString;
              QryConsultaAvancesAcumulados.ParamByName('Folio').AsString := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvancesAcumulados.Open;
              {$ENDREGION}

              {$REGION 'CONSULTA - PARTIDAS REPORTADAS'}
              QryPartidas.SQL.Clear;
              QryPartidas.SQL.Add('SELECT  a.sContrato, a.iIdDiario,	a.sNumeroActividad, o.dPonderado, a.sWbs, '+
                                  ' (SELECT   SUM(dAvance) ' +
                                  ' FROM bitacoradeactividades ' +
                                  ' WHERE ' +
                                  '   dIdFecha < a.dIdFecha ' +
                                  '   AND sIdTipoMovimiento = "E" ' +
                                  '   AND sWbs = a.sWbs ' +
                                  '   AND sContrato = a.sContrato ' +
                                  '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceAnterior_GeneralPartida, '+

                                  ' (SELECT  SUM(dAvance) ' +
                                  ' FROM bitacoradeactividades ' +
                                  ' WHERE ' +
                                  '   dIdFecha = a.dIdFecha ' +
                                  '   AND sIdTipoMovimiento = "E" ' +
                                  '   AND sWbs = a.sWbs ' +
                                  '   AND sContrato = a.sContrato ' +
                                  '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceActual_GeneralPartida, ' +

                                  '	CONCAT(a.sHoraInicio,"-",a.sHoraFinal)  AS Horario, ' +
                                  '	a.sIdClasificacion,	a.sHoraInicio, a.sHorafinal, ' +
                                  '	o.mdescripcion as sDescripcionActividad ' +
                                  'FROM bitacoradeactividades a ' +
                                  'INNER JOIN actividadesxorden o ' +
                                  '		ON (o.sContrato = a.sContrato AND o.sIdConvenio = :Convenio AND o.sNumeroOrden = a.sNumeroOrden AND a.sWbs = o.sWbs AND o.sNumeroActividad = a.sNumeroActividad AND o.sTipoActividad = "Actividad" and o.sTipoAnexo = "ADM") ' +

                                  'WHERE a.sContrato = :Orden AND a.dIdFecha = :Fecha AND a.sNumeroOrden = :Folio AND a.sIdTurno = :Turno AND a.sIdTipoMovimiento = "E" and a.lImprime ="Si" ' +
                                  'GROUP BY	a.sNumeroActividad ' +
                                  'ORDER BY a.sContrato, o.iItemOrden, a.sNumeroActividad, a.sHoraInicio ');
              QryPartidas.ParamByName('Convenio').AsString := global_convenio;
              QryPartidas.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sContrato').AsString;
              QryPartidas.ParamByName('Fecha').AsDateTime  := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryPartidas.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryPartidas.ParamByName('Turno').AsString    := ReporteDiario.FieldByName('sIdTurno').AsString;
              QryPartidas.Open;
              {$ENDREGION}

              while Not QryPartidas.Eof do
              begin
                  {Insertar Datos en RxPartidas}
                  rxPartidas.Append;
                  rxPartidas.FieldValues['sTipo']       := 'General';
                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                  rxPartidas.FieldValues['mDescripcionFolio'] := QryFolios.FieldValues['mDescripcion'];
                  rxPartidas.FieldValues['dAvanceAntFolio']   := QryConsultaAvanceAnterior.FieldByName('dAvanceAnterior').AsFloat;
                  rxPartidas.FieldValues['dAvanceActFolio']   := QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').AsFloat - QryConsultaAvanceAnterior.FieldByName('dAvanceAnterior').AsFloat;
                  rxPartidas.FieldValues['dAvanceAcumFolio']  := QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').AsFloat;
                  rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                  rxPartidas.FieldValues['mDescripcion']     := QryPartidas.FieldByName('sDescripcionActividad').AsString;
                  rxPartidas.FieldValues['dAvanceAntPart']   := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat;
                  rxPartidas.FieldValues['dAvanceActPart']   := QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                  rxPartidas.FieldValues['dAvanceAcumPart']  := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat + QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                  rxPartidas.Post;

                  {$REGION 'CONSULTA - DETALLE DE PARTIDAS'}
                  QryDetallePartidas.SQL.Clear;
                  QryDetallePartidas.SQL.Add( 'SELECT b.dIdFecha, b.sIdClasificacion, b.sHoraInicio, b.sHoraFinal, b.iIdDiario, b.sNumeroActividad, b.iIdTarea,b.iIdActividad,b.iHermano, b.lAplicaReporteUnico, '+
                                              '	( SELECT (ifnull(sum(ba.dAvance), 0)) ' +
                                              '		FROM ' +
                                              '			bitacoradeactividades AS ba ' +
                                              '		WHERE ' +
                                              '			ba.sContrato = b.sContrato ' +
                                              '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                              '		AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento ' +
                                              '		AND ba.swbs = b.swbs ' +
                                              '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                              '		AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
                                              '   < cast(b.sHoraInicio AS Time))  )	) AS dAvanceAnterior, ' +
                                              '	b.dAvance, b.mDescripcion AS descAct ' +
                                              'FROM ' +
                                              '	bitacoradeactividades b ' +
                                              'INNER JOIN tiposdemovimiento tm ON ( ' +
                                              '	tm.sContrato = :ContratoBarco ' +
                                              '	AND tm.sIdTipoMovimiento = b.sIdClasificacion ) ' +
                                              'INNER JOIN actividadesxorden ao ON ( ' +
                                              '	ao.sContrato = b.sContrato and ao.sIdConvenio =:Convenio and ao.sNumeroOrden = b.sNumeroOrden and ' +
                                              '	ao.sWbs = b.sWbs and ao.sTipoActividad = "Actividad" ) ' +
                                              'WHERE	b.sContrato = :Orden ' +
                                              'AND b.sNumeroOrden = :Folio ' +
                                              'AND b.sIdTipoMovimiento = "ED" ' +
                                              'AND b.sNumeroActividad = :Actividad ' +
                                              'AND dIdFecha = :Fecha ' +
                                              'and ((b.iHermano=-1) or (b.iHermano=b.iIdActividad)) '+
                                              'AND lImprime = "Si" group by b.iIdTarea,b.iIdActividad ' +
                                              'ORDER BY b.didFecha, ao.iItemOrden, Time(b.sHoraInicio) ');
                  QryDetallePartidas.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
                  QryDetallePartidas.ParamByName('Convenio').AsString      := Global_convenio;
                  QryDetallePartidas.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sContrato').AsString;
                  QryDetallePartidas.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                  QryDetallePartidas.ParamByName('Folio').AsString         := QryFolios.FieldByName('sNumeroOrden').AsString;
                  QryDetallePartidas.ParamByName('Actividad').AsString     := QryPartidas.FieldByName('sNumeroActividad').AsString;
                  QryDetallePartidas.Open;
                  {$ENDREGION}

                  {$REGION 'INSERCIÓN DE DATOS'}
                  while Not QryDetallePartidas.Eof do
                  begin
                      if QryDetallePartidas.FieldByName('iHermano').AsInteger=-1 then
                      begin
                        rxPartidas.Append;
                        rxPartidas.FieldValues['sTipo']       := 'DatosHorario';
                        rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                        rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                        rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                        rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        rxPartidas.FieldValues['sHoraInicio']  := QryDetallePartidas.FieldByName('sHoraInicio').AsString;
                        rxPartidas.FieldValues['sHoraFin']     := QryDetallePartidas.FieldByName('sHoraFinal').AsString;
                        rxPartidas.FieldValues['mDescripcionHorario']  := QryDetallePartidas.FieldByName('descAct').AsString;
                        rxPartidas.FieldValues['sAbierto']     := QryDetallePartidas.FieldByName('lAplicaReporteUnico').AsString;
                        rxPartidas.Post;
                      end
                      else
                      begin
                        QrAux.active:=False;
                        QrAux.sql.text:='select * from bitacoradeactividades where scontrato=:Orden and '+
                                        'didfecha=:fecha and sNumeroOrden=:Folio and sIdTipoMovimiento = "ED" and ' +
                                        'sNumeroActividad = :Actividad and iHermano=:IdVinculo '+
                                        'group by iIdTarea,iIdActividad ' +
                                        'ORDER BY Time(sHoraInicio) ';
                        QrAux.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sContrato').AsString;
                        QrAux.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                        QrAux.ParamByName('Folio').AsString         := QryFolios.FieldByName('sNumeroOrden').AsString;
                        QrAux.ParamByName('Actividad').AsString     := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        QrAux.ParamByName('IdVinculo').AsInteger     := QryDetallePartidas.FieldByName('iHermano').AsInteger;
                        QrAux.open;

                        rxPartidas.Append;
                        rxPartidas.FieldValues['sTipo']       := 'DatosHorario';
                        rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                        rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                        rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                        rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                        rxPartidas.FieldValues['sHoraInicio']  := QrAux.FieldByName('sHoraInicio').AsString;
                        rxPartidas.FieldValues['mDescripcionHorario']  := QryDetallePartidas.FieldByName('descAct').AsString;
                        rxPartidas.FieldValues['sAbierto']     := QryDetallePartidas.FieldByName('lAplicaReporteUnico').AsString;
                        sIidDiario:='';
                        siIdTarea:='';
                        siIdActividad:='';
                        while not QrAux.eof do
                        begin
                          if sIidDiario='' then
                            sIidDiario:=QrAux.FieldByName('iIdDiario').AsString
                          else
                            sIidDiario:= sIidDiario + ',' + QrAux.FieldByName('iIdDiario').AsString;

                          if siIdTarea='' then
                            siIdTarea:=QrAux.FieldByName('iIdTarea').AsString
                          else
                            siIdTarea:= siIdTarea + ',' + QrAux.FieldByName('iIdTarea').AsString;

                          if siIdActividad='' then
                            siIdActividad:=QrAux.FieldByName('iIdActividad').AsString
                          else
                            siIdActividad:= siIdActividad + ',' + QrAux.FieldByName('iIdActividad').AsString;

                          QrAux.Next;
                        end;
                        rxPartidas.FieldValues['sHoraFin']     := QrAux.FieldByName('sHoraFinal').AsString;
                        rxPartidas.Post;

                      end;

                      {$ENDREGION}
                      QryDetallePartidas.Next;
                  end;

                  {$ENDREGION}
                 QryPartidas.Next;
              end;

              QryFolios.Next;
          end;
          {Esta secccion es para mostrar la hoja vacia sino existen datos.}
          if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
          begin
             if QryFolios.RecordCount = 0 then
             begin
                 rxPartidas.Append;
                 rxPartidas.FieldValues['sOT'] := '';
                 rxPartidas.Post;
             end;
          end;
      end;

      {$ENDREGION}
      Td_Partidas.DataSet:= rxPartidas;
      Td_Partidas.FieldAliases.Clear;

      Td_ImpPartidas.DataSet:= rxPartidas;
      Td_ImpPartidas.FieldAliases.Clear;

      Reporte.DataSets.Add(Td_partidas);
      Reporte.DataSets.Add(Td_ImpPartidas);

    Finally
      QrAux.destroy;
    End;

end;


Procedure ReportePDF_ActividadesBitacora(ReporteDiario:TzREadOnlyQuery;Var Reporte: TfrxReport);
Var
  QryInfoFolio,
  QryPartidas,
  QryDetallePartidas,
  QryConsultaAvanceAnterior,
  QryConsultaAvancesAcumulados,
  QryPersonal,
  QryEquipo,
  QryMateriales,
  QryFolios,
  QryPU,
  QryAnexos,
  QryTiemposAct     : TZQuery;

  Td_partidas,
  Td_ImpPartidas,
  Td_ClasificacionTiempo : TfrxDBDataset;
  dDiaSiguiente  : TDateTime;
  rxPartidas,
  rxTiemposAct: TClientDataSet;
  iTotalPersonal, indice : integer;
  lContinua      : boolean;
  QrAux:TzReadOnlyQuery;
  sIidDiario,
  siIdTarea,
  siIdActividad:String;
begin
    QrAux:=TzReadOnlyQuery.create(nil);
    Try
      {Creacion de objetos de consulta}
      QryConsultaAvancesAcumulados := TZQuery.Create(nil);
      QryConsultaAvancesAcumulados.Connection := Connection.zConnection;
      QryConsultaAvanceAnterior := TZQuery.Create(nil);
      QryConsultaAvanceAnterior.Connection := connection.zConnection;
      QryFolios := TZQuery.Create(nil);
      QryFolios.Connection := Connection.zConnection;
      QryPartidas := TZQuery.Create(nil);
      QryPartidas.Connection := Connection.zConnection;
      QryDetallePartidas := TZQuery.Create(nil);
      QryDetallePartidas.Connection := Connection.zConnection;
      QryInfoFolio := TZQuery.Create(nil);
      QryInfoFolio.Connection := Connection.zConnection;
      QryPersonal := TZQuery.Create(nil);
      QryPersonal.Connection := Connection.zConnection;
      QryEquipo := TZQuery.Create(nil);
      QryEquipo.Connection := Connection.zConnection;
      QryMateriales := TZQuery.Create(nil);
      QryMateriales.Connection := Connection.zConnection;
      QryPU := TZQuery.Create(nil);
      QryPU.Connection := Connection.zConnection;
      QryAnexos := TZQuery.Create(nil);
      QryAnexos.Connection := Connection.zConnection;
      QryTiemposAct := TZQuery.Create(nil);
      QryTiemposAct.Connection := Connection.zConnection;

      Td_partidas:= TfrxDBDataset.Create(nil);
      Td_partidas.UserName:='td_partidas';

      Td_ImpPartidas:= TfrxDBDataset.Create(nil);
      Td_ImpPartidas.UserName:='Td_ImpPartidas';

      Td_ClasificacionTiempo := TfrxDBDataset.Create(nil);
      Td_ClasificacionTiempo.UserName:='Td_ClasificacionTiempo';

      QrAux.connection:=Connection.zconnection;
      rxPartidas := TClientDataSet.Create(nil);
      rxPartidas.Active := False;
      rxPartidas.FieldDefs.Add('sTipo', ftString, 25, False);
      rxPartidas.FieldDefs.Add('sOT', ftString, 15, False);
      rxPartidas.FieldDefs.Add('sFolio', ftString, 35, False);
      rxPartidas.FieldDefs.Add('sPlataforma', ftString, 50, False);
      rxPartidas.FieldDefs.Add('mDescripcionFolio', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumFolio', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sNumeroActividad', ftString, 20, False);
      rxPartidas.FieldDefs.Add('mDescripcion', ftMemo, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAntPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceActPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('dAvanceAcumPart', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('iIdDiario', ftInteger, 0, False);
      rxPartidas.FieldDefs.Add('sClasificacion', ftString, 10, False);
      rxPartidas.FieldDefs.Add('sIdRecurso', ftString, 30, False);
      rxPartidas.FieldDefs.Add('sDescripcionRecurso', ftString, 300, False);
      rxPartidas.FieldDefs.Add('dCantidadRecurso', ftFloat, 0, False);
      rxPartidas.FieldDefs.Add('sMedidaRecurso', ftString, 10, False);
      rxPartidas.CreateDataSet;
      rxPartidas.Active := True;

          {$REGION 'FOLIOS REPORTADOS'}
          QryFolios.SQL.Clear;
          QryFolios.SQL.Add('SELECT ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.mDescripcion, p.sDescripcion as plataforma '+
                            'FROM ordenesdetrabajo AS ot ' +
                            '	INNER JOIN contratos AS c ' +
                            '		ON (ot.sContrato=c.sContrato) ' +
                            '	INNER JOIN bitacoradeactividades AS ba ' +
                            '		ON (ba.sContrato=c.sContrato AND ba.sNumeroOrden=ot.sNumeroOrden) ' +
                            '	LEFT JOIN tiposdemovimiento AS tm ' +
                            '		ON (tm.sContrato= :ContratoBarco AND tm.sIdTipoMovimiento=ba.sIdTipoMovimiento AND tm.sClasificacion="Tarifa Diaria") ' +
                            'inner join plataformas  p on (p.sIdPlataforma = ot.sIdPlataforma) '+
                            'WHERE c.sContrato= :Orden AND ba.dIdFecha= :Fecha ' +
                            'GROUP BY ot.sContrato,	ot.sNumeroorden');
          QryFolios.ParamByName('Orden').AsString         := ReporteDiario.FieldByName('sOrden').AsString;
          QryFolios.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
          QryFolios.ParamByName('Fecha').AsDateTime       := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
          QryFolios.Open;

          while Not QryFolios.Eof do
          begin
              {$REGION 'AVANCES DEL FOLIO'}
              {Avances anteriores}
              QryConsultaAvanceAnterior.SQL.Clear;
              QryConsultaAvanceAnterior.SQL.Text := 'SELECT TRUNCATE(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', ReporteDiario.FieldByName('dIdFecha').AsDateTime)+'", :Orden, :Folio), 2) AS dAvanceAnterior;';
              QryConsultaAvanceAnterior.ParamByName('Orden').AsString := ReporteDiario.FieldByName('sOrden').AsString;
              QryConsultaAvanceAnterior.ParamByName('Folio').AsString := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvanceAnterior.Open;

              //Avances Acumulados
              dDiaSiguiente := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              dDiaSiguiente := IncDay(dDiaSiguiente);

              QryConsultaAvancesAcumulados.SQL.Clear;
              QryConsultaAvancesAcumulados.SQL.Text := 'SELECT TRUNCATE(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 2) AS dAvanceAcumulado;';
              QryConsultaAvancesAcumulados.ParamByName('Orden').AsString := ReporteDiario.FieldByName('sOrden').AsString;
              QryConsultaAvancesAcumulados.ParamByName('Folio').AsString := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryConsultaAvancesAcumulados.Open;
              {$ENDREGION}

              {$REGION 'CONSULTA - PARTIDAS REPORTADAS'}
              QryPartidas.SQL.Clear;
              QryPartidas.SQL.Add('SELECT  a.sContrato, a.iIdDiario,	a.sNumeroActividad, o.dPonderado, a.sWbs, '+
                                  ' (SELECT   SUM(dAvance) ' +
                                  ' FROM bitacoradeactividades ' +
                                  ' WHERE ' +
                                  '   dIdFecha < a.dIdFecha ' +
                                  '   AND sIdTipoMovimiento = "E" ' +
                                  '   AND sWbs = a.sWbs ' +
                                  '   AND sContrato = a.sContrato ' +
                                  '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceAnterior_GeneralPartida, '+

                                  ' (SELECT  SUM(dAvance) ' +
                                  ' FROM bitacoradeactividades ' +
                                  ' WHERE ' +
                                  '   dIdFecha = a.dIdFecha ' +
                                  '   AND sIdTipoMovimiento = "E" ' +
                                  '   AND sWbs = a.sWbs ' +
                                  '   AND sContrato = a.sContrato ' +
                                  '   AND sNumeroOrden = a.sNumeroOrden) AS dAvanceActual_GeneralPartida, ' +

                                  '	CONCAT(a.sHoraInicio,"-",a.sHoraFinal)  AS Horario, ' +
                                  '	a.sIdClasificacion,	a.sHoraInicio, a.sHorafinal, ' +
                                  '	o.mdescripcion as sDescripcionActividad ' +
                                  'FROM bitacoradeactividades a ' +
                                  'INNER JOIN actividadesxorden o ' +
                                  '		ON (o.sContrato = a.sContrato AND o.sIdConvenio = :Convenio AND o.sNumeroOrden = a.sNumeroOrden AND a.sWbs = o.sWbs AND o.sNumeroActividad = a.sNumeroActividad AND o.sTipoActividad = "Actividad" and o.sTipoAnexo = "ADM") ' +

                                  'WHERE a.sContrato = :Orden AND a.dIdFecha = :Fecha AND a.sNumeroOrden = :Folio AND a.sIdTurno = :Turno AND a.sIdTipoMovimiento = "E" and a.lImprime ="Si" ' +
                                  'GROUP BY	a.sNumeroActividad ' +
                                  'ORDER BY a.sContrato, o.iItemOrden, a.sNumeroActividad, a.sHoraInicio ');
              QryPartidas.ParamByName('Convenio').AsString := ReporteDiario.FieldByName('sIdConvenio').AsString;
              QryPartidas.ParamByName('Orden').AsString    := ReporteDiario.FieldByName('sOrden').AsString;
              QryPartidas.ParamByName('Fecha').AsDateTime  := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryPartidas.ParamByName('Folio').AsString    := QryFolios.FieldByName('sNumeroOrden').AsString;
              QryPartidas.ParamByName('Turno').AsString    := ReporteDiario.FieldByName('sIdTurno').AsString;
              QryPartidas.Open;
              {$ENDREGION}

              while Not QryPartidas.Eof do
              begin
                  {Insertar Datos en RxPartidas}
                  rxPartidas.Append;
                  rxPartidas.FieldValues['sTipo']       := 'General';
                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                  rxPartidas.FieldValues['mDescripcionFolio'] := QryFolios.FieldValues['mDescripcion'];
                  rxPartidas.FieldValues['dAvanceAntFolio']   := QryConsultaAvanceAnterior.FieldByName('dAvanceAnterior').AsFloat;
                  rxPartidas.FieldValues['dAvanceActFolio']   := QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').AsFloat - QryConsultaAvanceAnterior.FieldByName('dAvanceAnterior').AsFloat;
                  rxPartidas.FieldValues['dAvanceAcumFolio']  := QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').AsFloat;
                  rxPartidas.FieldValues['sNumeroActividad'] := QryPartidas.FieldByName('sNumeroActividad').AsString;
                  rxPartidas.FieldValues['mDescripcion']     := QryPartidas.FieldByName('sDescripcionActividad').AsString;
                  rxPartidas.FieldValues['dAvanceAntPart']   := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat;
                  rxPartidas.FieldValues['dAvanceActPart']   := QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                  rxPartidas.FieldValues['dAvanceAcumPart']  := QryPartidas.FieldByName('dAvanceAnterior_GeneralPartida').AsFloat + QryPartidas.FieldByName('dAvanceActual_GeneralPartida').AsFloat;
                  rxPartidas.Post;

                  QryPartidas.Next;
              end;

              QryFolios.Next;

              QryPersonal.SQL.Clear;
              QryPersonal.SQL.Add('select b.*, sum( b.dCantidad ) as Cantidad, (sum(b.dCantHH) + ' +
                              '  IFNULL((SELECT ' +
                              '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                              '          FROM ' +
                              '            bitacoradeajustes ba ' +
                              '          WHERE ' +
                              '            ba.sContrato = b.sContrato AND ba.sNumeroOrden = b.sNumeroOrden AND ' +
                              '            ba.sNumeroActividad = b.sNumeroActividad AND ' +
                              '            ba.sIdPartidaAnexo = b.sIdEquipo AND b.dFecha = b.dIdFecha), 0)) as dCantHHG, p.sMedida from bitacoradepersonal b '+
                              ' inner join personal p on (p.sContrato = :Contrato and b.sIdPersonal = p.sIdPersonal) '+
                              ' where b.sContrato =:Orden '+
                              'and b.dIdFecha =:fecha group by b.sIdPersonal order by b.iItemOrden ');
              QryPersonal.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryPersonal.ParamByName('Orden').AsString     := ReporteDiario.FieldByName('sOrden').AsString;
              QryPersonal.ParamByName('Contrato').AsString  := ReporteDiario.FieldByName('sContrato').AsString;
              QryPersonal.Open;

              rxPartidas.Append;
              rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
              rxPartidas.Post;

              rxPartidas.Append;
              rxPartidas.FieldValues['sTipo']       := 'TituloPersonal';
              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
              rxPartidas.Post;

              while not QryPersonal.Eof  do
              begin
                  rxPartidas.Append;
                  rxPartidas.FieldValues['sTipo']       := 'TituloDatosPersonal';
                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                  rxPartidas.FieldValues['sIdRecurso']          := QryPersonal.FieldByName('sIdPersonal').AsString;
                  rxPartidas.FieldValues['sDescripcionRecurso'] := QryPersonal.FieldByName('sDescripcion').AsString;
                  rxPartidas.FieldValues['dCantidadRecurso']    := QryPersonal.FieldByName('dCantHHG').AsFloat;
                  rxPartidas.FieldValues['sMedidaRecurso']      := QryPersonal.FieldByName('sMedida').AsString ;
                  rxPartidas.Post;
                  QryPersonal.Next;
              end;

              QryPersonal.SQL.Clear;
              QryPersonal.SQL.Add('select b.*, sum( b.dCantidad ) as Cantidad, (sum(b.dCantHH) + ' +
                              '  IFNULL((SELECT ' +
                              '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                              '          FROM ' +
                              '            bitacoradeajustes ba ' +
                              '          WHERE ' +
                              '            ba.sContrato = b.sContrato AND ba.sNumeroOrden = b.sNumeroOrden AND ' +
                              '            ba.sNumeroActividad = b.sNumeroActividad AND ' +
                              '            ba.sIdPartidaAnexo = b.sIdEquipo AND b.dFecha = b.dIdFecha), 0)) as dCantHHG, p.sMedida from bitacoradeequipos b '+
                              ' inner join equipos p on (p.sContrato = :Contrato and b.sIdEquipo = p.sIdEquipo) '+
                              ' where b.sContrato =:Orden '+
                              'and b.dIdFecha =:fecha group by b.sIdEquipo order by b.iItemOrden ');
              QryPersonal.ParamByName('Fecha').AsDateTime   := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryPersonal.ParamByName('Orden').AsString     := ReporteDiario.FieldByName('sOrden').AsString;
              QryPersonal.ParamByName('Contrato').AsString  := ReporteDiario.FieldByName('sContrato').AsString;
              QryPersonal.Open;

              rxPartidas.Append;
              rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
              rxPartidas.Post;

              rxPartidas.Append;
              rxPartidas.FieldValues['sTipo']       := 'TituloEquipo';
              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
              rxPartidas.Post;

              while not QryPersonal.Eof  do
              begin
                  rxPartidas.Append;
                  rxPartidas.FieldValues['sTipo']       := 'TituloDatosEquipo';
                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                  rxPartidas.FieldValues['sIdRecurso']          := QryPersonal.FieldByName('sIdEquipo').AsString;
                  rxPartidas.FieldValues['sDescripcionRecurso'] := QryPersonal.FieldByName('sDescripcion').AsString;
                  rxPartidas.FieldValues['dCantidadRecurso']    := QryPersonal.FieldByName('dCantHHG').AsString;
                  rxPartidas.FieldValues['sMedidaRecurso']      := QryPersonal.FieldByName('sMedida').AsString ;
                  rxPartidas.Post;
                  QryPersonal.Next;
              end;

              rxPartidas.Append;
              rxPartidas.FieldValues['sTipo']       := 'TituloEspacio';
              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
              rxPartidas.Post;

              rxPartidas.Append;
              rxPartidas.FieldValues['sTipo']       := 'TituloNota';
              rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
              rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
              rxPartidas.Post;

              QryPersonal.SQL.Clear;
              QryPersonal.SQL.Add('SELECT * FROM notas_generales WHERE sContrato = :Orden AND dIdFecha = :Fecha ORDER BY iOrden ASC');
              QryPersonal.ParamByName('Orden').AsString  := ReporteDiario.FieldByName('sOrden').AsString;
              QryPersonal.ParamByName('Fecha').AsDate    := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
              QryPersonal.Open;

              while not QryPersonal.Eof do
              begin
                  rxPartidas.Append;
                  rxPartidas.FieldValues['sTipo']       := 'TituloDatosNota';
                  rxPartidas.FieldValues['sOT']         := QryFolios.FieldByName('sContrato').AsString;
                  rxPartidas.FieldValues['sFolio']      := QryFolios.FieldByName('sIdFolio').AsString;
                  rxPartidas.FieldValues['sPlataforma'] := QryFolios.FieldByName('plataforma').AsString;
                  rxPartidas.FieldValues['mDescripcion'] := rxPartidas.FieldValues['mDescripcion'] + QryPersonal.FieldByName('sNotaGeneral').AsWideString;
                  rxPartidas.Post;
                  QryPersonal.Next;
              end;

          end;


          {Esta secccion es para mostrar la hoja vacia sino existen datos.}
          if connection.configuracion.FieldValues['eHojasBlanco'] = 'Si' then
          begin
             if QryFolios.RecordCount = 0 then
             begin
                 rxPartidas.Append;
                 rxPartidas.FieldValues['sOT'] := '';
                 rxPartidas.Post;
             end;
          end;


      {$ENDREGION}
      Td_Partidas.DataSet:= rxPartidas;
      Td_Partidas.FieldAliases.Clear;

      Td_ImpPartidas.DataSet:= rxPartidas;
      Td_ImpPartidas.FieldAliases.Clear;

      Td_ClasificacionTiempo.DataSet := rxTiemposAct;
      Td_ClasificacionTiempo.Clear;

      Reporte.DataSets.Add(Td_partidas);
      Reporte.DataSets.Add(Td_ImpPartidas);
      Reporte.DataSets.Add(Td_ClasificacionTiempo);

    Finally
      QrAux.destroy;
    End;

end;

function Td_maximoId(sParamContrato, sParamTabla, sParamCampo :string) : Integer;
var
    QryMaximo    : TZQuery;
begin
     QryMaximo:=TzQuery.create(nil);
    Try
      {Creacion de objetos de consulta}
      QryMaximo.Connection := Connection.zConnection;

      QryMaximo.Active := False;
      if trim(sParamContrato) <> '' then
         QryMaximo.SQL.Add('select max('+sParamCampo+') as Maximo from '+sParamTabla+' where sContrato =:Contrato group by sContrato ')
      else
         QryMaximo.SQL.Add('select max('+sParamCampo+') as Maximo from '+sParamTabla+' ');
      QryMaximo.ParamByName('contrato').AsString := sParamContrato;
      QryMaximo.Open;

      if QryMaximo.RecordCount > 0 then
         result := QryMaximo.FieldByName('Maximo').AsInteger + 1;

    Finally
        QryMaximo.Destroy;
    End;
end;


procedure BuscaElimina_datos(sParamTabla, sLlevaContrato, sLlevaFolio, sLlevaWbs, sLLevaAct, sParamContrato, sParamFolio, sParamWbs, sParamAct, sParamNuevoContrato, sParamNuevoFolio, sParamNuevaWbs, sParamNuevaAct : string; accion :string; lMensaje :boolean);
var
base, tabla, campo, cad : string;
datos : array[ 1..200] of String;
i,x,cuenta_campo  : Integer;
QryTablas, QryDatos    : TZQuery;
begin
    QryDatos := TzQuery.create(nil);
    QryDatos.Connection := Connection.zConnection;
    QryTablas:= TzQuery.create(nil);
    QryTablas.Connection := Connection.zConnection;
    try
     QryTablas.Active := False ;
     QryTablas.SQL.Clear ;
     QryTablas.SQL.Add('Show tables') ;
     QryTablas.Open ;
     base := 'Tables_in_'+global_db;
     i := 1;
     while not QryTablas.Eof do
     begin
         tabla :=  QryTablas.FieldValues[base];
         qryDatos.Active := False ;
         qryDatos.SQL.Clear ;
         qryDatos.SQL.Add('describe '+tabla+' ');
         qryDatos.Open ;

         if qryDatos.RecordCount > 0 then
         begin
             cuenta_campo := 0;
             while not qryDatos.Eof do
             begin
                 if tabla <> sParamTabla then
                 begin
                     //Seleccionamos primero los contrato, folio, wbs, act.
                     if (sLlevaContrato <> '') and (sLlevaFolio <> '') and (sLlevaWbs <> '') and (sLlevaAct <> '') then
                     begin
                         If qryDatos.FieldValues['Field'] = sLlevaContrato then
                            inc(cuenta_campo)
                         else
                            If qryDatos.FieldValues['Field'] = sLlevaFolio then
                               inc(cuenta_campo)
                            else
                               If qryDatos.FieldValues['Field'] = sLlevaWbs then
                                  inc(cuenta_campo)
                                 else
                                    If qryDatos.FieldValues['Field'] = sLlevaAct then
                                       inc(cuenta_campo);

                        if cuenta_campo  = 4 then
                        begin
                             datos[i] := tabla;
                             i:= i + 1;
                             cuenta_campo := 0;
                        end;
                     end
                     else
                         //Seleccionamos primero los contrato, folio, act.
                         if (sLlevaContrato <> '') and (sLlevaFolio <> '') and (sLlevaAct <> '') then
                         begin
                             If qryDatos.FieldValues['Field'] = sLlevaContrato then
                                inc(cuenta_campo)
                             else
                                If qryDatos.FieldValues['Field'] = sLlevaFolio then
                                   inc(cuenta_campo)
                                else
                                   If qryDatos.FieldValues['Field'] = sLlevaAct then
                                      inc(cuenta_campo);

                            if cuenta_campo  = 3 then
                            begin
                                 datos[i] := tabla;
                                 i:= i + 1;
                                 cuenta_campo := 0;
                            end;
                         end
                         else
                            //Seleccionamos primero los contrato, folio..
                           if (sLlevaContrato <> '') and (sLlevaFolio <> '') then
                           begin
                               If qryDatos.FieldValues['Field'] = sLlevaContrato then
                                  inc(cuenta_campo)
                               else
                                  If qryDatos.FieldValues['Field'] = sLlevaFolio then
                                     inc(cuenta_campo);

                              if cuenta_campo  = 2 then
                              begin
                                   datos[i] := tabla;
                                   i:= i + 1;
                                   cuenta_campo := 0;
                              end;
                           end
                           else
                             //Seleccionamos primero los contrato
                             if (sLlevaContrato <> '')  then
                             begin
                                 If qryDatos.FieldValues['Field'] = sLlevaContrato then
                                    inc(cuenta_campo);

                                if cuenta_campo  = 1 then
                                begin
                                     datos[i] := tabla;
                                     i:= i + 1;
                                     cuenta_campo := 0;
                                end;
                             end;
                 end;
                 qryDatos.Next;
             end;
         end;
         QryTablas.Next;
     end;

     //Elimina todos los registros...
     if accion = 'borrar' then
     begin
         x := i-1;
         while x >= 1 do
         begin
             tabla := datos[x];
             QryTablas.Active := False ;
             QryTablas.SQL.Clear ;
             if (sLlevaContrato <> '') and (sLlevaFolio <> '') and (sLlevaWbs <> '') and (sLlevaAct <> '') then
             begin
                 QryTablas.SQL.Add('delete from ' +tabla+ ' where sContrato = :Contrato and sNumeroOrden =:Folio and sWbs =:Wbs and sNumeroActividad =:Act ');
                 QryTablas.Params.ParamByName('Contrato').AsString  := sParamContrato ;
                 QryTablas.Params.ParamByName('Folio').AsString     := sParamFolio;
                 QryTablas.Params.ParamByName('Wbs').AsString       := sParamWbs;
                 QryTablas.Params.ParamByName('Act').AsString       := sParamAct;
             end
             else
                 if (sLlevaContrato <> '') and (sLlevaFolio <> '') and (sLlevaAct <> '') then
                 begin
                     QryTablas.SQL.Add('delete from ' +tabla+ ' where sContrato = :Contrato and sNumeroOrden =:Folio and sNumeroActividad =:Act ');
                     QryTablas.Params.ParamByName('Contrato').AsString  := sParamContrato ;
                     QryTablas.Params.ParamByName('Folio').AsString     := sParamFolio;
                     QryTablas.Params.ParamByName('Act').AsString       := sParamAct;
                 end
                 else
                     if (sLlevaContrato <> '') and (sLlevaFolio <> '') then
                     begin
                         QryTablas.SQL.Add('delete from ' +tabla+ ' where sContrato = :Contrato and sNumeroOrden =:Folio ');
                         QryTablas.Params.ParamByName('Contrato').AsString  := sParamContrato ;
                         QryTablas.Params.ParamByName('Folio').AsString     := sParamFolio;
                     end
                     else
                         if (sLlevaContrato <> '') then
                         begin
                             QryTablas.SQL.Add('delete from ' +tabla+ ' where sContrato = :Contrato ');
                             QryTablas.Params.ParamByName('Contrato').AsString  := sParamContrato ;
                         end;
             QryTablas.ExecSQL ;
             x := x-1;
         end;
     end;
     // Actualiza todos los registros..
     if accion = 'actualizar' then
     begin
         for x := 1 to i -1 do
         begin
             tabla := datos[x];
             QryTablas.Active := False ;
             QryTablas.SQL.Clear ;
             if (sLlevaContrato <> '') and (sLlevaFolio <> '') and (sLlevaWbs <> '') and (sLlevaAct <> '') then
             begin
                 QryTablas.SQL.Add('update ' +tabla+ ' set sWbs =:WbsNueva  where sContrato = :Contrato and sNumeroOrden =:Folio and sWbs =:Wbs and sNumeroActividad =:Act ');
                 QryTablas.Params.ParamByName('Contrato').AsString  := sParamContrato ;
                 QryTablas.Params.ParamByName('Folio').AsString     := sParamFolio;
                 QryTablas.Params.ParamByName('Wbs').AsString       := sParamWbs;
                 QryTablas.Params.ParamByName('Act').AsString       := sParamAct;
                 QryTablas.Params.ParamByName('WbsNueva').AsString  := sParamNuevaWbs;
             end
             else
                 if (sLlevaContrato <> '') and (sLlevaFolio <> '') and (sLlevaAct <> '') then
                 begin
                     QryTablas.SQL.Add('update ' +tabla+ ' set sNumeroActividad =:NuevaAct where sContrato = :Contrato and sNumeroOrden =:Folio and sNumeroActividad =:Act ');
                     QryTablas.Params.ParamByName('Contrato').AsString  := sParamContrato ;
                     QryTablas.Params.ParamByName('Folio').AsString     := sParamFolio;
                     QryTablas.Params.ParamByName('Act').AsString       := sParamAct;
                     QryTablas.Params.ParamByName('NuevaAct').AsString  := sParamNuevaAct;
                 end
                 else
                     if (sLlevaContrato <> '') and (sLlevaFolio <> '') then
                     begin
                         QryTablas.SQL.Add('update ' +tabla+ ' set sNumeroOrden =:NuevoFolio where sContrato = :Contrato and sNumeroOrden =:Folio ');
                         QryTablas.Params.ParamByName('Contrato').AsString   := sParamContrato ;
                         QryTablas.Params.ParamByName('Folio').AsString      := sParamFolio;
                         QryTablas.Params.ParamByName('NuevoFolio').AsString := sParamNuevoFolio;
                     end
                     else
                         if (sLlevaContrato <> '') then
                         begin
                             QryTablas.SQL.Add('update ' +tabla+ ' set sContrato =:NuevoContrato where sContrato = :Contrato ');
                             QryTablas.Params.ParamByName('Contrato').AsString       := sParamContrato ;
                             QryTablas.Params.ParamByName('NuevoContrato').AsString  := sParamNuevoContrato ;
                         end;
             QryTablas.ExecSQL ;
         end;
     end;
     if lMensaje = True then
        messageDLG('Proceso Terminado con Exito', mtInformation, [mbOk], 0);

    finally
        QryDatos.destroy;
        QryTablas.destroy;
    end;
end;

Function AvanceFolio(sParamcontrato, sParamFolio : string; dParamFecha : tDateTime) : double;
var
   dDiaSiguiente  : TDateTime;
begin
    {Avances anteriores}
    dDiaSiguiente := dParamFecha;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 4) AS dAvanceAnterior;';
    connection.QryBusca.ParamByName('Orden').AsString := global_contrato;
    connection.QryBusca.ParamByName('Folio').AsString := sParamFolio;
    connection.QryBusca.Open;

    //Avances Acumulados
    dDiaSiguiente := IncDay(dDiaSiguiente);
    connection.QryBusca2.SQL.Clear;
    connection.QryBusca2.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 4) AS dAvanceAcumulado;';
    connection.QryBusca2.ParamByName('Orden').AsString := global_contrato;
    connection.QryBusca2.ParamByName('Folio').AsString := sParamFolio;
    connection.QryBusca2.Open;

    result := connection.QryBusca2.FieldByName('dAvanceAcumulado').AsFloat +  connection.QryBusca.FieldByName('dAvanceAnterior').AsFloat;
end;

procedure procReporteGenerador(sParamContrato, sParamtipo, sParamCategoria :string; lParamContinuo :boolean; dParamFechaI, dParamFechaF: tDate; tOrigen: TComponent; tProcedure: tfrxGetValueEvent; FormatosExp: string = ''; PermisosExp: string = '');
var
    dsGeneradorFolio : TfrxDBDataSet;
    Categorias, qry_tripulacion : tzReadOnlyQuery;

    mDatos, mTotales, mSubTotal :TJvMemoryData;

    rDiario: TfrxReport;
    fechaAntes: tDate;
    sId, sFolio : string;
    CuentaCol : integer;

    dTotal       : double;
    dFechaActual : tDate;
begin
    dTotal := 0;
    dFechaActual := dParamFechaI;

    {$region 'Categorias'}
    Categorias := TzReadOnlyQuery.create(nil);
    Categorias.Connection := Connection.zConnection;
    Categorias.Active := False;
    Categorias.SQL.Clear;
    if (sParamTipo = 'Personal') or (sParamTipo = 'Equipo') then
    begin
        if sParamTipo = 'Personal' then
           Categorias.SQL.Add('select bp.sIdPersonal as sIdRecurso, bp.sTipoObra from bitacoradepersonal bp '+
                    'inner join personal p on (p.sContrato =:ContratoBarco and p.sIdPersonal  = bp.sIdPersonal) '+
                    'where bp.sContrato =:Contrato and bp.dIdFecha >=:FechaI and bp.dIdFecha <=:FechaF and bp.sIdPersonal like :Categoria '+
                    'group by bp.iItemOrden ')
        else
           Categorias.SQL.Add('select bp.sIdEquipo as sIdRecurso, bp.sTipoObra from bitacoradeequipos bp '+
                    'inner join equipos p on (p.sContrato =:ContratoBarco and p.sIdEquipo  = bp.sIdEquipo) '+
                    'where bp.sContrato =:contrato and bp.dIdFecha >=:FechaI and bp.dIdFecha <=:FechaF and bp.sIdEquipo like :categoria '+
                    'group by bp.iItemOrden ');
        Categorias.ParamByName('ContratoBarco').AsString := global_contrato_barco;
        Categorias.ParamByName('Contrato').AsString      := global_contrato;
        Categorias.ParamByName('FechaI').AsDate          := dParamFechaI;
        Categorias.ParamByName('FechaF').AsDate          := dParamFechaF;
    end;

    if sParamTipo = 'Barco' then
    begin
        Categorias.SQL.Add('select tm.sIdTipomovimiento as sIdRecurso, tm.sDescripcion as sTipoObra from tiposdemovimiento tm '+
                    'where tm.sContrato =:ContratoBarco and sClasificacion = "Movimiento de barco" and tm.sIdTipoMovimiento like :Categoria ');
        Categorias.ParamByName('ContratoBarco').AsString := global_contrato_barco;
    end;

    if sParamTipo = 'Pernocta' then
    begin
        Categorias.SQL.Add('select c.sIdCuenta as sIdRecurso, c.sDescripcion as sTipoObra from cuentas c  where c.sIdCuenta like :categoria ');
    end;

    if (sParamTipo = 'C6') or (sParamTipo = 'C7') or (sParamTipo = 'C8') then
    begin
        Categorias.SQL.Add('select aa.sNumeroActividad as sIdRecurso, substr(aa.mDescripcion,1,250) as sTipoObra from actividadesxanexo aa '+
                          'inner join bitacorademateriales ba on (ba.sContrato =:Contrato and ba.sIdMaterial = aa.sNumeroActividad and ba.dIdFecha >=:FechaI and ba.dIdFecha <=:FechaF) '+
                          'where aa.sContrato =:ContratoBarco and aa.sTipoActividad = "Actividad" and aa.sNumeroActividad like :categoria and aa.sAnexo =:Anexo group by aa.sWbs order by aa.iItemOrden');
        Categorias.ParamByName('ContratoBarco').AsString := global_contrato_barco;
        Categorias.ParamByName('Contrato').AsString      := global_contrato;
        Categorias.ParamByName('Anexo').AsString         := sParamTipo;
        Categorias.ParamByName('FechaI').AsDate          := dParamFechaI;
        Categorias.ParamByName('FechaF').AsDate          := dParamFechaF;
    end;

    categorias.params.ParamByName('categoria').DataType := ftString;
    if sParamCategoria = '<Todas>' then
       categorias.params.ParamByName('categoria').Value := '%'
    else
       categorias.params.ParamByName('categoria').Value := sParamCategoria;
    Categorias.Open;
    {$endregion}

    {$region 'Memorys Data'}
    rDiario := TfrxReport.Create(tOrigen);
    mDatos:=TJvMemoryData.Create(nil);
    with mDatos do
      begin
        Active:=false;
        FieldDefs.Add('sAnexo', ftString, 10, false);
        FieldDefs.Add('sIdRecurso', ftString, 100, True);
        FieldDefs.Add('sNumeroOrden', ftString, 50, True);
        FieldDefs.Add('sIdFolio', ftString, 100, True);
        FieldDefs.Add('sDescripcion', ftString, 1000, True);
        FieldDefs.Add('sMedida', ftString, 10, True);
        FieldDefs.Add('dcantTotal', FtFloat, 0, True);
        FieldDefs.Add('dSumSubTotal', FtFloat, 0, True);
        for CuentaCol:=1 to 31 do
        begin
          FieldDefs.Add('dCantidad' + Inttostr(CuentaCol), FtFloat, 0, false);
          FieldDefs.Add('iDia'+ Inttostr(CuentaCol), ftInteger, 0, false);
          FieldDefs.Add('dTotalDia' + Inttostr(CuentaCol), FtFloat, 0, false);
          FieldDefs.Add('dSubTotalDia' + Inttostr(CuentaCol), FtFloat, 0, false);
        end;
        Active:=true;
      end;

    mTotales:=TJvMemoryData.Create(nil);
    with mTotales do
      begin
        Active:=false;
        for CuentaCol:=1 to 31 do
        begin
          FieldDefs.Add('dTotal' + Inttostr(CuentaCol), FtFloat, 0, false);
          FieldDefs.Add('iDia'+ Inttostr(CuentaCol), ftInteger, 0, false);
        end;
        Active:=true;
      end;

    for CuentaCol := 1 to 31 do
    begin
        mTotales.Append;
        mTotales.FieldByName('dTotal'+ Inttostr(CuentaCol)).AsFloat := 0;
        mtotales.FieldByName('iDia'+ Inttostr(CuentaCol)).AsInteger := CuentaCol;
        mTotales.Post;
    end;

    mSubTotal:=TJvMemoryData.Create(nil);
    with mSubTotal do
      begin
        Active:=false;
        FieldDefs.Add('sIdRecurso', ftString, 100, True);
        for CuentaCol:=1 to 31 do
        begin
          FieldDefs.Add('iDia'+ Inttostr(CuentaCol), ftInteger, 0, false);
          FieldDefs.Add('dSubTotal' + Inttostr(CuentaCol), FtFloat, 0, false);
        end;
        Active:=true;
      end;

    Categorias.First;
    while not Categorias.Eof do
    begin
        mSubTotal.Append;
        mSubTotal.FieldByName('sIdRecurso').AsString := Categorias.FieldByName('sIdRecurso').AsString;
        for CuentaCol:=1 to 31 do
        begin
          mSubTotal.FieldByName('iDia'+ Inttostr(CuentaCol)).AsInteger := CuentaCol;
          mSubTotal.FieldByName('dSubTotal' + Inttostr(CuentaCol)).AsFloat := 0;
        end;
        mSubTotal.Post;
        Categorias.Next;
    end;

    {$endregion}

    {$REGION 'Consultas'}
    //Ahora vienen las consultas del generador..
    qry_tripulacion := TzReadOnlyQuery.create(nil);
    qry_tripulacion.Connection := Connection.zConnection;
    qry_tripulacion.Active := False;
    qry_tripulacion.SQL.Clear;
    if (sParamTipo = 'Personal') or (sParamTipo = 'Equipo') then
    begin
        //>>Personal<<
        if sParamTipo = 'Personal' then
           qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra,bp.sIdPersonal as sIdRecurso, bp.sDescripcion, '+
                          'concat(bp.sIdPersonal, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, e.sMedida, '+
                          'o.sIdFolio as sDescripcionFolio, '+
                          '(sum(xRound(bp.dCantHH, 2)) + ' +
                          '  IFNULL((SELECT ' +
                          '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                          '          FROM ' +
                          '            bitacoradeajustes ba ' +
                          '          WHERE ' +
                          '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                          '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                          '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as Total, '+
                          'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                          'mr.dCantidad as dSolicitado '+
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
                          'WHERE bp.scontrato =:Contrato AND bp.sNumeroOrden like :folio AND bp.dIdFecha >=:fechaI and bp.dIdFecha <=:fechaF '+
                          'and bp.sIdPersonal like :categoria GROUP BY bp.dIdFecha, bp.sNumeroOrden, bp.sIdPersonal,bp.sTipoObra order by e.iItemOrden, bp.sNumeroOrden, bp.dIdFecha ')
        else
        //>>Equipo<<
           qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, bp.sTipoObra,bp.sIdEquipo as sIdRecurso, bp.sDescripcion, '+
                          'concat(bp.sIdEquipo, " ", bp.sDescripcion) as sDescripcionRecurso, bp.sNumeroOrden, e.sMedida, '+
                          'o.sIdFolio as sDescripcionFolio, '+
                          '(sum(xround(bp.dCantHH, 2)) + ' +
                          '  IFNULL((SELECT ' +
                          '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                          '          FROM ' +
                          '            bitacoradeajustes ba ' +
                          '          WHERE ' +
                          '            ba.sContrato = bp.sContrato AND ba.sNumeroOrden = bp.sNumeroOrden AND ' +
                          '            ba.sNumeroActividad = bp.sNumeroActividad AND ' +
                          '            ba.sIdPartidaAnexo = bp.sIdPersonal AND ba.dFecha = bp.dIdFecha), 0)) as Total, '+
                          'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                          'mr.dCantidad as dSolicitado '+
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
                          'and bp.sIdEquipo like :categoria GROUP BY bp.dIdFecha, bp.sNumeroOrden, bp.sIdEquipo, bp.sTipoObra order by e.iItemOrden, bp.sNumeroOrden, bp.dIdFecha ');

    end;

    if sParamTipo = 'Barco' then
    begin
        qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, tp.sIdTipoMovimiento as sIdRecurso, tp.sDescripcion, '+
                        'concat(tp.sIdTipoMovimiento, " ", tp.sDescripcion) as sDescripcionRecurso, o.sNumeroOrden, " " as sMedida, '+
                        'o.sIdFolio as sDescripcionFolio, '+
                        'xround(sum(bp.sFactor),6) as Total, '+
                        'ifnull(SUM(bp.sFactorBarco),0) AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM movimientosxfolios bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sNumeroOrden and o.sNumeroOrden = bp.sFolio ) '+
                        'inner join movimientosdeembarcacion me on (me.sContrato =bp.sContrato and me.dIdFecha = bp.dIdFecha and me.iIdDiario = bp.iIddiario ) '+
                        'inner join tiposdemovimiento tp on (tp.sContrato = bp.sContrato and tp.sIdTipoMovimiento = me.sClasificacion and tp.sClasificacion = "Movimiento de barco") '+
                        'WHERE bp.scontrato =:ContratoBarco AND bp.sNumeroOrden =:Contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and me.sClasificacion like :categoria and bp.sfolio like :folio '+
                        'group by bp.dIdFecha, bp.sFolio, tp.sIdTipoMovimiento order by tp.sIdTipoMovimiento, o.iOrden, me.dIdFecha');

    end;

    if sParamTipo = 'Pernocta' then
    begin
       (* qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, bp.sIdCuenta as sIdRecurso, c.sDescripcion, '+
                        'concat(bp.sIdCuenta, " ", c.sDescripcion) as sDescripcionRecurso, o.sNumeroOrden, c.sMedida, '+
                        'o.sIdFolio as sDescripcionFolio, '+
                        'sum(bp.dCantidad) as Total,bpc.dCantidadc,bpc.Ajustec, '+
                        '0.00000 AS Ajuste, 0.00000 as dSolicitado '+
                        'FROM bitacoradepernocta bp '+
                        'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden ) '+
                        'inner join cuentas c on (c.sIdCuenta = bp.sIdCuenta) '+
                         'left join' + #13#10 + 
                          '(select bp2.sContrato,bp2.dIdFecha,sum( bp2.dCantHH) as dCantidadc, bp2.sNumeroOrden,bp2.stipopernocta,' + #13#10 +
                          'ifnull(SUM(bp2.dAjuste),0) AS Ajustec' + #13#10 + 
                          'from bitacoradepersonal bp2' + #13#10 + 
                          'where  bp2.lImprime = "Si" group by bp2.sContrato,bp2.sNumeroOrden,bp2.didfecha,bp2.stipopernocta ) as bpc' + #13#10 + 
                          'ON (bpc.sContrato =bp.sContrato' + #13#10 + 
                          'and bpc.dIdFecha =bp.didfecha' + #13#10 + 
                          'and bpc.sNumeroOrden=bp.sNumeroOrden and bpc.stipopernocta=c.sIdCuenta) '+
                        'WHERE bp.scontrato =:Contrato AND bp.didfecha >=:FechaI and bp.dIdFecha <=:FechaF '+
                        'and bp.sIdCuenta like :Categoria and bp.sNumeroOrden like :Folio and :ContratoBarco = :ContratoBarco '+
                        'group by bp.dIdFecha, bp.sNumeroOrden, bp.sIdCuenta order by c.sIdCuenta, o.iOrden, bp.dIdFecha ');*)



      qry_tripulacion.SQL.Add('SELECT bp.sContrato, bp.dIdFecha, bp.iIdDiario, "" as sTipoObra, bp.stipopernocta as sIdRecurso, c.sDescripcion,' + #13#10 +
                              'concat(bp.stipopernocta, " ", c.sDescripcion) as sDescripcionRecurso, o.sNumeroOrden, c.sMedida, ' + #13#10 +
                              'o.sIdFolio as sDescripcionFolio, ' + #13#10 +
                              'sum(bp.dCantidad) as Total,bpc.dCantidadc ,' + #13#10 +
                              'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS Ajuste, ' +
                              '0.00000 as dSolicitado ' + #13#10 +
                              'FROM' + #13#10 +
                              '( SELECT bp3.scontrato,bp3.didfecha, bp3.iIdDiario,bp3.sNumeroOrden,bp3.stipopernocta, ' + #13#10 +
                              '(sum(xround(bp3.dCanthh, 2)) + ' +
                              '  IFNULL((SELECT ' +
                              '            SUM(CAST(ba.dAjuste AS CHAR)) ' +
                              '          FROM ' +
                              '            bitacoradeajustes ba ' +
                              '          WHERE ' +
                              '            ba.sContrato = bp3.sContrato AND ba.sNumeroOrden = bp3.sNumeroOrden AND ' +
                              '            ba.sNumeroActividad = bp3.sNumeroActividad AND ' +
                              '            ba.sIdPartidaAnexo = bp3.sIdPersonal AND ba.dFecha = bp3.dIdFecha), 0)) as dCantidad, ' +
                              'IFNULL(SUM(CAST(bp.dAjuste AS CHAR)), 0) AS dAjuste' + #13#10 +
                              'FROM bitacoradepersonal bp3 ' + #13#10 +
                              'WHERE bp3.scontrato =:Contrato AND bp3.didfecha >=:FechaI and bp3.dIdFecha <=:FechaF ' + #13#10 +
                              'and bp3.stipopernocta like :Categoria and bp3.sNumeroOrden like :Folio' + #13#10 + 
                              'and bp3.lAplicaPernocta="Si" group by bp3.dIdFecha, bp3.sNumeroOrden, bp3.stipopernocta,bp3.sIdPersonal' + #13#10 + 
                              ') bp ' + #13#10 + 
                              '' + #13#10 +
                              'inner join ordenesdetrabajo o on (o.sContrato = bp.sContrato and o.sNumeroOrden = bp.sNumeroOrden )' + #13#10 + 
                              'inner join cuentas c on (c.sIdCuenta = bp.stipopernocta) ' + #13#10 + 
                              'left join      ' + #13#10 + 
                              '(select bp2.sContrato,bp2.sNumeroOrden,bp2.didfecha,bp2.sIdCuenta,sum( bp2.dCantidad) as dCantidadc' + #13#10 +
                              'from bitacoradepernocta bp2' + #13#10 + 
                              '' + #13#10 + 
                              'group by bp2.sContrato,bp2.sNumeroOrden,bp2.didfecha,bp2.sIdCuenta ) as bpc' + #13#10 +
                              'ON (bpc.sContrato =bp.sContrato' + #13#10 + 
                              'and bpc.dIdFecha =bp.didfecha' + #13#10 + 
                              'and bpc.sNumeroOrden=bp.sNumeroOrden and bpc.sIdCuenta=c.sIdCuenta)' + #13#10 + 
                              'where :ContratoBarco = :ContratoBarco' + #13#10 +
                              'group by bp.dIdFecha, bp.sNumeroOrden, bp.stipopernocta order by c.sIdCuenta, o.iOrden, bp.dIdFecha' );





    end;

    if (sParamTipo = 'C6') or (sParamTipo = 'C7') or (sParamTipo = 'C8') then
    begin
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
                        'group by bp.dIdFecha, bp.sNumeroOrden, bp.sIdMaterial order by ax.iItemOrden, o.iOrden, bp.dIdFecha');

         qry_tripulacion.params.ParamByName('Anexo').asString  := sParamTipo;
    end;


    qry_tripulacion.params.ParamByName('contratoBarco').asString  := global_contrato_barco;
    qry_tripulacion.params.ParamByName('contrato').asString       := sParamContrato;
    qry_tripulacion.params.ParamByName('fechaI').AsDate           := dParamFechaI;
    qry_tripulacion.params.ParamByName('fechaF').AsDate           := dParamFechaF;
    qry_tripulacion.params.ParamByName('categoria').DataType := ftString;
    if sParamCategoria = '<Todas>' then
       qry_tripulacion.params.ParamByName('categoria').Value := '%'
    else
       qry_tripulacion.params.ParamByName('categoria').Value := sParamCategoria;
    qry_tripulacion.params.ParamByName('folio').DataType := ftString;
    qry_tripulacion.params.ParamByName('folio').Value := '%';
    qry_tripulacion.Open;
    {$ENDREGION}

    {$REGION 'Proceso'}
    Categorias.First;
    while not Categorias.Eof do
    begin
       qry_tripulacion.Filtered := False;
       qry_tripulacion.Filter   := 'sIdRecurso = ' + QuotedStr( Categorias.FieldByName( 'sIdRecurso' ).AsString );
       qry_tripulacion.Filtered := True;
       qry_tripulacion.First;

       if qry_tripulacion.RecordCount = 0 then
       begin
           //Solo par pernocta y barco
           if (sParamTipo = 'Barco') or (sParamTipo = 'Pernocta') then
           begin
               qry_tripulacion.Filtered := False;
               sId := '';
               while not qry_tripulacion.Eof do
               begin
                   if sId <> qry_tripulacion.FieldByName('sNumeroOrden').AsString then
                   begin
                       sId := qry_tripulacion.FieldByName('sNumeroOrden').AsString;

                       mDatos.Append;
                       mDatos.FieldByName('sAnexo').AsString       := sParamtipo;
                       mDatos.FieldByName('sIdRecurso').AsString   := categorias.FieldByName('sIdRecurso').AsString;
                       mDatos.FieldByName('sNumeroOrden').AsString := qry_tripulacion.FieldByName('sNumeroOrden').AsString;
                       mDatos.FieldByName('sIdFolio').AsString     := qry_tripulacion.FieldByName('sDescripcionFolio').AsString;
                       mDatos.FieldByName('sDescripcion').AsString := qry_tripulacion.FieldByName('sDescripcion').AsString;
                       mDatos.FieldByName('sMedida').AsString      := qry_tripulacion.FieldByName('sMedida').AsString;
                       mDatos.FieldByName('dSumSubTotal').AsFloat  := 0;
                       mDatos.FieldByName('dcantTotal').AsFloat    := 0;
                       mDatos.Post;
                   end;
                   qry_tripulacion.Next;
               end;
           end;
       end
       else
       begin

           sId := '';
           while not qry_tripulacion.Eof do
           begin
               if sId <> (qry_tripulacion.FieldByName('sNumeroOrden').AsString +'-'+ qry_tripulacion.FieldByName('sIdRecurso').AsString) then
               begin
                   mSubTotal.Filtered := False;
                   mSubTotal.Filter   := 'sIdRecurso = '+QuotedStr(qry_tripulacion.FieldByName('sIdRecurso').AsString);
                   mSubTotal.Filtered := True;

                   dFechaActual := dParamFechaI;
                   CuentaCol := StrToInt(copy(DateToStr(dFechaActual),1,2));
                   dTotal    := 0;
                   sId    :=  qry_tripulacion.FieldByName('sNumeroOrden').AsString +'-'+ qry_tripulacion.FieldByName('sIdRecurso').AsString;
                   sFolio :=  qry_tripulacion.FieldByName('sNumeroOrden').AsString;
                   mDatos.Append;
                   mDatos.FieldByName('sAnexo').AsString  := sParamtipo;
                   mDatos.FieldByName('sIdRecurso').AsString   := qry_tripulacion.FieldByName('sIdRecurso').AsString;
                   mDatos.FieldByName('sNumeroOrden').AsString := qry_tripulacion.FieldByName('sNumeroOrden').AsString;
                   mDatos.FieldByName('sIdFolio').AsString     := qry_tripulacion.FieldByName('sDescripcionFolio').AsString;
                   mDatos.FieldByName('sDescripcion').AsString := qry_tripulacion.FieldByName('sDescripcion').AsString;
                   mDatos.FieldByName('sMedida').AsString      := qry_tripulacion.FieldByName('sMedida').AsString;
                   mDatos.FieldByName('dSumSubTotal').AsFloat  := 0;
               end;

               while (dFechaActual <= dParamFechaF) and (sFolio = qry_tripulacion.FieldByName('sNumeroOrden').AsString) do
               begin
                   if qry_tripulacion.FieldByName('dIdFecha').AsDateTime = dFechaActual then
                   begin
                     if sParamTipo = 'Pernocta' then
                       mDatos.FieldByName('dCantidad'+ Inttostr(CuentaCol)).AsFloat := (qry_tripulacion.FieldByName('Total').AsFloat + qry_tripulacion.FieldByName('Ajuste').AsFloat) +
                                                                                        qry_tripulacion.FieldByName('dCantidadc').AsFloat
                     else
                      mDatos.FieldByName('dCantidad'+ Inttostr(CuentaCol)).AsFloat := mDatos.FieldByName('dCantidad'+ Inttostr(CuentaCol)).AsFloat + (qry_tripulacion.FieldByName('Total').AsFloat + qry_tripulacion.FieldByName('Ajuste').AsFloat);

                       mDatos.FieldByName('iDia'+ Inttostr(CuentaCol)).AsInteger    := CuentaCol;

                       if sParamTipo = 'Pernocta' then
                         dTotal := dTotal + (qry_tripulacion.FieldByName('Total').AsFloat + qry_tripulacion.FieldByName('Ajuste').AsFloat) +
                                   qry_tripulacion.FieldByName('dCantidadc').AsFloat
                       else
                        dTotal := dTotal + (qry_tripulacion.FieldByName('Total').AsFloat + qry_tripulacion.FieldByName('Ajuste').AsFloat);




                       mTotales.Locate('iDia'+IntToStr(CuentaCol),CuentaCol,[]);
                       mTotales.Edit;
                       if sParamTipo = 'Pernocta' then
                        mTotales.FieldByName('dTotal'+IntToStr(CuentaCol)).AsFloat := mTotales.FieldByName('dTotal'+IntToStr(CuentaCol)).AsFloat + (qry_tripulacion.FieldByName('Total').AsFloat + qry_tripulacion.FieldByName('Ajuste').AsFloat) + qry_tripulacion.FieldByName('dCantidadc').AsFloat
                       else
                        mTotales.FieldByName('dTotal'+IntToStr(CuentaCol)).AsFloat := mTotales.FieldByName('dTotal'+IntToStr(CuentaCol)).AsFloat + (qry_tripulacion.FieldByName('Total').AsFloat + qry_tripulacion.FieldByName('Ajuste').AsFloat);
                       mTotales.Post;

                       mSubTotal.Locate('iDia'+IntToStr(CuentaCol),CuentaCol, [] );
                       mSubTotal.Edit;
                       if sParamTipo = 'Pernocta' then
                        mSubTotal.FieldByName('dSubTotal'+IntToStr(CuentaCol)).AsFloat := mSubTotal.FieldByName('dSubTotal'+IntToStr(CuentaCol)).AsFloat + ((qry_tripulacion.FieldByName('Total').AsFloat + qry_tripulacion.FieldByName('Ajuste').AsFloat) + qry_tripulacion.FieldByName('dCantidadc').AsFloat)
                       else
                        mSubTotal.FieldByName('dSubTotal'+IntToStr(CuentaCol)).AsFloat := mSubTotal.FieldByName('dSubTotal'+IntToStr(CuentaCol)).AsFloat + (qry_tripulacion.FieldByName('Total').AsFloat + qry_tripulacion.FieldByName('Ajuste').AsFloat);
                       mSubTotal.Post;

                       qry_tripulacion.Next;
                   end;
                   if (dFechaActual <> qry_tripulacion.FieldByName('dIdFecha').AsDateTime) or (qry_tripulacion.Eof) then
                   begin
                       inc(CuentaCol);
                       dFechaActual := dFechaActual + 1;
                   end;
               end;
               mDatos.FieldByName('dcantTotal').AsFloat := dTotal;
               mDatos.Post;
           end;
       end;
       Categorias.Next;
    end;

    mSubTotal.Filtered := False;
    mSubtotal.First;
    sId := '';
    while not mSubTotal.Eof do
    begin
        if sId <> mSubTotal.FieldByName('sIdRecurso').AsString then
        begin
            dTotal := 0;
            sId := mSubTotal.FieldByName('sIdRecurso').AsString;
            mDatos.Filtered := False;
            mDatos.Filter   := 'sIdRecurso = '+QuotedStr(mSubTotal.FieldByName('sIdRecurso').AsString);
            mDatos.Filtered := True;
            mDatos.First;
        end;
        if mDatos.RecordCount > 0 then
        for cuentaCol := 1 to 31 do
        begin
            dTotal := dTotal +  mSubTotal.FieldByName('dSubTotal'+IntToStr(cuentaCol)).AsFloat;
            mDatos.Edit;
            mDatos.FieldByName('dSubTotalDia'+IntToStr(cuentaCol)).AsFloat := mSubTotal.FieldByName('dSubTotal'+IntToStr(cuentaCol)).AsFloat;
            mDatos.FieldByName('dSumSubTotal').AsFloat := dTotal;
            mDatos.Post;
        end;

        mSubTotal.Next;
    end;
    mDatos.Filtered := False;

    dTotal := 0;
    mDatos.Last;
    mTotales.First;
    for cuentaCol := 1 to 31 do
    begin
        dTotal := dTotal + mTotales.FieldByName('dTotal'+IntToStr(cuentaCol)).AsFloat;
        mDatos.Edit;
        mDatos.FieldByName('dTotalDia'+IntToStr(cuentaCol)).AsFloat := mTotales.FieldByName('dTotal'+IntToStr(cuentaCol)).AsFloat;
        mDatos.FieldByName('dSumSubTotal').AsFloat := dTotal;
        mDatos.Post;
        mTotales.Next;
    end;
    {$ENDREGION}

    qryConfiguracion := TZReadOnlyQuery.Create(nil);
    qryConfiguracion.Connection := connection.zConnection;
    dsConfiguracion := TfrxDBDataSet.Create(tOrigen);
    dsConfiguracion.DataSet := QryConfiguracion;
    dsConfiguracion.UserName := 'dsConfiguracion';
    QryConfiguracion.Active := False;
    QryConfiguracion.SQL.Clear;
    QryConfiguracion.SQL.Add('select c2.sCodigo, c.iFirmas,c.sOrdenPerEq, c.sTipoPartida, c.sImprimePEP, c.sReportesCIA, c.sDireccion1, c.sDireccion2,  ' +
                    'c.sClaveSeguridad, c.cStatusProceso, c.sOrdenExtraordinaria, c.lLicencia, "" as sIdUsuarioValida, "" as sIdUsuarioResidente, "" as sIdUsuarioAutoriza, ' +
                    'c.bImagen, c.sContrato, c.sNombre, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, ' +
                    'c.lImprimeExtraordinario, c.sIdEmbarcacion, c.iRedondeoEmbarcacion, c.iFirmasBarco, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3,' +
                    'c2.mDescripcion, c2.mCliente, c2.bImagen as bImagenPEP, c3.dFechaInicio, c3.dFechaFinal ' +
                    'From contratos c2 INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
                    'inner join convenios c3 on (c2.sContrato = c3.sContrato and c.sIdConvenio = c3.sIdConvenio) ' +
                    'Where c2.sContrato = :Contrato');
    QryConfiguracion.Params.ParamByName('contrato').DataType := ftString;
    QryConfiguracion.Params.ParamByName('contrato').Value    := sParamContrato;
    QryConfiguracion.Open;

    dsGeneradorFolio:= TfrxDBDataSet.Create(Nil);
    dsGeneradorFolio.UserName   :='dsGeneradorFolio';
    dsGeneradorFolio.DataSet    := mDatos;
    dsGeneradorFolio.FieldAliases.Clear;

    rDiario.DataSets.Add(dsGeneradorFolio);
    rDiario.DataSets.Add(dsConfiguracion);

    rDiario.OnGetValue := tProcedure;
    {Pemdiente revision_ivan}
    rDiarioFirmas(sParamContrato, '', 'A', dParamFechaF, tOrigen);

    if lParamContinuo then
       rDiario.LoadFromFile(global_files + global_miReporte+'_cGeneradorFolio.fr3')
    else
       rDiario.LoadFromFile(global_files + global_miReporte+'_GeneradorFolio.fr3');
    rDiario.ShowReport;    //(FormatosExp, PermisosExp);


    // Destruye Objetos ...

   // rDiario.Destroy;
    dsGeneradorFolio.Destroy;
    qry_tripulacion.Destroy;
    mDatos.Destroy;
    mTotales.Destroy;
    mSubTotal.Destroy;
    categorias.Destroy;
end;


end.
