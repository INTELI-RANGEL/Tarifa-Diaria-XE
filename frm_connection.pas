unit frm_connection;

interface

uses
  SysUtils, Classes, DB, ADODB, frxExportMail, frxExportCSV, frxExportText,
  frxExportImage, frxExportPDF, frxExportXML, frxExportRTF, frxExportXLS,
  frxExportHTML, frxClass, frxExportTXT, frxDBSet, ImgList, Menus,
  ActnList, Controls, fqbClass, ZAbstractRODataset,
  ZDataset, ZConnection, ZAbstractDataset,  
  frxRich, frxGZip, frxDMPExport, global, ExtCtrls,
  IdMessage, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL,
  IdSSLOpenSSL, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient,
  IdExplicitTLSClientServerBase, IdMessageClient, IdSMTPBase, IdSMTP, cxGraphics,
  ZAbstractConnection;
type
  Tconnection = class(TDataModule)
    ds_setup: TDataSource;
    frxHTMLExport1: TfrxHTMLExport;
    frxXLSExport1: TfrxXLSExport;
    frxXMLExport1: TfrxXMLExport;
    frxRTFExport1: TfrxRTFExport;
    frxBMPExport1: TfrxBMPExport;
    frxJPEGExport1: TfrxJPEGExport;
    frxTIFFExport1: TfrxTIFFExport;
    frxPDFExport1: TfrxPDFExport;
    frxGIFExport1: TfrxGIFExport;
    a: TfrxSimpleTextExport;
    frxCSVExport1: TfrxCSVExport;
    frxMailExport1: TfrxMailExport;
    ImageList1: TImageList;
    QryBusca: TZReadOnlyQuery;
    zCommand: TZQuery;
    ds_ContratosxUsuario: TDataSource;
    ContratosxUsuario: TZReadOnlyQuery;
    rpt_contrato: TfrxDBDataset;
    rpt_setup: TfrxDBDataset;
    contrato: TZReadOnlyQuery;
    configuracion: TZReadOnlyQuery;
    ds_estimacionperiodo: TDataSource;
    EstimacionPeriodo: TZQuery;
    QryBusca2: TZReadOnlyQuery;
    UsuariosxPrograma: TZReadOnlyQuery;
    GruposxPrograma: TZReadOnlyQuery;
    zConnection: TZConnection;
    frxRichObject1: TfrxRichObject;
    frxDotMatrixExport1: TfrxDotMatrixExport;
    frxGZipCompressor1: TfrxGZipCompressor;
    qryROProrrateos: TZReadOnlyQuery;
    frxReport1: TfrxReport;
    Auxiliar: TZReadOnlyQuery;
    IdSMTP: TIdSMTP;
    IdSSLIOHandlerSocketOpenSSL: TIdSSLIOHandlerSocketOpenSSL;
    IdMessage: TIdMessage;
    ConnTrx: TZConnection;
    qryBuscaTrx: TZReadOnlyQuery;
    CommandTrx: TZQuery;
    ds_contrato: TDataSource;
    icnDevExpress32: TcxImageList;
    icnDevExpress16: TcxImageList;
    icnDevExpress24: TcxImageList;
    ZSentencia: TZReadOnlyQuery;
    //configuracionsSeccionImprime: TStringField;
    procedure rDiarioGetValue(const VarName: string; var Value: Variant);
    procedure frxReport1GetValue(const VarName: string; var Value: Variant);
    procedure zConnectionAfterConnect(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  connection: Tconnection;

implementation

{$R *.dfm}

procedure Tconnection.frxReport1GetValue(const VarName: string;
  var Value: Variant);
begin
  If CompareText(VarName, 'ORDEN') = 0 then
      Value := 'DE LA ORDEN DE TRABAJO ' + global_orden ;

  If CompareText(VarName, 'FECHA_INICIO') = 0 then
      Value := global_fecha  ;

  If CompareText(VarName, 'FECHA_FINAL') = 0 then
      Value := global_fecha  ;

  If CompareText(VarName, 'DESCRIPCION_CORTA') = 0 then
      Value := sDiarioDescripcionCorta ;

  If CompareText(VarName, 'IMPRIME_AVANCES') = 0 then
      Value := sDiarioComentario ;

  If CompareText(VarName, 'sNewTexto') = 0 then
      Value := sDiarioTitulo ;

  If CompareText(VarName, 'PERIODO') = 0 then
      Value := sDiarioPeriodo ;


  If CompareText(VarName, 'SUPERINTENDENTE') = 0 then
      If global_sOrigen_reporte = 'No' Then
          Value := sSuperIntendente
      Else
          Value := sSuperIntendentePatio ;

  If CompareText(VarName, 'SUPERVISOR') = 0 then
      If global_sOrigen_reporte = 'No' Then
          Value := sSupervisor
      Else
          Value := sSupervisorPatio ;

  If CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
      If global_sOrigen_reporte = 'No' Then
          Value := sSupervisorTierra
      Else
          Value := sResidente ;

  If CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
      If global_sOrigen_reporte = 'No' Then
          Value := sPuestoSuperIntendente
      Else
          Value := sPuestoSuperIntendentePatio ;

  If CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
      If global_sOrigen_reporte = 'No' Then
          Value := sPuestoSupervisor
      Else
          Value := sPuestoSupervisorPatio ;

  If CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
      If global_sOrigen_reporte = 'No' Then
          Value := sPuestoSupervisorTierra
      Else
          Value := sPuestoResidente ;


  If CompareText(VarName, 'DESCRIPCION_ORDEN') = 0 then
      Value := mDescripcionOrden  ;
  If CompareText(VarName, 'PLATAFORMA') = 0 then
      Value := sPlataformaOrden  ;

  If CompareText(VarName, 'JORNADAS_SUSPENDIDAS') = 0 then
      Value := sJornadasSuspendidas  ;

  If CompareText(VarName, 'TURNO') = 0 then
      Value := sDescripcionTurno ;
                
  If CompareText(VarName, 'REAL_ANTERIOR') = 0 then
      Value := dRealGlobalAnterior ;
  If CompareText(VarName, 'REAL_ACTUAL') = 0 then
      Value := dRealGlobalActual ;
  If CompareText(VarName, 'REAL_ACUMULADO') = 0 then
      Value := dRealGlobalAcumulado ;
  If CompareText(VarName, 'PROGRAMADO_ANTERIOR') = 0 then
      Value := dProgramadoGlobalAnterior ;
  If CompareText(VarName, 'PROGRAMADO_ACTUAL') = 0 then
      Value := dProgramadoGlobalActual ;
  If CompareText(VarName, 'PROGRAMADO_ACUMULADO') = 0 then
      Value := dProgramadoGlobalAcumulado;


  If CompareText(VarName, 'REAL_ANTERIOR_MULTIPLE') = 0 then
      Value := dRealOrdenAnterior ;
  If CompareText(VarName, 'REAL_ACTUAL_MULTIPLE') = 0 then
      Value := dRealOrdenActual ;
  If CompareText(VarName, 'REAL_ACUMULADO_MULTIPLE') = 0 then
      Value := dRealOrdenAcumulado ;
  If CompareText(VarName, 'PROGRAMADO_ANTERIOR_MULTIPLE') = 0 then
      Value := dProgramadoOrdenAnterior ;
  If CompareText(VarName, 'PROGRAMADO_ACTUAL_MULTIPLE') = 0 then
      Value := dProgramadoOrdenActual ;
  If CompareText(VarName, 'PROGRAMADO_ACUMULADO_MULTIPLE') = 0 then
      Value := dProgramadoOrdenAcumulado ;

end;

procedure Tconnection.rDiarioGetValue(const VarName: string;
  var Value: Variant);
begin
  if CompareText(VarName, 'ORDEN') = 0 then
    Value := 'DE LA ORDEN DE TRABAJO ' + global_orden;

  if CompareText(VarName, 'FECHA_INICIO') = 0 then
    Value := global_fecha;

  if CompareText(VarName, 'FECHA_FINAL') = 0 then
    Value := global_fecha;

  if CompareText(VarName, 'DESCRIPCION_CORTA') = 0 then
    Value := sDiarioDescripcionCorta;

  if CompareText(VarName, 'IMPRIME_AVANCES') = 0 then
    Value := sDiarioComentario;

  if CompareText(VarName, 'sNewTexto') = 0 then
    Value := sDiarioTitulo;

  if CompareText(VarName, 'PERIODO') = 0 then
    Value := sDiarioPeriodo;
  if CompareText(VarName, 'SUPERINTENDENTE') = 0 then
    if global_sOrigen_reporte = 'No' then
      Value := sSuperIntendente
    else
      Value := sSuperIntendentePatio;

  if CompareText(VarName, 'SUPERVISOR') = 0 then
    if global_sOrigen_reporte = 'No' then
      Value := sSupervisor
    else
      Value := sSupervisorPatio;

  if CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
    if global_sOrigen_reporte = 'No' then
      Value := sSupervisorTierra
    else
      Value := sResidente;

  if CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
    if global_sOrigen_reporte= 'No' then
      Value := sPuestoSuperIntendente
    else
      Value := sPuestoSuperIntendentePatio;

  if CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
    if global_sOrigen_reporte = 'No' then
      Value := sPuestoSupervisor
    else
      Value := sPuestoSupervisorPatio;

  if CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
    if global_sOrigen_reporte= 'No' then
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

procedure Tconnection.zConnectionAfterConnect(Sender: TObject);
begin
  // Connectar la base de datos alterna con los datos de la base original
  if ConnTrx.Connected then
    ConnTrx.Disconnect;
  ConnTrx.Catalog  := zConnection.Catalog;
  ConnTrx.Database := zConnection.DataBase;
  ConnTrx.HostName := zConnection.HostName;
  ConnTrx.PassWord := zConnection.PassWord;
  ConnTrx.Port     := zConnection.Port;
  ConnTrx.User     := zConnection.User;
  ConnTrx.Connect;
end;

end.









