unit frm_RequisicionPerf;

interface

uses
  Windows, Messages, Math, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, DB, ADODB, DBCtrls, global, strUtils,
  Mask, OleCtrls, Grids, DBGrids, frm_barra, ExtCtrls, Utilerias,
  Menus, frxClass, frxDBSet,  RXDBCtrl,  unitactivapop,
  RXCtrls, CheckLst, RxMemDS, ExcelXP, OleServer, ComObj,
  Excel2000, zDataset, ZAbstractRODataset, ZAbstractDataset, rxCurrEdit,
  rxToolEdit, RxRichEd, NxCollection, AdvGlowButton, udbgrid,
  unitexcepciones, unittbotonespermisos,UnitValidaTexto, UFunctionsGHH,
  UnitValidacion, JvExControls, JvDBLookup;

type
  TfrmRequisicionPerf = class(TForm)
    ds_anexo_requisicion: TDataSource;
    ds_anexo_prequisicion: TDataSource;
    frxDBReporte: TfrxDBDataset;
    GroupBox3: TGroupBox;
    frmBarra2: TfrmBarra;
    Grid_Entradas: TRxDBGrid;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    N4: TMenuItem;
    ComentariosAdicionales: TMenuItem;
    OpenXLS: TOpenDialog;
    ExcelApplication1: TExcelApplication;
    ExcelWorkbook1: TExcelWorkbook;
    ExcelWorksheet1: TExcelWorksheet;
    anexo_prequisicion: TZQuery;
    anexo_requisicion: TZQuery;
    Reporte: TZReadOnlyQuery;
    Maxfolio: TZQuery;
    ReportedAcumulado: TFloatField;
    ReportesContrato: TStringField;
    ReporteiFolioRequisicion: TIntegerField;
    ReportesNumeroOrden: TStringField;
    ReportedIdFecha: TDateField;
    ReportesReferencia: TStringField;
    ReportemComentarios: TMemoField;
    ReporteiItem: TIntegerField;
    ReportedCantidad: TFloatField;
    ReportemDescripcion: TMemoField;
    ReportesMedida: TStringField;
    ReportedFechaRequerimiento: TDateField;
    dtsPartidas: TDataSource;
    Partidas: TZReadOnlyQuery;
    PgControl: TPageControl;
    TabSheet1: TTabSheet;
    pgPartidas: TTabSheet;
    tsPlataforma: TLabel;
    imgNotas: TImage;
    frmBarra1: TfrmBarra;
    consumibles: TZReadOnlyQuery;
    dtConsumibles: TDataSource;
    anexo_prequisicionsContrato: TStringField;
    anexo_prequisicioniFolioRequisicion: TIntegerField;
    anexo_prequisicionsIdInsumo: TStringField;
    anexo_prequisicioniItem: TIntegerField;
    anexo_prequisicionmDescripcion: TMemoField;
    anexo_prequisicionsMedida: TStringField;
    anexo_prequisiciondFechaRequerimiento: TDateField;
    anexo_prequisiciondCantidad: TFloatField;
    anexo_prequisiciondCosto: TFloatField;
    anexo_prequisicionsNumeroActividad: TStringField;
    anexo_prequisicionsNumeroOrden: TStringField;
    anexo_prequisicionsMedidaAnexo: TStringField;
    anexo_prequisiciondCantidad_1: TFloatField;
    anexo_prequisicionsDescripcion: TStringField;
    anexo_prequisiciondMontoMN: TFloatField;
    ReportesNumeroActividad: TStringField;
    ReportedCosto: TFloatField;
    frxEntrada: TfrxReport;
    ReportedCostoMN: TFloatField;
    consumiblessDescripcion: TStringField;
    consumiblessIdInsumo: TStringField;
    consumiblesmDescripcion: TMemoField;
    consumiblesdCantidad: TFloatField;
    consumiblessMedida: TStringField;
    consumiblesdCostoMN: TFloatField;
    consumiblesdExistencia: TFloatField;
    Departamentos: TZReadOnlyQuery;
    ds_departamentos: TDataSource;
    NxHeaderPanel1: TNxHeaderPanel;
    Label7: TLabel;
    tiFolio: TCurrencyEdit;
    tdIdFecha: TDateTimePicker;
    Label3: TLabel;
    Label17: TLabel;
    Label8: TLabel;
    dbDepartamentos: TDBLookupComboBox;
    tsSolicitante: TEdit;
    Label9: TLabel;
    Label6: TLabel;
    tmComentarios: TMemo;
    NxHeaderPanel2: TNxHeaderPanel;
    lblPartida: TLabel;
    dbPartidas: TDBLookupComboBox;
    dbMatPart: TDBGrid;
    Label14: TLabel;
    tdCantidad: TRxCalcEdit;
    Label1: TLabel;
    tdFechaRequerimiento: TDateTimePicker;
    GridPartidas: TRxDBGrid;
    tsNumeroActividad: TLabel;
    lblBusca: TLabel;
    Label12: TLabel;
    dbFiltra: TComboBox;
    cmbRequisita: TDBComboBox;
    tsNumeroOrden: TDBComboBox;
    N5: TMenuItem;
    Imprimir1: TMenuItem;
    ImprimirConImportes1: TMenuItem;
    SeguimientoMaterialxPartida1: TMenuItem;
    SeguimientoMaterialxPartidasdeAnexo1: TMenuItem;
    ImprimirSeguimientoMaterial1: TMenuItem;
    N6: TMenuItem;
    rxSeguimiento_Mat: TRxMemoryData;
    rxSeguimiento_MatsContrato: TStringField;
    rxSeguimiento_MatsNumeroActividad: TStringField;
    rxSeguimiento_MatDescripcionAnexo: TStringField;
    rxSeguimiento_MatCantidadAnexo: TFloatField;
    rxSeguimiento_MatMedidaAnexo: TStringField;
    rxSeguimiento_MatCostoMNAnexo: TFloatField;
    rxSeguimiento_MatCostoDLLAnexo: TFloatField;
    rxSeguimiento_MatTipo: TStringField;
    rxSeguimiento_MatId: TStringField;
    rxSeguimiento_MatDescripcion: TStringField;
    rxSeguimiento_MatUnidad: TStringField;
    rxSeguimiento_MatCantidad: TFloatField;
    rxSeguimiento_MatCostoMN: TFloatField;
    rxSeguimiento_MatCostoDLL: TFloatField;
    frxSeguimiento_Mat: TfrxDBDataset;
    rxSeguimiento_MatdCantidadReq: TFloatField;
    rxSeguimiento_MatFolioReq: TIntegerField;
    rxSeguimiento_MatItemReq: TIntegerField;
    rxSeguimiento_MatPartida: TStringField;
    rxSeguimiento_MatdRestanteReq: TFloatField;
    rxSeguimiento_MatdExcedenteReq: TFloatField;
    rxSeguimiento_MatdPorcentajeReq: TFloatField;
    rxSeguimiento_MatFolioOC: TIntegerField;
    rxSeguimiento_MatItemOC: TIntegerField;
    rxSeguimiento_MatdCantidadOC: TFloatField;
    rxSeguimiento_MatdRestanteOC: TFloatField;
    rxSeguimiento_MatdExcedenteOC: TFloatField;
    rxSeguimiento_MatdPorcentajeOC: TFloatField;
    rxSeguimiento_MatFolioIn: TIntegerField;
    rxSeguimiento_MatItemIn: TIntegerField;
    rxSeguimiento_MatdCantidadIn: TFloatField;
    rxSeguimiento_MatdRestanteIn: TFloatField;
    rxSeguimiento_MatdExcedenteIn: TFloatField;
    rxSeguimiento_MatdPorcentajeIn: TFloatField;
    rxSeguimiento_MatFolioOut: TIntegerField;
    rxSeguimiento_MatItemOut: TIntegerField;
    rxSeguimiento_MatdCantidadOut: TFloatField;
    rxSeguimiento_MatdRestanteOut: TFloatField;
    rxSeguimiento_MatdExcedenteOut: TFloatField;
    rxSeguimiento_MatdPorcentajeOut: TFloatField;
    rxSeguimiento_MatdCantidadRD: TFloatField;
    rxSeguimiento_MatdRestanteRD: TFloatField;
    rxSeguimiento_MatdExcedenteRD: TFloatField;
    rxSeguimiento_MatdPorcentajeRD: TFloatField;
    rxSeguimiento_MatNumeroReporte: TStringField;
    rxSeguimiento_MatFechaRD: TDateField;
    rxSeguimiento_MatFrenteRD: TStringField;
    rxSeguimiento_MatiNumeroEstimacion: TIntegerField;
    rxSeguimiento_MatsNumeroOrden: TStringField;
    rxSeguimiento_MatsNumeroGenerador: TStringField;
    rxSeguimiento_MatdCantidadGen: TFloatField;
    rxSeguimiento_MatdExcedenteGen: TFloatField;
    rxSeguimiento_MatdRestanteGen: TFloatField;
    rxSeguimiento_MatdPorcentajeGen: TFloatField;
    rxSeguimiento_MatdPorcentajeReq_T: TFloatField;
    rxSeguimiento_MatdPorcentajeOC_T: TFloatField;
    rxSeguimiento_MatdPorcentajeRD_T: TFloatField;
    rxSeguimiento_Mat3: TRxMemoryData;
    StringField7: TStringField;
    FloatField4: TFloatField;
    IntegerField5: TIntegerField;
    IntegerField6: TIntegerField;
    FloatField17: TFloatField;
    FloatField18: TFloatField;
    FloatField19: TFloatField;
    FloatField20: TFloatField;
    frxSeguimiento_Mat2: TfrxDBDataset;
    rxSeguimiento_Mat2: TRxMemoryData;
    StringField20: TStringField;
    FloatField37: TFloatField;
    IntegerField12: TIntegerField;
    IntegerField13: TIntegerField;
    FloatField45: TFloatField;
    FloatField46: TFloatField;
    FloatField47: TFloatField;
    FloatField48: TFloatField;
    FloatField49: TFloatField;
    frxSeguimiento_Mat3: TfrxDBDataset;
    rxSeguimiento_Mat4: TRxMemoryData;
    StringField8: TStringField;
    FloatField5: TFloatField;
    IntegerField9: TIntegerField;
    IntegerField10: TIntegerField;
    FloatField26: TFloatField;
    FloatField27: TFloatField;
    FloatField28: TFloatField;
    FloatField29: TFloatField;
    frxSeguimiento_Mat4: TfrxDBDataset;
    rxSeguimiento_Mat5: TRxMemoryData;
    StringField22: TStringField;
    FloatField43: TFloatField;
    DateField2: TDateField;
    StringField26: TStringField;
    FloatField69: TFloatField;
    FloatField70: TFloatField;
    FloatField71: TFloatField;
    FloatField72: TFloatField;
    FloatField73: TFloatField;
    frxSeguimiento_Mat5: TfrxDBDataset;
    rxSeguimiento_Mat6: TRxMemoryData;
    StringField35: TStringField;
    FloatField81: TFloatField;
    IntegerField31: TIntegerField;
    StringField40: TStringField;
    StringField41: TStringField;
    FloatField107: TFloatField;
    FloatField108: TFloatField;
    FloatField109: TFloatField;
    FloatField110: TFloatField;
    frxSeguimiento_Mat6: TfrxDBDataset;
    rxSeguimiento_Mat6CantidadAnexo: TFloatField;
    GroupBox1: TGroupBox;
    frmBarra3: TfrmBarra;
    RxDBGrid1: TRxDBGrid;
    frxSeguimiento_Mat1: TfrxDBDataset;
    rxSeguimiento_Mat1: TRxMemoryData;
    StringField9: TStringField;
    FloatField6: TFloatField;
    IntegerField1: TIntegerField;
    IntegerField2: TIntegerField;
    FloatField9: TFloatField;
    FloatField10: TFloatField;
    FloatField11: TFloatField;
    FloatField12: TFloatField;
    FloatField13: TFloatField;
    rxSeguimiento_Mat1dCantidadReq_T: TFloatField;
    rxSeguimiento_Mat1dRestanteReq_T: TFloatField;
    rxSeguimiento_Mat1dExcedenteReq_T: TFloatField;
    rxSeguimiento_Mat1Unidad: TStringField;
    rxSeguimiento_Mat2dCantidadOC_T: TFloatField;
    rxSeguimiento_Mat2dRestanteOC_T: TFloatField;
    rxSeguimiento_Mat2dExcedenteOC_T: TFloatField;
    rxSeguimiento_Mat3dCantidadIn_T: TFloatField;
    rxSeguimiento_Mat3dExcedenteIn_T: TFloatField;
    rxSeguimiento_Mat4dCantidadOut_T: TFloatField;
    rxSeguimiento_Mat4dExcedenteOut_T: TFloatField;
    rxSeguimiento_Mat5dCantidadRD_T: TFloatField;
    rxSeguimiento_Mat5dRestanteRD_T: TFloatField;
    rxSeguimiento_Mat5dExcedenteRD_T: TFloatField;
    Registrar1: TMenuItem;
    Cancelar1: TMenuItem;
    Label16: TLabel;
    tdIdFecha2: TDateTimePicker;
    Label18: TLabel;
    tsLugarEntrega: TJvDBLookupCombo;
    QPernoctan: TZQuery;
    DsPernoctan: TDataSource;
    ReportedFechaSolicitado: TDateField;
    ReportedFechaRequerido: TDateField;
    ReportesRequisita: TStringField;
    ReportesRevision: TStringField;
    ReportesSolicito: TStringField;
    ReportesStatus: TStringField;
    ReportesAutorizo: TStringField;
    ReportesVerificacion: TStringField;
    ReportesRecibido: TStringField;
    ReportesidDepartamento: TStringField;
    ReportesMotivo: TStringField;
    ReportesEstado: TStringField;
    ReportesLugarEntrega: TStringField;
    ReporteiItemOrden: TStringField;
    Reporteesi: TStringField;
    Reporteeno: TStringField;
    Reportedepartamento: TStringField;
    Reportelugarentrega: TStringField;
    Reportedestino: TStringField;
    txtMaterial: TEdit;
    btnProveedores: TBitBtn;
    btnReferencia: TBitBtn;
    btnDepartamento: TBitBtn;
    consumiblesdStockMax: TFloatField;
    consumiblesdStockMin: TFloatField;
    anexo_prequisiciondExistencia: TFloatField;
    ReportesIdInsumo: TStringField;
    Label2: TLabel;
    tsSolicitud: TEdit;
    Label4: TLabel;
    tsCodigo: TEdit;
    Label5: TLabel;
    tsAnexo: TEdit;
    ExportaMaterialxPartida1: TMenuItem;
    SaveDialog1: TSaveDialog;
    anexo_prequisicionsAnexo: TStringField;
    function lExisteActividad ( sActividad : String ) : Boolean ;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BtnExitClick(Sender: TObject);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure anexo_requisicionAfterScroll(DataSet: TDataSet);
    procedure tsIsometricoReferenciaKeyPress(Sender: TObject;
      var Key: Char);
    procedure GridPartidasCellClick(Column: TColumn);
    procedure frxReport50GetValue(const VarName: String;
      var Value: Variant);
    procedure GridPartidasTitleBtnClick(Sender: TObject; ACol: Integer;
      Field: TField);
    procedure frmBarra2btnAddClick(Sender: TObject);
    procedure frmBarra2btnEditClick(Sender: TObject);
    procedure frmBarra2btnPostClick(Sender: TObject);
    procedure frmBarra2btnDeleteClick(Sender: TObject);
    procedure frmBarra2btnRefreshClick(Sender: TObject);
    procedure frmBarra2btnCancelClick(Sender: TObject);
    procedure frmBarra2btnExitClick(Sender: TObject);
    procedure tdIdFechaEnter(Sender: TObject);
    procedure tdIdFechaExit(Sender: TObject);
    procedure tdIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure Grid_EntradasEnter(Sender: TObject);
    procedure tmComentariosEnter(Sender: TObject);
    procedure tmComentariosExit(Sender: TObject);
    procedure tdCantidadEnter(Sender: TObject);
    procedure tdCantidadExit(Sender: TObject);
    procedure tdCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure anexo_prequisicionAfterScroll(DataSet: TDataSet);
    procedure GridPartidasEnter(Sender: TObject);
    procedure frmBarra2btnPrinterClick(Sender: TObject);
    procedure tsNumeroActividadEnter(Sender: TObject);
    procedure frxEntradaGetValue(const VarName: String;
      var Value: Variant);
    procedure imgNotasDblClick(Sender: TObject);
    procedure ComentariosAdicionalesClick(Sender: TObject);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tdFechaRequerimientoEnter(Sender: TObject);
    procedure tdFechaRequerimientoExit(Sender: TObject);
    procedure tdFechaRequerimientoKeyPress(Sender: TObject; var Key: Char);
    procedure anexo_prequisicionCalcFields(DataSet: TDataSet);
    procedure ReporteCalcFields(DataSet: TDataSet);
    procedure btnFilesClick(Sender: TObject);
    procedure btConsultaClick(Sender: TObject);
    procedure dbPartidasClick(Sender: TObject);
    procedure consumiblesCalcFields(DataSet: TDataSet);
    procedure dbMatPartKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure dbMatPartKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure dbMatPartCellClick(Column: TColumn);
    procedure dbMatPartEnter(Sender: TObject);
    procedure dbDepartamentosEnter(Sender: TObject);
    procedure dbDepartamentosExit(Sender: TObject);
    procedure dbDepartamentosKeyPress(Sender: TObject; var Key: Char);
    procedure tsSolicitanteKeyPress(Sender: TObject; var Key: Char);
    procedure tsSolicitanteEnter(Sender: TObject);
    procedure tsSolicitanteExit(Sender: TObject);
    procedure dbPartidasEnter(Sender: TObject);
    procedure dbPartidasExit(Sender: TObject);
    procedure tsNumeroActividadMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure validaciones();
    procedure dbFiltrosEnter(Sender: TObject);
    procedure dbFiltrosExit(Sender: TObject);
    procedure PgControlChange(Sender: TObject);
    procedure Grid_EntradasCellClick(Column: TColumn);
    procedure fnrequisita();
    procedure cmbRequisitaExit(Sender: TObject);
    procedure cmbRequisitaEnter(Sender: TObject);
    procedure cmbRequisitaKeyPress(Sender: TObject; var Key: Char);
    procedure dbFiltraKeyPress(Sender: TObject; var Key: Char);
    procedure dbFiltraEnter(Sender: TObject);
    procedure dbFiltraExit(Sender: TObject);
    procedure consulta_requisita();
    procedure ImprimirConImportes1Click(Sender: TObject);
    procedure InsertaPedidos();
    procedure InsertaPedidos2();
    procedure SeguimientoMaterialxPartidasdeAnexo1Click(Sender: TObject);
    procedure ImprimirSeguimientoMaterial1Click(Sender: TObject);
    procedure SeguimientoMaterialxPartida1Click(Sender: TObject);
    procedure Seguimiento_Material(dParamActividad : string);
    procedure frmBarra3btnExitClick(Sender: TObject);
    procedure frmBarra3btnPostClick(Sender: TObject);
    procedure frmBarra3btnAddClick(Sender: TObject);
    procedure frmBarra3btnEditClick(Sender: TObject);
    procedure Grid_EntradasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_EntradasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_EntradasTitleClick(Column: TColumn);
    procedure dbMatPartMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure dbMatPartMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dbMatPartTitleClick(Column: TColumn);
    procedure GridPartidasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure GridPartidasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GridPartidasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Cancelar1Click(Sender: TObject);
    procedure tdCantidadChange(Sender: TObject);
    procedure tiFolioChange(Sender: TObject);
    procedure txtMaterialKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure txtMaterialKeyPress(Sender: TObject; var Key: Char);
    procedure btnReferenciaClick(Sender: TObject);
    procedure btnProveedoresClick(Sender: TObject);
    procedure btnDepartamentoClick(Sender: TObject);
    procedure txtMaterialEnter(Sender: TObject);
    procedure txtMaterialExit(Sender: TObject);
    procedure dbMatPartKeyPress(Sender: TObject; var Key: Char);
    procedure dbFiltraChange(Sender: TObject);
    procedure tsSolicitudEnter(Sender: TObject);
    procedure tsSolicitudExit(Sender: TObject);
    procedure tsSolicitudKeyPress(Sender: TObject; var Key: Char);
    procedure tsCodigoKeyPress(Sender: TObject; var Key: Char);
    procedure tsCodigoEnter(Sender: TObject);
    procedure tsCodigoExit(Sender: TObject);
    procedure tsAnexoKeyPress(Sender: TObject; var Key: Char);
    procedure tsAnexoEnter(Sender: TObject);
    procedure tsAnexoExit(Sender: TObject);
    procedure ExportaMaterialxPartida1Click(Sender: TObject);
    procedure formatoEncabezado();

  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRequisicionPerf: TfrmRequisicionPerf;
  sDescripcion    : String ;
  txtAux          : String ;
  lNuevo          : Boolean ;
  OpcButton1      : String ;
  Valida          : boolean;
  TipoExplosion   : string;
  sAnexo          : string;

  sSuperintendente, sSupervisor              : String ;
  sPuestoSuperintendente, sPuestoSupervisor  : String ;
  sSupervisorTierra, sPuestoSupervisorTierra : String ;
  sBackup    : String ;
  MontoTotal : Currency ;
  lExiste    : boolean;

  utgrid:ticdbgrid;
  utgrid2:ticdbgrid;
  utgrid3:ticdbgrid;
  botonpermiso:tbotonespermisos;
implementation

uses frm_connection ,
     frm_comentariosxanexo,
     frm_Consumibles,
     frm_ordenes,
     frm_pernoctan,
     frm_deptos;

{$R *.dfm}

procedure filtra;
begin
  filtro := '';
  if length(trim(frmRequisicionPerf.txtMaterial.Text)) > 0  then
  begin
      buscar := frmRequisicionPerf.txtMaterial.Text;
      buscar := buscar + '*';
      filtro := ' mDescripcion like ' + QuotedStr(buscar);
  end;

  if filtro <> '' then
  begin
      frmRequisicionPerf.consumibles.Filtered := False;
      frmRequisicionPerf.consumibles.Filter := filtro;
      frmRequisicionPerf.consumibles.Filtered := True;
  end
  else
     frmRequisicionPerf.consumibles.Filtered := False;
end;

procedure TfrmRequisicionPerf.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  botonpermiso.Free;
  action := cafree ;
  utgrid.Destroy;
  utgrid2.Destroy;
  utgrid3.Destroy;
end;


procedure TfrmRequisicionPerf.FormShow(Sender: TObject);
begin
    sMenuP:=stMenu;
    BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'opRequisiciones', PopupPrincipal);
    UtGrid:=TicdbGrid.create(grid_entradas);
    UtGrid2:=TicdbGrid.create(dbmatpart);
    UtGrid3:=TicdbGrid.create(gridpartidas);
    tdIdFecha.Enabled            := False ;
    dbDepartamentos.Enabled      := False ;
    tsSolicitante.Enabled        := False ;
    tmComentarios.Enabled        := False ;
    tdCantidad.ReadOnly          := True ;
    tdFechaRequerimiento.Enabled := False ;
    tdFechaRequerimiento.Date    := Date;
    pgControl.ActivePageIndex    := 0 ;
    QPernoctan.active:=false;
    QPernoctan.Open;

    if connection.configuracion.FieldValues['sExplosion'] = 'Recursos por Concepto/Partida' then
       TipoExplosion := 'recursosanexo'
    else
       TipoExplosion := 'recursosanexosnuevos';

    tsNumeroOrden.Items.Clear ;
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select sNumeroOrden from ordenesdetrabajo where sContrato = :Contrato and ' +
                                'cIdStatus = :status order by sNumeroOrden') ;
    connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    connection.qryBusca.Params.ParamByName('Contrato').Value    := Global_Contrato ;
    connection.qryBusca.Params.ParamByName('status').DataType   := ftString ;
    connection.qryBusca.Params.ParamByName('status').Value      := connection.configuracion.FieldValues [ 'cStatusProceso' ];
    Connection.qryBusca.Open ;
    If (Connection.qryBusca.RecordCount = 1) Or (Connection.qryBusca.RecordCount > 1) Then
        While NOT Connection.qryBusca.Eof Do
        Begin
            tsNumeroOrden.Items.Add(Connection.qryBusca.FieldValues['sNumeroOrden']) ;
            Connection.qryBusca.Next
        End ;
    tsNumeroOrden.ItemIndex := 0 ;

    anexo_requisicion.Active := False ;
    anexo_requisicion.Params.ParamByName('Contrato').DataType := ftString ;
    anexo_requisicion.Params.ParamByName('Contrato').Value    := global_contrato ;
    anexo_requisicion.Open ;

    anexo_prequisicion.Active := False ;
    anexo_prequisicion.Params.ParamByName('Contrato').DataType := ftString ;
    anexo_prequisicion.Params.ParamByName('Contrato').Value    := global_contrato ;
    anexo_prequisicion.Params.ParamByName('Convenio').DataType := ftString ;
    anexo_prequisicion.Params.ParamByName('Convenio').Value    := global_convenio ;
    anexo_prequisicion.Params.ParamByName('Folio').DataType    := ftInteger ;
    If anexo_requisicion.RecordCount > 0 Then
        anexo_prequisicion.Params.ParamByName('Folio').Value := anexo_requisicion.FieldValues['iFolioRequisicion']
    Else
        anexo_prequisicion.Params.ParamByName('Folio').Value := 0 ;
    anexo_prequisicion.Open ;

    Grid_Entradas.SetFocus ;
    //Aqui mandamos a llamar el Usuario para el Text
    tsSolicitante.Text  := Global_Nombre ;
    departamentos.Open ;

    frmBarra2.btnAdd.Enabled     := true;
    frmBarra2.btnEdit.Enabled    := true;
    frmBarra2.btnDelete.Enabled  := true;
    frmBarra2.btnPrinter.Enabled := true;
    BotonPermiso.permisosBotones(frmBarra2);
    BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmRequisicionPerf.BtnExitClick(Sender: TObject);
begin
    Close ;
end;

procedure TfrmRequisicionPerf.frmBarra1btnAddClick(Sender: TObject);
begin
  Try
       validaciones;
       if valida then
       begin
            frmBarra1.btnCancel.Click ;
            exit;
       end;
       if anexo_requisicion.FieldValues['sRequisita'] = 'NUMERO DE ACTIVIDAD' then
          if partidas.RecordCount = 0  then
          begin
                messageDLG('No existen Partidas para el Frente de Trabajo " '+ anexo_requisicion.fieldValues['sNumeroOrden'] +' "', mtInformation, [mbOk], 0);
                exit;
          end;

       If (anexo_requisicion.RecordCount > 0) Then
       Begin
          Insertar1.Enabled := False ;
          Editar1.Enabled   := False ;
          Eliminar1.Enabled := False ;
          Refresh1.Enabled  := False ;
          frmBarra1.btnAddClick(Sender);
          tdCantidad.ReadOnly := False ;
          tdCantidad.Value    := 0;
          tdFechaRequerimiento.Enabled := True ;

          if anexo_requisicion.FieldValues['sRequisita'] = 'NUMERO DE ACTIVIDAD' then
          begin
              dbPartidas.Enabled := True;
              dbPartidas.SetFocus ;
          end;

          if anexo_requisicion.FieldValues['sRequisita'] = 'FAMILIA DE PRODUCTOS' then
          begin
              dbFiltra.Enabled := True;
              dbFiltra.SetFocus ;
          end;

          if anexo_requisicion.FieldValues['sRequisita'] = 'PROVEEDOR DE MATERIALES' then
          begin
              dbFiltra.Enabled := True;
              dbFiltra.SetFocus ;
          end;

          if anexo_requisicion.FieldValues['sRequisita'] = 'ID DE MATERIAL' then
          begin
              txtMaterial.Enabled := True;
              txtMaterial.SetFocus;
          end;

       End;
       BotonPermiso.permisosBotones(frmBarra1);
       BotonPermiso.permisosBotones(frmBarra2);
  except
   on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_RequisicionPerf', 'Al agregar registro ', 0);
   end;
  end;
end;

procedure TfrmRequisicionPerf.frmBarra1btnEditClick(Sender: TObject);
begin
    validaciones;
    if valida then
    begin
          frmBarra1.btnCancel.Click ;
          exit;
    end;
    If (anexo_prequisicion.RecordCount > 0 ) Then
    Begin
        dbFiltra.Items.Text := '';
        dbPartidas.KeyValue := null;

        Insertar1.Enabled := False ;
        Editar1.Enabled := False ;
        //Registrar1.Enabled := True ;
        //Can1.Enabled := True ;
        Eliminar1.Enabled := False ;
        Refresh1.Enabled := False;
        frmBarra1.btnEditClick(Sender);
        tdCantidad.ReadOnly := False ;
        tdFechaRequerimiento.Enabled := True ;
        tdCantidad.SetFocus ;
    End;
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;



procedure TfrmRequisicionPerf.frmBarra1btnPostClick(Sender: TObject);
Var
  SavePlace : TBookmark;
  dCantidad : Currency ;
begin
    If OpcButton = 'New' then
    Begin
        if (consumibles.FieldValues['sIdInsumo'] = '') or (consumibles.FieldValues['sIdInsumo'] = null) or (tdCantidad.Text = '') then
        begin
             messageDLG('Seleccione un Material para Agregar a la Requisicion.', mtInformation, [mbOK], 0);
             exit;
        end;

        if StrToFloat(tdCantidad.Text) > consumibles.FieldValues['dCantidad'] then
           If MessageDlg('La cantidad de Material Requisitado es mayor a la Cantidad Solicitada. Desea continuar ?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
              exit;

        // Consultamos todas las requsiciones que contienen Ordenes de Compra asignadas,..
        If dbPartidas.Text <> '' Then
        begin
             connection.QryBusca2.Active := False;
             Connection.QryBusca2.SQL.Clear ;
             Connection.QryBusca2.SQL.Add('Select distinct iFolioPedido from anexo_prequisicion where sContrato =:Contrato and iFolioRequisicion =:Folio and iFolioPedido <> 0 ');
             Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
             Connection.QryBusca2.Params.ParamByName('Contrato').value    := Global_Contrato ;
             Connection.QryBusca2.Params.ParamByName('Folio').DataType    := ftInteger ;
             Connection.QryBusca2.Params.ParamByName('Folio').value       := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
             Connection.QryBusca2.Open;
        end;
        lExiste := False;
        Try
            //Insertamos el material para la Requisicion,..
            Connection.zCommand.Active := False ;
            Connection.zcommand.SQL.Clear ;
            Connection.zCommand.SQL.Add('INSERT INTO anexo_prequisicion ( sContrato, iFolioRequisicion, sIdInsumo, iFolioPedido, mDescripcion, sMedida, dFechaRequerimiento, dCantidad, dCosto, sNumeroActividad, sWbs, sNumeroOrden) '  +
                                      'VALUES (:Contrato, :Folio, :Insumo, 0, :Descripcion, :Medida, :Requerido, :Cantidad, :Costo, :Actividad, :Wbs, :Orden) ' );
            Connection.zcommand.Params.ParamByName('Contrato').DataType     := ftString ;
            Connection.zcommand.Params.ParamByName('Contrato').value        := Global_Contrato ;
            Connection.zcommand.Params.ParamByName('Folio').DataType        := ftInteger ;
            Connection.zcommand.Params.ParamByName('Folio').value           := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
            Connection.zcommand.Params.ParamByName('Insumo').DataType       := ftString ;
            Connection.zcommand.Params.ParamByName('Insumo').value          := Consumibles.FieldValues['sIdInsumo'] ;
            Connection.zcommand.Params.ParamByName('Descripcion').DataType  := ftMemo ;
            Connection.zcommand.Params.ParamByName('Descripcion').value     := Consumibles.fieldValues['mDescripcion'] ;
            Connection.zcommand.Params.ParamByName('Medida').DataType       := ftString ;
            Connection.zcommand.Params.ParamByName('Medida').value          := Consumibles.fieldValues['sMedida'] ; ;
            Connection.zcommand.Params.ParamByName('Requerido').DataType    := ftDate ;
            Connection.zcommand.Params.ParamByName('Requerido').value       := tdFechaRequerimiento.Date ;
            Connection.zcommand.Params.ParamByName('Cantidad').DataType     := ftFloat ;
            Connection.zcommand.Params.ParamByName('Cantidad').value        := tdCantidad.Value ;
            Connection.zcommand.Params.ParamByName('Costo').DataType        := ftFloat ;
            Connection.zcommand.Params.ParamByName('Costo').value           := Consumibles.FieldValues['dCostoMN'];
            Connection.zcommand.Params.ParamByName('Actividad').DataType    := ftString ;
            Connection.zcommand.Params.ParamByName('Wbs').DataType          := ftString ;
            If dbPartidas.Text = '' Then
            Begin
                Connection.zcommand.Params.ParamByName('Actividad').value   := 'S/N' ;
                Connection.zcommand.Params.ParamByName('Wbs').value         := '' ;
            end
            Else
            Begin
                Connection.zcommand.Params.ParamByName('Actividad').value   := dbPartidas.Text ;
                Connection.zcommand.Params.ParamByName('Wbs').value         := partidas.FieldValues['sWbs'];
            end ;
            Connection.zcommand.Params.ParamByName('Orden').DataType        := ftString ;
            Connection.zcommand.Params.ParamByName('Orden').value           := anexo_requisicion.FieldValues['sNumeroOrden'] ;
            Connection.zcommand.ExecSQL  ;
        Except
            If MessageDlg('Se encontro una coincidencia de la partida en el paquete seleccionado en la requisición actual, desea modificar el registro encontrado?. Si asi no lo desea, se creara un registro nuevo.', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
            Begin
                lExiste := True;
                Connection.qryBusca.Active := False ;
                Connection.qryBusca.SQL.Clear ;
                Connection.qryBusca.SQL.Add('Select Max(iItem) as iItem From anexo_prequisicion Where sContrato = :Contrato And iFolioRequisicion = :Folio and iFolioPedido = 0 and sNumeroActividad = :Actividad') ;
                connection.qryBusca.Params.ParamByName('Contrato').DataType  := ftString ;
                connection.qryBusca.Params.ParamByName('Contrato').value     := Global_Contrato ;
                connection.qryBusca.Params.ParamByName('Folio').DataType     := ftInteger ;
                connection.qryBusca.Params.ParamByName('Folio').value        := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
                connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString ;
                If dbPartidas.Text = '' Then
                   connection.qryBusca.Params.ParamByName('Actividad').value := 'S/N'
                else
                   connection.qryBusca.Params.ParamByName('Actividad').value := dbPartidas.Text ;
                Connection.qryBusca.Open ;

                If NOT Connection.qryBusca.FieldByName('iItem').IsNull Then
                Begin
                    Connection.zCommand.Active := False ;
                    Connection.zCommand.SQL.Clear ;
                    Connection.zCommand.SQL.Add('INSERT INTO anexo_prequisicion ( sContrato, iFolioRequisicion, sIdInsumo, sNumeroActividad, iItem, iFolioPedido, mDescripcion, sMedida, dFechaRequerimiento, dCantidad, sNumeroOrden, dCosto ) ' +
                                                'VALUES (:Contrato, :Folio, :Insumo, :Actividad, :Item, 0, :Descripcion, :Medida, :Requerido, :Cantidad, :Orden, :Costo )');
                    Connection.zcommand.Params.ParamByName('Contrato').DataType   := ftString ;
                    Connection.zcommand.Params.ParamByName('Contrato').value      := Global_Contrato ;
                    Connection.zcommand.Params.ParamByName('Folio').DataType      := ftInteger ;
                    Connection.zcommand.Params.ParamByName('Folio').value         := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
                    Connection.zcommand.Params.ParamByName('Insumo').DataType     := ftString ;
                    Connection.zcommand.Params.ParamByName('Insumo').value        := Consumibles.FieldValues['sIdInsumo'] ;
                    Connection.zcommand.Params.ParamByName('Actividad').DataType  := ftString ;
                    If dbPartidas.Text <> '' Then
                       Connection.zcommand.Params.ParamByName('Actividad').value  := dbPartidas.Text
                    else
                       Connection.zcommand.Params.ParamByName('Actividad').value  := 'S/N' ;
                    Connection.zcommand.Params.ParamByName('Item').DataType       := ftInteger ;
                    Connection.zcommand.Params.ParamByName('Item').value          := Connection.qryBusca.FieldValues['iItem'] + 1 ;
                    Connection.zcommand.Params.ParamByName('Descripcion').DataType:= ftMemo ;
                    Connection.zcommand.Params.ParamByName('Descripcion').value   := Consumibles.fieldValues['mDescripcion'] ;
                    Connection.zcommand.Params.ParamByName('Medida').DataType     := ftString ;
                    Connection.zcommand.Params.ParamByName('Medida').value        := Consumibles.fieldValues['sMedida'] ;
                    Connection.zcommand.Params.ParamByName('Requerido').DataType  := ftDate ;
                    Connection.zcommand.Params.ParamByName('Requerido').value     := tdFechaRequerimiento.Date ;
                    Connection.zcommand.Params.ParamByName('Cantidad').DataType   := ftFloat ;
                    Connection.zcommand.Params.ParamByName('Cantidad').value      := tdCantidad.Value ;
                    Connection.zcommand.Params.ParamByName('Orden').DataType      := ftString ;
                    Connection.zcommand.Params.ParamByName('Orden').value         := anexo_requisicion.FieldValues['sNumeroOrden'] ;
                    Connection.zcommand.Params.ParamByName('Costo').DataType      := ftFloat ;
                    Connection.zcommand.Params.ParamByName('Costo').value         := Consumibles.FieldValues['dCostoMN'];
                    Connection.zcommand.ExecSQL ;

                    If dbPartidas.Text <> '' Then
                        if Connection.QryBusca2.RecordCount > 0 then
                           InsertaPedidos2;
                    MessageDlg('Se creo la partida ' + tsNumeroActividad.Caption + ' en el CPT #' + IntToStr(Connection.qryBusca.FieldValues['iItem'] + 1) , mtInformation, [mbOk], 0);
                End
                Else
                Begin
                    lExiste := True;
                    Connection.zCommand.Active := False ;
                    Connection.zCommand.SQL.Clear ;
                    Connection.zCommand.SQL.Add('INSERT INTO anexo_prequisicion ( sContrato, iFolioRequisicion, sIdInsumo, sNumeroActividad, iItem, iFolioPedido, mDescripcion, sMedida, dFechaRequerimiento, dCantidad, sNumeroOrden , dCosto) ' +
                                                'VALUES (:Contrato, :Folio, :Insumo, :Actividad, :Item, 0, :Descripcion, :Medida, :Requerido, :Cantidad, :Orden, :Costo )');
                    Connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
                    Connection.zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                    Connection.zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                    Connection.zcommand.Params.ParamByName('Folio').value          := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
                    Connection.zcommand.Params.ParamByName('Insumo').DataType      := ftString ;
                    Connection.zcommand.Params.ParamByName('Insumo').value         := Consumibles.FieldValues['sIdInsumo'] ;
                    Connection.zcommand.Params.ParamByName('Actividad').DataType   := ftString ;
                    If dbPartidas.Text <> '' Then
                       Connection.zcommand.Params.ParamByName('Actividad').value   := dbPartidas.Text
                    else
                       Connection.zcommand.Params.ParamByName('Actividad').value   := 'S/N' ;
                    Connection.zcommand.Params.ParamByName('Item').DataType        := ftInteger ;
                    Connection.zcommand.Params.ParamByName('Item').value           := Connection.qryBusca.FieldValues['iItem'] + 1;
                    Connection.zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
                    Connection.zcommand.Params.ParamByName('Descripcion').value    :=Consumibles.fieldValues['mDescripcion'] ;
                    Connection.zcommand.Params.ParamByName('Medida').DataType      := ftString ;
                    Connection.zcommand.Params.ParamByName('Medida').value         := Consumibles.fieldValues['sMedida'] ;
                    Connection.zcommand.Params.ParamByName('Requerido').DataType   := ftDate ;
                    Connection.zcommand.Params.ParamByName('Requerido').value      := tdFechaRequerimiento.Date ;
                    Connection.zcommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
                    Connection.zcommand.Params.ParamByName('Cantidad').value       := tdCantidad.Value ;
                    Connection.zcommand.Params.ParamByName('Orden').DataType       := ftString ;
                    Connection.zcommand.Params.ParamByName('Orden').value          := anexo_requisicion.FieldValues['sNumeroOrden'] ;
                    Connection.zcommand.Params.ParamByName('Costo').DataType       := ftFloat ;
                    Connection.zcommand.Params.ParamByName('Costo').value          := Consumibles.FieldValues['dCostoMN'];
                    Connection.zcommand.ExecSql ;

                    If dbPartidas.Text <> '' Then
                        if Connection.QryBusca2.RecordCount > 0 then
                           InsertaPedidos2;
                    MessageDlg('Se creo la partida ' + tsNumeroActividad.Caption + ' en el CPT #1', mtInformation, [mbOk], 0);
                End
            End
            Else
            Begin
                dCantidad := 0 ;
                lExiste   := True;
                Connection.qryBusca.Active := False ;
                Connection.qryBusca.SQL.Clear ;
                Connection.qryBusca.SQL.Add('Select dCantidad From anexo_prequisicion Where sContrato = :Contrato And ' +
                                            'iFolioRequisicion = :Folio and iFolioPedido = 0 and sNumeroActividad = :Actividad And iItem = 0') ;
                connection.qryBusca.Params.ParamByName('Contrato').DataType  := ftString ;
                connection.qryBusca.Params.ParamByName('Contrato').value     := Global_Contrato ;
                connection.qryBusca.Params.ParamByName('Folio').DataType     := ftInteger ;
                connection.qryBusca.Params.ParamByName('Folio').value        := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
                connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString ;
                connection.qryBusca.Params.ParamByName('Actividad').value    := anexo_requisicion.FieldValues['sNumeroActividad'];
                Connection.qryBusca.Open ;

                If Connection.qryBusca.RecordCount > 0 Then
                    dCantidad := Connection.qryBusca.FieldValues['dCantidad'] ;

                Connection.zCommand.Active := False ;
                Connection.zCommand.SQL.Clear ;
                Connection.zCommand.SQL.Add('UPDATE anexo_prequisicion SET mDescripcion = :Descripcion, sMedida = :Medida, ' +
                                            'dFechaRequerimiento = :Requerido, dCantidad = :Cantidad ' +
                                            'WHERE sContrato = :Contrato And iFolioRequisicion = :Folio and iFolioPedido = 0 And sNumeroActividad = :Actividad And iItem = 0');
                Connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
                Connection.zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                Connection.zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                Connection.zcommand.Params.ParamByName('Folio').value          := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
                Connection.zcommand.Params.ParamByName('Actividad').DataType   := ftString ;
                Connection.zcommand.Params.ParamByName('Actividad').value      := anexo_requisicion.FieldValues['sNumeroActividad'] ;
                Connection.zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
                Connection.zcommand.Params.ParamByName('Descripcion').value    := Consumibles.fieldValues['mDescripcion'] ;
                Connection.zcommand.Params.ParamByName('Medida').DataType      := ftString ;
                Connection.zcommand.Params.ParamByName('Medida').value         := Consumibles.fieldValues['sMedida'] ;
                Connection.zcommand.Params.ParamByName('Requerido').DataType   := ftDate ;
                Connection.zcommand.Params.ParamByName('Requerido').value      := tdFechaRequerimiento.Date ;
                Connection.zcommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
                Connection.zcommand.Params.ParamByName('Cantidad').value       := dCantidad + tdCantidad.Value ;
                Connection.zcommand.ExecSQL ;

                If dbPartidas.Text <> '' Then
                begin
                    if connection.QryBusca2.RecordCount > 0 then
                    begin
                        connection.QryBusca2.First;
                        while not connection.QryBusca2.Eof do
                        begin
                            Connection.zCommand.Active := False ;
                            Connection.zCommand.SQL.Clear ;
                            Connection.zCommand.SQL.Add('UPDATE anexo_prequisicion SET mDescripcion = :Descripcion, sMedida = :Medida, ' +
                                                       'dFechaRequerimiento = :Requerido, dCantidad = :Cantidad ' +
                                                       'WHERE sContrato = :Contrato And iFolioRequisicion = :Folio and iFolioPedido =:Pedido And sNumeroActividad = :Actividad And iItem = 0');
                            Connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
                            Connection.zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                            Connection.zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                            Connection.zcommand.Params.ParamByName('Folio').value          := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
                            Connection.zcommand.Params.ParamByName('Pedido').DataType      := ftInteger ;
                            Connection.zcommand.Params.ParamByName('Pedido').value         := connection.QryBusca2.FieldValues['iFolioPedido'];
                            Connection.zcommand.Params.ParamByName('Actividad').DataType   := ftString ;
                            Connection.zcommand.Params.ParamByName('Actividad').value      := anexo_requisicion.FieldValues['sNumeroActividad'] ;
                            Connection.zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
                            Connection.zcommand.Params.ParamByName('Descripcion').value    := Consumibles.fieldValues['mDescripcion'] ;
                            Connection.zcommand.Params.ParamByName('Medida').DataType      := ftString ;
                            Connection.zcommand.Params.ParamByName('Medida').value         := Consumibles.fieldValues['sMedida'] ;
                            Connection.zcommand.Params.ParamByName('Requerido').DataType   := ftDate ;
                            Connection.zcommand.Params.ParamByName('Requerido').value      := tdFechaRequerimiento.Date ;
                            Connection.zcommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
                            Connection.zcommand.Params.ParamByName('Cantidad').value       := dCantidad + tdCantidad.Value ;
                            Connection.zcommand.ExecSQL ;
                            connection.QryBusca2.Next;
                        end;
                    end;
                end;

            End
        end ;

        if lExiste = False then
        begin
            If dbPartidas.Text <> '' Then
               if Connection.QryBusca2.RecordCount > 0 then
                  InsertaPedidos;
        end;

        anexo_prequisicion.Refresh ;
    End
    Else
    try
        SavePlace := anexo_prequisicion.GetBookmark ;
        Connection.zCommand.Active  := False ;
        Connection.zCommand.SQL.Clear ;
        Connection.zCommand.SQL.Add('UPDATE anexo_prequisicion SET dFechaRequerimiento = :Requerido, dCantidad = :Cantidad ' +
                                    'WHERE sContrato = :Contrato And iFolioRequisicion = :Folio and sNumeroActividad  =:Actividad and iFolioPedido = 0 And sIdInsumo =:Insumo ');
        Connection.zcommand.Params.ParamByName('Contrato').DataType   := ftString ;
        Connection.zcommand.Params.ParamByName('Contrato').value      := Global_Contrato ;
        Connection.zcommand.Params.ParamByName('Folio').DataType      := ftInteger ;
        Connection.zcommand.Params.ParamByName('Folio').value         := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
        Connection.zcommand.Params.ParamByName('Insumo').DataType     := ftString ;
        Connection.zcommand.Params.ParamByName('Insumo').value        := anexo_prequisicion.FieldValues['sIdInsumo'] ;
        Connection.zcommand.Params.ParamByName('Requerido').DataType  := ftDate ;
        Connection.zcommand.Params.ParamByName('Actividad').DataType  := ftString ;
        Connection.zcommand.Params.ParamByName('Actividad').value     := anexo_prequisicion.FieldValues['sNumeroActividad'] ;
        Connection.zcommand.Params.ParamByName('Requerido').value     := tdFechaRequerimiento.Date ;
        Connection.zcommand.Params.ParamByName('Cantidad').DataType   := ftFloat ;
        Connection.zcommand.Params.ParamByName('Cantidad').value      := tdCantidad.Value ;
        Connection.zcommand.ExecSQL ;

        anexo_prequisicion.Refresh  ;
        anexo_prequisicion.GotoBookmark(SavePlace);
        anexo_prequisicion.FreeBookmark(SavePlace);
    except
      on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Requisiciones', 'Al actualizar registro', 0);
      end;
    End ;
    tdCantidad.ReadOnly := True ;
    tdFechaRequerimiento.Enabled := False ;
    Insertar1.Enabled := True ;
    Editar1.Enabled := True ;
    Eliminar1.Enabled := True ;
    Refresh1.Enabled := True ;
    frmBarra1.btnPostClick(Sender);
    BotonPermiso.permisosBotones(frmBarra1);
    BotonPermiso.permisosBotones(frmBarra2);
end;



procedure TfrmRequisicionPerf.frmBarra1btnCancelClick(Sender: TObject);
begin
  Insertar1.Enabled := True ;
  Editar1.Enabled     := True ;
  Eliminar1.Enabled := True ;
  Refresh1.Enabled := True ;
  tdCantidad.ReadOnly := True ;
  tdFechaRequerimiento.Enabled := False ;
  tsNumeroActividad.Caption := '';
  frmBarra1.btnCancelClick(Sender);
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmRequisicionPerf.frmBarra1btnExitClick(Sender: TObject);
begin
  Insertar1.Enabled := True ;
  Editar1.Enabled := True ;
  Eliminar1.Enabled := True ;
  Refresh1.Enabled := True ;
  frmBarra1.btnExitClick(Sender);
end;

procedure TfrmRequisicionPerf.frmBarra1btnDeleteClick(Sender: TObject);
Var
  SavePlace : TBookmark;
begin
    validaciones;
    if valida then
    begin
          frmBarra1.btnCancel.Click ;
          exit;
    end;
    If anexo_prequisicion.RecordCount > 0 Then
    Begin
        With connection do
        Begin
            try
                SavePlace := anexo_prequisicion.GetBookmark ;
                zCommand.Active  := False ;
                zCommand.SQL.Clear ;
                zCommand.SQL.Add('Delete from anexo_prequisicion where sContrato = :Contrato And ' +
                                 'iFolioRequisicion = :Folio and iItem =:Item And sNumeroActividad = :Actividad And sIdInsumo = :Insumo' );
                zcommand.Params.ParamByName('Contrato').DataType  := ftString ;
                zcommand.Params.ParamByName('Contrato').value     := Global_Contrato;
                zcommand.Params.ParamByName('Folio').DataType     := ftInteger ;
                zcommand.Params.ParamByName('Folio').value        := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
                zcommand.Params.ParamByName('Item').DataType      := ftInteger ;
                zcommand.Params.ParamByName('Item').value         := anexo_prequisicion.FieldValues['iItem'] ;
                zcommand.Params.ParamByName('Actividad').DataType := ftString ;
                zcommand.Params.ParamByName('Actividad').value    := anexo_prequisicion.FieldValues['sNumeroActividad'];
                zcommand.Params.ParamByName('Insumo').DataType    := ftstring ;
                zcommand.Params.ParamByName('Insumo').value       := anexo_prequisicion.FieldValues['sidInsumo'] ;
                zcommand.ExecSQL ;

                anexo_prequisicion.Refresh ;
                anexo_prequisicion.GotoBookmark(SavePlace);
                anexo_prequisicion.FreeBookmark(SavePlace);

            Except
              on e : exception do begin
              UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Requisiciones', 'Al eliminar registro', 0);
              end;
            End ;
            GridPartidas.SetFocus
        End
    End
end;

procedure TfrmRequisicionPerf.Insertar1Click(Sender: TObject);
begin
if grid_entradas.Focused then
   frmBarra2.btnAdd.Click;
if gridpartidas.Focused then
   frmBarra1.btnAdd.Click;
end;

procedure TfrmRequisicionPerf.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure TfrmRequisicionPerf.Registrar1Click(Sender: TObject);
begin
    frmBarra2.btnPost.Click 
end;

procedure TfrmRequisicionPerf.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmRequisicionPerf.Cancelar1Click(Sender: TObject);
begin
frmbarra2.btnCancel.Click;
end;

procedure TfrmRequisicionPerf.cmbRequisitaEnter(Sender: TObject);
begin
      cmbRequisita.Color := global_color_entrada;
end;

procedure TfrmRequisicionPerf.cmbRequisitaExit(Sender: TObject);
begin
     cmbRequisita.Color := global_color_salida;
end;

procedure TfrmRequisicionPerf.cmbRequisitaKeyPress(Sender: TObject;
  var Key: Char);
begin
      If Key = #13 Then
         tmComentarios.SetFocus
end;

procedure TfrmRequisicionPerf.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TfrmRequisicionPerf.ExportaMaterialxPartida1Click(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// PLANTILAS DE IMPORTACION //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena: string;
      fs: tStream;
      Alto: Extended;
      Ren, numero : integer;
      dAvance     : double;
    begin
      Ren := 3;
      // Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 100;
      Excel.Columns['A:A'].ColumnWidth := 9.71;
      Excel.Columns['B:B'].ColumnWidth := 8.86;
      Excel.Columns['C:C'].ColumnWidth := 57.29;
      Excel.Columns['D:D'].ColumnWidth := 11.29;
      Excel.Columns['E:E'].ColumnWidth := 7.14;
      Excel.Columns['F:F'].ColumnWidth := 13.29;
      Excel.Columns['G:J'].ColumnWidth := 15;

      // Colocar los encabezados de la plantilla...
      Hoja.Range['A2:A2'].Select;
      Excel.Selection.Value := 'PARTIDA';
      FormatoEncabezado;
      Hoja.Range['B2:B2'].Select;
      Excel.Selection.Value := 'CLAVE';
      FormatoEncabezado;
      Hoja.Range['C2:C2'].Select;
      Excel.Selection.Value := 'MATERIAL';
      FormatoEncabezado;
      Hoja.Range['D2:D2'].Select;
      Excel.Selection.Value := 'CANTIDAD';
      FormatoEncabezado;
      Hoja.Range['E2:E2'].Select;
      Excel.Selection.Value := 'UNIDAD';
      FormatoEncabezado;
      Hoja.Range['F2:F2'].Select;
      Excel.Selection.Value := 'AVANCE';
      FormatoEncabezado;
      Hoja.Range['G2:G2'].Select;
      Excel.Selection.Value := 'No. DE SOLICITUD';
      FormatoEncabezado;
      Hoja.Range['H2:H2'].Select;
      Excel.Selection.Value := 'CODIGO DE MATERIALES';
      FormatoEncabezado;
      Hoja.Range['I2:I2'].Select;
      Excel.Selection.Value := 'CANTIDAD';
      FormatoEncabezado;
      Hoja.Range['J2:J2'].Select;
      Excel.Selection.Value := 'FECHA';
      FormatoEncabezado;

      connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select * from actividadesxanexo where sContrato =:Contrato and sIdConvenio =:Convenio and sTipoActividad ="Actividad" order by iItemOrden');
      connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Contrato').Value    := global_contrato;
      connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Convenio').Value    := global_convenio;
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then
      begin
        sAnexo := connection.QryBusca.FieldValues['sAnexo'];
        while not connection.QryBusca.Eof do
        begin
          Hoja.Cells[Ren, 1].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sNumeroActividad'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 8;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Arial';

          Hoja.Cells[Ren, 3].Select;
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 8;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Value := 'Análisis: '+ connection.QryBusca.FieldValues['sNumeroActividad'] + '     Unidad: '+ connection.QryBusca.FieldValues['sMedida'];

          Hoja.Range['A'+IntToStr(Ren)+':J'+IntToStr(Ren)].Select;
          Excel.Selection.Interior.ColorIndex := 6 ;

          //Aqui consultamos los datos de los insumos,,
          Connection.QryBusca2.Active := False;
          Connection.QryBusca2.Filtered := False;
          connection.QryBusca2.SQL.Clear;
          connection.QryBusca2.SQL.Add('select r.*, i.mDescripcion, i.sMedida from recursosanexosnuevos r '+
                                       'inner join insumos i on (r.sContrato = i.sContrato and r.sIdInsumo = i.sIdInsumo) '+
                                       'where r.sContrato =:Contrato and r.sWbs =:Wbs and r.sNumeroActividad =:Actividad order by r.sIdInsumo');
          connection.QryBusca2.Params.ParamByName('Contrato').DataType  := ftString;
          connection.QryBusca2.Params.ParamByName('Contrato').Value     := global_contrato;
          connection.QryBusca2.Params.ParamByName('Wbs').DataType       := ftString;
          connection.QryBusca2.Params.ParamByName('Wbs').Value          := connection.QryBusca.FieldValues['sWbs'];
          connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString;
          connection.QryBusca2.Params.ParamByName('Actividad').Value    := connection.QryBusca.FieldValues['sNumeroActividad'];
          connection.QryBusca2.Open;

          while not connection.QryBusca2.Eof do
          begin
              Inc(Ren);
              Hoja.Cells[Ren, 1].Select;
              Excel.Selection.Value := connection.QryBusca2.FieldValues['sNumeroActividad'];
              Excel.Selection.HorizontalAlignment := xlCenter;
              Excel.Selection.VerticalAlignment := xlCenter;
              Excel.Selection.Font.Size := 8;
              Excel.Selection.Font.Bold := False;
              Excel.Selection.Font.Name := 'Arial';

              Hoja.Cells[Ren, 2].Select;
              Excel.Selection.Value := connection.QryBusca2.FieldValues['sIdInsumo'];
              Excel.Selection.HorizontalAlignment := xlCenter;
              Excel.Selection.VerticalAlignment := xlCenter;

              Hoja.Cells[Ren, 3].Select;
              Excel.Selection.Value := connection.QryBusca2.FieldValues['mDescripcion'];
              Alto := Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight;
              Hoja.Cells[Ren, 3].Value := '';

              if Alto > 15 then
                Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := Alto
              else
                Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := 15;

              Excel.Selection.Value := connection.QryBusca2.FieldValues['mDescripcion'];

              Hoja.Cells[Ren, 4].Select;
              Excel.Selection.Value :=  FloatToStrF(connection.QryBusca2.FieldValues['dCantidad'], ffNumber, 8, 2 );
              Excel.Selection.HorizontalAlignment := xlRight;
              Excel.Selection.VerticalAlignment   := xlCenter;

              Hoja.Cells[Ren, 5].Select;
              Excel.Selection.Value := connection.QryBusca2.FieldValues['sMedida'];
              Excel.Selection.HorizontalAlignment := xlCenter;
              Excel.Selection.VerticalAlignment := xlCenter;

              //En esta otra parte consultamos los datos de las solicitudes de materiales...
              Connection.zCommand.Active := False;
              Connection.zCommand.Filtered := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('select pr.dCantidad as Cantidad, pr.*, r.sNumeroSolicitud, r.sCodigoMaterial from anexo_prequisicion pr '+
                        'inner join anexo_requisicion r on (pr.sContrato = r.sContrato and pr.iFolioRequisicion = r.iFolioRequisicion) '+
                        'where pr.sContrato =:Contrato and pr.sWbs =:Wbs and pr.sNumeroActividad =:Actividad and sIdInsumo =:Insumo group by sIdInsumo');
              connection.zCommand.Params.ParamByName('Contrato').DataType  := ftString;
              connection.zCommand.Params.ParamByName('Contrato').Value     := global_contrato;
              connection.zCommand.Params.ParamByName('Wbs').DataType       := ftString;
              connection.zCommand.Params.ParamByName('Wbs').Value          := connection.QryBusca.FieldValues['sWbs'];
              connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
              connection.zCommand.Params.ParamByName('Actividad').Value    := connection.QryBusca.FieldValues['sNumeroActividad'];
              connection.zCommand.Params.ParamByName('Insumo').DataType    := ftString;
              connection.zCommand.Params.ParamByName('Insumo').Value       := connection.QryBusca2.FieldValues['sIdInsumo'];
              connection.zCommand.Open;

              if connection.zCommand.RecordCount > 0 then
              begin
                  Hoja.Cells[Ren, 6].Select;

                  dAvance := 0;
                  if connection.QryBusca2.FieldValues['dCantidad'] > 0 then
                  begin
                      dAvance := ((100 / connection.QryBusca2.FieldValues['dCantidad']) * connection.zCommand.FieldValues['Cantidad']);
                      if dAvance > 100 then
                         dAvance:= 100;
                  end;
                  Excel.Selection.Value := FloatToStrF(dAvance, ffNumber, 8, 2 )+ ' %';
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment := xlCenter;

                  Hoja.Cells[Ren, 7].Select;
                  Excel.Selection.Value := connection.zCommand.FieldValues['sNumeroSolicitud'];
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;

                  Hoja.Cells[Ren, 8].Select;
                  Excel.Selection.Value := connection.zCommand.FieldValues['sCodigoMaterial'];
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;

                  Hoja.Cells[Ren, 9].Select;
                  Excel.Selection.Value := connection.zCommand.FieldValues['Cantidad'];
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;

                  Hoja.Cells[Ren, 10].Select;
                  Excel.Selection.Value := connection.zCommand.FieldValues['dFechaRequerimiento'];
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;

              end;
              connection.QryBusca2.Next;
          end;
          if connection.QryBusca2.RecordCount = 0 then
             Inc(Ren);         

          connection.QryBusca.Next;
          Inc(Ren);
        end;
      end;
      Hoja.Cells[2, 2].Select;
      Hoja.Range['A2:J2'].Select;
      // Formato general de encabezado de datos..
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment := xlCenter;
      Excel.Selection.Interior.ColorIndex := 24;
      Excel.Selection.Interior.Pattern := xlSolid;
    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      Hoja.Name := 'ANEXO ' + sAnexo;
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al generar Material por partida' + #10 + #10 + e.Message;
      end;
    end;

    Result := Resultado;
  end;

begin
  // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
  if not SaveDialog1.Execute then
    Exit;

  // Generar el ambiente de excel
  try
    Excel := CreateOleObject('Excel.Application');
  except
    FreeAndNil(Excel);
    showmessage('No es posible generar el ambiente de EXCEL, informe de esto al administrador del sistema.');
    Exit;
  end;

  Excel.Visible := True;
  Excel.DisplayAlerts := False;
  Excel.ScreenUpdating := True;

  Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

  // Verificar si cuenta con las hojas necesarias
  while Libro.Sheets.Count < 2 do
    Libro.Sheets.Add;

  // Verificar si se pasa de hojas necesarias
  Libro.Sheets[1].Select;
  while Libro.Sheets.Count > 1 do
    Excel.ActiveWindow.SelectedSheets.Delete;

  // Proceder a generar la hoja REPORTE
  CadError := '';

  if GenerarPlantilla then
    // Grabar el archivo de excel con el nombre dado
    messageDlg('El Archivo se Genero Correctamente!', mtConfirmation, [mbOk], 0);

  Excel := '';

  if CadError <> '' then
    showmessage(CadError);
end;

procedure TfrmRequisicionPerf.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmRequisicionPerf.Imprimir1Click(Sender: TObject);
begin
   if (grid_entradas.DataSource.DataSet.IsEmpty) or (grid_entradas.DataSource.DataSet.RecordCount<=0) then
       exit;
  try
    If (tsNumeroOrden.Text = 'CONTRATO No. ' + global_contrato) Then
        rDiarioFirmas (global_contrato, '', 'a',anexo_requisicion.FieldValues['dIdFecha'], frmRequisicionPerf )
    Else
        rDiarioFirmas (global_contrato, tsNumeroOrden.Text, 'A',anexo_requisicion.FieldValues['dIdFecha'], frmRequisicionPerf) ;


    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select Sum(r.dCantidad * r.dCosto) as dMontoMN From anexo_prequisicion r ' +
                                'Where r.sContrato = :Contrato And r.iFolioRequisicion = :Folio and r.iFolioPedido = 0 Group By r.iFolioRequisicion');
    connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato ;
    connection.qryBusca.Params.ParamByName('Folio').DataType := ftInteger ;
    connection.qryBusca.Params.ParamByName('Folio').Value := Anexo_Requisicion.FieldValues['iFolioRequisicion'] ;
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 Then
        MontoTotal :=  Connection.qryBusca.fieldValues['dMontoMN'] ;

    Reporte.Active := False ;
    Reporte.Params.ParamByName('Contrato').DataType := ftString ;
    Reporte.Params.ParamByName('Contrato').Value    := global_contrato ;
    Reporte.Params.ParamByName('Convenio').DataType := ftString ;
    Reporte.Params.ParamByName('Convenio').Value    := global_convenio ;
    Reporte.Params.ParamByName('Folio').DataType    := ftInteger ;
    Reporte.Params.ParamByName('Folio').Value       := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
    Reporte.Open ;

    frxEntrada.PreviewOptions.MDIChild := False ;
    frxEntrada.PreviewOptions.Modal := True ;
    frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
    frxEntrada.PreviewOptions.ShowCaptions := False ;
    frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
    frxEntrada.LoadFromFile (global_files + global_miReporte +'_requisicion_simporte.fr3') ;
    frxentrada.ShowReport;    //connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
   if not FileExists(global_files + global_miReporte + '_requisicion_simporte.fr3') then
       showmessage('El archivo de reporte requisicion_simporte.fr3 no existe, notifique al administrador del sistema');
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Requisicion de Materiales', 'Al imprimir', 0);
    end;

  end;
end;

procedure TfrmRequisicionPerf.ImprimirConImportes1Click(Sender: TObject);
begin
  try
     if (grid_entradas.DataSource.DataSet.IsEmpty) or (grid_entradas.DataSource.DataSet.RecordCount<=0) then
 begin
   exit
 end;
    If (tsNumeroOrden.Text = 'CONTRATO No. ' + global_contrato) Then
        rDiarioFirmas (global_contrato, '', 'A',anexo_requisicion.FieldValues['dIdFecha'], frmRequisicionPerf )
    Else
        rDiarioFirmas (global_contrato, tsNumeroOrden.Text, 'A',anexo_requisicion.FieldValues['dIdFecha'], frmRequisicionPerf) ;


    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select Sum(r.dCantidad * r.dCosto) as dMontoMN From anexo_prequisicion r ' +
                                'Where r.sContrato = :Contrato And r.iFolioRequisicion = :Folio and r.iFolioPedido = 0 Group By r.iFolioRequisicion');
    connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato ;
    connection.qryBusca.Params.ParamByName('Folio').DataType := ftInteger ;
    connection.qryBusca.Params.ParamByName('Folio').Value := Anexo_Requisicion.FieldValues['iFolioRequisicion'] ;
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 Then
       MontoTotal :=  Connection.qryBusca.fieldValues['dMontoMN'] ;

    Reporte.Active := False ;
    Reporte.Params.ParamByName('Contrato').DataType := ftString ;
    Reporte.Params.ParamByName('Contrato').Value    := global_contrato ;
    Reporte.Params.ParamByName('Convenio').DataType := ftString ;
    Reporte.Params.ParamByName('Convenio').Value    := global_convenio ;
    Reporte.Params.ParamByName('Folio').DataType    := ftInteger ;
    Reporte.Params.ParamByName('Folio').Value       := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
    Reporte.Open ;

    frxEntrada.PreviewOptions.MDIChild := False ;
    frxEntrada.PreviewOptions.Modal := True ;
    frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
    frxEntrada.PreviewOptions.ShowCaptions := False ;
    frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
    frxEntrada.LoadFromFile (global_files + global_MiReporte +'_requisicionsp.fr3') ;
    frxentrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
    if not FileExists(global_files + global_MiReporte + '_requisicionsp.fr3') then
       showmessage('El archivo de reporte requisicionsp.fr3 no existe, notifique al administrador del sistema');
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Requisicion de Materiales', 'Al imprimir con importes', 0);
    end;
  end;
end;

procedure TfrmRequisicionPerf.ImprimirSeguimientoMaterial1Click(Sender: TObject);
var
   i, x, contador : integer;
   SumTotal, SumExcedente, SumRestante : double;
   Cadena   : string;
begin
  try
    if anexo_prequisicion.RecordCount = 0 then
    begin
        messageDLG('No se encontro ningun registro para imprimir Reporte', mtInformation, [mbOK], 0);
        exit;
    end;
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select sContrato, sWbs, sNumeroActividad, mDescripcion as DescripcionAnexo, '+
                                'dVentaMN, dVentaDLL, dCantidadAnexo, sMedida as sMedidaAnexo  from actividadesxanexo '+
                                'where sContrato =:Contrato and sNumeroActividad =:Actividad and sTipoActividad = "Actividad" and sIdConvenio =:Convenio order by iItemOrden ');
    connection.zCommand.ParamByName('Contrato').AsString  := global_contrato;
    connection.zCommand.ParamByName('Convenio').AsString  := global_convenio;
    connection.zCommand.ParamByName('Actividad').AsString := anexo_prequisicion.FieldValues['sNumeroActividad'];
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
    begin
         rxSeguimiento_Mat.Active := True;
         rxSeguimiento_Mat.EmptyTable;
         while not connection.zCommand.Eof do
         begin
              rxSeguimiento_Mat.Append;
              rxSeguimiento_Mat.FieldValues['sContrato']        := global_contrato;
              rxSeguimiento_Mat.FieldValues['Partida']          := connection.zCommand.FieldValues['sNumeroActividad'];
              rxSeguimiento_Mat.FieldValues['sNumeroActividad'] := connection.zCommand.FieldValues['sNumeroActividad'];
              rxSeguimiento_Mat.FieldValues['DescripcionAnexo'] := connection.zCommand.FieldValues['DescripcionAnexo'];
              rxSeguimiento_Mat.FieldValues['CantidadAnexo']    := connection.zCommand.FieldValues['dCantidadAnexo'];
              rxSeguimiento_Mat.FieldValues['MedidaAnexo']      := connection.zCommand.FieldValues['sMedidaAnexo'];
              rxSeguimiento_Mat.FieldValues['CostoMNAnexo']     := connection.zCommand.FieldValues['dVentaMN'];
              rxSeguimiento_Mat.FieldValues['CostoDLLAnexo']    := connection.zCommand.FieldValues['dVentaDLL'];
              rxSeguimiento_Mat.FieldValues['Tipo']             := 'Anexo';
              rxSeguimiento_Mat.Post;

              //R E Q U I S I C I O N E S .... <<ivan>>
              rxSeguimiento_Mat1.Active := True;
              rxSeguimiento_Mat1.EmptyTable;

              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select  ra.sIdInsumo, ra.dCantidad, ra.dCostoMN, ra.dCostoDLL, i.mDescripcion as Descripcion, '+
                                          'i.sMedida, ap.iFolioRequisicion, ap.iItem, SUM(ap.dCantidad) as dCantidadReq  from recursosanexosnuevos ra '+
                                          'left join insumos i '+
                                          'on (i.sContrato = ra.sContrato and i.sIdInsumo = ra.sIdInsumo ) '+
                                          'left join anexo_prequisicion ap '+
                                          'on (ap.sContrato = ra.sContrato and ap.sNumeroActividad = ra.sNumeroActividad and ap.sIdInsumo = ra.sIdInsumo and ap.iFolioPedido = 0 ) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo, ap.iFolioRequisicion, ap.iItem order by ra.sIdInsumo, ap.iFolioRequisicion, ap.iItem ');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   i  := 1;
                   SumTotal     := 0;
                   SumExcedente := 0;
                   SumRestante  := 0;
                   contador := connection.QryBusca.RecordCount;
                   cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat1.Append;
                        rxSeguimiento_Mat1.FieldValues['Id']              := connection.QryBusca.FieldValues['sIdInsumo'];
                        rxSeguimiento_Mat1.FieldValues['Cantidad']        := connection.QryBusca.FieldValues['dCantidad'];
                        rxSeguimiento_Mat1.FieldValues['Unidad']          := connection.QryBusca.FieldValues['sMedida'];
                        rxSeguimiento_Mat1.FieldValues['FolioReq']        := connection.QryBusca.FieldValues['iFolioRequisicion'];
                        rxSeguimiento_Mat1.FieldValues['ItemReq']         := connection.QryBusca.FieldValues['iItem'];
                        rxSeguimiento_Mat1.FieldValues['dCantidadReq']    := connection.QryBusca.FieldValues['dCantidadReq'];
                        rxSeguimiento_Mat1.FieldValues['dRestanteReq']    := 0;
                        rxSeguimiento_Mat1.FieldValues['dExcedenteReq']   := 0;
                        rxSeguimiento_Mat1.FieldValues['dPorcentajeReq']  := 100;

                        if connection.QryBusca.FieldValues['dCantidadReq'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat1.FieldValues['dRestanteReq']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadReq'];

                        if connection.QryBusca.FieldValues['dCantidadReq'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat1.FieldValues['dExcedenteReq'] := connection.QryBusca.FieldValues['dCantidadReq'] - connection.QryBusca.FieldValues['dCantidad'];

                        if connection.QryBusca.FieldValues['dCantidadReq'] < connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat1.FieldValues['dPorcentajeReq']:= (connection.QryBusca.FieldValues['dCantidadReq'] / connection.QryBusca.FieldValues['dCantidad']) * 100;

                        if Not (rxSeguimiento_Mat1.FieldValues['dCantidadReq'] = null ) then
                           SumTotal     := SumTotal + rxSeguimiento_Mat1.FieldValues['dCantidadReq'];

                        if Not (rxSeguimiento_Mat1.FieldValues['dExcedenteReq'] = null) then
                           SumExcedente := SumExcedente + rxSeguimiento_Mat1.FieldValues['dExcedenteReq'];

                        if Not (rxSeguimiento_Mat1.FieldValues['dRestanteReq'] = null) then
                           SumRestante := SumRestante + rxSeguimiento_Mat1.FieldValues['dRestanteReq'];

                        rxSeguimiento_Mat1.Post;
                        connection.QryBusca.Next;
                        i := i + 1;

                        if (Cadena <> connection.QryBusca.FieldValues['sIdInsumo']) or (contador = 1) then
                        begin
                            for x := 1 to i - 1 do
                                rxSeguimiento_Mat1.Prior;

                            for x := 1 to i -1 do
                            begin
                                rxSeguimiento_Mat1.Edit;
                                rxSeguimiento_Mat1.FieldValues['dCantidadReq_T']    := SumTotal;
                                if SumTotal <= rxSeguimiento_Mat1.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat1.FieldValues['dRestanteReq_T'] := rxSeguimiento_Mat1.FieldValues['Cantidad'] - SumTotal
                                else
                                   rxSeguimiento_Mat1.FieldValues['dRestanteReq_T'] := 0;

                                if SumTotal > rxSeguimiento_Mat1.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat1.FieldValues['dExcedenteReq_T']:= SumTotal - rxSeguimiento_Mat1.FieldValues['Cantidad']
                                else
                                   rxSeguimiento_Mat1.FieldValues['dExcedenteReq_T'] := 0;

                                if SumTotal < rxSeguimiento_Mat1.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat1.FieldValues['dPorcentajeReq_T'] := ((SumTotal - SumExcedente)/rxSeguimiento_Mat1.FieldValues['Cantidad'])* 100
                                else
                                   rxSeguimiento_Mat1.FieldValues['dPorcentajeReq_T'] := 100;
                                rxSeguimiento_Mat1.Post;
                                rxSeguimiento_Mat1.Next;
                            end;
                            SumTotal     := 0;
                            SumExcedente := 0;
                            SumRestante  := 0;
                            i := 0;
                            Cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                        end;
                        contador := contador - 1;
                   end;
              end;

              //O R D E N E S  D E   C O M P R A ....
              rxSeguimiento_Mat2.Active := True;
              rxSeguimiento_Mat2.EmptyTable;

              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select  ra.sIdInsumo, ra.dCantidad, ap.iFolioPedido, ap.iItem, SUM(ap.dCantidad) as dCantidadOC  from recursosanexosnuevos ra '+
                                          'left join anexo_ppedido ap '+
                                          'on (ap.sContrato = ra.sContrato and ap.sNumeroActividad = ra.sNumeroActividad and ap.sIdInsumo = ra.sIdInsumo) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo, ap.iFolioPedido, ap.iItem order by ra.sIdInsumo, ap.iFolioPedido, ap.iItem');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   i  := 1;
                   SumTotal     := 0;
                   SumExcedente := 0;
                   SumRestante  := 0;
                   contador := connection.QryBusca.RecordCount;
                   cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat2.Append;
                        rxSeguimiento_Mat2.FieldValues['Id']             := connection.QryBusca.FieldValues['sIdInsumo'];
                        rxSeguimiento_Mat2.FieldValues['Cantidad']       := connection.QryBusca.FieldValues['dCantidad'];
                        rxSeguimiento_Mat2.FieldValues['FolioOC']        := connection.QryBusca.FieldValues['iFolioPedido'];
                        rxSeguimiento_Mat2.FieldValues['ItemOC']         := connection.QryBusca.FieldValues['iItem'];
                        rxSeguimiento_Mat2.FieldValues['dCantidadOC']    := connection.QryBusca.FieldValues['dCantidadOC'];
                        rxSeguimiento_Mat2.FieldValues['dRestanteOC']    := 0;
                        rxSeguimiento_Mat2.FieldValues['dExcedenteOC']   := 0;
                        rxSeguimiento_Mat2.FieldValues['dPorcentajeOC']  := 100;

                        if connection.QryBusca.FieldValues['dCantidadOC'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat2.FieldValues['dRestanteOC']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadOC'];

                        if connection.QryBusca.FieldValues['dCantidadOC'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat2.FieldValues['dExcedenteOC'] := connection.QryBusca.FieldValues['dCantidadOC'] - connection.QryBusca.FieldValues['dCantidad'];

                        if Not (rxSeguimiento_Mat2.FieldValues['dCantidadOC'] = null ) then
                           SumTotal     := SumTotal + rxSeguimiento_Mat2.FieldValues['dCantidadOC'];

                        if Not (rxSeguimiento_Mat2.FieldValues['dExcedenteOC'] = null ) then
                           SumExcedente := SumExcedente + rxSeguimiento_Mat2.FieldValues['dExcedenteOC'];

                        if Not (rxSeguimiento_Mat2.FieldValues['dRestanteOC'] = null) then
                           SumRestante := SumRestante + rxSeguimiento_Mat2.FieldValues['dRestanteOC'];

                        rxSeguimiento_Mat2.Post;
                        connection.QryBusca.Next;

                        i := i + 1;

                        if (Cadena <> connection.QryBusca.FieldValues['sIdInsumo']) or (contador = 1) then
                        begin
                            for x := 1 to i - 1 do
                                rxSeguimiento_Mat2.Prior;

                            for x := 1 to i -1 do
                            begin
                                rxSeguimiento_Mat2.Edit;
                                rxSeguimiento_Mat2.FieldValues['dCantidadOC_T']    := SumTotal;
                                if SumTotal <= rxSeguimiento_Mat2.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat2.FieldValues['dRestanteOC_T'] := rxSeguimiento_Mat2.FieldValues['Cantidad'] - SumTotal
                                else
                                   rxSeguimiento_Mat2.FieldValues['dRestanteOC_T'] := 0;

                                if SumTotal > rxSeguimiento_Mat2.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat2.FieldValues['dExcedenteOC_T']:= SumTotal - rxSeguimiento_Mat2.FieldValues['Cantidad']
                                else
                                   rxSeguimiento_Mat2.FieldValues['dExcedenteOC_T'] := 0;

                                if SumTotal < rxSeguimiento_Mat2.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat2.FieldValues['dPorcentajeOC_T'] := ((SumTotal - SumExcedente)/rxSeguimiento_Mat2.FieldValues['Cantidad'])* 100
                                else
                                   rxSeguimiento_Mat2.FieldValues['dPorcentajeOC_T'] := 100;
                                rxSeguimiento_Mat2.Post;
                                rxSeguimiento_Mat2.Next;
                            end;
                            SumTotal     := 0;
                            SumExcedente := 0;
                            SumRestante  := 0;
                            i := 0;
                            Cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                        end;
                        contador := contador - 1;
                   end;
              end;

              // E N T R A D A  D E  M A T E R I A L E S ....
              rxSeguimiento_Mat3.Active := True;
              rxSeguimiento_Mat3.EmptyTable;

              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select  ra.sIdInsumo, ra.dCantidad, ap.iFolioPedido, ap.iItem, SUM(ap.dCantidad) as dCantidadIn  from recursosanexosnuevos ra '+
                                          'left join bitacoradeentrada  ap '+
                                          'on (ap.sContrato = ra.sContrato and ap.sNumeroActividad = ra.sNumeroActividad and ap.sIdInsumo = ra.sIdInsumo) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo, ap.iFolioPedido, ap.iItem order by ra.sIdInsumo, ap.iFolioPedido, ap.iItem');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   i  := 1;
                   SumTotal     := 0;
                   SumExcedente := 0;
                   contador := connection.QryBusca.RecordCount;
                   cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat3.Append;
                        rxSeguimiento_Mat3.FieldValues['Id']             := connection.QryBusca.FieldValues['sIdInsumo'];
                        rxSeguimiento_Mat3.FieldValues['Cantidad']       := connection.QryBusca.FieldValues['dCantidad'];
                        rxSeguimiento_Mat3.FieldValues['FolioIn']        := connection.QryBusca.FieldValues['iFolioPedido'];
                        rxSeguimiento_Mat3.FieldValues['ItemIn']         := connection.QryBusca.FieldValues['iItem'];
                        rxSeguimiento_Mat3.FieldValues['dCantidadIn']    := connection.QryBusca.FieldValues['dCantidadIn'];
                        rxSeguimiento_Mat3.FieldValues['dExcedenteIn']   := 0;

                        if connection.QryBusca.FieldValues['dCantidadIn'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat3.FieldValues['dRestanteIn']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadIn'];

                        if connection.QryBusca.FieldValues['dCantidadIn'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat3.FieldValues['dExcedenteIn'] := connection.QryBusca.FieldValues['dCantidadIn'] - connection.QryBusca.FieldValues['dCantidad'];

                        if Not (rxSeguimiento_Mat3.FieldValues['dCantidadIn'] = null ) then
                           SumTotal     := SumTotal + rxSeguimiento_Mat3.FieldValues['dCantidadIn'];

                        if Not (rxSeguimiento_Mat3.FieldValues['dExcedenteIn'] = null ) then
                           SumExcedente := SumExcedente + rxSeguimiento_Mat3.FieldValues['dExcedenteIn'];

                        rxSeguimiento_Mat3.Post;
                        connection.QryBusca.Next;

                        i := i + 1;

                        if (Cadena <> connection.QryBusca.FieldValues['sIdInsumo']) or (contador = 1) then
                        begin
                            for x := 1 to i - 1 do
                                rxSeguimiento_Mat3.Prior;

                            for x := 1 to i -1 do
                            begin
                                rxSeguimiento_Mat3.Edit;
                                rxSeguimiento_Mat3.FieldValues['dCantidadIn_T']    := SumTotal;

                                if SumTotal > rxSeguimiento_Mat3.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat3.FieldValues['dExcedenteIn_T']:= SumTotal - rxSeguimiento_Mat3.FieldValues['Cantidad']
                                else
                                   rxSeguimiento_Mat3.FieldValues['dExcedenteIn_T'] := 0;

                                rxSeguimiento_Mat3.Post;
                                rxSeguimiento_Mat3.Next;
                            end;
                            SumTotal     := 0;
                            SumExcedente := 0;
                            i := 0;
                            Cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                        end;
                        contador := contador - 1;

                   end;
              end;

              // S A L I D A  D E  M A T E R I A L E S ....
              rxSeguimiento_Mat4.Active := True;
              rxSeguimiento_Mat4.EmptyTable;

              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select  ra.sIdInsumo, ra.dCantidad, ap.iFolioSalida, SUM(ap.dCantidad) as dCantidadOut  from recursosanexosnuevos ra '+
                                          'left join bitacoradesalida  ap '+
                                          'on (ap.sContrato = ra.sContrato and ap.sNumeroActividad = ra.sNumeroActividad and ap.sIdInsumo = ra.sIdInsumo) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo, ap.iFolioSalida order by ra.sIdInsumo, ap.iFolioSalida ');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   i  := 1;
                   SumTotal     := 0;
                   SumExcedente := 0;
                   contador := connection.QryBusca.RecordCount;
                   cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat4.Append;
                        rxSeguimiento_Mat4.FieldValues['Id']              := connection.QryBusca.FieldValues['sIdInsumo'];
                        rxSeguimiento_Mat4.FieldValues['Cantidad']        := connection.QryBusca.FieldValues['dCantidad'];
                        rxSeguimiento_Mat4.FieldValues['FolioOut']        := connection.QryBusca.FieldValues['iFolioSalida'];
                        rxSeguimiento_Mat4.FieldValues['dCantidadOut']    := connection.QryBusca.FieldValues['dCantidadOut'];
                        rxSeguimiento_Mat4.FieldValues['dExcedenteOut']   := 0;

                        if connection.QryBusca.FieldValues['dCantidadOut'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat4.FieldValues['dRestanteOut']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadOut'];

                        if connection.QryBusca.FieldValues['dCantidadOut'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat4.FieldValues['dExcedenteOut'] := connection.QryBusca.FieldValues['dCantidadOut'] - connection.QryBusca.FieldValues['dCantidad'];

                         if Not (rxSeguimiento_Mat4.FieldValues['dCantidadOut'] = null ) then
                           SumTotal     := SumTotal + rxSeguimiento_Mat4.FieldValues['dCantidadOut'];

                        if Not (rxSeguimiento_Mat4.FieldValues['dExcedenteOut'] = null ) then
                           SumExcedente := SumExcedente + rxSeguimiento_Mat4.FieldValues['dExcedenteOut'];

                        rxSeguimiento_Mat4.Post;
                        connection.QryBusca.Next;

                        i := i + 1;

                        if (Cadena <> connection.QryBusca.FieldValues['sIdInsumo']) or (contador = 1) then
                        begin
                            for x := 1 to i - 1 do
                                rxSeguimiento_Mat4.Prior;

                            for x := 1 to i -1 do
                            begin
                                rxSeguimiento_Mat4.Edit;
                                rxSeguimiento_Mat4.FieldValues['dCantidadOut_T']    := SumTotal;

                                if SumTotal > rxSeguimiento_Mat4.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat4.FieldValues['dExcedenteOut_T']:= SumTotal - rxSeguimiento_Mat4.FieldValues['Cantidad']
                                else
                                   rxSeguimiento_Mat4.FieldValues['dExcedenteOut_T'] := 0;

                                rxSeguimiento_Mat4.Post;
                                rxSeguimiento_Mat4.Next;
                            end;
                            SumTotal     := 0;
                            SumExcedente := 0;
                            i := 0;
                            Cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                        end;
                        contador := contador - 1;
                   end;
              end;

              // R E P O R T E S   D I A R I O S ....
              rxSeguimiento_Mat5.Active := True;
              rxSeguimiento_Mat5.EmptyTable;

              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select ra.sIdInsumo, ra.dCantidad, SUM(bmp.dCantidad) as dCantidadRD, rd.sNumeroReporte, rd.dIdFecha, rd.sNumeroOrden  from recursosanexosnuevos ra '+
                                          'inner join bitacoradeactividades ba '+
                                          'on (ba.sContrato = ra.sContrato  and ba.sWbs = ra.sWbs and ba.sNumeroActividad = ra.sNumeroActividad) '+
                                          'left join bitacorademateriales  bmp '+
                                          'on(bmp.sContrato = ra.sContrato and bmp.dIdFecha = ba.dIdFecha and bmp.iIdDiario = ba.iIdDiario and bmp.sIdMaterial = ra.sIdInsumo) '+
                                          'inner join reportediario rd '+
                                          'on (rd.sContrato = ba.sContrato and rd.dIdFecha = ba.dIdFecha and rd.sIdTurno = ba.sIdTurno and rd.sNumeroOrden = ba.sNumeroOrden ) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo, rd.sNumeroOrden, rd.dIdFecha order by ra.sIdInsumo, rd.sNumeroOrden, rd.dIdFecha ');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;


              if connection.QryBusca.RecordCount > 0 then
              begin
                   i  := 1;
                   SumTotal     := 0;
                   SumExcedente := 0;
                   SumRestante  := 0;
                   contador := connection.QryBusca.RecordCount;
                   cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat5.Append;
                        rxSeguimiento_Mat5.FieldValues['Id']              := connection.QryBusca.FieldValues['sIdInsumo'];
                        rxSeguimiento_Mat5.FieldValues['Cantidad']        := connection.QryBusca.FieldValues['dCantidad'];
                        rxSeguimiento_Mat5.FieldValues['FechaRD']         := connection.QryBusca.FieldValues['dIdFecha'];
                        rxSeguimiento_Mat5.FieldValues['FrenteRD']        := connection.QryBusca.FieldValues['sNumeroOrden'];
                        rxSeguimiento_Mat5.FieldValues['dCantidadRD']     := connection.QryBusca.FieldValues['dCantidadRD'];
                        rxSeguimiento_Mat5.FieldValues['dRestanteRD']     := 0;
                        rxSeguimiento_Mat5.FieldValues['dExcedenteRD']    := 0;
                        rxSeguimiento_Mat5.FieldValues['dPorcentajeRD']   := 100;

                        if connection.QryBusca.FieldValues['dCantidadRD'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat5.FieldValues['dRestanteRD']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadRD'];

                        if connection.QryBusca.FieldValues['dCantidadRD'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat5.FieldValues['dExcedenteRD'] := connection.QryBusca.FieldValues['dCantidadRD'] - connection.QryBusca.FieldValues['dCantidad'];

                        if connection.QryBusca.FieldValues['dCantidadRD'] < connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat5.FieldValues['dPorcentajeRD']:= (connection.QryBusca.FieldValues['dCantidadRD'] / connection.QryBusca.FieldValues['dCantidad']) * 100;

                        if Not (rxSeguimiento_Mat5.FieldValues['dCantidadRD'] = null ) then
                           SumTotal     := SumTotal + rxSeguimiento_Mat5.FieldValues['dCantidadRD'];

                        if Not (rxSeguimiento_Mat5.FieldValues['dExcedenteRD'] = null ) then
                           SumExcedente := SumExcedente + rxSeguimiento_Mat5.FieldValues['dExcedenteRD'];

                        if Not (rxSeguimiento_Mat5.FieldValues['dRestanteRD'] = null) then
                           SumRestante := SumRestante + rxSeguimiento_Mat5.FieldValues['dRestanteRD'];

                        rxSeguimiento_Mat5.Post;
                        connection.QryBusca.Next;

                        i := i + 1;

                        if (Cadena <> connection.QryBusca.FieldValues['sIdInsumo']) or (contador = 1) then
                        begin
                            for x := 1 to i - 1 do
                                rxSeguimiento_Mat5.Prior;

                            for x := 1 to i -1 do
                            begin
                                rxSeguimiento_Mat5.Edit;
                                rxSeguimiento_Mat5.FieldValues['dCantidadRD_T']    := SumTotal;
                                if SumTotal <= rxSeguimiento_Mat5.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat5.FieldValues['dRestanteRD_T'] := rxSeguimiento_Mat5.FieldValues['Cantidad'] - SumTotal
                                else
                                   rxSeguimiento_Mat5.FieldValues['dRestanteRD_T'] := 0;

                                if SumTotal > rxSeguimiento_Mat5.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat5.FieldValues['dExcedenteRD_T']:= SumTotal - rxSeguimiento_Mat5.FieldValues['Cantidad']
                                else
                                   rxSeguimiento_Mat5.FieldValues['dExcedenteRD_T'] := 0;

                                if SumTotal < rxSeguimiento_Mat5.FieldValues['Cantidad'] then
                                   rxSeguimiento_Mat5.FieldValues['dPorcentajeRD_T'] := ((SumTotal - SumExcedente)/rxSeguimiento_Mat5.FieldValues['Cantidad'])* 100
                                else
                                   rxSeguimiento_Mat5.FieldValues['dPorcentajeRD_T'] := 100;
                                rxSeguimiento_Mat5.Post;
                                rxSeguimiento_Mat5.Next;
                            end;
                            SumTotal     := 0;
                            SumExcedente := 0;
                            SumRestante  := 0;
                            i := 0;
                            Cadena := connection.QryBusca.FieldValues['sIdInsumo'];
                        end;
                        contador := contador - 1;
                   end;
              end;

              // G E N E R A D O R E S  D E  O B R A ....
              rxSeguimiento_Mat6.Active := True;
              rxSeguimiento_Mat6.EmptyTable;

              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('Select e.sContrato, aa.sNumeroActividad, sum(e.dCantidad) as dCantidad, '+
                                          'e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador '+
                                          'from actividadesxanexo aa '+
                                          'inner join  estimacionxpartida e '+
                                          'on (e.sContrato = aa.sContrato and e.sWbs = aa.sWbs and e.sNumeroActividad = aa.sNumeroActividad) '+
                                          'inner join estimaciones e2 '+
                                          'on (e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And e.sNumeroGenerador = e2.sNumeroGenerador) '+
                                          'inner join estimacionperiodo e3 '+
                                          'on (e2.sContrato = e3.sContrato And e2.iNumeroEstimacion = e3.iNumeroEstimacion) '+
                                          'where aa.sContrato =:Contrato and aa.sNumeroActividad =:Actividad and aa.sWbs =:Wbs and sIdConvenio =:Convenio '+
                                          'group by aa.sNumeroActividad ');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Convenio').AsString   := global_convenio;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.ParamByName('Wbs').AsString        := connection.zCommand.FieldValues['sWbs'];
              connection.QryBusca.Open;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat6.Append;
                        rxSeguimiento_Mat6.FieldValues['Id']                := connection.QryBusca.FieldValues['sNumeroActividad'];
                        rxSeguimiento_Mat6.FieldValues['CantidadAnexo']     := connection.zCommand.FieldValues['dCantidadAnexo'];
                        rxSeguimiento_Mat6.FieldValues['dCantidadGen']      := connection.QryBusca.FieldValues['dCantidad'];
                        rxSeguimiento_Mat6.FieldValues['iNumeroEstimacion'] := connection.QryBusca.FieldValues['iNumeroEstimacion'];
                        rxSeguimiento_Mat6.FieldValues['sNumeroOrden']      := connection.QryBusca.FieldValues['sNumeroOrden'];
                        rxSeguimiento_Mat6.FieldValues['sNumeroGenerador']  := connection.QryBusca.FieldValues['sNumeroGenerador'];
                        rxSeguimiento_Mat6.FieldValues['dExcedenteGen']     := 0;

                        rxSeguimiento_Mat6.Post;
                        connection.QryBusca.Next;
                   end;
              end;
              connection.zCommand.Next;
         end;
    end;
    frxEntrada.PreviewOptions.MDIChild := False ;
    frxEntrada.PreviewOptions.Modal := True ;
    frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
    frxEntrada.PreviewOptions.ShowCaptions := False ;
    frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
    frxEntrada.LoadFromFile (global_files + 'seguimiento_material.fr3') ;
    frxEntrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
   if not FileExists(global_files + 'seguimiento_material.fr3') then
       showmessage('El archivo de reporte seguimiento_material.fr3 no existe, notifique al administrador del sistema');
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Requisicion de materiales', 'En el proceso, seguimiento Material x Partida Detalle', 0);
    end;
  end;
end;

procedure TfrmRequisicionPerf.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

procedure TfrmRequisicionPerf.SeguimientoMaterialxPartida1Click(Sender: TObject);
begin
  try
    if anexo_prequisicion.RecordCount > 0 then
       Seguimiento_Material(anexo_prequisicion.FieldValues['sNumeroActividad'])
    else
    begin
         messageDLG('No existe partida para Mostra Reporte.', mtInformation, [mbOk], 0);
         exit;
    end;
    frxEntrada.PreviewOptions.MDIChild := False ;
    frxEntrada.PreviewOptions.Modal := True ;
    frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
    frxEntrada.PreviewOptions.ShowCaptions := False ;
    frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
    frxEntrada.LoadFromFile (global_files + 'seguimiento_materialxpartida_1.fr3') ;
    frxEntrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Requisicion de Materiales', 'En el proceso Seguimiento Material x Partida', 0);
    end;
  end;
end;

procedure TfrmRequisicionPerf.SeguimientoMaterialxPartidasdeAnexo1Click(Sender: TObject);
begin
  try
     if (grid_entradas.DataSource.DataSet.IsEmpty) or (grid_entradas.DataSource.DataSet.RecordCount<=0) then
 begin
   exit
 end;
    Seguimiento_Material('');
    frxEntrada.PreviewOptions.MDIChild := False ;
    frxEntrada.PreviewOptions.Modal := True ;
    frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
    frxEntrada.PreviewOptions.ShowCaptions := False ;
    frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
    frxEntrada.LoadFromFile (global_files + 'seguimiento_materialxpartida.fr3') ;
    frxEntrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
   if not FileExists(global_files + 'seguimiento_materialxpartida.fr3') then
       showmessage('El archivo de reporte seguimiento_materialxpartida.fr3 no existe, notifique al administrador del sistema');
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Requisicion de Materiales', 'En el proceso seguimiento Material General', 0);
    end;
  end;
end;

procedure TfrmRequisicionPerf.anexo_requisicionAfterScroll(DataSet: TDataSet);
begin
    If frmbarra2.btnCancel.Enabled = False then
       frmBarra2.btnCancel.Click ;

    if ((OpcButton1 <> 'New') and (OpcButton1 <> 'Edit')) then
    begin
        tiFolio.Value            := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
        tdIdFecha.Date           := anexo_requisicion.FieldValues['dIdFecha'] ;
        tsNumeroOrden.Text       := anexo_requisicion.FieldValues['sNumeroOrden'] ;
        tsSolicitante.Text       := anexo_requisicion.FieldValues['sSolicito'];
        tmComentarios.Text       := anexo_requisicion.FieldValues['mComentarios'] ;
        dbDepartamentos.KeyValue := anexo_requisicion.FieldValues['sIdDepartamento'];
        tsSolicitud.Text         := anexo_requisicion.FieldValues['sNumeroSolicitud'];
        tsCodigo.Text            := anexo_requisicion.FieldValues['sCodigoMaterial'];
        tsAnexo.Text             := anexo_requisicion.FieldValues['sAnexo'];
        tdIdFecha2.Date           := anexo_requisicion.FieldValues['dFechaRequerido'] ;

        If anexo_requisicion.RecordCount > 0 Then
        Begin
            anexo_prequisicion.Active := False ;
            anexo_prequisicion.Params.ParamByName('Contrato').DataType := ftString ;
            anexo_prequisicion.Params.ParamByName('Contrato').Value    := global_contrato ;
            anexo_prequisicion.Params.ParamByName('Convenio').DataType := ftString ;
            anexo_prequisicion.Params.ParamByName('Convenio').Value    := global_convenio ;
            anexo_prequisicion.Params.ParamByName('Folio').DataType    := ftInteger ;
            anexo_prequisicion.Params.ParamByName('Folio').Value       := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
            anexo_prequisicion.Open ;

            If anexo_prequisicion.RecordCount > 0 then
            Begin
                tdCantidad.Value := anexo_prequisicion.FieldValues['dCantidad'] ;
                Connection.qryBusca.Active := False ;
                Connection.qryBusca.SQL.Clear ;
                Connection.qryBusca.SQL.Add('Select sContrato From comentariosxanexo Where sContrato = :Contrato And sNumeroActividad = :Actividad') ;
                connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
                connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
                connection.qryBusca.Params.ParamByName('actividad').DataType := ftString ;
                connection.qryBusca.Params.ParamByName('actividad').Value := tsNumeroActividad.Caption ;
                Connection.qryBusca.Open ;
                If Connection.qryBusca.RecordCount > 0 Then
                    imgNotas.Visible := True ;
            End
            Else
            Begin
                tsNumeroActividad.Caption := '' ;
                tdCantidad.Value := 0 ;
            End
        End
        Else
        Begin
            tiFolio.Value      := 0 ;
            tdIdFecha.Date     := Date ;
            tmComentarios.Text := '' ;
            tdCantidad.Value   := 0 ;
            tsNumeroOrden.Text := '' ;

            anexo_prequisicion.Active := False ;
            anexo_prequisicion.Params.ParamByName('Contrato').DataType := ftString ;
            anexo_prequisicion.Params.ParamByName('Contrato').Value    := global_contrato ;
            anexo_prequisicion.Params.ParamByName('Convenio').DataType := ftString ;
            anexo_prequisicion.Params.ParamByName('Convenio').Value    := global_convenio ;
            anexo_prequisicion.Params.ParamByName('Folio').DataType    := ftInteger ;
            anexo_prequisicion.Params.ParamByName('Folio').Value       := 0 ;
            anexo_prequisicion.Open ;
        End;
   End
end;

procedure TfrmRequisicionPerf.tsAnexoEnter(Sender: TObject);
begin
    tsAnexo.Color := global_color_entrada;
end;

procedure TfrmRequisicionPerf.tsAnexoExit(Sender: TObject);
begin
    tsAnexo.Color := global_color_salida;
end;

procedure TfrmRequisicionPerf.tsAnexoKeyPress(Sender: TObject; var Key: Char);
begin
     If Key = #13 Then
        tsSolicitud.SetFocus
end;

procedure TfrmRequisicionPerf.tsCodigoEnter(Sender: TObject);
begin
    tsCodigo.Color := global_Color_entrada;
end;

procedure TfrmRequisicionPerf.tsCodigoExit(Sender: TObject);
begin
    tsCodigo.Color := global_Color_salida;
end;

procedure TfrmRequisicionPerf.tsCodigoKeyPress(Sender: TObject; var Key: Char);
begin
    if key = #13 then    
       tsNumeroOrden.SetFocus
end;

procedure TfrmRequisicionPerf.tsIsometricoReferenciaKeyPress(
  Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tmComentarios.SetFocus
end;


procedure TfrmRequisicionPerf.GridPartidasCellClick(Column: TColumn);
begin
     if anexo_prequisicion.RecordCount > 0 Then
         tdCantidad.Text    := anexo_prequisicion.FieldValues['dCantidad'] ;
end;

procedure TfrmRequisicionPerf.frxReport50GetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'ANEXO') = 0 then
  Begin
      Connection.qryBusca.Active := False ;
      Connection.qryBusca.SQL.Clear ;
      Connection.qryBusca.SQL.Add('Select sAnexo From convenios Where sContrato = :Contrato And sIdConvenio = :convenio') ;
      connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
      connection.qryBusca.Params.ParamByName('convenio').DataType := ftString ;
      connection.qryBusca.Params.ParamByName('convenio').Value := global_convenio ;
      Connection.qryBusca.Open ;
      If Connection.qryBusca.RecordCount > 0 Then
          Value := Connection.qryBusca.FieldValues ['sAnexo']
      Else
          Value := '' ;
  End ;

  If CompareText(VarName, 'SUPERINTENDENTE') = 0 then
      Value := sSuperIntendente ;
  If CompareText(VarName, 'SUPERVISOR') = 0 then
      Value := sSupervisor ;
  If CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
      Value := sSupervisorTierra ;

  If CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
      Value := sPuestoSuperIntendente ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
      Value := sPuestoSupervisor ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
      Value := sPuestoSupervisorTierra ;
end;


procedure TfrmRequisicionPerf.GridPartidasTitleBtnClick(Sender: TObject;
  ACol: Integer; Field: TField);
begin

    anexo_prequisicion.Active := False ;
    anexo_prequisicion.Params.ParamByName('Contrato').DataType := ftString ;
    anexo_prequisicion.Params.ParamByName('Contrato').Value    := global_contrato ;
    anexo_prequisicion.Params.ParamByName('Convenio').DataType := ftString ;
    anexo_prequisicion.Params.ParamByName('Convenio').Value    := global_convenio ;
    anexo_prequisicion.Params.ParamByName('Folio').DataType    := ftInteger ;
    anexo_prequisicion.Params.paramByName('Folio').Value       := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
    anexo_prequisicion.Open ;
end;

procedure TfrmRequisicionPerf.GridPartidasTitleClick(Column: TColumn);
begin
    if gridpartidas.datasource.DataSet.IsEmpty=false  then
       if gridpartidas.DataSource.DataSet.RecordCount>0 then
          UtGrid3.DbGridTitleClick(Column);
end;

procedure TfrmRequisicionPerf.frmBarra2btnAddClick(Sender: TObject);
begin
  try
    OpcButton1 := 'New' ;
    frmBarra2.btnAddClick(Sender);
    frmBarra1.btnCancel.Click ;
    pgControl.ActivePageIndex := 0 ;
    tdIdFecha.Enabled         := True ;
    tsNumeroOrden.Enabled     := True ;
    tsSolicitante.Enabled     := True ;
    dbDepartamentos.Enabled   := True ;
    tmComentarios.Enabled     := True ;
    cmbRequisita.Enabled      := True ;
    tiFolio.Value             := 0 ;
    tdIdFecha.Date            := Date ;
    tmComentarios.Text        := '' ;
    tdCantidad.Value          := 0 ;
    tsNumeroOrden.ItemIndex   := 0;
    cmbRequisita.ItemIndex    := 0;
    tdIdFecha2.Date           := date ;
    tsCodigo.Text             := '';
    tsSolicitud.Text          := '';
    tsAnexo.Text              := '';

    anexo_requisicion.Append;
    anexo_requisicion.FieldValues['sContrato']     := global_contrato;
    anexo_requisicion.FieldValues['sMotivo']       := '';
    anexo_requisicion.FieldValues['sEstado']       := '';
    anexo_requisicion.FieldValues['sLugarEntrega'] := 'No';
    anexo_requisicion.FieldValues['mComentarios']  := '*';
    tdIdFecha.SetFocus;
    BotonPermiso.permisosBotones(frmBarra1);
    BotonPermiso.permisosBotones(frmBarra2);
    Grid_Entradas.Enabled := False;
  Except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_RequisicionPerf', 'Al agregar registro ', 0);
    end;
  end;
end;

procedure TfrmRequisicionPerf.frmBarra2btnEditClick(Sender: TObject);
begin
 if grid_entradas.DataSource.DataSet.IsEmpty=false then
 begin
  try
   validaciones;
   if valida then
   begin
          frmBarra1.btnCancel.Click ;
          exit;
   end;
   If anexo_requisicion.RecordCount > 0 then
   Begin
       OpcButton1 := 'Edit' ;
       anexo_requisicion.Edit;
       frmBarra2.btnEditClick(Sender);
       pgControl.ActivePageIndex := 0 ;
       tdIdFecha.Enabled         := True ;
       tsNumeroOrden.Enabled     := True ;
       tsSolicitante.Enabled     := True ;
       dbDepartamentos.Enabled   := True ;
       tmComentarios.Enabled     := True ;
       cmbRequisita.Enabled      := True ;
       tiFolio.SetFocus;
       activapop(frmRequisicionPerf,popupprincipal)
   End;
   BotonPermiso.permisosBotones(frmBarra1);
   BotonPermiso.permisosBotones(frmBarra2);
   Grid_Entradas.Enabled := False;
  except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Requisicion de MAteriales', 'Al editar registro', 0);
  end;
  end;
 end;
end;

procedure TfrmRequisicionPerf.frmBarra2btnPostClick(Sender: TObject);
Var
   nombres, cadenas : TStringList;
begin
    //empieza validacion
    nombres:=TStringList.Create;cadenas:=TStringList.Create;
    nombres.Add('No de frente');
    nombres.Add('Requisita por');
    nombres.Add('Departamento');
    nombres.Add('solicitante');
    nombres.Add('Comentarios');

    cadenas.Add(tsnumeroorden.Text); cadenas.Add(cmbRequisita.Text);
    cadenas.Add(dbdepartamentos.Text);cadenas.Add(tssolicitante.Text);
    cadenas.Add(tmComentarios.Text);

    if not validaTexto(nombres, cadenas,'','') then
    begin
        MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
        exit;
    end;
    //continuainserccion de datos

    if tdIdFecha2.Date < tdIdFecha.Date then
    begin
        messageDLG('La Fecha de entrega es menor a la Fecha de Requisición', mtInformation, [mbOk],0);
        exit;
    end;

    If OpcButton1 = 'New' then
    Begin
        //Verificacion si existe o no el numero de solicitud,,
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('select * from anexo_requisicion Where sContrato =:Contrato and sNumeroOrden =:orden and sNumeroSolicitud =:solicitud ');
        connection.QryBusca.ParamByName('Contrato').AsString  := global_contrato;
        connection.QryBusca.ParamByName('Orden').AsString     := tsNumeroOrden.Text;
        connection.QryBusca.ParamByName('Solicitud').AsString := tsSolicitud.Text;
        connection.QryBusca.Open;

        if connection.QryBusca.RecordCount > 0 then
        begin
            messageDLG('El numero de solicitud '+tsSolicitud.Text+' ya Existe! Favor de Verificar.', mtInformation, [mbOk], 0);
            exit;
        end;

          try
              anexo_requisicion.FieldValues['sNumeroOrden']     := tsNumeroOrden.Text;
              anexo_requisicion.FieldValues['sReferencia']      := tsNumeroOrden.Text;
              anexo_requisicion.FieldValues['dIdFecha']         := tdIdFecha.date;
              anexo_requisicion.FieldValues['dFechaSolicitado'] := tdIdFecha.date;
              anexo_requisicion.FieldValues['dFechaRequerido']  := tdIdFecha2.date;
              anexo_requisicion.FieldValues['sRequisita']       := cmbRequisita.Text;
              anexo_requisicion.FieldValues['sIdDepartamento']  := dbDepartamentos.KeyValue;
              anexo_requisicion.FieldValues['sStatus']          := 'PENDIENTE';
              anexo_requisicion.FieldValues['sEstado']          := 'PENDIENTE';
              anexo_requisicion.FieldValues['mComentarios']     := tmComentarios.Text;
              anexo_requisicion.FieldValues['sSolicito']        := tsSolicitante.Text;
              anexo_requisicion.FieldValues['sNumeroSolicitud'] := tsSolicitud.Text;
              anexo_requisicion.FieldValues['sCodigoMaterial']  := tsCodigo.Text;
              anexo_requisicion.FieldValues['sAnexo']           := tsAnexo.Text;

              // Actualizo Kardex del Sistema ....
              Connection.zCommand.Active := False ;
              Connection.zCommand.SQL.Clear ;
              Connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                          'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
              connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
              connection.zcommand.Params.ParamByName('Usuario').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Usuario').Value := Global_Usuario ;
              connection.zcommand.Params.ParamByName('Fecha').DataType := ftDate ;
              connection.zcommand.Params.ParamByName('Fecha').Value := Date ;
              connection.zcommand.Params.ParamByName('Hora').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now) ;
              connection.zcommand.Params.ParamByName('Descripcion').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Descripcion').Value := 'Registro de Requisición ' + tiFolio.Text +  '] Registrado el [' + DateToStr(tdIdFecha.Date) +  '] Usuario [ ' + global_usuario + ']' ;
              connection.zcommand.Params.ParamByName('Origen').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Origen').Value := 'Requisiciones y Compras' ;
              connection.zcommand.ExecSQL () ;

              anexo_requisicion.Post ;
              anexo_requisicion.Refresh ;
              tdIdFecha.Enabled        := False ;
              tsNumeroOrden.Enabled    := False ;
              tsSolicitante.Enabled    := False ;
              dbDepartamentos.Enabled  := False ;
              tmComentarios.Enabled    := False ;
              cmbRequisita.Enabled     := False ;
              desactivapop(popupprincipal);
              frmBarra2.btnCancelClick(Sender);
          Except
           on e : exception do begin
            UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Requisiciones', 'Al salvar registro', 0);
           end;
          End
    End
    Else
        If OpcButton1 = 'Edit' then
        Begin
              try
                  anexo_requisicion.FieldValues['sNumeroOrden']     := tsNumeroOrden.Text;
                  anexo_requisicion.FieldValues['sReferencia']      := tsNumeroOrden.Text;
                  anexo_requisicion.FieldValues['dIdFecha']         := tdIdFecha.date;
                  anexo_requisicion.FieldValues['dFechaSolicitado'] := tdIdFecha.date;
                  anexo_requisicion.FieldValues['dFechaRequerido']  := tdIdFecha2.date;
                  anexo_requisicion.FieldValues['sRequisita']       := cmbRequisita.Text;
                  anexo_requisicion.FieldValues['sIdDepartamento']  := dbDepartamentos.KeyValue;
                  anexo_requisicion.FieldValues['mComentarios']     := tmComentarios.Text;
                  anexo_requisicion.FieldValues['sSolicito']        := tsSolicitante.Text;
                  anexo_requisicion.FieldValues['sNumeroSolicitud'] := tsSolicitud.Text;
                  anexo_requisicion.FieldValues['sCodigoMaterial']  := tsCodigo.Text;
                  anexo_requisicion.FieldValues['sAnexo']           := tsAnexo.Text;

                  // Actualizo Kardex del Sistema ....
                  Connection.zCommand.Active := False ;
                  Connection.zCommand.SQL.Clear ;
                  Connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                              'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
                  connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                  connection.zcommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
                  connection.zcommand.Params.ParamByName('Usuario').DataType := ftString ;
                  connection.zcommand.Params.ParamByName('Usuario').Value := Global_Usuario ;
                  connection.zcommand.Params.ParamByName('Fecha').DataType := ftDate ;
                  connection.zcommand.Params.ParamByName('Fecha').Value := Date ;
                  connection.zcommand.Params.ParamByName('Hora').DataType := ftString ;
                  connection.zcommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now) ;
                  connection.zcommand.Params.ParamByName('Descripcion').DataType := ftString ;
                  connection.zcommand.Params.ParamByName('Descripcion').Value := 'Modificación de Requisición ' + tiFolio.Text +  '] Registrado el [' + DateToStr(tdIdFecha.Date) +  '] Usuario [ ' + global_usuario + ']' ;
                  connection.zcommand.Params.ParamByName('Origen').DataType := ftString ;
                  connection.zcommand.Params.ParamByName('Origen').Value := 'Requisiciones y Compras' ;
                  connection.zcommand.ExecSQL () ;

                  anexo_requisicion.Post;
                  anexo_requisicion.Refresh ;
                  tdIdFecha.Enabled        := False ;
                  tsNumeroOrden.Enabled    := False ;
                  tsSolicitante.Enabled    := False ;
                  dbDepartamentos.Enabled  := False ;
                  tmComentarios.Enabled    := False ;
                  cmbRequisita.Enabled     := False ;
                  frmBarra2.btnCancelClick(Sender);
              except
                on e : exception do begin
                UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Requisiciones', 'Al salvar registro', 0);
                end;
              end;
        End ;
   OpcButton1 := '' ;
   BotonPermiso.permisosBotones(frmBarra1);
   BotonPermiso.permisosBotones(frmBarra2);
   grid_entradas.Enabled:=true;
end;



procedure TfrmRequisicionPerf.frmBarra2btnDeleteClick(Sender: TObject);
begin
 if grid_entradas.DataSource.DataSet.IsEmpty=false then
 begin
       validaciones;
       if valida then
       begin
            frmBarra1.btnCancel.Click ;
            exit;
       end;
       If anexo_requisicion.RecordCount > 0 Then
          If MessageDlg('Desea eliminar el Folio seleccionado?. Se Eliminaran todos los Materiales!', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          Begin
              // Actualizo Kardex del Sistema ....
              Connection.zCommand.Active := False   ;
              Connection.zCommand.SQL.Clear ;
              Connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                          'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
              connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
              connection.zcommand.Params.ParamByName('Usuario').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Usuario').Value := Global_Usuario ;
              connection.zcommand.Params.ParamByName('Fecha').DataType := ftDate ;
              connection.zcommand.Params.ParamByName('Fecha').Value := Date ;
              connection.zcommand.Params.ParamByName('Hora').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now) ;
              connection.zcommand.Params.ParamByName('Descripcion').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Descripcion').Value := 'Eliminación de Requisición ' + tiFolio.Text +  '] Registrado el [' + DateToStr(tdIdFecha.Date) +  '] Usuario [ ' + global_usuario + ']' ;
              connection.zcommand.Params.ParamByName('Origen').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Origen').Value := 'Requisiciones y Compras' ;
              connection.zcommand.ExecSQL ;

              With connection do
              Begin
                 try
                     zCommand.Sql.Clear ;
                     zcommand.SQL.Add('Delete from anexo_prequisicion where sContrato = :Contrato And iFolioRequisicion = :Folio') ;
                     zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                     zcommand.Params.ParamByName('Contrato').value := Global_Contrato ;
                     zcommand.Params.ParamByName('Folio').DataType := ftInteger ;
                     zcommand.Params.ParamByName('Folio').value := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
                     zcommand.ExecSQL() ;

                     zCommand.Sql.Clear ;
                     zcommand.SQL.Add('Delete from anexo_requisicion where sContrato = :Contrato And iFolioRequisicion = :Folio') ;
                     zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                     zcommand.Params.ParamByName('Contrato').value := Global_Contrato ;
                     zcommand.Params.ParamByName('Folio').DataType := ftInteger ;
                     zcommand.Params.ParamByName('Folio').value := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
                     zcommand.ExecSQL() ;

                     anexo_requisicion.Refresh ;
                 Except
//                     MessageDlg('Ocurrio un error al eliminar el registro.', mtInformation, [mbOk], 0);
                    on e : exception do begin
                    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Requisiciones', 'Al eliminar registro', 0);
                    end;
                 End
              End
          End
 end;
end;


procedure TfrmRequisicionPerf.frmBarra2btnRefreshClick(Sender: TObject);
begin
 if grid_entradas.DataSource.DataSet.IsEmpty=false then
 begin
    anexo_requisicion.Refresh ;
 end;
end;

procedure TfrmRequisicionPerf.frmBarra3btnAddClick(Sender: TObject);
begin
  frmBarra3.btnAddClick(Sender);

end;

procedure TfrmRequisicionPerf.frmBarra3btnEditClick(Sender: TObject);
begin
  frmBarra3.btnEditClick(Sender);

end;

procedure TfrmRequisicionPerf.frmBarra3btnExitClick(Sender: TObject);
begin
  frmBarra3.btnExitClick(Sender);
  close;
end;

procedure TfrmRequisicionPerf.frmBarra3btnPostClick(Sender: TObject);
begin
  frmBarra3.btnPostClick(Sender);

end;

procedure TfrmRequisicionPerf.frmBarra2btnCancelClick(Sender: TObject);
begin
 try
  anexo_requisicion.Cancel;
  tdIdFecha.Enabled       := False ;
  tsNumeroOrden.Enabled   := False ;
  tsSolicitante.Enabled   := False ;
  dbDepartamentos.Enabled := False ;
  cmbRequisita.Enabled    := False ;
  frmBarra2.btnCancelClick(Sender);
  OpcButton1 := '' ;
  //Grid_Entradas.SetFocus ;
  //Grid_Entradas.Enabled:=True;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Requisicion de Materiales', 'Al cancelar', 0);
  end;
 end;
 grid_entradas.Enabled:=true;
end;

procedure TfrmRequisicionPerf.frmBarra2btnExitClick(Sender: TObject);
begin
  frmBarra2.btnExitClick(Sender);
  close
end;

procedure TfrmRequisicionPerf.tdIdFechaEnter(Sender: TObject);
begin
    tdIdFecha.Color := global_color_entrada
end;

procedure TfrmRequisicionPerf.tdIdFechaExit(Sender: TObject);
begin
    tdIdFecha.Color := global_color_salida
end;

procedure TfrmRequisicionPerf.tdIdFechaKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsAnexo.SetFocus
end;

procedure TfrmRequisicionPerf.tiFolioChange(Sender: TObject);
begin
 //  TCurrenCyEdit(sender).Value:=abs(TCurrenCyEdit(sender).Value);
end;

procedure TfrmRequisicionPerf.tsSolicitanteEnter(Sender: TObject);
begin
  tsSolicitante.Color := global_color_entrada ;
end;

procedure TfrmRequisicionPerf.tsSolicitanteExit(Sender: TObject);
begin
  tsSolicitante.Color := global_color_salida ;
end;

procedure TfrmRequisicionPerf.tsSolicitanteKeyPress(Sender: TObject;
  var Key: Char);
begin
      If Key = #13 Then
        cmbRequisita.SetFocus
end;

procedure TfrmRequisicionPerf.tsSolicitudEnter(Sender: TObject);
begin
    tsSolicitud.Color := global_Color_entrada;
end;

procedure TfrmRequisicionPerf.tsSolicitudExit(Sender: TObject);
begin
    tsSolicitud.Color := global_Color_salida;
end;

procedure TfrmRequisicionPerf.tsSolicitudKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key = #13 then
       tsCodigo.SetFocus;
end;

procedure TfrmRequisicionPerf.txtMaterialEnter(Sender: TObject);
begin
    txtMaterial.Color := global_color_entrada
end;

procedure TfrmRequisicionPerf.txtMaterialExit(Sender: TObject);
begin
    txtMaterial.Color := global_color_salida
end;

procedure TfrmRequisicionPerf.txtMaterialKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key = #13 then
       dbMatPart.SetFocus;
end;

procedure TfrmRequisicionPerf.txtMaterialKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     filtra;
end;

procedure TfrmRequisicionPerf.Grid_EntradasCellClick(Column: TColumn);
begin
      fnRequisita;
end;

procedure TfrmRequisicionPerf.Grid_EntradasEnter(Sender: TObject);
begin
      frmBarra1.btnCancel.Click ;
      If frmbarra2.btnCancel.Enabled = True then
          frmBarra2.btnCancel.Click ;

      dbPartidas.KeyValue := null;

      If anexo_requisicion.RecordCount > 0 Then
      Begin
          anexo_prequisicion.Active := False ;
          anexo_prequisicion.Params.ParamByName('Contrato').DataType := ftString ;
          anexo_prequisicion.Params.ParamByName('Contrato').Value    := global_contrato ;
          anexo_prequisicion.Params.ParamByName('Convenio').DataType := ftString ;
          anexo_prequisicion.Params.ParamByName('Convenio').Value    := global_convenio ;
          anexo_prequisicion.Params.ParamByName('Folio').DataType    := ftInteger ;
          anexo_prequisicion.Params.ParamByName('Folio').Value       := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
          anexo_prequisicion.Open ;
          If anexo_prequisicion.RecordCount > 0 then
              tdCantidad.Value := anexo_prequisicion.FieldValues['dCantidad']
          Else
          Begin
              tsNumeroActividad.Caption := '' ;
              tdCantidad.Value := 0 ;
          End
      End
      Else
      Begin
          tiFolio.Value := 0 ;
          tdIdFecha.Date := Date ;
          tmComentarios.Text := '' ;
          tdCantidad.Value := 0 ;
          anexo_prequisicion.Active := False ;
          anexo_prequisicion.Params.ParamByName('Contrato').DataType := ftString ;
          anexo_prequisicion.Params.ParamByName('Contrato').Value    := global_contrato ;
          anexo_prequisicion.Params.ParamByName('Convenio').DataType := ftString ;
          anexo_prequisicion.Params.ParamByName('Convenio').Value    := global_convenio ;
          anexo_prequisicion.Params.ParamByName('Folio').DataType    := ftInteger ;
          anexo_prequisicion.Params.ParamByName('Folio').Value       := 0 ;
          anexo_prequisicion.Open ;
      End
    
end;

procedure TfrmRequisicionPerf.Grid_EntradasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
if grid_entradas.datasource.DataSet.IsEmpty=false  then
if grid_entradas.DataSource.DataSet.RecordCount>0 then
UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmRequisicionPerf.Grid_EntradasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if grid_entradas.datasource.DataSet.IsEmpty=false  then
if grid_entradas.DataSource.DataSet.RecordCount>0 then
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmRequisicionPerf.Grid_EntradasTitleClick(Column: TColumn);
begin
    if grid_entradas.datasource.DataSet.IsEmpty=false  then
       if grid_entradas.DataSource.DataSet.RecordCount>0 then
          UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmRequisicionPerf.tmComentariosEnter(Sender: TObject);
begin
    tmComentarios.Color := global_color_entrada
end;

procedure TfrmRequisicionPerf.tmComentariosExit(Sender: TObject);
begin
    tmComentarios.Color := global_color_salida
end;


procedure TfrmRequisicionPerf.tdCantidadChange(Sender: TObject);
begin
TRxCalcEditChangef(tdCantidad,'Cantidad');
end;

procedure TfrmRequisicionPerf.tdCantidadEnter(Sender: TObject);
begin
    tdCantidad.Color := global_color_entrada
end;

procedure TfrmRequisicionPerf.tdCantidadExit(Sender: TObject);
begin
    tdCantidad.Color := global_color_salida
end;

procedure TfrmRequisicionPerf.tdCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
if not keyFiltroTRxCalcEdit(tdCantidad,key) then
   key:=#0;
    If Key = #13 Then
        tdFechaRequerimiento.SetFocus 
end;

procedure TfrmRequisicionPerf.anexo_prequisicionAfterScroll(
  DataSet: TDataSet);
begin
    ImgNotas.Visible := False ;
    If anexo_prequisicion.RecordCount > 0 then
    Begin

        tdCantidad.Value := anexo_prequisicion.FieldValues['dCantidad'] ;
        tdFechaRequerimiento.Date := anexo_prequisicion.FieldValues['dFechaRequerimiento'] ;

        Connection.qryBusca.Active := False ;
        Connection.qryBusca.SQL.Clear ;
        Connection.qryBusca.SQL.Add('Select sContrato From comentariosxanexo Where sContrato = :Contrato And sNumeroActividad = :Actividad') ;
        connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        connection.qryBusca.Params.ParamByName('actividad').DataType := ftString ;
        connection.qryBusca.Params.ParamByName('actividad').Value := tsNumeroActividad.Caption ;
        Connection.qryBusca.Open ;
        If Connection.qryBusca.RecordCount > 0 Then
            imgNotas.Visible := True ;
    End
    Else
    Begin
        tsNumeroActividad.Caption := '' ;
        tdCantidad.Value := 0 ;
    End
end;

procedure TfrmRequisicionPerf.GridPartidasEnter(Sender: TObject);
begin
    If frmBarra1.btnCancel.Enabled = True Then
        frmBarra1.btnCancel.Click ;

    If anexo_prequisicion.RecordCount > 0 then
     begin
        tdCantidad.Value := anexo_prequisicion.FieldValues['dCantidad'] ;
        tdFechaRequerimiento.Date := anexo_prequisicion.FieldValues['dFechaRequerimiento'] ;
     end
    Else
    Begin
        tsNumeroActividad.Caption := '' ;
        tdCantidad.Value := 0 ;
        tdFechaRequerimiento.Date := Date ;
    End
end;

procedure TfrmRequisicionPerf.GridPartidasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    if gridpartidas.datasource.DataSet.IsEmpty=false  then
       if gridpartidas.DataSource.DataSet.RecordCount>0 then
          UtGrid3.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmRequisicionPerf.GridPartidasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
    if gridpartidas.datasource.DataSet.IsEmpty=false  then
       if gridpartidas.DataSource.DataSet.RecordCount>0 then
          UtGrid3.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmRequisicionPerf.frmBarra2btnPrinterClick(Sender: TObject);
begin
     Imprimir1.OnClick(sender);
end;

procedure TfrmRequisicionPerf.tsNumeroActividadEnter(Sender: TObject);
begin
    tsNumeroActividad.Color := global_Color_Entrada
end;


procedure TfrmRequisicionPerf.tsNumeroActividadMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
       if dbPartidas.Text <> '' then
          tsNumeroActividad.Hint := partidas.FieldValues['mDescripcion'];
end;

function TfrmRequisicionPerf.lExisteActividad ( sActividad : String ) : Boolean ;
Begin
      connection.qryBusca.Active := False ;
      connection.qryBusca.SQL.Clear ;
      connection.qryBusca.SQL.Add('select mDescripcion, dCantidadAnexo, sMedida from actividadesxanexo where sContrato = :Contrato ' +
                                'And sIdConvenio = :Convenio and sNumeroActividad = :Actividad' ) ;
      connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
      connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString ;
      connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio ;
      connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString ;
      connection.qryBusca.Params.ParamByName('Actividad').Value := sActividad ;
      connection.qryBusca.Open ;
      If connection.qryBusca.RecordCount > 0 then
      Begin
          sDescripcion := connection.qryBusca.FieldValues['mDescripcion'] ;
          lExisteActividad := True
      end
      else
      Begin
          sDescripcion := '' ;
          lExisteActividad := False
      end
End ;

procedure TfrmRequisicionPerf.frxEntradaGetValue(const VarName: String;
  var Value: Variant);
  Var
    sCadena : String ;
    iValorNumerico   : LongInt  ;
    Resultado        : Real     ;
    zConsulta  : TZQuery;
    sSQL       : string;
begin

  If CompareText(VarName, 'SUPERINTENDENTE') = 0 then
      Value := sSuperIntendente ;
  If CompareText(VarName, 'SUPERVISOR') = 0 then
      Value := sSupervisor ;
  If CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
      Value := sSupervisorTierra ;

  If CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
      Value := sPuestoSuperIntendente ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
      Value := sPuestoSupervisor ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
      Value := sPuestoSupervisorTierra ;

  If CompareText(VarName, 'CANTIDAD_LETRA')= 0 Then
    Begin
         iValorNumerico   := Trunc(Montototal) ;
         sCadena := xIntToLletres (iValorNumerico) +' PESOS CON ';
         Resultado := roundto(Montototal - iValorNumerico, -2) ;
         Resultado := Resultado * 100;
         iValorNumerico := Trunc(Resultado);
         sCadena := sCadena + xIntToLletres(iValorNumerico) + ' CENTAVOS /100 M.N.';
         Value :=   sCadena
    end ;

  sSQL:='SELECT * FROM firmas WHERE sContrato = :contrato AND dIdFecha <= :fecha ORDER BY dIdFecha DESC';
  zConsulta := TZQuery.Create(self);
  zConsulta.Connection := connection.zConnection;
  zConsulta.Active := False;
  zConsulta.SQL.Clear;
  zConsulta.SQL.Add(sSQL);
  zConsulta.Params.ParamByName('contrato').DataType := ftString;
  zConsulta.Params.ParamByName('contrato').Value := global_contrato;
  zConsulta.Params.ParamByName('fecha').DataType := ftDate;
  zConsulta.Params.ParamByName('fecha').Value := anexo_requisicion.FieldValues['dIdFecha'];
  zConsulta.Open;
  if zConsulta.RecordCount>0 then begin
    If CompareText(VarName, 'REALIZO_PUESTO') = 0 then
        Value := zConsulta.FieldValues['sPuesto11'] ;
    If CompareText(VarName, 'REVISO_PUESTO') = 0 then
        Value := zConsulta.FieldValues['sPuesto12'] ;
    If CompareText(VarName, 'AUTORIZO_PUESTO') = 0 then
        Value := zConsulta.FieldValues['sPuesto13'] ;
    If CompareText(VarName, 'REALIZO_FIRMA') = 0 then
        Value := zConsulta.FieldValues['sFirmante11'] ;
    If CompareText(VarName, 'REVISO_FIRMA') = 0 then
        Value := zConsulta.FieldValues['sFirmante12'] ;
    If CompareText(VarName, 'AUTORIZO_FIRMA') = 0 then
        Value := zConsulta.FieldValues['sFirmante13'] ;
  end
  else
  begin
    If CompareText(VarName, 'REALIZO_PUESTO') = 0 then
        Value := '' ;
    If CompareText(VarName, 'REVISO_PUESTO') = 0 then
        Value := '' ;
    If CompareText(VarName, 'AUTORIZO_PUESTO') = 0 then
        Value := '' ;
    If CompareText(VarName, 'REALIZO_FIRMA') = 0 then
        Value := '' ;
    If CompareText(VarName, 'REVISO_FIRMA') = 0 then
        Value := '' ;
    If CompareText(VarName, 'AUTORIZO_FIRMA') = 0 then
        Value := '' ;
  end;
  zConsulta.free;

end;

procedure TfrmRequisicionPerf.imgNotasDblClick(Sender: TObject);
begin
    ComentariosAdicionales.Click
end;

procedure TfrmRequisicionPerf.ComentariosAdicionalesClick(Sender: TObject);
begin
   if grid_entradas.DataSource.DataSet.IsEmpty=false then
   begin
       if grid_entradas.DataSource.DataSet.RecordCount>0 then
       begin
           global_partida := tsNumeroActividad.Caption ;
           Application.CreateForm(TfrmComentariosxAnexo, frmComentariosxAnexo);
           frmComentariosxAnexo.show ;
       end;
   end;
end;
procedure TfrmRequisicionPerf.tsNumeroOrdenEnter(Sender: TObject);
begin
      tsNumeroOrden.Color := global_color_entrada;
end;

procedure TfrmRequisicionPerf.tsNumeroOrdenExit(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_salida;
end;

procedure TfrmRequisicionPerf.tdFechaRequerimientoEnter(
  Sender: TObject);
begin
    tdFechaRequerimiento.Color := global_color_entrada
end;

procedure TfrmRequisicionPerf.tdFechaRequerimientoExit(Sender: TObject);
begin
    tdFechaRequerimiento.Color := global_color_salida
end;

procedure TfrmRequisicionPerf.tdFechaRequerimientoKeyPress(
  Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        frmBarra1.btnPost.SetFocus
end;

procedure TfrmRequisicionPerf.anexo_prequisicionCalcFields(
  DataSet: TDataSet);
begin
    anexo_prequisiciondMontoMN.Value := anexo_prequisiciondCantidad.Value * anexo_prequisiciondCosto.Value ;
//  anexo_prequisiciondMontoDLL.Value := anexo_prequisiciondCantidad.Value * anexo_prequisiciondVentaDLL.Value ;
    anexo_prequisicionsDescripcion.Value := MidStr(Anexo_pRequisicion.FieldValues['mDescripcion'], 1 , 250) ;
end;

procedure TfrmRequisicionPerf.ReporteCalcFields(DataSet: TDataSet);
begin
  Connection.qryBusca.Active := False ;
  Connection.qryBusca.SQL.Clear ;
  Connection.qryBusca.SQL.Add('Select Sum(p.dCantidad) as dCantidad From anexo_prequisicion p ' +
                              'INNER JOIN anexo_requisicion a ON (a.sContrato = p.sContrato And a.iFolioRequisicion = p.iFolioRequisicion And a.dIdFecha <= :Fecha) ' +
                              'Where p.sContrato = :Contrato And p.sNumeroActividad = :Actividad And sTipoActividad="Actividad" Group By p.sNumeroActividad ') ;
  connection.qryBusca.Params.ParamByName('Contrato').DataType  := ftString ;
  connection.qryBusca.Params.ParamByName('Contrato').Value     := Global_Contrato ;
  connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString ;
  connection.qryBusca.Params.ParamByName('Actividad').Value    := Reporte.FieldValues['sNumeroActividad'] ;
  connection.qryBusca.Params.ParamByName('Fecha').DataType     := ftDate ;
  connection.qryBusca.Params.ParamByName('Fecha').Value        := Anexo_Requisicion.FieldValues['dIdFecha'] ;
  Connection.qryBusca.Open ;
  If Connection.qryBusca.RecordCount > 0 Then
     ReportedAcumulado.Value := Connection.qryBusca.FieldValues['dCantidad'] ;
end;

procedure TfrmRequisicionPerf.btnFilesClick(Sender: TObject);
Var
    flcid, Fila    : Integer ;
    sValue         : String ;
    ImpsNumeroActividad,
    ImpsPaquete,
    ImpmDescripcion,
    ImpdCantidad,
    ImpdFecha             : String ;
begin
    OpenXLS.Title := 'Inserta Archivo de Excel';
    If OpenXLS.Execute then
    Begin
       flcid:=GetUserDefaultLCID;
       ExcelApplication1.Connect;
       ExcelApplication1.Visible[flcid]:= true;
       ExcelApplication1.UserControl   := true;
       Try
           ExcelWorkbook1.ConnectTo(ExcelApplication1.Workbooks.Open(OpenXLS.FileName,
                    emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam,
                    emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, flcid));

          ExcelWorksheet1.ConnectTo(ExcelWorkbook1.Sheets.Item[1] as ExcelWorkSheet);
          Fila := 1 ;
          sValue:=ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2 ;
          If StrToInt(sValue) = Anexo_Requisicion.FieldValues['iFolioRequisicion']  Then
          Begin
               Connection.zCommand.Active := False ;
               Connection.zCommand.SQL.Clear ;
               Connection.zCommand.SQL.Add('DELETE FROM anexo_prequisicion Where sContrato = :contrato And iFolioRequisicion = :Folio') ;
               connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
               connection.zcommand.Params.ParamByName('Contrato').Value    := Global_Contrato ;
               connection.zcommand.Params.ParamByName('Folio').DataType    := ftInteger ;
               connection.zcommand.Params.ParamByName('Folio').Value       := Anexo_Requisicion.FieldValues['iFolioRequisicion'] ;
               connection.zcommand.ExecSQL ;
               While (sValue <> '') DO
               Begin
                  ImpsNumeroActividad := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2 ;
                  ImpsPaquete := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2 ;
                  ImpmDescripcion := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2 ;
                  ImpdCantidad := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2 ;
                  ImpdFecha := DateToStr(ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2) ;
                  // Inserto Datos a la Tabla .....
                  With connection do
                  Begin
                      try
                          zCommand.Active := False ;
                          zCommand.SQL.Clear ;
                          Zcommand.SQL.Add('INSERT INTO anexo_prequisicion ( sContrato, iFolioRequisicion, sNumeroActividad, iItem, mDescripcion, dCantidad, dFechaRequerimiento) ' +
                                           ' VALUES (:contrato, :Folio, :Partida, :Item, :Descripcion, :Cantidad, :Fecha )');
                          zcommand.Params.ParamByName('contrato').DataType    := ftString ;
                          zcommand.Params.ParamByName('contrato').value       := global_contrato ;
                          zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                          zcommand.Params.ParamByName('Folio').value          := Anexo_Requisicion.FieldValues['iFolioRequisicion'] ;
                          zcommand.Params.ParamByName('Partida').DataType     := ftString ;
                          zcommand.Params.ParamByName('Partida').value        := ImpsNumeroActividad ;
                          zcommand.Params.ParamByName('Item').DataType        := ftString ;
                          zcommand.Params.ParamByName('Item').value           := StrToInt(ImpsPaquete) ;
                          zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
                          zcommand.Params.ParamByName('Descripcion').value    := ImpmDescripcion  ;
                          zcommand.Params.ParamByName('Cantidad').DataType    := ftMemo ;
                          zcommand.Params.ParamByName('Cantidad').value       := Trim(ImpdCantidad) ;
                          zcommand.Params.ParamByName('Fecha').DataType       := ftDate ;
                          zcommand.Params.ParamByName('Fecha').value          := StrToDate(Trim(ImpdFecha)) ;
                          zcommand.ExecSQL ;
                      Except
                          MessageDlg('Ocurrio un  error al tratar de actualizar la partida ' + ImpsNumeroActividad , mtInformation, [mbOk], 0);
                      End
                  End ;
                  Fila := Fila + 1 ;
                  sValue :=ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2 ;
               End ;
               MessageDlg('Proceso Terminado.', mtInformation, [mbOk], 0)
          End
       Except
          MessageDlg('No se pudo lograr la conexion con aplicación externa' , mtInformation, [mbOk], 0) ;
       End ;
       Anexo_PRequisicion.Refresh ;
       ExcelApplication1.Quit;
       ExcelApplication1.Disconnect;
    end
end;



procedure TfrmRequisicionPerf.btnProveedoresClick(Sender: TObject);
begin
    if True then
    
    global_frmActivo := 'frm_requisicionPerf';
    Application.CreateForm(TfrmPernoctan, frmPernoctan);
    frmPernoctan.show
end;

procedure TfrmRequisicionPerf.btnReferenciaClick(Sender: TObject);
begin
    global_frmActivo := 'frm_requisicionPerf';
    Application.CreateForm(TfrmOrdenes, frmOrdenes);
    frmOrdenes.show;
end;

procedure TfrmRequisicionPerf.btConsultaClick(Sender: TObject);
begin
   Application.CreateForm(TfrmConsumibles, frmConsumibles);
   frmConsumibles.Show ;
end;


procedure TfrmRequisicionPerf.btnDepartamentoClick(Sender: TObject);
begin
    global_frmActivo := 'frm_requisicionPerf';
    Application.CreateForm(TfrmDeptos, frmDeptos);
    frmDeptos.show
end;

procedure TfrmRequisicionPerf.dbPartidasClick(Sender: TObject);
var
Id : String;
begin

    tsNumeroActividad.Caption  := Partidas.FieldValues['mDescripcion'] ;
    if Partidas.RecordCount > 0 then
    begin
           if anexo_requisicion.FieldValues['sRequisita'] = 'FAMILIA DE PRODUCTOS' then
           begin
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('select sIdFamilia from familias where sDescripcion =:Familia ');
                connection.zCommand.ParamByName('Familia').AsString := dbFiltra.Text;
                connection.zCommand.Open;

                if connection.zCommand.RecordCount > 0 then
                   Id := ' and b.sIdGrupo = "'+connection.zCommand.FieldValues['sIdFamilia']+ '" '
                else
                   Id := ' and b.sIdGrupo = "" ';
           end;

           if anexo_requisicion.FieldValues['sRequisita'] = 'PROVEEDOR DE MATERIALES' then
           begin
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('select sIdProveedor from proveedores where sRazon =:Razon ');
                connection.zCommand.ParamByName('Razon').AsString := dbFiltra.Text;
                connection.zCommand.Open;

                if connection.zCommand.RecordCount > 0 then
                   Id := ' and b.sIdProveedor = "'+connection.zCommand.FieldValues['sIdProveedor']+'" '
                else
                   Id := ' and b.sIdProveedor = "" ';
           end;

           if anexo_requisicion.FieldValues['sRequisita'] = 'ID DE MATERIAL' then
              Id := ' and b.sIdInsumo = "'+dbFiltra.Text + '" ';   

           Consumibles.Active := False ;
           Consumibles.SQL.Clear ;
           Consumibles.SQL.Add('select a.sNumeroActividad, b.sIdInsumo, b.mDescripcion, b.dCostoMN, a.dCantidad, b.sMedida, ' +
                               'b.dExistencia, b.dStockMax, b.dStockMin from '+TipoExplosion+' a inner join insumos b ' +
                               'On (a.sContrato=b.sContrato and a.sIdInsumo =b.sIdInsumo ) where a.scontrato = :Contrato ' +
                               'And a.sNumeroActividad = :Actividad order by a.sNumeroActividad') ;
           Consumibles.Params.ParamByName('Contrato').DataType  := ftString ;
           Consumibles.Params.ParamByName('Contrato').Value     := global_contrato ;
           Consumibles.Params.ParamByName('Actividad').DataType := ftString ;
           Consumibles.Params.ParamByName('Actividad').Value    := Partidas.fieldValues['sNumeroActividad'] ;
           Consumibles.Open ;

           If Consumibles.RecordCount > 0  Then
              Consumibles.First ;
    end
    else
    begin
           Consumibles.Active := False ;
           Consumibles.SQL.Clear ;
           Consumibles.SQL.Add('select a.sNumeroActividad, b.sIdInsumo, b.mDescripcion, b.dCostoMN, a.dCantidad, b.sMedida, ' +
                               'b.dExistencia, b.dStockMax, b.dStockMin from '+TipoExplosion+' a inner join insumos b ' +
                               'On (a.sContrato=b.sContrato and a.sIdInsumo =b.sIdInsumo ) where a.scontrato = :Contrato ' +
                               'And a.sNumeroActividad = :Actividad order by a.sNumeroActividad') ;
           Consumibles.Params.ParamByName('Contrato').DataType  := ftString ;
           Consumibles.Params.ParamByName('Contrato').Value     := global_contrato ;
           Consumibles.Params.ParamByName('Actividad').DataType := ftString ;
           Consumibles.Params.ParamByName('Actividad').Value    := '' ;
           Consumibles.Open ;
    end;

end;

procedure TfrmRequisicionPerf.dbPartidasEnter(Sender: TObject);
begin
      dbPartidas.Color := global_color_entrada
end;

procedure TfrmRequisicionPerf.dbPartidasExit(Sender: TObject);
begin
        dbPartidas.Color := global_color_salida
end;

procedure TfrmRequisicionPerf.consumiblesCalcFields(DataSet: TDataSet);
begin
      If consumibles.RecordCount > 0 Then
      begin
          consumiblessDescripcion.Text := consumibles.FieldValues['mDescripcion'];

          if consumibles.FieldValues['dExistencia'] >= consumibles.FieldValues['dStockMax'] then
             dbMatPart.Columns[4].Font.Color := clBlue
          else
             dbMatPart.Columns[4].Font.Color := clBlack;

          if consumibles.FieldValues['dExistencia'] <= consumibles.FieldValues['dStockMin'] then
             dbMatPart.Columns[5].Font.Color := clRed
          else
             dbMatPart.Columns[5].Font.Color := clBLack;
      end;
end;

procedure TfrmRequisicionPerf.Copy1Click(Sender: TObject);
begin
 if grid_entradas.DataSource.DataSet.IsEmpty=false then
 begin
  if grid_entradas.Focused=true then
    begin
      UtGrid.CopyRowsToClip;
    end;
  if dbmatpart.Focused=true then
    begin
      if dbmatpart.datasource.DataSet.IsEmpty=false  then
      if dbmatpart.DataSource.DataSet.RecordCount>0 then
      UtGrid2.CopyRowsToClip;
    end;
  if gridpartidas.Focused=true then
    begin
      UtGrid3.CopyRowsToClip;
    end;
 end;

end;

procedure TfrmRequisicionPerf.dbDepartamentosEnter(Sender: TObject);
begin
   dbDepartamentos.Color := global_color_entrada
end;

procedure TfrmRequisicionPerf.dbDepartamentosExit(Sender: TObject);
begin
   dbDepartamentos.Color := global_color_salida ;
end;

procedure TfrmRequisicionPerf.dbDepartamentosKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key =  #13 then
    tsSolicitante.SetFocus ;

end;

procedure TfrmRequisicionPerf.dbFiltraChange(Sender: TObject);
begin
    if pgControl.ActivePageIndex = 1 then
       begin
            if anexo_requisicion.RecordCount > 0 then
            begin
                 if anexo_requisicion.FieldValues['sRequisita'] = 'FAMILIA DE PRODUCTOS' then
                 begin
                     connection.zCommand.Active := False;
                     connection.zCommand.SQL.Clear;
                     connection.zCommand.SQL.Add('select sIdFamilia from familias where sDescripcion =:Descripcion');
                     connection.zCommand.ParamByName('Descripcion').AsString := dbFiltra.Text;
                     connection.zCommand.Open;

                     if connection.zCommand.RecordCount > 0 then
                     begin
                         Consumibles.Active := False ;
                         Consumibles.SQL.Clear ;
                         Consumibles.SQL.Add('select "" as sNumeroActividad, sIdInsumo, mDescripcion, dCostoMN, dCantidad, sMedida, ' +
                                             'dExistencia, dStockMax, dStockMin from insumos where sContrato =:Contrato and sIdGrupo =:Familia') ;
                         Consumibles.Params.ParamByName('Contrato').DataType  := ftString ;
                         Consumibles.Params.ParamByName('Contrato').Value     := global_contrato ;
                         Consumibles.Params.ParamByName('Familia').DataType   := ftString ;
                         Consumibles.Params.ParamByName('Familia').Value      := connection.zCommand.FieldValues['sIdFamilia'];
                         Consumibles.Open ;
                     end;
                 end;
                 if anexo_requisicion.FieldValues['sRequisita'] = 'PROVEEDOR DE MATERIALES' then
                 begin
                     connection.zCommand.Active := False;
                     connection.zCommand.SQL.Clear;
                     connection.zCommand.SQL.Add('select sIdProveedor from proveedores where sRazon =:Descripcion');
                     connection.zCommand.ParamByName('Descripcion').AsString := dbFiltra.Text;
                     connection.zCommand.Open;

                     if connection.zCommand.RecordCount > 0 then
                     begin
                         Consumibles.Active := False ;
                         Consumibles.SQL.Clear ;
                         Consumibles.SQL.Add('select "" as sNumeroActividad, sIdInsumo, mDescripcion, dCostoMN, dCantidad, sMedida, ' +
                                             'dExistencia, dStockMax, dStockMin from insumos where sContrato =:Contrato and sIdProveedor =:Proveedor') ;
                         Consumibles.Params.ParamByName('Contrato').DataType  := ftString ;
                         Consumibles.Params.ParamByName('Contrato').Value     := global_contrato ;
                         Consumibles.Params.ParamByName('Proveedor').DataType := ftString ;
                         Consumibles.Params.ParamByName('Proveedor').Value    := connection.zCommand.FieldValues['sIdProveedor'] ;
                         Consumibles.Open ;
                     end;
                 end;
            end;
       end;
end;

procedure TfrmRequisicionPerf.dbFiltraEnter(Sender: TObject);
begin
      dbFiltra.Color := global_color_entrada
end;

procedure TfrmRequisicionPerf.dbFiltraExit(Sender: TObject);
begin
       dbFiltra.Color := global_color_salida;
       consulta_requisita;
end;

procedure TfrmRequisicionPerf.dbFiltraKeyPress(Sender: TObject; var Key: Char);
begin
       if key = #13 then
          dbPartidas.SetFocus;
end;

procedure TfrmRequisicionPerf.dbFiltrosEnter(Sender: TObject);
begin
      dbFiltra.Color := global_color_entrada
end;

procedure TfrmRequisicionPerf.dbFiltrosExit(Sender: TObject);
begin
      dbFiltra.Color := global_color_salida
end;

procedure TfrmRequisicionPerf.dbMatPartCellClick(Column: TColumn);
begin
    if dbmatpart.datasource.DataSet.IsEmpty=false  then
    begin
        if dbmatpart.DataSource.DataSet.RecordCount>0 then
        begin
            if Consumibles.RecordCount > 0 Then
            begin
                if anexo_requisicion.FieldValues['sRequisita'] = 'NUMERO DE ACTIVIDAD' then
                   tdCantidad.Value  := Consumibles.FieldValues['dCantidad']
                else
                   tdCantidad.Value  := 0;
            end;
        end;
    end;
end;

procedure TfrmRequisicionPerf.dbMatPartEnter(Sender: TObject);
begin
    if dbmatpart.datasource.DataSet.IsEmpty = false  then
     begin
         if dbmatpart.DataSource.DataSet.RecordCount>0 then
         begin
             if Consumibles.RecordCount > 0 Then
             begin
                  if anexo_requisicion.FieldValues['sRequisita'] = 'NUMERO DE ACTIVIDAD' then
                     tdCantidad.Value  := Consumibles.FieldValues['dCantidad']
                  else
                     tdCantidad.Value  := 0;
             end;
         end;
     end;
end;

procedure TfrmRequisicionPerf.dbMatPartKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if Consumibles.RecordCount > 0 Then
       tdCantidad.Text    := Consumibles.FieldValues['dCantidad'] ;

end;

procedure TfrmRequisicionPerf.dbMatPartKeyPress(Sender: TObject; var Key: Char);
begin
     if key = #13 then
        tdCantidad.SetFocus;
end;

procedure TfrmRequisicionPerf.dbMatPartKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     if Consumibles.RecordCount > 0 Then
        tdCantidad.Text   := Consumibles.FieldValues['dCantidad'] ;

end;

procedure TfrmRequisicionPerf.dbMatPartMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    if dbmatpart.datasource.DataSet.IsEmpty=false  then
       if dbmatpart.DataSource.DataSet.RecordCount>0 then
           UtGrid2.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmRequisicionPerf.dbMatPartMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
    if dbmatpart.datasource.DataSet.IsEmpty=false  then
       if dbmatpart.DataSource.DataSet.RecordCount>0 then
           UtGrid2.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmRequisicionPerf.dbMatPartTitleClick(Column: TColumn);
begin
if dbmatpart.datasource.DataSet.IsEmpty=false  then
if dbmatpart.DataSource.DataSet.RecordCount>0 then
 UtGrid2.DbGridTitleClick(Column);
end;

procedure TfrmRequisicionPerf.Paste1Click(Sender: TObject);
begin
  if grid_entradas.Focused=true then
    begin
      UtGrid.AddRowsFromClip;
    end;
  if dbmatpart.Focused=true then
    begin
      UtGrid2.AddRowsFromClip;
    end;
  if gridpartidas.Focused=true then
    begin
      UtGrid3.AddRowsFromClip;
    end;
end;

procedure TfrmRequisicionPerf.PgControlChange(Sender: TObject);
begin
      fnRequisita;
end;

procedure TfrmRequisicionPerf.validaciones;
begin
     //Verificamos el estatus de la requisicion..
     valida := False;
     connection.zCommand.Active := False;
     connection.zCommand.SQL.Clear;
     connection.zCommand.SQL.Add('select iFolioRequisicion from anexo_requisicion where sContrato =:Contrato and iFolioRequisicion =:Requisicion and sStatus = "AUTORIZADO" ');
     connection.zCommand.ParamByName('Contrato').AsString    := global_contrato;
     connection.zCommand.ParamByName('Requisicion').AsString := anexo_requisicion.FieldValues['iFolioRequisicion'];
     connection.zCommand.Open;

     if connection.zCommand.RecordCount > 0 then
     begin
          messageDLG(' La Requiscion '+IntToStr(anexo_requisicion.FieldValues['iFolioRequisicion'])+ ' se encuentra en estatus de Autorizado', mtInformation, [mbOk], 0);
          valida := True;
     end;
end;

procedure TfrmRequisicionPerf.fnrequisita;
begin
     if pgControl.ActivePageIndex = 1 then
       begin
            Consumibles.Active := False ;
            if anexo_requisicion.RecordCount > 0 then
            begin
                 if anexo_requisicion.FieldValues['sRequisita'] = 'FAMILIA DE PRODUCTOS' then
                 begin
                       dbFiltra.Visible    := True;
                       dbPartidas.Enabled  := False;
                       txtMaterial.Visible := False;
                       lblBusca.Caption := 'Grupo/Familia' ;

                       connection.QryBusca.Active := False;
                       connection.QryBusca.SQL.Clear;
                       connection.QryBusca.SQL.Add('select sDescripcion from familias');
                       connection.QryBusca.Open;

                       if connection.QryBusca.RecordCount > 0 then
                       begin
                           dbFiltra.Clear;
                           while not connection.QryBusca.Eof do
                           begin
                                dbFiltra.Items.Add(connection.QryBusca.FieldValues['sDescripcion']);
                                connection.QryBusca.Next;
                           end;
                       end;
                 end;
                 if anexo_requisicion.FieldValues['sRequisita'] = 'PROVEEDOR DE MATERIALES' then
                 begin
                       dbFiltra.Visible    := True;
                       dbPartidas.Enabled  := False;
                       txtMaterial.Visible := False;
                       lblBusca.Caption := 'Proveedor' ;

                       connection.QryBusca.Active := False;
                       connection.QryBusca.SQL.Clear;
                       connection.QryBusca.SQL.Add('select sRazon from proveedores');
                       connection.QryBusca.Open;

                       if connection.QryBusca.RecordCount > 0 then
                       begin
                           dbFiltra.Clear;
                           while not connection.QryBusca.Eof do
                           begin
                                dbFiltra.Items.Add(connection.QryBusca.FieldValues['sRazon']);
                                connection.QryBusca.Next;
                           end;
                       end;
                 end;
                 if anexo_requisicion.FieldValues['sRequisita'] = 'ID DE MATERIAL' then
                 begin
                       dbFiltra.Visible    := False;
                       dbPartidas.Enabled  := False;
                       txtMaterial.Visible := True;
                       lblBusca.Caption := 'Id Material' ;

                       Consumibles.Active := False ;
                       Consumibles.SQL.Clear ;
                       Consumibles.SQL.Add('select "" as sNumeroActividad, sIdInsumo, mDescripcion, dCostoMN, dCantidad, sMedida, ' +
                                           'dExistencia, dStockMax, dStockMin from insumos where sContrato =:Contrato') ;
                       Consumibles.Params.ParamByName('Contrato').DataType  := ftString ;
                       Consumibles.Params.ParamByName('Contrato').Value     := global_contrato ;
                       Consumibles.Open ;
                 end;
                 if anexo_requisicion.FieldValues['sRequisita'] = 'NUMERO DE ACTIVIDAD' then
                 begin
                       lblBusca.Caption    := 'Desactivado' ;
                       dbFiltra.Enabled    := False;
                       dbPartidas.Enabled  := True;
                       txtMaterial.Enabled := False;
                       Partidas.Active := False;
                       Partidas.SQL.Clear;
                       Partidas.SQL.Add(' select sNumeroActividad, mDescripcion, sWbs from actividadesxorden '+
                                        ' where scontrato =:Contrato and sNumeroOrden = :Orden '+
                                        ' And sIdConvenio =:Convenio '+
                                        ' And sTipoActividad="Actividad" order by iItemOrden');
                       Partidas.ParamByName('Contrato').AsString := global_contrato;
                       Partidas.ParamByName('Convenio').AsString := global_convenio;
                       Partidas.ParamByName('Orden').AsString    := anexo_requisicion.FieldValues['sNumeroOrden'];
                       Partidas.Open;
                 end;
            end
       end;

end;

procedure TfrmRequisicionPerf.consulta_requisita;
var Id : string;
begin
      if anexo_requisicion.RecordCount > 0 then
      begin
           if anexo_requisicion.FieldValues['sRequisita'] = 'FAMILIA DE PRODUCTOS' then
           begin
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('select sIdFamilia from familias where sDescripcion =:Familia ');
                connection.zCommand.ParamByName('Familia').AsString := dbFiltra.Text;
                connection.zCommand.Open;

                if connection.zCommand.RecordCount > 0 then
                   Id := connection.zCommand.FieldValues['sIdFamilia']
                else
                   Id := '';

                Partidas.Active := False;
                Partidas.SQL.Clear;
                Partidas.SQL.Add('select r.sNumeroActividad, a.mDescripcion, a.sWbs  from insumos i '+
                                 'inner join '+TipoExplosion+' r '+
                                 'on (r.sContrato = i.sContrato and r.sIdInsumo = i.sIdInsumo ) '+
                                 'inner join actividadesxorden a '+
                                 'on (a.sContrato = i.sContrato and a.sIdConvenio =:Convenio and a.sNumeroOrden =:Orden and a.sNumeroActividad = r.sNumeroActividad ) '+
                                 'where i.sContrato =:Contrato and i.sIdGrupo =:Familia group by sNumeroActividad order by iItemOrden ');
                Partidas.ParamByName('Contrato').AsString := global_contrato;
                Partidas.ParamByName('Convenio').AsString := global_convenio;
                Partidas.ParamByName('Orden').AsString    := anexo_requisicion.FieldValues['sNumeroOrden'];
                Partidas.ParamByName('Familia').AsString  := Id;
                Partidas.Open;
           end;


           if anexo_requisicion.FieldValues['sRequisita'] = 'PROVEEDOR DE MATERIALES' then
           begin
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('select sIdProveedor from proveedores where sRazon =:Razon ');
                connection.zCommand.ParamByName('Razon').AsString := dbFiltra.Text;
                connection.zCommand.Open;

                if connection.zCommand.RecordCount > 0 then
                   Id := connection.zCommand.FieldValues['sIdProveedor']
                else
                   Id := '';

                Partidas.Active := False;
                Partidas.SQL.Clear;
                Partidas.SQL.Add('select r.sNumeroActividad, a.mDescripcion, a.sWbs  from insumos i '+
                                 'inner join '+TipoExplosion+' r '+
                                 'on (r.sContrato = i.sContrato and r.sIdInsumo = i.sIdInsumo ) '+
                                 'inner join actividadesxorden a '+
                                 'on (a.sContrato = i.sContrato and a.sIdConvenio =:Convenio and a.sNumeroOrden =:Orden and a.sNumeroActividad = r.sNumeroActividad ) '+
                                 'where i.sContrato =:Contrato and i.sIdProveedor =:Proveedor group by sNumeroActividad order by iItemOrden ');
                Partidas.ParamByName('Contrato').AsString  := global_contrato;
                Partidas.ParamByName('Convenio').AsString  := global_convenio;
                Partidas.ParamByName('Orden').AsString     := anexo_requisicion.FieldValues['sNumeroOrden'];
                Partidas.ParamByName('Proveedor').AsString := Id;
                Partidas.Open;

           end;
           if anexo_requisicion.FieldValues['sRequisita'] = 'ID DE MATERIAL' then
           begin
                Partidas.Active := False;
                Partidas.SQL.Clear;
                Partidas.SQL.Add('select r.sNumeroActividad, a.mDescripcion, a.sWbs  from insumos i '+
                                 'inner join '+TipoExplosion+' r '+
                                 'on (r.sContrato = i.sContrato and r.sIdInsumo = i.sIdInsumo ) '+
                                 'inner join actividadesxorden a '+
                                 'on (a.sContrato = i.sContrato and a.sIdConvenio =:Convenio and a.sNumeroOrden =:Orden and a.sNumeroActividad = r.sNumeroActividad ) '+
                                 'where i.sContrato =:Contrato and i.sIdInsumo =:Insumo group by sNumeroActividad order by iItemOrden ');
                Partidas.ParamByName('Contrato').AsString := global_contrato;
                Partidas.ParamByName('Convenio').AsString := global_convenio;
                Partidas.ParamByName('Orden').AsString    := anexo_requisicion.FieldValues['sNumeroOrden'];
                Partidas.ParamByName('Insumo').AsString   := dbFiltra.Text;
                Partidas.Open;
           end;

           if anexo_requisicion.FieldValues['sRequisita'] = 'NUMERO ACTIVIDAD' then
           begin
                Partidas.Active := False;
                Partidas.SQL.Clear;
                Partidas.SQL.Add(' select sNumeroActividad, mDescripcion, a.sWbs from actividadesxorden '+
                                 ' where scontrato =:Contrato and sNumeroOrden = :Orden '+
                                 ' And sIdConvenio =:Convenio '+
                                 ' And sTipoActividad="Actividad" order by iItemOrden');
                Partidas.ParamByName('Contrato').AsString := global_contrato;
                Partidas.ParamByName('Convenio').AsString := global_convenio;
                Partidas.ParamByName('Orden').AsString    := anexo_requisicion.FieldValues['sNumeroOrden'];
                Partidas.Open;
           end;

      end;

end;

procedure TfrmRequisicionPerf.InsertaPedidos;
begin
     connection.QryBusca2.First;
     While not connection.QryBusca2.Eof do
     begin
          Connection.zCommand.Active := False ;
          Connection.zcommand.SQL.Clear ;
          Connection.zCommand.SQL.Add('INSERT INTO anexo_prequisicion ( sContrato, iFolioRequisicion, sIdInsumo, iFolioPedido, mDescripcion, sMedida, dFechaRequerimiento, dCantidad, dCosto, sNumeroActividad, sNumeroOrden) '  +
                                            'VALUES (:Contrato, :Folio, :Insumo, :Pedido, :Descripcion, :Medida, :Requerido, :Cantidad, :Costo, :Actividad, :Orden) ' );
          Connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
          Connection.zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
          Connection.zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
          Connection.zcommand.Params.ParamByName('Folio').value          := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
          Connection.zcommand.Params.ParamByName('Pedido').DataType      := ftInteger ;
          Connection.zcommand.Params.ParamByName('Pedido').value         := connection.QryBusca2.FieldValues['iFolioPedido'];
          Connection.zcommand.Params.ParamByName('Insumo').DataType      := ftString ;
          Connection.zcommand.Params.ParamByName('Insumo').value         := Consumibles.FieldValues['sIdInsumo'] ;
          Connection.zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
          Connection.zcommand.Params.ParamByName('Descripcion').value    := Consumibles.fieldValues['mDescripcion'] ;
          Connection.zcommand.Params.ParamByName('Medida').DataType      := ftString ;
          Connection.zcommand.Params.ParamByName('Medida').value         := Consumibles.fieldValues['sMedida'] ; ;
          Connection.zcommand.Params.ParamByName('Requerido').DataType   := ftDate ;
          Connection.zcommand.Params.ParamByName('Requerido').value      := tdFechaRequerimiento.Date ;
          Connection.zcommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
          Connection.zcommand.Params.ParamByName('Cantidad').value       := tdCantidad.Value ;
          Connection.zcommand.Params.ParamByName('Costo').DataType       := ftFloat ;
          Connection.zcommand.Params.ParamByName('Costo').value          := Consumibles.FieldValues['dCostoMN'];
          Connection.zcommand.Params.ParamByName('Actividad').DataType   := ftString ;
          If dbPartidas.Text <> '' Then
             Connection.zcommand.Params.ParamByName('Actividad').value   := dbPartidas.Text
          else
             Connection.zcommand.Params.ParamByName('Actividad').value   := null;
          Connection.zcommand.Params.ParamByName('Orden').DataType       := ftString ;
          Connection.zcommand.Params.ParamByName('Orden').value          := anexo_requisicion.FieldValues['sNumeroOrden'] ;
          Connection.zcommand.ExecSQL  ;
          connection.QryBusca2.Next;
     end;
end;

procedure TfrmRequisicionPerf.InsertaPedidos2;
begin
      connection.QryBusca2.First;
      while not connection.QryBusca2.Eof do
      begin
           Connection.zCommand.Active := False ;
           Connection.zCommand.SQL.Clear ;
           Connection.zCommand.SQL.Add('INSERT INTO anexo_prequisicion ( sContrato, iFolioRequisicion, sIdInsumo, sNumeroActividad, iItem, iFolioPedido, mDescripcion, sMedida, dFechaRequerimiento, dCantidad, sNumeroOrden, dCosto ) ' +
                                       'VALUES (:Contrato, :Folio, :Insumo, :Actividad, :Item, :Pedido, :Descripcion, :Medida, :Requerido, :Cantidad, :Orden, :Costo )');
           Connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
           Connection.zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
           Connection.zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
           Connection.zcommand.Params.ParamByName('Folio').value          := anexo_requisicion.FieldValues['iFolioRequisicion'] ;
           Connection.zcommand.Params.ParamByName('Insumo').DataType      := ftString ;
           Connection.zcommand.Params.ParamByName('Insumo').value         := Consumibles.FieldValues['sIdInsumo'] ;
           Connection.zcommand.Params.ParamByName('Pedido').DataType      := ftInteger ;
           Connection.zcommand.Params.ParamByName('Pedido').value         := connection.QryBusca2.FieldValues['iFolioPedido'] ;
           Connection.zcommand.Params.ParamByName('Actividad').DataType   := ftString ;
           If dbPartidas.Text <> '' Then
              Connection.zcommand.Params.ParamByName('Actividad').value   := dbPartidas.Text
           else
              Connection.zcommand.Params.ParamByName('Actividad').value   := null;
           Connection.zcommand.Params.ParamByName('Item').DataType        := ftInteger ;
           Connection.zcommand.Params.ParamByName('Item').value           := Connection.qryBusca.FieldValues['iItem'] + 1;
           Connection.zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
           Connection.zcommand.Params.ParamByName('Descripcion').value    := Consumibles.fieldValues['mDescripcion'] ;
           Connection.zcommand.Params.ParamByName('Medida').DataType      := ftString ;
           Connection.zcommand.Params.ParamByName('Medida').value         := Consumibles.fieldValues['sMedida'] ;
           Connection.zcommand.Params.ParamByName('Requerido').DataType   := ftDate ;
           Connection.zcommand.Params.ParamByName('Requerido').value      := tdFechaRequerimiento.Date ;
           Connection.zcommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
           Connection.zcommand.Params.ParamByName('Cantidad').value       := tdCantidad.Value ;
           Connection.zcommand.Params.ParamByName('Orden').DataType       := ftString ;
           Connection.zcommand.Params.ParamByName('Orden').value          := anexo_requisicion.FieldValues['sNumeroOrden'] ;
           Connection.zcommand.Params.ParamByName('Costo').DataType       := ftFloat ;
           Connection.zcommand.Params.ParamByName('Costo').value          := Consumibles.FieldValues['dCostoMN'];
           Connection.zcommand.ExecSql ;
           connection.QryBusca2.Next;
      end;
end;

procedure TfrmRequisicionPerf.Seguimiento_Material(dParamActividad: string);
var
   x, y, num, i : integer;
   SumCantidad, SumTotal, SumExcedente : double;
   linea : string;
begin
    if dParamActividad <> '' then
       linea := ' and sNumeroActividad =:Actividad '
    else
       linea := '';

    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select sContrato, sWbs, sNumeroActividad, mDescripcion as DescripcionAnexo, '+
                                'dVentaMN, dVentaDLL, dCantidadAnexo, sMedida as sMedidaAnexo  from actividadesxanexo '+
                                'where sContrato =:Contrato '+ linea +'and sTipoActividad = "Actividad" and sIdConvenio =:Convenio order by iItemOrden ');
    connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
    connection.zCommand.ParamByName('Convenio').AsString := global_convenio;
    if dParamActividad <> '' then
       connection.zCommand.ParamByName('Actividad').AsString := dParamActividad;
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
    begin
         rxSeguimiento_Mat.Active := True;
         rxSeguimiento_Mat.EmptyTable;
         //Inicualizacion de vector en 0...
         i := 1;

         while not connection.zCommand.Eof do
         begin
              SumCantidad  := 0;
              SumTotal     := 0;
              SumExcedente := 0;

              rxSeguimiento_Mat.Append;
              rxSeguimiento_Mat.FieldValues['sContrato']        := global_contrato;
              rxSeguimiento_Mat.FieldValues['Partida']          := connection.zCommand.FieldValues['sNumeroActividad'];
              rxSeguimiento_Mat.FieldValues['sNumeroActividad'] := connection.zCommand.FieldValues['sNumeroActividad'];
              rxSeguimiento_Mat.FieldValues['DescripcionAnexo'] := connection.zCommand.FieldValues['DescripcionAnexo'];
              rxSeguimiento_Mat.FieldValues['CantidadAnexo']    := connection.zCommand.FieldValues['dCantidadAnexo'];
              rxSeguimiento_Mat.FieldValues['MedidaAnexo']      := connection.zCommand.FieldValues['sMedidaAnexo'];
              rxSeguimiento_Mat.FieldValues['CostoMNAnexo']     := connection.zCommand.FieldValues['dVentaMN'];
              rxSeguimiento_Mat.FieldValues['CostoDLLAnexo']    := connection.zCommand.FieldValues['dVentaDLL'];
              rxSeguimiento_Mat.FieldValues['Tipo']             := 'Anexo';
              rxSeguimiento_Mat.Post;

              //R E Q U I S I C I O N E S .... <<ivan>>
              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select  ra.sIdInsumo, ra.dCantidad, ra.dCostoMN, ra.dCostoDLL, i.mDescripcion as Descripcion, '+
                                          'i.sMedida, ap.iFolioRequisicion, ap.iItem, SUM(ap.dCantidad) as dCantidadReq  from recursosanexosnuevos ra '+
                                          'left join insumos i '+
                                          'on (i.sContrato = ra.sContrato and i.sIdInsumo = ra.sIdInsumo and i.sIdAlmacen = "") '+
                                          'left join anexo_prequisicion ap '+
                                          'on (ap.sContrato = ra.sContrato and ap.sNumeroActividad = ra.sNumeroActividad and ap.sIdInsumo = ra.sIdInsumo and ap.iFolioPedido = 0 ) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo ');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;


              if connection.QryBusca.RecordCount > 0 then
              begin
                   num  := 0;
                   while not connection.QryBusca.Eof do
                   begin
                        if num = 0 then
                            rxSeguimiento_Mat.Edit
                        else
                            rxSeguimiento_Mat.Append;
                        rxSeguimiento_Mat.FieldValues['Tipo']            := 'Requisicion';
                        rxSeguimiento_Mat.FieldValues['Id']              := connection.QryBusca.FieldValues['sIdInsumo'];
                        rxSeguimiento_Mat.FieldValues['Descripcion']     := connection.QryBusca.FieldValues['Descripcion'];
                        rxSeguimiento_Mat.FieldValues['Unidad']          := connection.QryBusca.FieldValues['sMedida'];
                        rxSeguimiento_Mat.FieldValues['Cantidad']        := connection.QryBusca.FieldValues['dCantidad'];
                        rxSeguimiento_Mat.FieldValues['CostoMN']         := connection.QryBusca.FieldValues['dCostoMN'];
                        rxSeguimiento_Mat.FieldValues['CostoDLL']        := connection.QryBusca.FieldValues['dCostoDLL'];
                        rxSeguimiento_Mat.FieldValues['FolioReq']        := connection.QryBusca.FieldValues['iFolioRequisicion'];
                        rxSeguimiento_Mat.FieldValues['ItemReq']         := connection.QryBusca.FieldValues['iItem'];
                        rxSeguimiento_Mat.FieldValues['dCantidadReq']    := connection.QryBusca.FieldValues['dCantidadReq'];
                        rxSeguimiento_Mat.FieldValues['dRestanteReq']    := 0;
                        rxSeguimiento_Mat.FieldValues['dExcedenteReq']   := 0;
                        rxSeguimiento_Mat.FieldValues['dPorcentajeReq']  := 100;

                        if connection.QryBusca.FieldValues['dCantidadReq'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dRestanteReq']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadReq'];

                        if connection.QryBusca.FieldValues['dCantidadReq'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dExcedenteReq'] := connection.QryBusca.FieldValues['dCantidadReq'] - connection.QryBusca.FieldValues['dCantidad'];

                        if connection.QryBusca.FieldValues['dCantidadReq'] < connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dPorcentajeReq']:= (connection.QryBusca.FieldValues['dCantidadReq'] / connection.QryBusca.FieldValues['dCantidad']) * 100;

                        rxSeguimiento_Mat.FieldValues['sNumeroActividad'] := connection.zCommand.FieldValues['sNumeroActividad'];

                        if Not (rxSeguimiento_Mat.FieldValues['Cantidad'] = null ) then
                           SumCantidad  := SumCantidad + rxSeguimiento_Mat.FieldValues['Cantidad'];

                        if Not (rxSeguimiento_Mat.FieldValues['dCantidadReq'] = null ) then
                           SumTotal     := SumTotal + rxSeguimiento_Mat.FieldValues['dCantidadReq'];

                        if Not (rxSeguimiento_Mat.FieldValues['dExcedenteReq'] = null) then
                           SumExcedente := SumExcedente + rxSeguimiento_Mat.FieldValues['dExcedenteReq'];

                        rxSeguimiento_Mat.Post;
                        connection.QryBusca.Next;
                        num := 1;
                   end;
              end;
              num := connection.QryBusca.RecordCount - 1;
              for x := 1 to num do
                  rxSeguimiento_Mat.Prior;

              for x := 0 to num do
              begin
                  rxSeguimiento_Mat.Edit;
                  rxSeguimiento_Mat.FieldValues['dPorcentajeReq_T']  := ((SumTotal - SumExcedente)/SumCantidad)* 100;
                  rxSeguimiento_Mat.Post;
                  rxSeguimiento_Mat.Next;
              end;

              for x := 1 to num do
                  rxSeguimiento_Mat.Prior;

              //O R D E N E S  D E   C O M P R A ....
              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select  ra.sIdInsumo, ra.dCantidad, ap.iFolioPedido, ap.iItem, SUM(ap.dCantidad) as dCantidadOC  from recursosanexosnuevos ra '+
                                          'left join anexo_ppedido ap '+
                                          'on (ap.sContrato = ra.sContrato and ap.sNumeroActividad = ra.sNumeroActividad and ap.sIdInsumo = ra.sIdInsumo) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;

              SumTotal     := 0;
              SumExcedente := 0;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat.Edit;
                        rxSeguimiento_Mat.FieldValues['dCantidadOC']    := connection.QryBusca.FieldValues['dCantidadOC'];
                        rxSeguimiento_Mat.FieldValues['dRestanteOC']    := 0;
                        rxSeguimiento_Mat.FieldValues['dExcedenteOC']   := 0;
                        if not connection.QryBusca.Fieldbyname('dCantidadOC').IsNull then
                          rxSeguimiento_Mat.FieldValues['dPorcentajeOC']  := 100;

                        if connection.QryBusca.FieldValues['dCantidadOC'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dRestanteOC']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadOC'];

                        if connection.QryBusca.FieldValues['dCantidadOC'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dExcedenteOC'] := connection.QryBusca.FieldValues['dCantidadOC'] - connection.QryBusca.FieldValues['dCantidad'];

                        if Not (rxSeguimiento_Mat.FieldValues['dCantidadOC'] = null ) then
                           SumTotal     := SumTotal + rxSeguimiento_Mat.FieldValues['dCantidadOC'];

                        if Not (rxSeguimiento_Mat.FieldValues['dExcedenteOC'] = null ) then
                           SumExcedente := SumExcedente + rxSeguimiento_Mat.FieldValues['dExcedenteOC'];

                        rxSeguimiento_Mat.Post;
                        rxSeguimiento_Mat.Next;
                        connection.QryBusca.Next;
                   end;
              end;

              for x := 1 to num do
                  rxSeguimiento_Mat.Prior;

              for x := 0 to num do
              begin
                  rxSeguimiento_Mat.Edit;
                  rxSeguimiento_Mat.FieldValues['dPorcentajeOC_T']  := ((SumTotal - SumExcedente)/SumCantidad)* 100;
                  rxSeguimiento_Mat.Post;
                  rxSeguimiento_Mat.Next;
              end;

              for x := 1 to num do
                  rxSeguimiento_Mat.Prior;

               // E N T R A D A  D E  M A T E R I A L E S ....
              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select  ra.sIdInsumo, ra.dCantidad, ap.iFolioPedido, ap.iItem, SUM(ap.dCantidad) as dCantidadIn  from recursosanexosnuevos ra '+
                                          'left join bitacoradeentrada  ap '+
                                          'on (ap.sContrato = ra.sContrato and ap.sNumeroActividad = ra.sNumeroActividad and ap.sIdInsumo = ra.sIdInsumo) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat.Edit;
                        rxSeguimiento_Mat.FieldValues['dCantidadIn']    := connection.QryBusca.FieldValues['dCantidadIn'];
                        rxSeguimiento_Mat.FieldValues['dExcedenteIn']   := 0;

                        if connection.QryBusca.FieldValues['dCantidadIn'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dRestanteIn']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadIn'];

                        if connection.QryBusca.FieldValues['dCantidadIn'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dExcedenteIn'] := connection.QryBusca.FieldValues['dCantidadIn'] - connection.QryBusca.FieldValues['dCantidad'];

                        rxSeguimiento_Mat.Post;
                        rxSeguimiento_Mat.Next;
                        connection.QryBusca.Next;
                   end;
              end;

              for x := 1 to num do
                  rxSeguimiento_Mat.Prior;
               // S A L I D A  D E  M A T E R I A L E S ....
              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select  ra.sIdInsumo, ra.dCantidad, ap.iFolioSalida, SUM(ap.dCantidad) as dCantidadOut  from recursosanexosnuevos ra '+
                                          'left join bitacoradesalida  ap '+
                                          'on (ap.sContrato = ra.sContrato and ap.sNumeroActividad = ra.sNumeroActividad and ap.sIdInsumo = ra.sIdInsumo) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat.Edit;
                        rxSeguimiento_Mat.FieldValues['dCantidadOut']    := connection.QryBusca.FieldValues['dCantidadOut'];
                        rxSeguimiento_Mat.FieldValues['dExcedenteOut']   := 0;

                        if connection.QryBusca.FieldValues['dCantidadOut'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dRestanteOut']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadOut'];

                        if connection.QryBusca.FieldValues['dCantidadOut'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dExcedenteOut'] := connection.QryBusca.FieldValues['dCantidadOut'] - connection.QryBusca.FieldValues['dCantidad'];

                        rxSeguimiento_Mat.Post;
                        rxSeguimiento_Mat.Next;
                        connection.QryBusca.Next;
                   end;
              end;

              for x := 1 to num do
                  rxSeguimiento_Mat.Prior;
              // R E P O R T E S   D I A R I O S ....
              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('select ra.dCantidad, SUM(bmp.dCantidad) as dCantidadRD, COUNT(rd.sNumeroReporte) as total  from recursosanexosnuevos ra '+
                                          'inner join bitacoradeactividades ba '+
                                          'on (ba.sContrato = ra.sContrato  and ba.sWbs = ra.sWbs and ba.sNumeroActividad = ra.sNumeroActividad) '+
                                          'left join bitacorademateriales  bmp '+
                                          'on(bmp.sContrato = ra.sContrato and bmp.dIdFecha = ba.dIdFecha and bmp.iIdDiario = ba.iIdDiario and bmp.sIdMaterial = ra.sIdInsumo) '+
                                          'inner join reportediario rd '+
                                          'on (rd.sContrato = ba.sContrato and rd.dIdFecha = ba.dIdFecha and rd.sIdTurno = ba.sIdTurno and rd.sNumeroOrden = ba.sNumeroOrden ) '+
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.Open;

              SumTotal     := 0;
              SumExcedente := 0;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat.Edit;
                        rxSeguimiento_Mat.FieldValues['dCantidadRD']     := connection.QryBusca.FieldValues['dCantidadRD'];
                        rxSeguimiento_Mat.FieldValues['dRestanteRD']     := 0;
                        rxSeguimiento_Mat.FieldValues['dExcedenteRD']    := 0;
                        rxSeguimiento_Mat.FieldValues['dPorcentajeRD']   := 100;

                        if connection.QryBusca.FieldValues['dCantidadRD'] <= connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dRestanteRD']  := connection.QryBusca.FieldValues['dCantidad'] - connection.QryBusca.FieldValues['dCantidadRD'];

                        if connection.QryBusca.FieldValues['dCantidadRD'] > connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dExcedenteRD'] := connection.QryBusca.FieldValues['dCantidadRD'] - connection.QryBusca.FieldValues['dCantidad'];

                        if connection.QryBusca.FieldValues['dCantidadRD'] < connection.QryBusca.FieldValues['dCantidad'] then
                           rxSeguimiento_Mat.FieldValues['dPorcentajeRD']:= (connection.QryBusca.FieldValues['dCantidadRD'] / connection.QryBusca.FieldValues['dCantidad']) * 100;

                        if Not (rxSeguimiento_Mat.FieldValues['dCantidadRD'] = null ) then
                           SumTotal     := SumTotal + rxSeguimiento_Mat.FieldValues['dCantidadRD'];

                        if Not (rxSeguimiento_Mat.FieldValues['dExcedenteRD'] = null ) then
                           SumExcedente := SumExcedente + rxSeguimiento_Mat.FieldValues['dExcedenteRD'];

                        rxSeguimiento_Mat.Post;
                        rxSeguimiento_Mat.Next;
                        connection.QryBusca.Next;
                   end;
              end;

              for x := 1 to num do
                  rxSeguimiento_Mat.Prior;

              for x := 0 to num do
              begin
                  rxSeguimiento_Mat.Edit;
                  rxSeguimiento_Mat.FieldValues['dPorcentajeRD_T']  := ((SumTotal - SumExcedente)/SumCantidad)* 100;
                  rxSeguimiento_Mat.Post;
                  rxSeguimiento_Mat.Next;
              end;

              for x := 1 to num do
                  rxSeguimiento_Mat.Prior;
              // G E N E R A D O R E S  D E  O B R A ....
              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('Select e.sContrato, aa.sNumeroActividad, sum(e.dCantidad) as dCantidad, '+
                                          'e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador '+
                                          'from actividadesxanexo aa '+
                                          'inner join  estimacionxpartida e '+
                                          'on (e.sContrato = aa.sContrato and e.sWbs = aa.sWbs and e.sNumeroActividad = aa.sNumeroActividad) '+
                                          'inner join estimaciones e2 '+
                                          'on (e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And e.sNumeroGenerador = e2.sNumeroGenerador) '+
                                          'inner join estimacionperiodo e3 '+
                                          'on (e2.sContrato = e3.sContrato And e2.iNumeroEstimacion = e3.iNumeroEstimacion) '+
                                          'where aa.sContrato =:Contrato and aa.sNumeroActividad =:Actividad and aa.sWbs =:Wbs and sIdConvenio =:Convenio '+
                                          'group by aa.sNumeroActividad ');
              connection.QryBusca.ParamByName('Contrato').AsString   := global_contrato;
              connection.QryBusca.ParamByName('Convenio').AsString   := global_convenio;
              connection.QryBusca.ParamByName('Actividad').AsString  := connection.zCommand.FieldValues['sNumeroActividad'];
              connection.QryBusca.ParamByName('Wbs').AsString        := connection.zCommand.FieldValues['sWbs'];
              connection.QryBusca.Open;

              if connection.QryBusca.RecordCount > 0 then
              begin
                   while not connection.QryBusca.Eof do
                   begin
                        rxSeguimiento_Mat.Edit;
                        rxSeguimiento_Mat.FieldValues['dCantidadGen']      := connection.QryBusca.FieldValues['dCantidad'];
                        rxSeguimiento_Mat.FieldValues['iNumeroEstimacion'] := connection.QryBusca.FieldValues['iNumeroEstimacion'];
                        rxSeguimiento_Mat.FieldValues['sNumeroOrden']      := connection.QryBusca.FieldValues['sNumeroOrden'];
                        rxSeguimiento_Mat.FieldValues['sNumeroGenerador']  := connection.QryBusca.FieldValues['sNumeroGenerador'];
                        rxSeguimiento_Mat.FieldValues['dExcedenteGen']     := 0;


                        rxSeguimiento_Mat.Post;
                        rxSeguimiento_Mat.Next;
                        connection.QryBusca.Next;
                   end;
              end;
              connection.zCommand.Next;
              i := i + 1;
         end;
    end;
end;

procedure TfrmRequisicionPerf.formatoEncabezado;
begin
    Excel.Selection.MergeCells := False;
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.VerticalAlignment := xlCenter;
    Excel.Selection.Font.Size := 12;
    Excel.Selection.Font.Bold := False;
    Excel.Selection.Font.Name := 'Calibri';
end;

End.
