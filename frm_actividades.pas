unit frm_actividades;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, global, frm_barra, db, Grids, DBGrids, frm_Connection, StdCtrls,
  DBCtrls, ComCtrls, Mask, Utilerias, masUtilerias, StrUtils, Menus, 
  ExtCtrls, DateUtils, frxClass, frxDBSet, RXDBCtrl, RxLookup, Math,
  Newpanel, Buttons, ZDataset, ZAbstractRODataset, ZAbstractDataset, Gauges,
  rxToolEdit, rxCurrEdit, OleServer, ComObj, Excel2000, RxMemDS,
  UnitExcel, unitexcepciones, udbgrid,ShellAPI, unitMetodos,
  UnitTBotonesPermisos, unitactivapop, UnitValidacion, Frm_PopUpReprogramacion,
  AdvDateTimePicker, AdvDBDateTimePicker, ShlObj,jpeg, AdvGlowButton,
  AdvSmoothPanel, AdvCombo, AdvProgressBar, JvExDBGrids, JvDBGrid, JvDBUltimGrid,
  JvExComCtrls, JvDateTimePicker, JvDBDateTimePicker, JvExMask, JvToolEdit,
  JvMaskEdit, JvDBControls, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore, dxSkinDevExpressStyle,
  dxSkinFoggy, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxDBEdit;

type
  Tfirmas = class
    private
      FInicialCia,FInicialPep:String;
      FFinalCia,FFinalPep:String;
  end;

  TfrmActividades = class(TForm)
    frmBarra1: TfrmBarra;
    ds_ordenesdetrabajo: TDataSource;
    Label1: TLabel;
    tsNumeroOrden: TDBLookupComboBox;
    ds_actividadesxorden: TDataSource;
    dbActividadesxOrden: TfrxDBDataset;
    grid_actividades: TRxDBGrid;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    Imprimir1: TMenuItem;
    Salir1: TMenuItem;
    ds_Paquetes: TDataSource;
    BuscarPartida1: TMenuItem;
    dsOrden: TfrxDBDataset;
    dsResumen: TfrxDBDataset;
    ActividadesxOrden: TZQuery;
    ordenesdetrabajo: TZReadOnlyQuery;
    ordenesdetrabajosNumeroOrden: TStringField;
    ordenesdetrabajosDescripcionCorta: TStringField;
    ordenesdetrabajosIdTipoOrden: TStringField;
    Paquetes: TZReadOnlyQuery;
    PaquetesiNivel: TIntegerField;
    PaquetessWBS: TStringField;
    PaquetessWBSAnterior: TStringField;
    PaquetessPaquete: TStringField;
    PaquetessNumeroActividad: TStringField;
    PaquetesmDescripcion: TMemoField;
    PaquetesdFechaInicio: TDateField;
    PaquetesdFechaFinal: TDateField;
    PaquetesdDuracion: TFloatField;
    PaquetessDescripcion: TStringField;
    PonderarConceptos: TMenuItem;
    Progress: TGauge;
    PaquetesiItemOrden: TStringField;
    ActividadesxOrdensContrato: TStringField;
    ActividadesxOrdensIdConvenio: TStringField;
    ActividadesxOrdensNumeroOrden: TStringField;
    ActividadesxOrdeniNivel: TIntegerField;
    ActividadesxOrdensSimbolo: TStringField;
    ActividadesxOrdensWbs: TStringField;
    ActividadesxOrdensWbsAnterior: TStringField;
    ActividadesxOrdensPaquete: TStringField;
    ActividadesxOrdensNumeroActividad: TStringField;
    ActividadesxOrdensTipoActividad: TStringField;
    ActividadesxOrdeniItemOrden: TStringField;
    ActividadesxOrdenmDescripcion: TMemoField;
    ActividadesxOrdensHoraInicio: TStringField;
    ActividadesxOrdendDuracion: TFloatField;
    ActividadesxOrdensHoraFinal: TStringField;
    ActividadesxOrdendPonderado: TFloatField;
    ActividadesxOrdendCantidad: TFloatField;
    ActividadesxOrdendCargado: TFloatField;
    ActividadesxOrdendInstalado: TFloatField;
    ActividadesxOrdendExcedente: TFloatField;
    ActividadesxOrdendVentaMN: TFloatField;
    ActividadesxOrdendVentaDLL: TFloatField;
    ActividadesxOrdensIdPlataforma: TStringField;
    ActividadesxOrdensIdPernocta: TStringField;
    ActividadesxOrdenmComentarios: TMemoField;
    ActividadesxOrdenlGerencial: TStringField;
    ActividadesxOrdenlCalculo: TStringField;
    ActividadesxOrdeniColor: TIntegerField;
    ActividadesxOrdenlGenerado: TStringField;
    ActividadesxOrdenlCancelada: TStringField;
    ActividadesxOrdendMontoMN: TCurrencyField;
    ActividadesxOrdendMontoDLL: TCurrencyField;
    ActividadesxOrdensWbsSpace: TStringField;
    tNewGroupBox1: tNewGroupBox;
    Label2: TLabel;
    Label18: TLabel;
    EtiquetaPU2: TLabel;
    Label17: TLabel;
    Label5: TLabel;
    Label11: TLabel;
    tdVentaMN: TRxDBCalcEdit;
    tdCantidad: TRxDBCalcEdit;
    tsNumeroActividad: TDBEdit;
    tmDescripcion: TDBMemo;
    tiColor: TDBComboBox;
    tiColores: TColorBox;
    tsUnidad: TDBEdit;
    tNewGroupBox2: tNewGroupBox;
    Label4: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label3: TLabel;
    tdCostoMN: TRxDBCalcEdit;
    tNewGroupBox3: tNewGroupBox;
    tmComentarios: TDBMemo;
    ActividadesxOrdendCostoMN: TFloatField;
    ActividadesxOrdendCostoDLL: TFloatField;
    ActividadesxOrdendMontoCostoMN: TCurrencyField;
    ActividadesxOrdendMontoCostoDLL: TCurrencyField;
    InsertarConceptos: TMenuItem;
    tlGerencial: TDBCheckBox;
    ordenesdetrabajodFechaInicioT: TDateField;
    ordenesdetrabajodFechaSitioM: TDateField;
    ordenesdetrabajosDepSolicitante: TStringField;
    ordenesdetrabajodfiProgramado: TDateField;
    ordenesdetrabajodffProgramado: TDateField;
    ordenesdetrabajosIdPlataforma: TStringField;
    ordenesdetrabajosEquipo: TStringField;
    ordenesdetrabajosPozo: TStringField;
    ordenesdetrabajodFechaElaboracion: TDateField;
    frxCentrosCosto: TfrxDBDataset;
    qryCentrocostos: TZReadOnlyQuery;
    dtsCentroCostos: TDataSource;
    ordenesdetrabajosPuestoPEP: TStringField;
    ordenesdetrabajosFirmantePEP: TStringField;
    ordenesdetrabajosPuestoCia: TStringField;
    ordenesdetrabajosFirmantecia: TStringField;
    QryConcentradoOt: TZReadOnlyQuery;
    frxDBDataset1: TfrxDBDataset;
    QryConcentradoOtsNumeroActividad: TStringField;
    QryConcentradoOtsNumeroOrden: TStringField;
    QryConcentradoOtsMedida: TStringField;
    QryConcentradoOtdCantidad: TFloatField;
    QryConcentradoOtdVentaMn: TFloatField;
    frxReporte: TfrxReport;
    Label20: TLabel;
    tlCalculo: TDBComboBox;
    ActividadesxOrdensMedida: TStringField;
    tdPonderado: TDBEdit;
    Label21: TLabel;
    AdministraciondelPrograma1: TMenuItem;
    N7: TMenuItem;
    ActividadesxOrdensWbsContrato: TStringField;
    progreso: TProgressBar;
    Label10: TLabel;
    ActividadesxOrdensTipoAnexo: TStringField;
    ExportaaExcel1: TMenuItem;
    SaveDialog1: TSaveDialog;
    frxDBValida: TfrxDBDataset;
    RxMDValida: TRxMemoryData;
    RxMDValidasNumeroActividad: TStringField;
    RxMDValidasWbs: TStringField;
    RxMDValidadCantidad: TStringField;
    RxMDValidasuma: TStringField;
    RxMDValidaaMN: TStringField;
    RxMDValidaaDLL: TStringField;
    RxMDValidabMN: TStringField;
    RxMDValidabDLL: TStringField;
    RxMDValidadCantidadAnexo: TStringField;
    RxMDValidadescripcion: TStringField;
    RxMDValidamensaje: TStringField;
    RxMDValidasNumeroOrden: TStringField;
    RxMDValidasWbs2: TStringField;
    frxReport1: TfrxReport;
    Label12: TLabel;
    tsActAnterior: TDBEdit;
    ActividadesxOrdensActividadAnterior: TStringField;
    ExportaVolumenesExcel1: TMenuItem;
    ActividadesxOrdensAnexo: TStringField;
    Label13: TLabel;
    N6: TMenuItem;
    N8: TMenuItem;
    PanelProgress: TPanel;
    Label15: TLabel;
    Label16: TLabel;
    Label14: TLabel;
    Label19: TLabel;
    BarraEstado: TProgressBar;
    ActividadesxOrdenNewSimbol: TStringField;
    ActividadesxOrdenDescripcion: TStringField;
    N9: TMenuItem;
    ProgramaDiariodelConceptodelaCia1: TMenuItem;
    ActividadesxOrdensIdFase: TStringField;
    Com1: TMenuItem;
    Label27: TLabel;
    tsIdFase: TDBLookupComboBox;
    zFasesProyecto: TZReadOnlyQuery;
    dsFasesxProyecto: TDataSource;
    tsTipoAnexo: TDBComboBox;
    Label35: TLabel;
    tdFechaInicio: TAdvDBDateTimePicker;
    tdFechaFinal: TAdvDBDateTimePicker;
    CampoDuracion: TEdit;
    ActividadesxOrdendDiferenciaDuracion: TStringField;
    ActividadesxOrdendFechaInicio: TDateTimeField;
    ActividadesxOrdendFechaFinal: TDateTimeField;
    SdgExcel: TSaveDialog;
    DBEdit1: TDBEdit;
    lbl1: TLabel;
    ds_ProgramaDeActividad: TDataSource;
    zq_ProgramaDeActividad: TZQuery;
    lbl2: TLabel;
    lbl3: TLabel;
    lbl4: TLabel;
    lbl5: TLabel;
    edt1: TEdit;
    btn1: TButton;
    ActividadesxOrdensDuracionHoras: TStringField;
    chklFactorBarco: TDBCheckBox;
    ActividadesxOrdenlFactorBarco: TStringField;
    tsItemOrden: TDBEdit;
    Label34: TLabel;
    chkItemOrden: TCheckBox;
    AVANCESGLOBALESXFOLIOPARTIDAS1: TMenuItem;
    PnlRango: TAdvSmoothPanel;
    DFechaInicio: TAdvDateTimePicker;
    DFechaFin: TAdvDateTimePicker;
    BtnImprime: TAdvGlowButton;
    CmbFolio: TAdvComboBox;
    BtnSalir: TAdvGlowButton;
    ChbPrevisualizar: TCheckBox;
    GuardaExcel: TSaveDialog;
    PbExcel: TAdvProgressBar;
    Label6: TLabel;
    ChkAcumulativo: TCheckBox;
    rxdbgrd1: TJvDBUltimGrid;
    jDbDtpFechaI: TJvDBDateTimePicker;
    jDbDtpFechaT: TJvDBDateTimePicker;
    JDbMEdtHInicio: TJvDBMaskEdit;
    JDbMEdtHTermino: TJvDBMaskEdit;
    DBRadioGroup1: TDBRadioGroup;
    ActividadesxOrdenlExtraordinario: TStringField;
    GraficadeGantt1: TMenuItem;
    Label9: TLabel;
    ds_plataformas: TDataSource;
    zqPlataformas: TZReadOnlyQuery;
    Label22: TLabel;
    tsReprogramacion: TDBLookupComboBox;
    zqReprogramacion: TZReadOnlyQuery;
    ds_reprogramacion: TDataSource;
    ReprogramarFolio1: TMenuItem;
    zqCopiaReprogramacion: TZReadOnlyQuery;
    tsPlataforma: TDBLookupComboBox;
    tsPaquete: TRxDBLookupCombo;
    procedure FormShow(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroActividadKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaInicioKeyPress(Sender: TObject; var Key: Char);
    procedure tdDuracionKeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tdDuracionExit(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tdCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure tdCantidadAnexoKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroActividadExit(Sender: TObject);
    procedure tdFechaInicioExit(Sender: TObject);
    procedure tsPaqueteKeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra1btnPrinterClick(Sender: TObject);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsPaqueteEnter(Sender: TObject);
    procedure tsPaqueteExit(Sender: TObject);
    procedure tmDescripcionEnter(Sender: TObject);
    procedure tmDescripcionExit(Sender: TObject);
    procedure tsNumeroActividadEnter(Sender: TObject);
    procedure tdFechaInicioEnter(Sender: TObject);
    procedure tdDuracionEnter(Sender: TObject);
    procedure tdFechaFinalEnter(Sender: TObject);
    procedure tdFechaFinalExit(Sender: TObject);
    procedure tdCantidadEnter(Sender: TObject);
    procedure tdCantidadExit(Sender: TObject);
    procedure tiColorEnter(Sender: TObject);
    procedure tiColorExit(Sender: TObject);
    procedure MenuItem9Click(Sender: TObject);
    procedure PartidasBeforeDelete(DataSet: TDataSet);
    procedure ActividadesxOrdenAfterScroll(DataSet: TDataSet);
    procedure grid_actividadesGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure PartidasBeforeInsert(DataSet: TDataSet);
    procedure tiColoresEnter(Sender: TObject);
    procedure tiColoresExit(Sender: TObject);
    procedure tiColoresKeyPress(Sender: TObject; var Key: Char);
    procedure tiColoresChange(Sender: TObject);
    procedure PaquetesCalcFields(DataSet: TDataSet);
    procedure BuscarPartida1Click(Sender: TObject);
    procedure ActividadesxOrdenAfterInsert(DataSet: TDataSet);
    procedure tdVentaMNEnter(Sender: TObject);
    procedure tdVentaMNExit(Sender: TObject);
    procedure tdVentaMNKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaFinalKeyPress(Sender: TObject; var Key: Char);
    procedure tmComentariosEnter(Sender: TObject);
    procedure tmComentariosExit(Sender: TObject);
    procedure ActividadesxOrdenCalcFields(DataSet: TDataSet);
    procedure ActividadesxOrdensWbsAnteriorChange(Sender: TField);
    procedure tsUnidadEnter(Sender: TObject);
    procedure tsUnidadExit(Sender: TObject);
    procedure tsUnidadKeyPress(Sender: TObject; var Key: Char);
    procedure grid_actividadesEnter(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure grid_actividadesDblClick(Sender: TObject);
    procedure btnImportarClick(Sender: TObject);
    procedure PonderarConceptosClick(Sender: TObject);
    function lExisteMedida(sMedida: string): Boolean;
    procedure tdCostoMNEnter(Sender: TObject);
    procedure tdCostoMNExit(Sender: TObject);
    procedure tdCostoMNKeyPress(Sender: TObject; var Key: Char);
    procedure InsertarConceptosClick(Sender: TObject);
    procedure ActividadesxOrdensTipoActividadChange(Sender: TField);
    procedure tlCalculoEnter(Sender: TObject);
    procedure tlCalculoExit(Sender: TObject);
    procedure tlCalculoKeyPress(Sender: TObject; var Key: Char);
    procedure zProcCancelaInsert(DataSet: TDataSet);
    procedure InsertaActividad(Sender: TObject);
    //*****************************BRITO 25-03-11*******************************
    procedure SeleccionarNuevaActividad(Sender: TObject);
    procedure NuevoPaquete(Sender: TObject);
    procedure PonderadoAnterior;
    procedure PonderadoHorarios;

    procedure procBuscaPartida(Sender: TObject);
    procedure ImportaXLS(Sender: TObject);

    procedure FormClick(Sender: TObject);
    procedure ExportaaExcel1Click(Sender: TObject);
    procedure formatoEncabezado(); overload;
    procedure ExportaVolumenesExcel1Click(Sender: TObject);
    procedure grid_actividadesMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_actividadesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_actividadesTitleClick(Column: TColumn);
    procedure tdPonderadoKeyPress(Sender: TObject; var Key: Char);
    procedure tiColorKeyPress(Sender: TObject; var Key: Char);
    procedure tdPonderadoEnter(Sender: TObject);
    procedure tdPonderadoExit(Sender: TObject);
    procedure tsActAnteriorEnter(Sender: TObject);
    procedure tsActAnteriorExit(Sender: TObject);
    procedure tsActAnteriorKeyPress(Sender: TObject; var Key: Char);
    function SumaCantidades(): boolean;
    procedure grid_actividadesCellClick(Column: TColumn);
    procedure tdCantidadChange(Sender: TObject);
    procedure tdVentaMNChange(Sender: TObject);
    procedure tdCostoMNChange(Sender: TObject);
    procedure ActividadesxOrdendPonderadoSetText(Sender: TField;
      const Text: string);
    procedure tdPonderadoChange(Sender: TObject);
    procedure tdDuracionChange(Sender: TObject);
    procedure ActividadesxOrdenNewSimbolGetText(Sender: TField;
      var Text: string; DisplayText: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure ProgramaDiariodelConceptodelaCia1Click(Sender: TObject);
    procedure ProgramarActividad();
    procedure Com1Click(Sender: TObject);
    procedure GenerarNotadeCampo1Click(Sender: TObject);
    procedure zq_ProgramaDeActividadAfterInsert(DataSet: TDataSet);
    procedure tsItemOrdenEnter(Sender: TObject);
    procedure tsItemOrdenExit(Sender: TObject);
    procedure tsItemOrdenKeyPress(Sender: TObject; var Key: Char);

    Procedure ActualiaFactorGeneradorPER(sParamEmbarcacion, sParamOrden, sParamFolio : string) ;
    Procedure ActualiaFactorGeneradorEQ( sParamEmbarcacion, sParamOrden, sParamFolio : string) ;
    procedure AVANCESGLOBALESXFOLIOPARTIDAS1Click(Sender: TObject);
    procedure BtnImprimeClick(Sender: TObject);
    procedure BtnSalirClick(Sender: TObject);
    procedure CmbFolioChange(Sender: TObject);
    procedure ActividadesxOrdendDuracionChange(Sender: TField);
    procedure ActividadesxOrdendFechaInicioChange(Sender: TField);
    procedure ActividadesxOrdendFechaFinalChange(Sender: TField);
    procedure AgregarObservaciones1Click(Sender: TObject);
    procedure frxReport1GetValue(const VarName: string; var Value: Variant);
    procedure GraficadeGantt1Click(Sender: TObject);
    procedure tsPlataformaEnter(Sender: TObject);
    procedure tsPlataformaExit(Sender: TObject);
    procedure tsPlataformaKeyPress(Sender: TObject; var Key: Char);
    procedure tsReprogramacionExit(Sender: TObject);
    procedure tsReprogramacionEnter(Sender: TObject);
    procedure tsReprogramacionKeyPress(Sender: TObject; var Key: Char);
    procedure ConsultaFolios;
    procedure ConsultaReprogramacion;
    procedure ReprogramarFolio1Click(Sender: TObject);
    procedure tsTipoAnexoKeyPress(Sender: TObject; var Key: Char);
  private
    sMenuP: string;
    Paq: TstringList;
    isOpen,Ciclar:Boolean;
    procedure ImprimirAvxFP(FInicio,FFin:TDateTime;CmbFl:TAdvComboBox;Ver:Boolean);
    procedure ImprimirAvxFPAcumulativo(FInicio, FFin: TDateTime;
      CmbFl: TAdvComboBox; Ver: Boolean);
    procedure GenerarNotaPdasSinBarco(var Excel, Hoja: Variant; var QrDatos,
      QrPdas: TZReadOnlyQuery);

    { Private declarations }
  public
    { Public declarations }
    sIdFrente:string;
    procedure CalcDiferenciasOT(lista: TStringList);
    procedure ventasDiferentes(sActividad, suma: string);
    function cantidadesDiferentes(sActividad: string): string;
    procedure acumularDiferencia(suma, sMensaje: string);
    procedure PopUpNuevoRegistro;
    procedure PonerEncabezado(var Excel,Hoja:Variant;var QrDatos:TZReadOnlyQuery);
    procedure GenerarNotaPdas(var Excel,Hoja:Variant;var QrDatos:TZReadOnlyQuery;var QrPdas:TZReadOnlyQuery);
    Procedure NotaCampoExcel;
    Procedure FormatoEncabezado(var Excel: Variant;Cadena:string; Align: Integer;Negrita:Boolean);overload;
    Procedure FormatoNormal(var Excel: Variant;Cadena:string; Align: Integer;Negrita,Ajustar:Boolean);overload;
    Procedure FormatoNormal(var Excel: Variant;Cadena:VAriant; Align: Integer;Negrita,Ajustar:Boolean;Formato:String;Column:Integer=0);overload;
    procedure GenerarMarco(var Excel: Variant);
    procedure ConfigurarHoja(var excel: Variant; var Hoja: Variant);
    procedure AjustarTexto(var rangoE: Variant;TotalR:Integer);
    Procedure ActualizaDuracion(Paramwbs:String;ParamInicio,ParamTermino:TDateTime);Overload;
    Procedure ActualizaDuracion;Overload;
  end;

var
  frmActividades: TfrmActividades;
  sIdPlataforma: string;
  sIdPernocta: string;
  sNumeroOrden: string;
  sPaquete: string;
  sPaqueteDesc: string;
  iItemOrden: Integer;
  sFiltro: string;
  SavePlace: TBookmark;
  GridProgConstExist: TRxDBGrid;
  GridActividadesxAnexo: TRxDBGrid;
  zActividadesxAnexo: TZReadOnlyQuery;
  lYaPregunto: Boolean;
  sWbsAnterior: string;
  sItemOrden: string;
  iNivel: Byte;
  WbsAnt, ActividadAnt, DescripcionAnt, UnidadAnt, sWbsOrig: string;
  VentaAnt, CostoAnt: double;

  NivelAnt: integer;

  OrdenPaqueteItem, OrdenPaqueteWbs: string;
  OrdenPaqueteNivel: integer;

  utgrid: ticdbgrid;
  //Exporta elementos a Excel..
  Excel, Libro, Hoja: Variant;

  //Matriz de colores
  Colores: array[1..10, 1..2] of integer;
  columnas: array[1..1400] of string;
  BotonPermiso: TBotonesPermisos;
  TablasConsult: array[1..2,1..4] of string=(('bitacoradepersonal','sIdPersonal','personal','PERSONAL'),
                                             ('bitacoradeequipos','sIdEquipo','equipos','EQUIPO'){,
                                             ('bitacoradebarcoxfases','sIdEmbarcacion','BARCO')});
  sOpcion : string;
implementation

uses frm_GraficaGerencial, frm_GraficaGerencialDX, UnitValidaTexto,
  UFunctionsGHH, frm_CostoFrente, Frm_ProgramacionPartidasCia, UnitTarifa,
  Frm_NotaCampoObservaciones;

{$R *.dfm}

Procedure TfrmActividades.ActualizaDuracion;
var
  scalcula:String;
begin
  scalcula := 'Si';
  Connection.QryBusca2.Active := False;
  Connection.QryBusca2.SQL.Clear;
  Connection.QryBusca2.SQL.Add('Select Distinct sWBS From actividadesxorden ' +
  'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Paquete" Order By iNivel DESC');
  Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
  Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_contrato;
  Connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString;
  Connection.QryBusca2.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
  Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
  Connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
  Connection.QryBusca2.Open;
  while not Connection.QryBusca2.Eof do
  begin
    Connection.QryBusca.Active := False;
    Connection.QryBusca.Filtered := False;
    Connection.QryBusca.SQL.Clear;
    Connection.QryBusca.SQL.Add('Select Min(dFechaInicio) as dFechaInicio, Max(dFechaFinal) as dFechaFinal, sum(dPonderado) as dPonderado, ' +
      'sum(dCantidad * dVentaMN) as dMontoMN, sum(dCantidad * dVentaDLL) as dMontoDLL From actividadesxorden ' +
      'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sWBSAnterior = :Paquete ' +
      'and lcalculo=:calculo Group By sWBSAnterior');
    Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
    Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
    Connection.QryBusca.Params.ParamByName('Paquete').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Paquete').Value := Connection.QryBusca2.FieldValues['sWBS'];
    Connection.QryBusca.Params.ParamByName('calculo').AsString := sCalcula;
    Connection.QryBusca.Open;
    if Connection.QryBusca.RecordCount > 0 then
      if (not Connection.QryBusca.FieldByName('dFechaInicio').IsNull) and (not Connection.QryBusca.FieldByName('dFechaFinal').IsNull) then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.Filtered := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Update actividadesxorden SET dFechaInicio = :Inicio, dFechaFinal = :Final, dPonderado = :Ponderado, dVentaMN = :MontoMN, dVentaDLL = :MontoDLL ' +
          ',dDuracion=:Duracion Where sContrato = :Contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden And ' +
          'sWBS = :Paquete And sTipoActividad = "Paquete"');
        connection.zcommand.params.ParamByName('Contrato').DataType := ftString;
        connection.zcommand.params.ParamByName('Contrato').Value := global_contrato;
        connection.zcommand.params.ParamByName('Convenio').DataType := ftString;
        connection.zcommand.params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        connection.zcommand.params.ParamByName('Orden').DataType := ftString;
        connection.zcommand.params.ParamByName('Orden').Value := tsNumeroOrden.Text;
        connection.zcommand.params.ParamByName('Paquete').DataType := ftString;
        connection.zcommand.params.ParamByName('Paquete').Value := Connection.QryBusca2.FieldValues['sWBS'];
        connection.zcommand.params.ParamByName('Inicio').DataType := ftDate;
        connection.zcommand.params.ParamByName('Inicio').Value := Connection.QryBusca.FieldValues['dFechaInicio'];
        connection.zcommand.params.ParamByName('Final').DataType := ftDate;
        connection.zcommand.params.ParamByName('Final').Value := Connection.QryBusca.FieldValues['dFechaFinal'];
        connection.zcommand.ParamByName('Duracion').AsInteger:=DaysBetween(Connection.QryBusca.FieldByName('dFechaFinal').AsDateTime,Connection.QryBusca.FieldByName('dFechaInicio').AsDateTime);
        connection.zcommand.params.ParamByName('Ponderado').DataType := ftFloat;
       // if roundTo(Connection.QryBusca.FieldValues['dPonderado'], -2) >= 100 then
       //   connection.zcommand.params.ParamByName('Ponderado').Value := 100
       // else
        connection.zcommand.params.ParamByName('Ponderado').Value := Connection.QryBusca.FieldValues['dPonderado'];
        connection.zcommand.params.ParamByName('MontoMN').DataType := ftFloat;
        connection.zcommand.params.ParamByName('MontoMN').Value := Connection.QryBusca.FieldValues['dMontoMN'];
        connection.zcommand.params.ParamByName('MontoDLL').DataType := ftFloat;
        connection.zcommand.params.ParamByName('MontoDLL').Value := Connection.QryBusca.FieldValues['dMontoDLL'];
        Connection.zCommand.ExecSQL;
      end;
    Connection.QryBusca2.Next
  end;


end;

Procedure TfrmActividades.ActualizaDuracion(Paramwbs: string;ParamInicio,ParamTermino:TDateTime);
var
  QPadre:TzQuery;
  bAct:Boolean;
begin
  QPadre:=TzQuery.Create(nil);
  bAct:=false;
  try
    QPadre.Connection:=Connection.zConnection;
    QPadre.SQL.Text:= 'select * from actividadesxorden where sContrato=:Contrato and '+
                      'sIdConvenio=:Convenio and sNumeroOrden=:Orden and swbs=:wbs';
    QPadre.ParamByName('Contrato').AsString:=Global_Contrato;
    QPadre.ParamByName('Convenio').AsString:=zqReprogramacion.FieldByName('sIdConvenio').AsString;
    QPadre.ParamByName('Orden').AsString:=TsNumeroOrden.KeyValue;
    QPadre.ParamByName('wbs').AsString:=Paramwbs;
    QPadre.Open;
    if QPadre.RecordCount=1 then
    begin
      QPadre.Edit;
      if ParamInicio<QPadre.FieldByName('dFechaInicio').AsDateTime then
      begin
        QPadre.FieldByName('dFechaInicio').AsDateTime:=ParamInicio;
        bAct:=true;
      end;

      if ParamTermino>QPadre.FieldByName('dFechaFinal').AsDateTime then
      begin
        QPadre.FieldByName('dFechaFinal').AsDateTime:=ParamTermino;
        bAct:=true;
      end;

      if bAct then
      begin
         QPadre.FieldByName('dDuracion').AsInteger:=DaysBetween(QPadre.FieldByName('dFechaFinal').NewValue,
                                                    QPadre.FieldByName('dFechaInicio').NewValue) + 1;
         QPadre.Post;
         ActualizaDuracion(QPadre.FieldByName('swbsAnterior').AsString,
                        ParamInicio,ParamTermino);
      end;

    end;
  finally
    QPadre.Destroy;
  end;


end;
procedure TfrmActividades.AjustarTexto(var rangoE: Variant;TotalR:Integer);
var
  sngAnchoTotal,sngAnchoCelda,sngAlto:Extended;
  n:Integer;
begin
  sngAnchoTotal:=0;
  For n := 1 To TotalR do
    sngAnchoTotal := sngAnchoTotal + rangoE.columns.columns[n].ColumnWidth;

  sngAnchoCelda :=rangoE.columns.columns[1].ColumnWidth;
//  rangoE.HorizontalAlignment := xlJustify;
//  rangoE.VerticalAlignment := xlcenter;
  rangoE.MergeCells := False;

  if sngAnchoTotal>255 then
    rangoE.columns.columns[1].ColumnWidth :=255
  else
    rangoE.columns.columns[1].ColumnWidth := sngAnchoTotal;

  rangoE.parent.rows[rangoE.row].Autofit;
  sngAlto :=rangoE.RowHeight;

  rangoE.Merge;
  rangoE.Columns[1].EntireColumn.ColumnWidth := sngAnchoCelda;
  if sngAlto + 5 >409 then
    rangoE.Columns[1].RowHeight :=409
  else
  rangoE.Columns[1].RowHeight := sngAlto + 5;
  application.ProcessMessages;
end;

Procedure TfrmActividades.GenerarMarco(var Excel:Variant);
begin
  Excel.Selection.Borders[xlEdgeLeft].LineStyle         := xlContinuous;
  Excel.Selection.Borders[xlEdgeLeft].Weight            := xlThin;
  Excel.Selection.Borders[xlEdgeTop].LineStyle          := xlContinuous;
  Excel.Selection.Borders[xlEdgeTop].Weight             := xlThin;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle       := xlContinuous;
  Excel.Selection.Borders[xlEdgeBottom].Weight          := xlThin;
  Excel.Selection.Borders[xlEdgeRight].LineStyle        := xlContinuous;
  Excel.Selection.Borders[xlEdgeRight].Weight           := xlThin;
  Excel.Selection.Borders[xlInsideVertical].LineStyle   := xlContinuous ;
  Excel.Selection.Borders[xlInsideVertical].Weight      := xlThin;
  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideHorizontal].Weight    := xlThin;
  Excel.Selection.Borders[xlDiagonalDown].LineStyle := xlNone;
  Excel.Selection.Borders[xlDiagonalUp].LineStyle := xlNone;
end;

Procedure TfrmActividades.FormatoEncabezado(var Excel: Variant;Cadena:string; Align: Integer;Negrita:Boolean);
begin
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment :=Align ;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.WrapText:=True;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value :=Cadena;
  Excel.Selection.Interior.Pattern := xlSolid;
  Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
  Excel.Selection.Interior.ThemeColor := 2;//xlThemeColorLight2;
  Excel.Selection.Interior.TintAndShade := 0.799981688894314;
  Excel.Selection.Interior.PatternTintAndShade := 0;
end;

Procedure TfrmActividades.FormatoNormal(var Excel: Variant;Cadena:string; Align: Integer;Negrita,Ajustar:Boolean);
begin
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment :=Align ;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := Negrita;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value :=Cadena;
  if Ajustar then


end;

Procedure TfrmActividades.FormatoNormal(var Excel: Variant;Cadena:Variant; Align: Integer;Negrita,Ajustar:Boolean;Formato:String;Column:Integer=0);
var
  RangoE:Variant;
begin
  RangoE:=Excel.Selection;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment :=Align ;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.NumberFormat:=Formato;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := Negrita;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value :=Cadena;
  if Ajustar then
    AjustarTexto(RangoE,column);

end;


procedure TfrmActividades.ConfigurarHoja(var excel: Variant; var Hoja: Variant);
var
  pfHoja: Byte;
  SubCad,CadError: String;
  Difer, AcumDifer: Extended;
  sFirmante1,sFirmante2,sPuesto1,sPuesto2:string;
  QryBuscarFirmas: tzReadOnlyQuery;
  fs:TStream;
  //tmpNombreC:string;
  imgAux:TImage;
  Pic : TJpegImage;
  TempPath: array [0..MAX_PATH-1] of Char;
  FNombre1,FNombre2:TFileName;
  sCadT:string;
begin
//sfnSumaHoras()
  // Seleccionar el periodo de firmantes
  application.ProcessMessages;
  imgAux:=TImage.Create(nil);
  QryBuscarFirmas:=TZReadOnlyQuery.Create(nil);
  {QryBuscarFirmas.Connection := connection.zconnection;
  QryBuscarFirmas.SQL.Add('Select * from firmas where sContrato = :contrato and sIdTurno =:Turno and sNumeroOrden = :Orden And dIdFecha = :fecha');
  QryBuscarFirmas.Params.ParamByName('Orden').DataType := ftString;
  QryBuscarFirmas.Params.ParamByName('Orden').Value := ReporteDiario.FieldByName('sNumeroOrden').AsString;
  QryBuscarFirmas.Params.ParamByName('Contrato').DataType := ftString;
  QryBuscarFirmas.Params.ParamByName('Contrato').Value :=ReporteDiario.FieldByName('sContrato').AsString;
  QryBuscarFirmas.Params.ParamByName('Turno').DataType := ftString;
  QryBuscarFirmas.Params.ParamByName('Turno').Value := ReporteDiario.FieldByName('sIdTurno').AsString;
  QryBuscarFirmas.Params.ParamByName('fecha').DataType := ftDate;
  QryBuscarFirmas.Params.ParamByName('fecha').Value := ReporteDiario.FieldByName('didfecha').AsDateTime;
  QryBuscarFirmas.Open;
  Application.ProcessMessages;
  if QryBuscarFirmas.RecordCount=0 then
  begin
    QryBuscarFirmas.Active := False;
    QryBuscarFirmas.SQL.Clear;
    QryBuscarFirmas.SQL.Add('Select ImgFirma1,ImgFirma5 from firmas where sContrato = :contrato and sNumeroOrden = :Orden and sIdTurno =:Turno And dIdFecha <= :fecha Order By dIdFecha DESC');
    QryBuscarFirmas.Params.ParamByName('Orden').DataType := ftString;
    QryBuscarFirmas.Params.ParamByName('Orden').Value := ReporteDiario.FieldByName('sNumeroOrden').AsString;
    QryBuscarFirmas.Params.ParamByName('Contrato').DataType := ftString;
    QryBuscarFirmas.Params.ParamByName('Contrato').Value :=ReporteDiario.FieldByName('sContrato').AsString;
    QryBuscarFirmas.Params.ParamByName('Turno').DataType := ftString;
    QryBuscarFirmas.Params.ParamByName('Turno').Value := ReporteDiario.FieldByName('sIdTurno').AsString;
    QryBuscarFirmas.Params.ParamByName('fecha').DataType := ftDate;
    QryBuscarFirmas.Params.ParamByName('fecha').Value := ReporteDiario.FieldByName('didfecha').AsDateTime;
    QryBuscarFirmas.Open;

    if QryBuscarFirmas.RecordCount > 0 then
    begin
      GetTempPath(SizeOf(TempPath), TempPath);
      FNombre1:=TempPath +'imgtempAby'+formatdatetime('dddddd hhnnss',now)+'.jpg';

      fs := QryBuscarFirmas.CreateBlobStream(QryBuscarFirmas.FieldByName('ImgFirma1'), bmRead) ;
      // fs := QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagen'), bmRead);
      If fs.Size > 1 Then
      Begin
          try
              Pic:=TJpegImage.Create;
              try
                 Pic.LoadFromStream(fs);
                 imgAux.Picture.Graphic := Pic;
              finally
                 Pic.Free;
              end;
          finally
              fs.Free;
          End;
        imgAux.Picture.SaveToFile(FNombre1);
      End;
      application.ProcessMessages;
      GetTempPath(SizeOf(TempPath), TempPath);
      FNombre2:=TempPath +'imgtempAby2'+formatdatetime('dddddd hhnnss',now)+'.jpg';

      fs := QryBuscarFirmas.CreateBlobStream(QryBuscarFirmas.FieldByName('ImgFirma5'), bmRead) ;
      // fs := QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagen'), bmRead);
      If fs.Size > 1 Then
      Begin
          try
              Pic:=TJpegImage.Create;
              try
                 Pic.LoadFromStream(fs);
                 imgAux.Picture.Graphic := Pic;
              finally
                 Pic.Free;
              end;
          finally
              fs.Free;
          End;
        imgAux.Picture.SaveToFile(FNombre2);
      End ;

      application.ProcessMessages;

    end;

  end;  }




   Application.ProcessMessages;

    // Poner las firmas en todas las hojas del libro generado
    try

      //for pfHoja := 1 to Excel.Sheets.Count do
      begin

        //AcumDifer := AcumDifer + Difer;
        //ProgressBar1.Position := Trunc(AcumDifer);
        //Application.ProcessMessages;

        //Excel.Sheets[pfHoja].Select;
        Excel.ActiveSheet.PageSetup.PaperSize := xlPaperLetter;
        Excel.ActiveWindow.View :=xlPageLayoutView;

        //Excel.PrintCommunication := true;
        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := False;


          Excel.ActiveSheet.PageSetup.PrintTitleRows := '$1:$9';
          sCadT:='$1:$9';
        

        Application.ProcessMessages;
        Excel.ActiveSheet.PageSetUp.CenterFooter :='';
        Excel.ActiveSheet.PageSetUp.LeftFooter := '';
        Excel.ActiveSheet.PageSetUp.RightFooter := '';
        Excel.ActiveSheet.PageSetUp.LeftMargin     := 0;
        Excel.ActiveSheet.PageSetUp.RightMargin    := 0;
        Excel.ActiveSheet.PageSetUp.TopMargin      := 14;
        Excel.ActiveSheet.PageSetUp.BottomMargin   := 120;
        Excel.ActiveSheet.PageSetUp.HeaderMargin   := 0;
        Excel.ActiveSheet.PageSetUp.FooterMargin   :=Excel.InchesToPoints(0.393700787401575); //56;
        Excel.ActiveSheet.PageSetUp.Zoom := 32;
        application.ProcessMessages;



         Excel.ActiveSheet.PageSetUp.ScaleWithDocHeaderFooter := True;
        Excel.ActiveSheet.PageSetUp.AlignMarginsHeaderFooter := True;
        Excel.ActiveSheet.PageSetUp.EvenPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.EvenPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.EvenPage.RightFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.FirstPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.FirstPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetUp.FirstPage.RightFooter.Text := '';

        Excel.ActiveSheet.PageSetup.Orientation := xlPortrait;
        Excel.ActiveSheet.PageSetUp.Zoom           := False;
        Excel.ActiveSheet.PageSetUp.FitToPagesWide := 1;
        Excel.ActiveSheet.PageSetUp.FitToPagesTall := False;

        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := True;

        ///////////////////////////////////////
        application.ProcessMessages;
        Excel.ActiveWindow.View := xlPageLayoutView;
        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := False;
        Excel.ActiveSheet.PageSetup.LeftHeader := '';
        Excel.ActiveSheet.PageSetup.CenterHeader := '';
        Excel.ActiveSheet.PageSetup.RightHeader := '';
        Excel.ActiveSheet.PageSetup.LeftFooter := '';
        Excel.ActiveSheet.PageSetup.CenterFooter := '';
        Excel.ActiveSheet.PageSetup.RightFooter := '';
        Excel.ActiveSheet.PageSetup.LeftMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.TopMargin := Excel.InchesToPoints(0.194444444444444);
        Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(1.66666666666667);
        Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0.393700787401575);
        Excel.ActiveSheet.PageSetup.Zoom := 32;
        Excel.ActiveSheet.PageSetup.PrintErrors := xlPrintErrorsDisplayed;
        Excel.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
        Excel.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := true;
        Excel.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightFooter.Text := '';
        if (Excel.Application.version >= 14) then
        begin
          Excel.PrintCommunication := True;
          Excel.PrintCommunication := False;
        end;
        application.ProcessMessages;
        Excel.ActiveSheet.PageSetup.LeftHeader := '';
        Excel.ActiveSheet.PageSetup.CenterHeader := '';
        Excel.ActiveSheet.PageSetup.RightHeader := '';
        Excel.ActiveSheet.PageSetup.LeftFooter := '';
        Excel.ActiveSheet.PageSetup.CenterFooter := '';
        Excel.ActiveSheet.PageSetup.RightFooter := '';
        Excel.ActiveSheet.PageSetup.LeftMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.TopMargin := Excel.InchesToPoints(0.194444444444444);
        Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(1.66666666666667);
        Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0.393700787401575);
        Excel.ActiveSheet.PageSetup.Zoom := 32;
        Excel.ActiveSheet.PageSetup.PrintErrors := xlPrintErrorsDisplayed;
        Excel.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
        Excel.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := true;
        Excel.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightFooter.Text := '';
        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := True;

        application.ProcessMessages;

        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := False;

        Excel.ActiveSheet.PageSetup.PrintTitleRows :=sCadT;
        Excel.ActiveSheet.PageSetup.PrintTitleColumns := '';

        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := True;
        Excel.ActiveSheet.PageSetup.PrintArea := '';
        if (Excel.Application.version >= 14) then
          Excel.PrintCommunication := False;

        application.ProcessMessages;
        Excel.ActiveSheet.PageSetup.LeftHeader := '';
        Excel.ActiveSheet.PageSetup.CenterHeader := '';
        Excel.ActiveSheet.PageSetup.RightHeader := '';
       // Excel.ActiveSheet.PageSetup.LeftFooter :='&G';
        Excel.ActiveSheet.PageSetup.CenterFooter :='&"Arial,Normal"&'+inttostr(TamFont)+'&P de &#';//'&Z&G&P de &#&D&G'; //'&P de &N';
        //"&""Arial,Normal""&14&P de &#"
       // Excel.ActiveSheet.PageSetup.RightFooter :='&G';
        Excel.ActiveSheet.PageSetup.LeftMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.TopMargin := Excel.InchesToPoints(0.196850393700787);
        Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(1.65354330708661);
        Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0.393700787401575);
        Excel.ActiveSheet.PageSetup.PrintHeadings := False;
        Excel.ActiveSheet.PageSetup.PrintGridlines := False;
       // Excel.ActiveSheet.PageSetup.PrintComments := xlPrintNoComments;  //
        Excel.ActiveSheet.PageSetup.PrintQuality := 600;
        Excel.ActiveSheet.PageSetup.CenterHorizontally := False;
        Excel.ActiveSheet.PageSetup.CenterVertically := False;
      { if   pfHoja>(Excel.Sheets.Count-iTotalLandscape) then
        Excel.ActiveSheet.PageSetup.Orientation := xlLandscape
        else
         Excel.ActiveSheet.PageSetup.Orientation := xlPortrait; }
        Excel.ActiveSheet.PageSetup.Draft := False;
        try
          Excel.ActiveSheet.PageSetup.PaperSize := xlPaperLetter;
        except
          Excel.ActiveSheet.PageSetup.PaperSize :=119;
        end;
        Excel.ActiveSheet.PageSetup.FirstPageNumber := xlAutomatic;
        Excel.ActiveSheet.PageSetup.Order := xlDownThenOver;
        Excel.ActiveSheet.PageSetup.BlackAndWhite := False;
        Excel.ActiveSheet.PageSetup.Zoom := 32;
        Excel.ActiveSheet.PageSetup.PrintErrors := xlPrintErrorsDisplayed;
        Excel.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter := False;
        Excel.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
        Excel.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := True;
        Excel.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.EvenPage.RightFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightHeader.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text := '';
        Excel.ActiveSheet.PageSetup.FirstPage.RightFooter.Text := '';
         Excel.ActiveSheet.PageSetUp.Zoom           := False;
        Excel.ActiveSheet.PageSetUp.FitToPagesWide := 1;
        Excel.ActiveSheet.PageSetUp.FitToPagesTall := False;
//        if (Excel.Application.version >= 14) then
//          Excel.PrintCommunication := True;
        Application.ProcessMessages;
        Excel.ActiveWindow.View := xlPageBreakPreview;
        Excel.ActiveWindow.Zoom := 115;

      end;
       
    except
      on e:exception do
      begin
        ShowMessage(e.Message + ', ' + e.ClassName);
        CadError := 'No se puede generar el pie de página';

      end;
    end;

end;

procedure TfrmActividades.GenerarNotaPdasSinBarco(var Excel,Hoja:Variant;var QrDatos:TZReadOnlyQuery;var QrPdas:TZReadOnlyQuery);
var
  Ren,IniciaR,I:Integer;
  QrConsulta,
  QryMovimientos :TZReadOnlyQuery;

  sEfectivo,sAfectaciones:string;
  dCostoDll,dCostoDllT,dCostoMn,dCostoMnT:Extended;
  OriFecha:TDate;
  OriHora,OriClasif,sHoraIntervalo:string;
  OridAvanceAnt,OridAvanceAct:Extended;
  dFechaBucle: TDateTime;
  dProrrateoPernocta, dAjustePernocta: Double;
  iDiferenciaFechas, x: Integer;
  bPasoAjuste: Boolean;

  //Variables nota de campo..
  dCantidadPartida, dCantidadTotal : Double;
  dFactorMovimiento : Double;
  dCantidadPartidaAux, dCantidadTotalAux,
  dCantPernoctaAdicional : double;

  Progreso, TotalProgreso: real;

  //Horas extras - Personal
  qrHorasExras : TZReadOnlyQuery;
  sumaHE : string;
  TotHoras,
  TotMinutos,
  dCantHETot : Double;


  procedure getHM(cadena : string ; var h, m : Double);
  var
    x, xant : integer;
    sHoras, sMinutos : string;
  begin
    sHoras := '';
    sMinutos := '';
    xant := 1;
    for x := xant to Length(cadena) do
    begin
      if cadena[x] <> ':' then
        sHoras := sHoras + cadena[x]
      else
        Break;
    end;
    xant := x + 1;
    for x := xant to Length(cadena) do
    begin
      if cadena[x] <> ':' then
      begin
        sMinutos := sMinutos + cadena[x];
      end;
    end;

    h := StrToInt(sHoras);
    m := StrToInt(sMinutos);
  end;

begin
  qrHorasExras := TZReadOnlyQuery.Create(nil);
  qrHorasExras.Connection := connection.zConnection;

  Ren:=12;
  QrConsulta:=TZReadOnlyQuery.Create(nil);
  QrConsulta.Connection:=connection.zConnection;

  QryMovimientos:=TZReadOnlyQuery.Create(nil);
  QryMovimientos.Connection:=connection.zConnection;

  {$REGION 'BUSCAR BARCO'}
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.sql.text:='select ev.*,e.sDescripcion from embarcacion_vigencia ev ' +
                                'inner join embarcaciones e on (e.scontrato=ev.scontrato and ev.sidembarcacion=e.sIdembarcacion) ' +
                                'where ev.scontrato=:Contrato and ev.dFechaInicio='+
                                '(select max(dFechaInicio) from embarcacion_vigencia ' +
                                'where scontrato=:Contrato and dFechaInicio<=:Fecha)';
  connection.QryBusca.ParamByName('contrato').AsString := global_Contrato_Barco;
  connection.QryBusca.ParamByName('Fecha').AsDate      := ActividadesxOrden.FieldValues['dFechaFinal'];
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount = 1 then
     global_barco := connection.QryBusca.FieldByName('sIdEmbarcacion').AsString;
  {$ENDREGION}
  PanelProgress.Visible := True;
  grid_actividades.Enabled := False;
  bPasoAjuste := False;
  dAjustePernocta := 0;
  while not QrPdas.Eof do
  begin
      Progreso := (1 / (QrPdas.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
      TotalProgreso := TotalProgreso + Progreso;
      Label15.Caption := 'Procesando Nota de Campo...';
      BarraEstado.Position := Trunc(TotalProgreso);

      dCostoMnT:=0;
      dCostoDllT:=0;

    if QrPdas.FieldByName('sTipoActividad').AsString='Actividad' then
    begin
      {$REGION 'ACTIVIDADES' }
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 12.9;
      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PARTIDA',xlCenter,True);

      Application.ProcessMessages;

      Hoja.Range['D'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'ACTIVIDAD',xlCenter,True);
      Application.ProcessMessages;

      inc(Ren);

      //Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 12.9;
      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoNormal(Excel,QrPdas.FieldByName('sNumeroActividad').AsString,xlCenter,True,false,'@');
      //FormatoNormal(var Excel: Variant;Cadena:Variant; Align: Integer;Negrita,Ajustar:Boolean;Formato:String;Column:Integer=0);

      Application.ProcessMessages;

      Hoja.Range['D'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,QrPdas.FieldByName('mDescripcion').AsString,xlJustify,True,true,'@',9);
      Application.ProcessMessages;

      Hoja.Range['B'+IntToStr(ren-1)+':L'+IntToStr(ren)].Select;
      GenerarMarco(Excel);

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+4)].RowHeight := 10.2;

      Inc(ren,3);
      Hoja.Range['B'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'PERIODOS DE EJECUCION DE LA ACTIVIDAD',xlCenter,False,false);
      Application.ProcessMessages;

      Inc(ren,2);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 16.8;
      Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'FECHA',xlCenter,false);

      Hoja.Range['C'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'INICIO',xlCenter,false);

      Hoja.Range['D'+IntToStr(ren)+':D'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'TERMINO',xlCenter,false);

      Hoja.Range['E'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'AFECTACIÓN',xlCenter,false);

      Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'INTERVALO TIEMPO',xlCenter,false);

      Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'AVANCE ANTERIOR',xlCenter,false);

      Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'AVANCE ACTUAL',xlCenter,false);

      Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'AVANCE ACUMULADO',xlCenter,false);
      IniciaR:=Ren;

      inc(ren);
      QrConsulta.Active:=False;
      QrConsulta.SQL.Text:= 'select b.*,cast(b.sHoraInicio as Time) as HoraI,' + #10 +
                          '(select ifnull(sum(ba.dAvance),0) from bitacoradeactividades as ba where ba.sContrato=b.sContrato and ba.sNumeroOrden=b.sNumeroOrden and ba.sIdTipoMovimiento=b.sIdTipoMovimiento' + #10 +
                          'and ba.swbs=b.swbs and ba.sNumeroActividad=b.sNumeroActividad and (ba.didfecha<b.didfecha or (ba.didfecha=b.didfecha and cast(ba.sHoraInicio as Time)<cast(b.sHoraInicio as Time) ))) as AvAnterior' + #10 +
                          'from bitacoradeactividades b inner join tiposdemovimiento tm on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=b.sIdClasificacion) ' +
                          'where b.sContrato=:Contrato and b.sNumeroOrden=:Orden and b.sIdTipoMovimiento=''ED''' + #10 +
                          'and b.swbs=:wbs and b.sNumeroActividad=:Actividad and tm.lAplicaNotaCampo="Si"  order by b.didFecha,HoraI';
      QrConsulta.ParamByName('Contrato').AsString:=QrPdas.FieldByName('sContrato').AsString;
      QrConsulta.ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;
      QrConsulta.ParamByName('Orden').AsString:=QrPdas.FieldByName('sNumeroOrden').AsString;
      QrConsulta.ParamByName('wbs').AsString:=QrPdas.FieldByName('swbs').AsString;
      QrConsulta.ParamByName('Actividad').AsString:=QrPdas.FieldByName('sNumeroActividad').AsString;
      QrConsulta.Open;

      sEfectivo:='00:00';
      sAfectaciones:='00:00';

      OriFecha:=StrToDate('23/02/1984');
      OriHora:='SAHL';
      OriClasif:='ABY';
      OridAvanceAnt:=0;
      OridAvanceAct:=0;
      sHoraIntervalo:='00:00';

      while not QrConsulta.Eof do
      begin
        if (OriFecha<>QrConsulta.FieldByName('dIdFecha').AsDateTime) or (OriClasif<>QrConsulta.FieldByName('sIdClasificacion').AsString)
        or(OriHora<>QrConsulta.FieldByName('sHoraInicio').AsString) then
        begin

          Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('dIdFecha').Value,xlCenter,false,False,'dd/mm/aaaa');

          OriFecha:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
          //FormatoNormal(var Excel: Variant;Cadena:string; Align: Integer;Negrita,Ajustar:Boolean;Formato:String);

          Hoja.Range['C'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sHoraInicio').Value,xlCenter,false,False,'@');

          Hoja.Range['D'+IntToStr(ren)+':D'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sHoraFinal').Value,xlCenter,false,False,'@');
          OriHora:=QrConsulta.FieldByName('sHoraFinal').AsString;

          Hoja.Range['E'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sIdClasificacion').Value,xlCenter,false,False,'@');
          OriClasif:=QrConsulta.FieldByName('sIdClasificacion').AsString;


          sHoraIntervalo:=sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString);
          Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
          FormatoNormal(Excel,sHoraIntervalo,xlCenter,false,False,'@');

          if QrConsulta.FieldByName('sIdClasificacion').AsString='TE' then
            sEfectivo:=sfnSumaHoras(sEfectivo,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString))
          else
            sAfectaciones:=sfnSumaHoras(sAfectaciones,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString));

          Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('AvAnterior').Value,xlCenter,false,False,'0.00%');
          OridAvanceAnt:=QrConsulta.FieldByName('AvAnterior').AsFloat;

          Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('dAvance').Value,xlCenter,false,False,'0.00%');
          OridAvanceAct:=QrConsulta.FieldByName('dAvance').AsFloat;

          Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('AvAnterior').Value + QrConsulta.FieldByName('dAvance').Value,xlCenter,false,False,'0.00%');

          Inc(ren);
        end
        else
        begin
          Hoja.Range['D'+IntToStr(ren-1)+':D'+IntToStr(ren-1)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sHoraFinal').Value,xlCenter,false,False,'@');
          OriHora:=QrConsulta.FieldByName('sHoraFinal').AsString;


          sHoraIntervalo:=sfnSumaHoras(sHoraIntervalo,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString));
          Hoja.Range['F'+IntToStr(ren-1)+':F'+IntToStr(ren-1)].Select;
          FormatoNormal(Excel,sHoraIntervalo,xlCenter,false,False,'@');


          Hoja.Range['H'+IntToStr(ren-1)+':H'+IntToStr(ren-1)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('dAvance').Value + OridAvanceAct ,xlCenter,false,False,'0.00%');
          OridAvanceAct:=OridAvanceAct + QrConsulta.FieldByName('dAvance').AsFloat;

          Hoja.Range['I'+IntToStr(ren-1)+':I'+IntToStr(ren-1)].Select;
          FormatoNormal(Excel,OridAvanceAnt + OridAvanceAct,xlCenter,false,False,'0.00%');

          if QrConsulta.FieldByName('sIdClasificacion').AsString='TE' then
            sEfectivo:=sfnSumaHoras(sEfectivo,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString))
          else
            sAfectaciones:=sfnSumaHoras(sAfectaciones,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString));

        end;
        QrConsulta.Next;

      end;

      Hoja.Range['B'+IntToStr(IniciaR)+':I'+IntToStr(ren-1)].Select;
      GenerarMarco(Excel);

      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+2)].RowHeight := 15;

      Hoja.Range['B'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'DURACION TIEMPO EFECTIVO(HRS):',xlRight,false,False,'@');

      Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,sEfectivo,xlCenter,true,False,'@');


      inc(ren);

      Hoja.Range['B'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'DURACION TIEMPO AFECTACIONES(HRS):',xlRight,false,False,'@');

      Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,sAfectaciones,xlCenter,True,False,'@');

      inc(ren);

      Hoja.Range['B'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'TIEMPO TOTAL(HRS):',xlRight,false,False,'@');

      Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,sfnSumaHoras(sEfectivo,sAfectaciones),xlCenter,true,False,'@');


      inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+1)].RowHeight := 10.2;

      Inc(ren,2);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 12.9;
      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PARTIDA',xlCenter,True);

      Application.ProcessMessages;

      Hoja.Range['D'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'ACTIVIDAD',xlCenter,True);
      Application.ProcessMessages;

      inc(Ren);

      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoNormal(Excel,QrPdas.FieldByName('sNumeroActividad').AsString,xlCenter,True,false,'@');

      Application.ProcessMessages;

      Hoja.Range['D'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,QrPdas.FieldByName('mDescripcion').AsString,xlJustify,True,true,'@',9);
      Application.ProcessMessages;

      Hoja.Range['B'+IntToStr(ren-1)+':L'+IntToStr(ren)].Select;
      GenerarMarco(Excel);

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 10.2;

      Inc(ren);
      {$ENDREGION}

      {$REGION 'PERSONAL Y EQUIPO'}
      for I := 1 to 2 do
      begin
        dCostoDll:=0;
        dCostoMn:=0;


        IniciaR:=Ren;
        QrConsulta.Active:=False;
        if i=1 then
        begin

          QrConsulta.SQL.Text:= 'select r.iItemOrden, gp.sIdGrupo as '+ TablasConsult[I,2] +', SUM(r.dCantHHGenerador) as HH, p.dVentaMn as CostoMn,p.dVentaDll as CostoDll,gp.sDescripcion as DescRecurso,p.smedida from ' + TablasConsult[I,1] + ' r' + #10 +
                                'inner join bitacoradeactividades ba on(ba.sContrato=r.sContrato and ba.dIdFecha = r.dIdFecha and ba.iIdDiario=r.iIdDiario)' + #10 +
                                'inner join ' + TablasConsult[I,3] + ' p on(p.sContrato=:ContratoBarco and p.' + TablasConsult[I,2] + '=r.' + TablasConsult[I,2] + ')' + #10 +
                                'inner join grupospersonal gp on (gp.sIdGrupo=p.sAgrupaPersonal) ' +
                                'where r.sContrato=:Contrato AND r.sNumeroOrden = :Orden AND ba.sNumeroOrden=:Orden and ba.swbs=:wbs' + #10 +
                                'group by r.sContrato,gp.sIdGrupo order by p.iItemOrden';
        end
        else
        begin
          QrConsulta.SQL.Text:= 'select r.*, SUM(r.dCantHHGenerador) as HH, p.dVentaMN as CostoMn,p.dVentaDll as CostoDll,p.sDescripcion as DescRecurso,p.smedida from ' + TablasConsult[I,1] + ' r' + #10 +
                                'inner join bitacoradeactividades ba on(ba.sContrato=r.sContrato and ba.dIdFecha = r.dIdFecha and ba.iIdDiario=r.iIdDiario)' + #10 +
                                'inner join ' + TablasConsult[I,3] + ' p on(p.sContrato=:ContratoBarco and p.' + TablasConsult[I,2] + '=r.' + TablasConsult[I,2] + ')' + #10 +
                                'where r.sContrato=:Contrato AND r.sNumeroOrden = :Orden and ba.sNumeroOrden=:Orden and ba.swbs=:wbs' + #10 +
                                'group by r.sContrato,r.' + TablasConsult[I,2] +' order by p.iItemOrden';
        end;

        QrConsulta.ParamByName('Contrato').AsString:=QrPdas.FieldByName('sContrato').AsString;
        QrConsulta.ParamByName('Orden').AsString:=QrPdas.FieldByName('sNumeroOrden').AsString;
        QrConsulta.ParamByName('wbs').AsString:=QrPdas.FieldByName('swbs').AsString;
        QrConsulta.ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;
        QrConsulta.Open;

        Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+1)].RowHeight := 12.8;
        Hoja.Range['B'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,TablasConsult[I,4],xlCenter,false);

        Inc(ren);

        Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'PARTIDA',xlCenter,false);

        Hoja.Range['C'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'DESCRIPCIÓN',xlCenter,false);

        Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'UNIDAD',xlCenter,false);

        Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'CANTIDAD',xlCenter,false);

        Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'PU MN',xlCenter,false);

        Hoja.Range['J'+IntToStr(ren)+':J'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'PU USD',xlCenter,false);

        Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'IMP MN',xlCenter,false);

        Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'IMP USD',xlCenter,false);

        while not QrConsulta.Eof do
        begin          
          Inc(Ren);
          Hoja.Range['A'+IntToStr(ren)+':A'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('iItemOrden').Value,xlCenter,false,False,'@');

          Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName(TablasConsult[I,2]).Value,xlCenter,false,False,'@');

          Hoja.Range['C'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('DescRecurso').Value,xlJustify,false,true,'@',4);

          //DescRecurso
          Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sMedida').Value,xlCenter,false,False,'@');


          Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('HH').Value,xlCenter,false,False,'#0.00000000');

          Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('CostoMn').Value,xlCenter,false,False,'#,##0.00');


          Hoja.Range['J'+IntToStr(ren)+':J'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('CostoDll').Value,xlCenter,false,False,'#,##0.00');


          Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
          FormatoNormal(Excel, xRound(QrConsulta.FieldByName('CostoMn').Value * QrConsulta.FieldByName('HH').Value, 2),xlCenter,false,False,'#,##0.00');
          dCostoMn:=dCostoMn + xRound((QrConsulta.FieldByName('CostoMn').AsFloat * QrConsulta.FieldByName('HH').AsFloat), 2) ;

          Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
          FormatoNormal(Excel, xRound(QrConsulta.FieldByName('CostoDll').Value * QrConsulta.FieldByName('HH').Value, 2),xlCenter,false,False,'#,##0.00');
          dCostoDll:=dCostoDll + xRound((QrConsulta.FieldByName('CostoDll').AsFloat * QrConsulta.FieldByName('HH').AsFloat), 2);

          if i =  1 then
          begin
            qrHorasExras.Active := False;
            qrHorasExras.SQL.Clear;
            qrHorasExras.SQL.Add('select he.sIdPersonal, he.sDescripcion, p.sMedida, he.sCantHH, p.dVentaMN, p.dVentaDLL, p.dCostoMN, p.dCostoDLL '+
                                 'from horasextras he '+
                                 'inner join personal p '+
                                 'on (he.sContrato = p.sContrato and he.sIdPersonal = p.sIdPersonal) '+
                                 'where he.sContrato = :contrato '+
                                 'and he.sNumeroOrden = :folio '+
                                 'and he.sIdPersonal like :personal '+
                                 'and he.sNumeroActividad = :partida '+
                                 'and he.sWbs = :wbs');
            qrHorasExras.ParamByName('contrato').AsString := global_contrato;
            qrHorasExras.ParamByName('folio').AsString := tsNumeroOrden.Text;
            qrHorasExras.ParamByName('personal').AsString := QrConsulta.FieldByName('sIdPersonal').AsString + '%';
            qrHorasExras.ParamByName('partida').AsString := QrPdas.FieldByName('sNumeroActividad').AsString;
            qrHorasExras.ParamByName('wbs').AsString := QrPdas.FieldByName('sWbs').AsString;
            qrHorasExras.Open;

            {Horas extras}
            //Solo aplica personal...
            {$REGION 'HORAS EXTRAS'}

            if qrHorasExras.RecordCount > 0 then
            begin
              sumaHE := '00:00';
              qrHorasExras.First;
              while not qrHorasExras.Eof do
              begin
                sumaHE := sfnSumaHoras(qrHorasExras.FieldByName('sCantHH').AsString, sumaHE);
                qrHorasExras.Next;
              end;

              //Convierte horas extras a cantidad
              getHM(sumaHE, TotHoras, TotMinutos);
              TotHoras := (TotHoras * 60) / 100;
              TotMinutos := TotMinutos / 100;
              dCantHETot := TotHoras + TotMinutos;
              dCantHETot := (dCantHETot * 100) / 60;

              Inc(Ren);

              Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
              FormatoNormal(Excel,qrHorasExras.FieldByName('sIdPersonal').AsString,xlCenter,false,False,'@');

              Hoja.Range['C'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
              FormatoNormal(Excel,qrHorasExras.FieldByName('sDescripcion').AsString,xlJustify,false,true,'@',4);

              Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
              FormatoNormal(Excel,qrHorasExras.FieldByName('sMedida').AsString,xlCenter,false,False,'@');

              Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
              FormatoNormal(Excel,dCantHETot,xlCenter,false,False,'#0.00000000');

              Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
              FormatoNormal(Excel,qrHorasExras.FieldByName('dVentaMN').AsFloat,xlCenter,false,False,'#,##0.00');

              Hoja.Range['J'+IntToStr(ren)+':J'+IntToStr(ren)].Select;
              FormatoNormal(Excel,qrHorasExras.FieldByName('dVentaDLL').AsFloat,xlCenter,false,False,'#,##0.00');

              Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
              FormatoNormal(Excel, xRound(qrHorasExras.FieldByName('dVentaMN').AsFloat * dCantHETot, 2),xlCenter,false,False,'#,##0.00');
              dCostoMn:=dCostoMn + xRound((qrHorasExras.FieldByName('dVentaMN').AsFloat * dCantHETot), 2) ;

              Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
              FormatoNormal(Excel, xRound(qrHorasExras.FieldByName('dVentaDLL').AsFloat * dCantHETot, 2),xlCenter,false,False,'#,##0.00');
              dCostoDll:=dCostoDll + xRound((qrHorasExras.FieldByName('dVentaDLL').AsFloat * dCantHETot), 2);
            end;
            
            {$ENDREGION}

          end;

          QrConsulta.Next;
        end;

        Hoja.Range['B'+IntToStr(IniciaR)+':L'+IntToStr(ren)].Select;
        GenerarMarco(Excel);

        Inc(ren);
        Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 14.3;

        Hoja.Range['B'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
        FormatoNormal(Excel,'IMPORTE '+TablasConsult[I,4]+':',xlCenter,false,False,'@');

        Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
        FormatoNormal(Excel,xRound(dCostoMn, 2),xlCenter,true,False,'$#,##0.00');

        Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
        FormatoNormal(Excel,xRound(dCostoDll, 2),xlCenter,true,False,'$#,##0.00');

        dCostoDllT:=dCostoDllT + dCostoDll;
        dCostoMnT:=dCostoMnT + dCostoMn;

        Inc(ren);
        Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 10.2;

        Inc(ren);
      end;
      {$ENDREGION}

      {$REGION 'MATERIALES'}
      
      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 10.2;

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+1)].RowHeight := 11.3;
      Hoja.Range['B'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'MATERIAL',xlCenter,false);

      Inc(ren);

      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'TRAZABILIDAD',xlCenter,false);

      Hoja.Range['D'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'DESCRIPCIÓN',xlCenter,false);

      Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'UNIDAD',xlCenter,false);

      Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'CANTIDAD',xlCenter,false);

      IniciaR:=Ren;
      QrConsulta.Active:=False;
      QrConsulta.SQL.Text:= 'select ba.sNumeroOrden,ba.sNumeroActividad,if(length(trim(bm.sTrazabilidadaux))=0 , '+
      '"S/T",bm.sTrazabilidadaux) as sTrazabilidad,i.mDescripcion,i.sMedida as smedida ,sum(bm.dCantidad) as dcantidad  from bitacoradeactividades ba '+
      'inner join bitacorademateriales bm on '+
      '(ba.sContrato = bm.sContrato and ba.dIdFecha = bm.dIdFecha and ba.iiddiario = bm.iIdDiario and ba.sWbs = bm.sWbs) '+
      'inner join insumos i on '+
      '(i.scontrato = :Contrato and i.sIdInsumo = bm.sIdMaterial and i.sTrazabilidad = bm.strazabilidad) '+
      'where ba.sContrato = :Orden and ba.sNumeroOrden = :Folio and ba.sNumeroActividad = :Actividad  '+
      'group by bm.sidmaterial,bm.sTrazabilidadaux order by bm.sTrazabilidadAux';
      QrConsulta.ParamByName('Contrato').AsString  := global_Contrato_Barco;
      QrConsulta.ParamByName('Orden').AsString     := QrPdas.FieldByName('sContrato').AsString;
      QrConsulta.ParamByName('Folio').AsString     := QrPdas.FieldByName('sNumeroOrden').AsString;
      QrConsulta.ParamByName('Actividad').AsString := QrPdas.FieldByName('sNumeroActividad').AsString;
      QrConsulta.Open;

      while not QrConsulta.Eof do
      begin
        Inc(Ren);

        Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
        FormatoNormal(Excel,QrConsulta.FieldByName('sTrazabilidad').Value,xlCenter,false,False,'@');
        //FormatoNormal(Excel,QrConsulta.FieldByName('AvAnterior').Value + QrConsulta.FieldByName('dAvance').Value,xlCenter,false,False,'@');

        Hoja.Range['D'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
        FormatoNormal(Excel,QrConsulta.FieldByName('mDescripcion').Value,xlJustify,false,true,'@',4);

        //DescRecurso
        Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
        FormatoNormal(Excel,QrConsulta.FieldByName('sMedida').Value,xlCenter,false,False,'@');


        Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
        FormatoNormal(Excel,QrConsulta.FieldByName('dCantidad').Value,xlCenter,false,False,'#,##0.00');

        QrConsulta.Next;

      end;

      Hoja.Range['B'+IntToStr(IniciaR-1)+':I'+IntToStr(ren)].Select;
      GenerarMarco(Excel);

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight :=10.2;

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight :=11.3;

      Hoja.Range['B'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'COSTO TOTAL DE LA ACTIVIDAD::',xlCenter,false,False,'@');

      Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
      FormatoNormal(Excel,xRound(dCostoMnT, 2),xlCenter,true,False,'$#,##0.00');

      Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,xRound(dCostoDllT, 2),xlCenter,true,False,'$#,##0.00');

      Inc(ren);
      {$ENDREGION}

    end;
    QrPdas.Next;
  end;
  PanelProgress.Visible := False;
  grid_actividades.Enabled := True;
end;


procedure TfrmActividades.GraficadeGantt1Click(Sender: TObject);
begin
  ExportOTProject(Global_Contrato,zqReprogramacion.FieldByName('sIdConvenio').AsString,TsNumeroOrden.KeyValue);
end;

procedure TfrmActividades.GenerarNotaPdas(var Excel,Hoja:Variant;var QrDatos:TZReadOnlyQuery;var QrPdas:TZReadOnlyQuery);
var
  Ren,IniciaR,I:Integer;
  QrConsulta,
  QryMovimientos :TZReadOnlyQuery;

  sEfectivo,sAfectaciones:string;
  dCostoDll,dCostoDllT,dCostoMn,dCostoMnT:Extended;
  OriFecha:TDate;
  OriHora,OriClasif,sHoraIntervalo:string;
  OridAvanceAnt,OridAvanceAct:Extended;
  dFechaBucle: TDateTime;
  dProrrateoPernocta, dAjustePernocta: Double;
  iDiferenciaFechas, x: Integer;
  bPasoAjuste: Boolean;

  //Variables nota de campo..
  dCantidadPartida, dCantidadTotal : Double;
  dFactorMovimiento : Double;
  dCantidadPartidaAux, dCantidadTotalAux,
  dCantPernoctaAdicional : double;

  Progreso, TotalProgreso: real;
begin
  Ren:=12;
  QrConsulta:=TZReadOnlyQuery.Create(nil);
  QrConsulta.Connection:=connection.zConnection;

  QryMovimientos:=TZReadOnlyQuery.Create(nil);
  QryMovimientos.Connection:=connection.zConnection;

  {$REGION 'BUSCAR BARCO'}
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.sql.text:='select ev.*,e.sDescripcion from embarcacion_vigencia ev ' +
                                'inner join embarcaciones e on (e.scontrato=ev.scontrato and ev.sidembarcacion=e.sIdembarcacion) ' +
                                'where ev.scontrato=:Contrato and ev.dFechaInicio='+
                                '(select max(dFechaInicio) from embarcacion_vigencia ' +
                                'where scontrato=:Contrato and dFechaInicio<=:Fecha)';
  connection.QryBusca.ParamByName('contrato').AsString := global_Contrato_Barco;
  connection.QryBusca.ParamByName('Fecha').AsDate      := ActividadesxOrden.FieldValues['dFechaFinal'];
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount = 1 then
     global_barco := connection.QryBusca.FieldByName('sIdEmbarcacion').AsString;
  {$ENDREGION}
  PanelProgress.Visible := True;
  grid_actividades.Enabled := False;
  bPasoAjuste := False;
  dAjustePernocta := 0;
  while not QrPdas.Eof do
  begin
      Progreso := (1 / (QrPdas.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
      TotalProgreso := TotalProgreso + Progreso;
      Label15.Caption := 'Procesando Nota de Campo...';
      BarraEstado.Position := Trunc(TotalProgreso);

      dCostoMnT:=0;
      dCostoDllT:=0;

    if QrPdas.FieldByName('sTipoActividad').AsString='Actividad' then
    begin
      {$REGION 'ACTIVIDADES' }
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 12.9;
      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PARTIDA',xlCenter,True);

      Application.ProcessMessages;

      Hoja.Range['D'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'ACTIVIDAD',xlCenter,True);
      Application.ProcessMessages;

      inc(Ren);

      //Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 12.9;
      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoNormal(Excel,QrPdas.FieldByName('sNumeroActividad').AsString,xlCenter,True,false,'@');
      //FormatoNormal(var Excel: Variant;Cadena:Variant; Align: Integer;Negrita,Ajustar:Boolean;Formato:String;Column:Integer=0);

      Application.ProcessMessages;

      Hoja.Range['D'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,QrPdas.FieldByName('mDescripcion').AsString,xlJustify,True,true,'@',9);
      Application.ProcessMessages;

      Hoja.Range['B'+IntToStr(ren-1)+':L'+IntToStr(ren)].Select;
      GenerarMarco(Excel);

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+4)].RowHeight := 10.2;

      Inc(ren,3);
      Hoja.Range['B'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'PERIODOS DE EJECUCION DE LA ACTIVIDAD',xlCenter,False,false);
      Application.ProcessMessages;

      Inc(ren,2);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 16.8;
      Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'FECHA',xlCenter,false);

      Hoja.Range['C'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'INICIO',xlCenter,false);

      Hoja.Range['D'+IntToStr(ren)+':D'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'TERMINO',xlCenter,false);

      Hoja.Range['E'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'AFECTACIÓN',xlCenter,false);

      Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'INTERVALO TIEMPO',xlCenter,false);

      Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'AVANCE ANTERIOR',xlCenter,false);

      Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'AVANCE ACTUAL',xlCenter,false);

      Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'AVANCE ACUMULADO',xlCenter,false);
      IniciaR:=Ren;

      inc(ren);
      QrConsulta.Active:=False;
      QrConsulta.SQL.Text:= 'select b.*,cast(b.sHoraInicio as Time) as HoraI,' + #10 +
                          '(select ifnull(sum(ba.dAvance),0) from bitacoradeactividades as ba where ba.sContrato=b.sContrato and ba.sNumeroOrden=b.sNumeroOrden and ba.sIdTipoMovimiento=b.sIdTipoMovimiento' + #10 +
                          'and ba.swbs=b.swbs and ba.sNumeroActividad=b.sNumeroActividad and (ba.didfecha<b.didfecha or (ba.didfecha=b.didfecha and cast(ba.sHoraInicio as Time)<cast(b.sHoraInicio as Time) ))) as AvAnterior' + #10 +
                          'from bitacoradeactividades b inner join tiposdemovimiento tm on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=b.sIdClasificacion) ' +
                          'where b.sContrato=:Contrato and b.sNumeroOrden=:Orden and b.sIdTipoMovimiento=''ED''' + #10 +
                          'and b.swbs=:wbs and b.sNumeroActividad=:Actividad and tm.lAplicaNotaCampo="Si"  order by b.didFecha,HoraI';
      QrConsulta.ParamByName('Contrato').AsString:=QrPdas.FieldByName('sContrato').AsString;
      QrConsulta.ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;
      QrConsulta.ParamByName('Orden').AsString:=QrPdas.FieldByName('sNumeroOrden').AsString;
      QrConsulta.ParamByName('wbs').AsString:=QrPdas.FieldByName('swbs').AsString;
      QrConsulta.ParamByName('Actividad').AsString:=QrPdas.FieldByName('sNumeroActividad').AsString;
      QrConsulta.Open;

      sEfectivo:='00:00';
      sAfectaciones:='00:00';

      OriFecha:=StrToDate('23/02/1984');
      OriHora:='SAHL';
      OriClasif:='ABY';
      OridAvanceAnt:=0;
      OridAvanceAct:=0;
      sHoraIntervalo:='00:00';

      while not QrConsulta.Eof do
      begin
        if (OriFecha<>QrConsulta.FieldByName('dIdFecha').AsDateTime) or (OriClasif<>QrConsulta.FieldByName('sIdClasificacion').AsString)
        or(OriHora<>QrConsulta.FieldByName('sHoraInicio').AsString) then
        begin

          Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('dIdFecha').Value,xlCenter,false,False,'dd/mm/aaaa');

          OriFecha:=QrConsulta.FieldByName('dIdFecha').AsDateTime;
          //FormatoNormal(var Excel: Variant;Cadena:string; Align: Integer;Negrita,Ajustar:Boolean;Formato:String);

          Hoja.Range['C'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sHoraInicio').Value,xlCenter,false,False,'@');

          Hoja.Range['D'+IntToStr(ren)+':D'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sHoraFinal').Value,xlCenter,false,False,'@');
          OriHora:=QrConsulta.FieldByName('sHoraFinal').AsString;

          Hoja.Range['E'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sIdClasificacion').Value,xlCenter,false,False,'@');
          OriClasif:=QrConsulta.FieldByName('sIdClasificacion').AsString;


          sHoraIntervalo:=sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString);
          Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
          FormatoNormal(Excel,sHoraIntervalo,xlCenter,false,False,'@');

          if QrConsulta.FieldByName('sIdClasificacion').AsString='TE' then
            sEfectivo:=sfnSumaHoras(sEfectivo,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString))
          else
            sAfectaciones:=sfnSumaHoras(sAfectaciones,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString));

          Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('AvAnterior').Value,xlCenter,false,False,'0.00%');
          OridAvanceAnt:=QrConsulta.FieldByName('AvAnterior').AsFloat;

          Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('dAvance').Value,xlCenter,false,False,'0.00%');
          OridAvanceAct:=QrConsulta.FieldByName('dAvance').AsFloat;

          Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('AvAnterior').Value + QrConsulta.FieldByName('dAvance').Value,xlCenter,false,False,'0.00%');

          Inc(ren);
        end
        else
        begin
          Hoja.Range['D'+IntToStr(ren-1)+':D'+IntToStr(ren-1)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sHoraFinal').Value,xlCenter,false,False,'@');
          OriHora:=QrConsulta.FieldByName('sHoraFinal').AsString;


          sHoraIntervalo:=sfnSumaHoras(sHoraIntervalo,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString));
          Hoja.Range['F'+IntToStr(ren-1)+':F'+IntToStr(ren-1)].Select;
          FormatoNormal(Excel,sHoraIntervalo,xlCenter,false,False,'@');


          Hoja.Range['H'+IntToStr(ren-1)+':H'+IntToStr(ren-1)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('dAvance').Value + OridAvanceAct ,xlCenter,false,False,'0.00%');
          OridAvanceAct:=OridAvanceAct + QrConsulta.FieldByName('dAvance').AsFloat;

          Hoja.Range['I'+IntToStr(ren-1)+':I'+IntToStr(ren-1)].Select;
          FormatoNormal(Excel,OridAvanceAnt + OridAvanceAct,xlCenter,false,False,'0.00%');

          if QrConsulta.FieldByName('sIdClasificacion').AsString='TE' then
            sEfectivo:=sfnSumaHoras(sEfectivo,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString))
          else
            sAfectaciones:=sfnSumaHoras(sAfectaciones,sfnRestaHoras(QrConsulta.FieldByName('sHoraFinal').AsString,QrConsulta.FieldByName('sHoraInicio').AsString));

        end;
        QrConsulta.Next;

      end;

      Hoja.Range['B'+IntToStr(IniciaR)+':I'+IntToStr(ren-1)].Select;
      GenerarMarco(Excel);

      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+2)].RowHeight := 15;

      Hoja.Range['B'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'DURACION TIEMPO EFECTIVO(HRS):',xlRight,false,False,'@');

      Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,sEfectivo,xlCenter,true,False,'@');


      inc(ren);

      Hoja.Range['B'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'DURACION TIEMPO AFECTACIONES(HRS):',xlRight,false,False,'@');

      Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,sAfectaciones,xlCenter,True,False,'@');

      inc(ren);

      Hoja.Range['B'+IntToStr(ren)+':E'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'TIEMPO TOTAL(HRS):',xlRight,false,False,'@');

      Hoja.Range['F'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,sfnSumaHoras(sEfectivo,sAfectaciones),xlCenter,true,False,'@');


      inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+1)].RowHeight := 10.2;

      Inc(ren,2);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 12.9;
      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PARTIDA',xlCenter,True);

      Application.ProcessMessages;

      Hoja.Range['D'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'ACTIVIDAD',xlCenter,True);
      Application.ProcessMessages;

      inc(Ren);

      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoNormal(Excel,QrPdas.FieldByName('sNumeroActividad').AsString,xlCenter,True,false,'@');

      Application.ProcessMessages;

      Hoja.Range['D'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,QrPdas.FieldByName('mDescripcion').AsString,xlJustify,True,true,'@',9);
      Application.ProcessMessages;

      Hoja.Range['B'+IntToStr(ren-1)+':L'+IntToStr(ren)].Select;
      GenerarMarco(Excel);

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 10.2;

      Inc(ren);
      {$ENDREGION}

      {$REGION 'BARCO'}
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+1)].RowHeight := 12.8;
      Hoja.Range['B'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'MOVIMIENTO DE EMBARCACION',xlCenter,false);

      Inc(ren);
      dCostoDll:=0;
      dCostoMn :=0;

      Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PARTIDA',xlCenter,false);

      Hoja.Range['C'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'DESCRIPCIÓN',xlCenter,false);

      Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'CLAS.',xlCenter,false);

      Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'CANTIDAD',xlCenter,false);

      Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PU MN',xlCenter,false);

      Hoja.Range['J'+IntToStr(ren)+':J'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PU USD',xlCenter,false);

      Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'IMP MN',xlCenter,false);

      Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'IMP USD',xlCenter,false);

      Hoja.Range['B'+IntToStr(ren-1)+':L'+IntToStr(ren)].Select;
      GenerarMarco(Excel);

      QrConsulta.Active:=False;
      QrConsulta.SQL.Text:= 'select b.sContrato, b.dIdFecha, b.sNumeroOrden, b.sNumeroActividad, b.sHoraInicio, b.shoraFinal,cast(b.sHoraInicio as Time) as HoraI, ' + #10 +
                          '(select ifnull(sum(ba.dAvance),0) from bitacoradeactividades as ba where ba.sContrato=b.sContrato and ba.sNumeroOrden=b.sNumeroOrden and ba.sIdTipoMovimiento=b.sIdTipoMovimiento' + #10 +
                          'and ba.swbs=b.swbs and ba.sNumeroActividad=b.sNumeroActividad and (ba.didfecha<b.didfecha or (ba.didfecha=b.didfecha and cast(ba.sHoraInicio as Time)<cast(b.sHoraInicio as Time) ))) as AvAnterior' + #10 +
                          'from bitacoradeactividades b inner join tiposdemovimiento tm on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=b.sIdClasificacion) ' +
                          'where b.sContrato=:Contrato and b.sNumeroOrden=:Orden and b.sIdTipoMovimiento=''ED''' + #10 +
                          'and b.swbs=:wbs and b.sNumeroActividad=:Actividad and tm.lAplicaNotaCampo="Si" group by b.didFecha ';
      QrConsulta.ParamByName('Contrato').AsString       := QrPdas.FieldByName('sContrato').AsString;
      QrConsulta.ParamByName('ContratoBarco').AsString  := global_Contrato_Barco;
      QrConsulta.ParamByName('Orden').AsString          := QrPdas.FieldByName('sNumeroOrden').AsString;
      QrConsulta.ParamByName('wbs').AsString            := QrPdas.FieldByName('swbs').AsString;
      QrConsulta.ParamByName('Actividad').AsString      := QrPdas.FieldByName('sNumeroActividad').AsString;
      QrConsulta.Open;

      QryMovimientos.Active := False;
      QryMovimientos.SQL.Clear;
      QryMovimientos.SQL.Add('select t.sContrato, t.sIdTipoMovimiento, t.sTipo, t.sDescripcion, t.dVentaMN, t.dVentaDLL '+
                       ' from tiposdemovimiento t '+
                       ' where t.sContrato =:contrato and t.sClasificacion = "Movimiento de Barco" and t.lAplicaNotaCampo = "Si" order by t.iOrden');
      QryMovimientos.ParamByName('contrato').AsString := global_Contrato_Barco;
      QryMovimientos.Open;

      IniciaR := Ren;
      while not QryMovimientos.Eof do
      begin
          Inc(Ren);
          Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QryMovimientos.FieldByName('sIdTipoMovimiento').Value,xlCenter,false,False,'@');

          Hoja.Range['C'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QryMovimientos.FieldByName('sDescripcion').Value,xlJustify,false,true,'@',4);

          Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QryMovimientos.FieldByName('sTipo').Value,xlCenter,false,False,'@');

          Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
          FormatoNormal(Excel,0,xlCenter,false,False,'#0.000000');

          Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QryMovimientos.FieldByName('dVentaMN').Value,xlCenter,false,False,'#,##0.00');

          Hoja.Range['J'+IntToStr(ren)+':J'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QryMovimientos.FieldByName('dVentaDLL').Value,xlCenter,false,False,'#,##0.00');

          Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QryMovimientos.FieldByName('dVentaMN').Value * 0 ,xlCenter,false,False,'#,##0.00');

          Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QryMovimientos.FieldByName('dVentaDLL').Value * 0 ,xlCenter,false,False,'#,##0.00');

          if QrPdas.FieldByName('lFactorBarco').AsString = 'Si' then
          begin
              {Aqui vamos a recorrer los dias en que se reportó la partida para consultar por horarios el movimiento de barco..}
              dFactorMovimiento := 0;
              dCantidadTotal    := 0;
              dCantidadPartida  := 0;
              QrConsulta.First;
              while not QrConsulta.Eof do
              begin
                  connection.QryBusca2.Active := False ;
                  connection.QryBusca2.SQL.Clear ;
                  connection.QryBusca2.SQL.Add(' select me.sClasificacion, mf.sFolio, mf.sNumeroOrden, sum(mf.sFactor) as sFactor '+
                                      'from movimientosdeembarcacion me '+
                                      'inner join movimientosxfolios mf on (mf.sContrato = me.sContrato and mf.dIdFecha = me.dIdFecha and mf.iIdDiario = me.iIdDiario '+
                                      'and mf.sNumeroOrden = me.sOrden and mf.sFolio = :Folio ) '+
                                      'where me.sContrato =:Contrato and me.sOrden =:Orden and me.sIdEmbarcacion =:Embarcacion and me.dIdFecha =:Fecha '+
                                      'and me.sClasificacion =:Clasificacion ' +
                                      'Group By me.sClasificacion order By me.sClasificacion');
                  connection.QryBusca2.Params.ParamByName('Contrato').DataType    := ftString ;
                  connection.QryBusca2.Params.ParamByName('Contrato').Value       := global_Contrato_Barco ;
                  connection.QryBusca2.Params.ParamByName('Embarcacion').DataType := ftString ;
                  connection.QryBusca2.Params.ParamByName('Embarcacion').Value    := global_barco ;
                  connection.QryBusca2.Params.ParamByName('Orden').DataType       := ftString ;
                  connection.QryBusca2.Params.ParamByName('Orden').Value          := Global_Contrato;
                  connection.QryBusca2.Params.ParamByName('Folio').DataType       := ftString ;
                  connection.QryBusca2.Params.ParamByName('Folio').Value          := QrPdas.FieldByName('sNumeroOrden').AsString;
                  connection.QryBusca2.Params.ParamByName('Fecha').DataType       := ftDate ;
                  connection.QryBusca2.Params.ParamByName('Fecha').Value          := QrConsulta.FieldByName('dIdFecha').AsDateTime;
                  connection.QryBusca2.Params.ParamByName('Clasificacion').DataType := ftString ;
                  connection.QryBusca2.Params.ParamByName('Clasificacion').Value    := QryMovimientos.FieldByName('sIdTipoMovimiento').AsString;
                  connection.QryBusca2.Open;

                  while not connection.QryBusca2.Eof do
                  begin
                      dFactorMovimiento := dFactorMovimiento +  connection.QryBusca2.FieldByName('sFactor').AsFloat;
                      {Aqui consultaos el tipo de movimiento por día}
                      connection.zCommand.SQL.Clear ;
                      connection.zCommand.SQL.Add(' select sNumeroActividad, ROUND(sum(dCantHHGenerador), 6) as dCantidad, sum(dCantHHGenerador) as dCantidadHH from bitacoradepersonal where sContrato =:Orden '+
                                                   'and sNumeroOrden  = :Folio and dIdFecha =:fecha group by sNumeroActividad');
                      connection.zCommand.Params.ParamByName('Orden').DataType  := ftString ;
                      connection.zCommand.Params.ParamByName('Orden').Value     := Global_Contrato;
                      connection.zCommand.Params.ParamByName('Folio').DataType  := ftString ;
                      connection.zCommand.Params.ParamByName('Folio').Value     := QrPdas.FieldByName('sNumeroOrden').AsString;
                      connection.zCommand.Params.ParamByName('Fecha').DataType  := ftDate ;
                      connection.zCommand.Params.ParamByName('Fecha').Value     := QrConsulta.FieldByName('dIdFecha').AsDateTime;
                      connection.zCommand.Open;

                      while not connection.zCommand.Eof do
                      begin
                          dCantidadTotal      := dCantidadTotal + connection.zCommand.FieldValues['dCantidad'];
                          if (QrPdas.FieldByName('sNumeroActividad').AsString = connection.zCommand.FieldValues['sNumeroActividad']) then
                              dCantidadPartida := dCantidadPartida + connection.zCommand.FieldValues['dCantidad'];
                          connection.zCommand.Next;
                          dCantidadTotalAux         := dCantidadTotal;
                          dCantidadPartidaAux       := dCantidadPartida;
                      end;
                      connection.QryBusca2.Next;
                  end;
                  QrConsulta.Next;
              end;
              Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
              if dCantidadTotal > 0 then
                 dFactorMovimiento := ((dCantidadPartida / dCantidadTotal) * dFactorMovimiento);

              FormatoNormal(Excel, xRound(dFactorMovimiento, 6) ,xlCenter,false,False,'#0.000000');

              Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
              FormatoNormal(Excel, xRound(QryMovimientos.FieldByName('dVentaMN').Value * xRound(dFactorMovimiento, 6), 2),xlCenter,false,False,'#,##0.00');
              dCostoMn:=dCostoMn + xRound((QryMovimientos.FieldByName('dVentaMN').AsFloat * xRound(dFactorMovimiento, 6)), 2) ;

              Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
              FormatoNormal(Excel, xRound(QryMovimientos.FieldByName('dVentaDLL').Value * xRound(dFactorMovimiento, 6), 2) ,xlCenter,false,False,'#,##0.00');
              dCostoDll:=dCostoDll + xRound((QryMovimientos.FieldByName('dVentaDLL').AsFloat * xRound(dFactorMovimiento, 6)), 2);
          end;
          QryMovimientos.Next;
      end;

      Hoja.Range['B'+IntToStr(IniciaR)+':L'+IntToStr(ren)].Select;
      GenerarMarco(Excel);

      Inc(Ren);

      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight :=15;

      Hoja.Range['B'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'IMPORTE BARCO:',xlCenter,false,False,'@');

      Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
      FormatoNormal(Excel,xRound(dCostoMn, 2),xlCenter,true,False,'$#,##0.00');

      Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,xRound(dCostoDll, 2),xlCenter,true,False,'$#,##0.00');

      dCostoDllT:=dCostoDllT + dCostoDll;
      dCostoMnT:=dCostoMnT + dCostoMn;

      Inc(ren,2);
      {$ENDREGION}

      {$REGION 'PERSONAL Y EQUIPO'}
      for I := 1 to 2 do
      begin
        dCostoDll:=0;
        dCostoMn:=0;


        IniciaR:=Ren;
        QrConsulta.Active:=False;
        if i=1 then
        begin

          QrConsulta.SQL.Text:= 'select r.iItemOrden, gp.sIdGrupo as '+ TablasConsult[I,2] +', SUM(r.dCantHHGenerador) as HH, p.dVentaMn as CostoMn,p.dVentaDll as CostoDll,gp.sDescripcion as DescRecurso,p.smedida from ' + TablasConsult[I,1] + ' r' + #10 +
                                'inner join bitacoradeactividades ba on(ba.sContrato=r.sContrato and ba.dIdFecha = r.dIdFecha and ba.iIdDiario=r.iIdDiario)' + #10 +
                                'inner join ' + TablasConsult[I,3] + ' p on(p.sContrato=:ContratoBarco and p.' + TablasConsult[I,2] + '=r.' + TablasConsult[I,2] + ')' + #10 +
                                'inner join grupospersonal gp on (gp.sIdGrupo=p.sAgrupaPersonal) ' +
                                'where r.sContrato=:Contrato AND r.sNumeroOrden = :Orden AND ba.sNumeroOrden=:Orden and ba.swbs=:wbs' + #10 +
                                'group by r.sContrato,gp.sIdGrupo order by p.iItemOrden';
        end
        else
        begin
          QrConsulta.SQL.Text:= 'select r.*, SUM(r.dCantHHGenerador) as HH, p.dVentaMN as CostoMn,p.dVentaDll as CostoDll,p.sDescripcion as DescRecurso,p.smedida from ' + TablasConsult[I,1] + ' r' + #10 +
                                'inner join bitacoradeactividades ba on(ba.sContrato=r.sContrato and ba.dIdFecha = r.dIdFecha and ba.iIdDiario=r.iIdDiario)' + #10 +
                                'inner join ' + TablasConsult[I,3] + ' p on(p.sContrato=:ContratoBarco and p.' + TablasConsult[I,2] + '=r.' + TablasConsult[I,2] + ')' + #10 +
                                'where r.sContrato=:Contrato AND r.sNumeroOrden = :Orden and ba.sNumeroOrden=:Orden and ba.swbs=:wbs' + #10 +
                                'group by r.sContrato,r.' + TablasConsult[I,2] +' order by p.iItemOrden';
        end;

        QrConsulta.ParamByName('Contrato').AsString:=QrPdas.FieldByName('sContrato').AsString;
        QrConsulta.ParamByName('Orden').AsString:=QrPdas.FieldByName('sNumeroOrden').AsString;
        QrConsulta.ParamByName('wbs').AsString:=QrPdas.FieldByName('swbs').AsString;
        QrConsulta.ParamByName('ContratoBarco').AsString:=global_Contrato_Barco;
        QrConsulta.Open;

        Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+1)].RowHeight := 12.8;
        Hoja.Range['B'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,TablasConsult[I,4],xlCenter,false);

        Inc(ren);

        Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'PARTIDA',xlCenter,false);

        Hoja.Range['C'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'DESCRIPCIÓN',xlCenter,false);

        Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'UNIDAD',xlCenter,false);

        Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'CANTIDAD',xlCenter,false);

        Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'PU MN',xlCenter,false);

        Hoja.Range['J'+IntToStr(ren)+':J'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'PU USD',xlCenter,false);

        Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'IMP MN',xlCenter,false);

        Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
        FormatoEncabezado(Excel,'IMP USD',xlCenter,false);

        while not QrConsulta.Eof do
        begin
          Inc(Ren);
          Hoja.Range['A'+IntToStr(ren)+':A'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('iItemOrden').Value,xlCenter,false,False,'@');

          Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName(TablasConsult[I,2]).Value,xlCenter,false,False,'@');

          Hoja.Range['C'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('DescRecurso').Value,xlJustify,false,true,'@',4);

          //DescRecurso
          Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sMedida').Value,xlCenter,false,False,'@');


          Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('HH').Value,xlCenter,false,False,'#0.00000000');

          Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('CostoMn').Value,xlCenter,false,False,'#,##0.00');


          Hoja.Range['J'+IntToStr(ren)+':J'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('CostoDll').Value,xlCenter,false,False,'#,##0.00');


          Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
          FormatoNormal(Excel, xRound(QrConsulta.FieldByName('CostoMn').Value * QrConsulta.FieldByName('HH').Value, 2),xlCenter,false,False,'#,##0.00');
          dCostoMn:=dCostoMn + xRound((QrConsulta.FieldByName('CostoMn').AsFloat * QrConsulta.FieldByName('HH').AsFloat), 2) ;

          Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
          FormatoNormal(Excel, xRound(QrConsulta.FieldByName('CostoDll').Value * QrConsulta.FieldByName('HH').Value, 2),xlCenter,false,False,'#,##0.00');
          dCostoDll:=dCostoDll + xRound((QrConsulta.FieldByName('CostoDll').AsFloat * QrConsulta.FieldByName('HH').AsFloat), 2);

          QrConsulta.Next;
        end;

        Hoja.Range['B'+IntToStr(IniciaR)+':L'+IntToStr(ren)].Select;
        GenerarMarco(Excel);

        Inc(ren);
        Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 14.3;

        Hoja.Range['B'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
        FormatoNormal(Excel,'IMPORTE '+TablasConsult[I,4]+':',xlCenter,false,False,'@');

        Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
        FormatoNormal(Excel,xRound(dCostoMn, 2),xlCenter,true,False,'$#,##0.00');

        Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
        FormatoNormal(Excel,xRound(dCostoDll, 2),xlCenter,true,False,'$#,##0.00');

        dCostoDllT:=dCostoDllT + dCostoDll;
        dCostoMnT:=dCostoMnT + dCostoMn;

        Inc(ren);
        Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 10.2;

        Inc(ren);
      end;
      {$ENDREGION}

      {$REGION 'PERNOCTAS'}
      IniciaR:=Ren;
      QrConsulta.Active:=False;
      QrConsulta.SQL.Text:= 'select c.*,ifnull((select sum(r.dCantidad) from bitacoradepersonal r' + #10 +
                            'inner join bitacoradeactividades ba on(ba.sContrato=r.sContrato and ba.iIdDiario=r.iIdDiario)' + #10 +
                            'where r.sContrato=:Contrato and ba.sNUmeroOrden=:Orden and ba.swbs=:wbs' + #10 +
                            'and r.sTipoPernocta=c.sIdCuenta' + #10 +
                            'group by r.sTipoPernocta),0) as TotalP from cuentas c';
      QrConsulta.ParamByName('Contrato').AsString:=QrPdas.FieldByName('sContrato').AsString;
      QrConsulta.ParamByName('Orden').AsString:=QrPdas.FieldByName('sNumeroOrden').AsString;
      QrConsulta.ParamByName('wbs').AsString:=QrPdas.FieldByName('swbs').AsString;
      QrConsulta.Open;

      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+1)].RowHeight := 12.8;
      Hoja.Range['B'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PERNOCTAS',xlCenter,false);

      Inc(ren);

      Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PARTIDA',xlCenter,false);

      Hoja.Range['C'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'DESCRIPCIÓN',xlCenter,false);

      Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'UNIDAD',xlCenter,false);

      Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'CANTIDAD',xlCenter,false);

      Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PU MN',xlCenter,false);

      Hoja.Range['J'+IntToStr(ren)+':J'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'PU USD',xlCenter,false);

      Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'IMP MN',xlCenter,false);

      Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'IMP USD',xlCenter,false);

      dCostoDll:=0;
      dCostoMn :=0;

      while not QrConsulta.Eof do
      begin
          Inc(Ren);
          Hoja.Range['B'+IntToStr(ren)+':B'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sIdCuenta').Value,xlCenter,false,False,'@');

          Hoja.Range['C'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sDescripcion').Value,xlJustify,false,true,'@',4);

          //DescRecurso
          Hoja.Range['G'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('sMedida').Value,xlCenter,false,False,'@');
          dProrrateoPernocta := 0;

          QryMovimientos.Active:=False;
          QryMovimientos.SQL.Text:= 'select b.sContrato, b.dIdFecha, b.sNumeroOrden, b.sNumeroActividad ' + #10 +
                                'from bitacoradeactividades b inner join tiposdemovimiento tm on(tm.sContrato=:ContratoBarco and tm.sIdTipoMovimiento=b.sIdClasificacion) ' +
                                'where b.sContrato=:Contrato and b.sNumeroOrden=:Orden and b.sIdTipoMovimiento=''ED''' + #10 +
                                'and b.swbs=:wbs and b.sNumeroActividad=:Actividad and tm.lAplicaNotaCampo="Si" group by b.didFecha ';
          QryMovimientos.ParamByName('Contrato').AsString       := QrPdas.FieldByName('sContrato').AsString;
          QryMovimientos.ParamByName('ContratoBarco').AsString  := global_Contrato_Barco;
          QryMovimientos.ParamByName('Orden').AsString          := QrPdas.FieldByName('sNumeroOrden').AsString;
          QryMovimientos.ParamByName('wbs').AsString            := QrPdas.FieldByName('swbs').AsString;
          QryMovimientos.ParamByName('Actividad').AsString      := QrPdas.FieldByName('sNumeroActividad').AsString;
          QryMovimientos.Open;

          {Aqui vamos a recorrer los dias en que se reportó la partida para consultar por horarios el movimiento de barco..}
          while not QryMovimientos.Eof do
          begin
              dCantidadTotal    := 0;
              dCantidadPartida  := 0;
              //Consultamos las pernoctas asignadas directamente..
              connection.QryBusca2.Active := False ;
              connection.QryBusca2.SQL.Clear ;
              connection.QryBusca2.SQL.Add('select ROUND(sum(dCantidad), 6) as dCantidad '+
                                          'from bitacoradepersonal_cuadre where sContrato =:Orden '+
                                          'and sNumeroOrden  = :Folio and dIdFecha =:Fecha and sTipoPernocta =:Pernocta group by sTipoPernocta');
              connection.QryBusca2.Params.ParamByName('Orden').DataType    := ftString ;
              connection.QryBusca2.Params.ParamByName('Orden').Value       := Global_Contrato;
              connection.QryBusca2.Params.ParamByName('Folio').DataType    := ftString ;
              connection.QryBusca2.Params.ParamByName('Folio').Value       := QrPdas.FieldByName('sNumeroOrden').AsString;
              connection.QryBusca2.Params.ParamByName('Fecha').DataType    := ftDate ;
              connection.QryBusca2.Params.ParamByName('Fecha').Value       := QryMovimientos.FieldByName('dIdFecha').AsDateTime;
              connection.QryBusca2.Params.ParamByName('Pernocta').DataType := ftString ;
              connection.QryBusca2.Params.ParamByName('Pernocta').Value    := QrConsulta.FieldByName('sIdCuenta').AsString;
              connection.QryBusca2.Open;

              dCantPernoctaAdicional := 0;
              if connection.QryBusca2.RecordCount > 0 then
                 dCantPernoctaAdicional := connection.QryBusca2.FieldValues['dCantidad'];

              {Aqui consultaos el tipo de movimiento por día}
              connection.zCommand.SQL.Clear ;
              connection.zCommand.SQL.Add(' select sNumeroActividad, ROUND(sum(dCantHHGenerador), 6) as dCantidad, sum(dCantHHGenerador) as dCantidadHH from bitacoradepersonal where sContrato =:Orden '+
                                           'and sNumeroOrden  = :Folio and dIdFecha =:fecha and sTipoPernocta =:Pernocta group by sNumeroActividad');
              connection.zCommand.Params.ParamByName('Orden').DataType  := ftString ;
              connection.zCommand.Params.ParamByName('Orden').Value     := Global_Contrato;
              connection.zCommand.Params.ParamByName('Folio').DataType  := ftString ;
              connection.zCommand.Params.ParamByName('Folio').Value     := QrPdas.FieldByName('sNumeroOrden').AsString;
              connection.zCommand.Params.ParamByName('Fecha').DataType  := ftDate ;
              connection.zCommand.Params.ParamByName('Fecha').Value     := QryMovimientos.FieldByName('dIdFecha').AsDateTime;
              connection.zCommand.Params.ParamByName('Pernocta').DataType := ftString ;
              connection.zCommand.Params.ParamByName('Pernocta').Value    := QrConsulta.FieldByName('sIdCuenta').AsString;
              connection.zCommand.Open;

              while not connection.zCommand.Eof do
              begin
                  dCantidadTotal      := dCantidadTotal + connection.zCommand.FieldValues['dCantidadHH'];
                  if (QrPdas.FieldByName('sNumeroActividad').AsString = connection.zCommand.FieldValues['sNumeroActividad']) then
                      dCantidadPartida := dCantidadPartida + connection.zCommand.FieldValues['dCantidadHH'];
                  connection.zCommand.Next;
                  dCantidadTotalAux         := dCantidadTotal;
                  dCantidadPartidaAux       := dCantidadPartida;
              end;

              if dCantidadPartida > 0 then
                 dProrrateoPernocta := dProrrateoPernocta + ((dCantidadPartidaAux / dCantidadTotalAux) * (dCantidadTotalAux + dCantPernoctaAdicional));

              {Posteriormente sumamos o restamos las penoctas pertenecientes a la bitacoradepernocta}
              connection.QryBusca2.Active := False ;
              connection.QryBusca2.SQL.Clear ;
              connection.QryBusca2.SQL.Add('select ROUND(sum(dCantidad), 6) as dCantidad '+
                                          'from bitacoradepernocta where sContrato =:Orden '+
                                          'and sNumeroOrden  = :Folio and dIdFecha =:Fecha group by sContrato');
              connection.QryBusca2.Params.ParamByName('Orden').DataType    := ftString ;
              connection.QryBusca2.Params.ParamByName('Orden').Value       := Global_Contrato;
              connection.QryBusca2.Params.ParamByName('Folio').DataType    := ftString ;
              connection.QryBusca2.Params.ParamByName('Folio').Value       := QrPdas.FieldByName('sNumeroOrden').AsString;
              connection.QryBusca2.Params.ParamByName('Fecha').DataType    := ftDate ;
              connection.QryBusca2.Params.ParamByName('Fecha').Value       := QryMovimientos.FieldByName('dIdFecha').AsDateTime;
              connection.QryBusca2.Open;

              if connection.QryBusca2.RecordCount > 0 then
                 if dProrrateoPernocta > 0 then
                    dProrrateoPernocta := dProrrateoPernocta + connection.QryBusca2.FieldValues['dCantidad'];

              if dProrrateoPernocta < 0 then
                 dProrrateoPernocta := 0;

              QryMovimientos.Next;
          end;

          Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
          FormatoNormal(Excel,dProrrateoPernocta, xlCenter, False, False, '#0.00000000');

          Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('dVentaMn').Value,xlCenter,false,False,'#,##0.00');

          Hoja.Range['J'+IntToStr(ren)+':J'+IntToStr(ren)].Select;
          FormatoNormal(Excel,QrConsulta.FieldByName('dVentaDll').Value,xlCenter,false,False,'#,##0.00');

          Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
          FormatoNormal(Excel, xRound(QrConsulta.FieldByName('dVentaMn').Value * dProrrateoPernocta, 2) ,xlCenter,false,False,'#,##0.00');
          dCostoMn:=dCostoMn + xRound((QrConsulta.FieldByName('dVentaMn').AsFloat *dProrrateoPernocta), 2);

          Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
          FormatoNormal(Excel, xRound(QrConsulta.FieldByName('dVentaDll').Value * dProrrateoPernocta, 2),xlCenter,false,False,'#,##0.00');
          dCostoDll:=dCostoDll + xRound((QrConsulta.FieldByName('dVentaDll').AsFloat * dProrrateoPernocta), 2);

          QrConsulta.Next;
      end;

      Hoja.Range['B'+IntToStr(IniciaR)+':L'+IntToStr(ren)].Select;
      GenerarMarco(Excel);

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight :=15;

      Hoja.Range['B'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'IMPORTE PERNOCTAS:',xlCenter,false,False,'@');

      Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
      FormatoNormal(Excel,xRound(dCostoMn, 2),xlCenter,true,False,'$#,##0.00');

      Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,xRound(dCostoDll, 2),xlCenter,true,False,'$#,##0.00');

      dCostoDllT:=dCostoDllT + dCostoDll;
      dCostoMnT:=dCostoMnT + dCostoMn;
      {$ENDREGION}

      {$REGION 'MATERIALES'}

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight := 10.2;

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren+1)].RowHeight := 11.3;
      Hoja.Range['B'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'MATERIAL',xlCenter,false);

      Inc(ren);

      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'TRAZABILIDAD',xlCenter,false);

      Hoja.Range['D'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'DESCRIPCIÓN',xlCenter,false);

      Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'UNIDAD',xlCenter,false);

      Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
      FormatoEncabezado(Excel,'CANTIDAD',xlCenter,false);

      IniciaR:=Ren;
      QrConsulta.Active:=False;
      QrConsulta.SQL.Text:= 'select s.sNumeroOrden, s.sNumeroActividad, b.sTrazabilidad, i.mDescripcion, i.sMedida, sum(b.dCantidad) as dCantidad from almacen_salida s '+
                            'inner join bitacoradesalida b '+
                            'on(b.sContrato = s.sContrato and b.iFolioSalida = s.iFolioSalida and b.sext = s.sext) '+
                            'inner join insumos i '+
                            'on (i.sContrato = :Contrato and i.sIdInsumo = b.sIdInsumo ) '+
                            'where s.sContrato =:Orden and s.sNumeroOrden =:Folio and s.sNumeroActividad =:Actividad '+
                            'group by b.sIdInsumo, b.sTrazabilidad Order by b.sTrazabilidad';
      QrConsulta.ParamByName('Contrato').AsString  := global_Contrato_Barco;
      QrConsulta.ParamByName('Orden').AsString     := QrPdas.FieldByName('sContrato').AsString;
      QrConsulta.ParamByName('Folio').AsString     := QrPdas.FieldByName('sNumeroOrden').AsString;
      QrConsulta.ParamByName('Actividad').AsString := QrPdas.FieldByName('sNumeroActividad').AsString;
      QrConsulta.Open;

      while not QrConsulta.Eof do
      begin
        Inc(Ren);

        Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
        FormatoNormal(Excel,QrConsulta.FieldByName('sTrazabilidad').Value,xlCenter,false,False,'@');
        //FormatoNormal(Excel,QrConsulta.FieldByName('AvAnterior').Value + QrConsulta.FieldByName('dAvance').Value,xlCenter,false,False,'@');

        Hoja.Range['D'+IntToStr(ren)+':G'+IntToStr(ren)].Select;
        FormatoNormal(Excel,QrConsulta.FieldByName('mDescripcion').Value,xlJustify,false,true,'@',4);

        //DescRecurso
        Hoja.Range['H'+IntToStr(ren)+':H'+IntToStr(ren)].Select;
        FormatoNormal(Excel,QrConsulta.FieldByName('sMedida').Value,xlCenter,false,False,'@');


        Hoja.Range['I'+IntToStr(ren)+':I'+IntToStr(ren)].Select;
        FormatoNormal(Excel,QrConsulta.FieldByName('dCantidad').Value,xlCenter,false,False,'#,##0.00');

        QrConsulta.Next;

      end;

      Hoja.Range['B'+IntToStr(IniciaR-1)+':I'+IntToStr(ren)].Select;
      GenerarMarco(Excel);

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight :=10.2;

      Inc(ren);
      Excel.Rows[IntToStr(ren)+':'+IntToStr(ren)].RowHeight :=11.3;

      Hoja.Range['B'+IntToStr(ren)+':F'+IntToStr(ren)].Select;
      FormatoNormal(Excel,'COSTO TOTAL DE LA ACTIVIDAD::',xlCenter,false,False,'@');

      Hoja.Range['K'+IntToStr(ren)+':K'+IntToStr(ren)].Select;
      FormatoNormal(Excel,xRound(dCostoMnT, 2),xlCenter,true,False,'$#,##0.00');

      Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
      FormatoNormal(Excel,xRound(dCostoDllT, 2),xlCenter,true,False,'$#,##0.00');

      Inc(ren);
      {$ENDREGION}

    end;
    QrPdas.Next;
  end;
  PanelProgress.Visible := False;
  grid_actividades.Enabled := True;
end;

Procedure TfrmActividades.NotaCampoExcel;
var
  Excel,Hoja:Variant;
  sFileName:string;
  sDescFrente:string;
  pidl: PItemIDList;
  InFolder: array[0..MAX_PATH] of Char;
  QrConfiguracion,QrAux : TZReadOnlyQuery;

  Embarcacion : string;
begin
  sFileName:='';
  sDescFrente:='';
  tmpNombre:='';
  tmpNombreC:='';
  // Armar el nombre de archivo
  SHGetSpecialFolderLocation(application.Handle, CSIDL_PERSONAL, pidl);
  SHGetPathFromIDList(PIDL, InFolder);
  sFileName := InFolder;

  if sFileName[Length(sFileName)] <> '\' then
    sFileName := sFileName + '\';

  //sFileName := sFileName + 'Reporte Diario ' + IntToStr(YearOf(ReporteDiario.FieldByName('dIdFecha').AsDateTime)) + '-' + IntToStr(MonthOf(ReporteDiario.FieldByName('dIdFecha').AsDateTime)) + '-' + IntToStr(DayOf(ReporteDiario.FieldByName('dIdFecha').AsDateTime)) + '.xls';

  SdgExcel.InitialDir:= sFileName;
  SdgExcel.FileName:='Nota_Campo_Folio ' + ActividadesxOrden.FieldByName('sNumeroOrden').AsString + '.xlsx';
 // SdgExcel.DefaultExt:='XLSX';
  if SdgExcel.Execute then
  begin
    try
        Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el ambiente de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    try
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;

      Libro := Excel.Workbooks.Add;    // Crear el libro sobre el que se ha de trabajar
      while Libro.Sheets.Count < 9 do
        Libro.Sheets.Add;                                  // Verificar si cuenta con las hojas necesarias          

      Hoja := Libro.Sheets[1];
        Hoja.Select;
        //Libro.Sheets[1].Select;
        while Libro.Sheets.Count > 9 do
          Excel.ActiveWindow.SelectedSheets.Delete;

        QrConfiguracion := TZReadOnlyQuery.Create(Nil);
        QrConfiguracion.Connection := Connection.zConnection;
        QrConfiguracion.SQL.Text:='select c.sMostrarAvances,c.iFirmas, c.sOrdenPerEq, c.sTipoPartida, c.sImprimePEP,ot.sIdplataforma as localizacion, ' +
            ' (select sContrato from contratos where sContrato=:ContratoBarco and sTipoObra = "BARCO" ) as sContratoBarco, ' +
            ' (select mDescripcion from contratos where sContrato=:ContratoBarco and sTipoObra = "BARCO" ) as mDescripcionBarco, ' +
            'c.sClaveSeguridad, c.cStatusProceso, c.sOrdenExtraordinaria, c.lLicencia, c.sReportesCIA, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3,' +
            'ot.bAvanceFrente, ot.bAvanceContrato, ot.bComentarios, ot.bPermisos, ot.lMostrarAvanceProgramado, ot.lImprimePersonalTM, ot.lPersonalxPartida, ' +
            'c.bImagen, c.sContrato, c.sNombre, c2.sCodigo, c2.sProrrateoBarco, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, c.lImprimeExtraordinario, ' +
            'c2.mDescripcion, c2.sTitulo, c2.mCliente, c2.bImagen as bImagenPEP, ot.lImprimeFases, cv.dFechaInicio, cv.dfechaFinal, ' +
            'ot.mdescripcion as DescFolio, ot.sNumeroOrden From contratos c2 INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
            'inner join ordenesdetrabajo ot on (ot.sContrato = c2.sContrato and ot.sNumeroOrden =:Orden ) ' +
            'inner join convenios cv on (cv.sContrato = c2.sContrato and cv.sIdConvenio =:convenio) '+
            'Where c2.sContrato = :Contrato';
        QrConfiguracion.ParamByName('contrato').AsString:= global_contrato;
        QrConfiguracion.ParamByName('convenio').AsString:= zqReprogramacion.FieldByName('sIdConvenio').AsString;
        QrConfiguracion.ParamByName('Orden').AsString:= tsNumeroOrden.KeyValue;
        QrConfiguracion.ParamByName('ContratoBarco').AsString:= global_Contrato_barco;
        //QrConfiguracion.ParamByName('Fecha').AsDate:= ReporteDiario.FieldByName('dIdFecha').AsDateTime;
        QrConfiguracion.Open;
        Application.ProcessMessages;
        TamFont:=7;

        Application.ProcessMessages;
        //nHoja:=0;
        PonerEncabezado(Excel,Hoja,QrConfiguracion);

        QrAux:=TZReadOnlyQuery.Create(nil);
        QrAux.Connection:=connection.zConnection;
        QrAux.Active := False;
        QrAux.SQL.Clear;
        QrAux.SQL.Add('select * from actividadesxorden where scontrato=:Contrato and sIdConvenio=:Convenio and sNumeroOrden=:Orden and '+
                      ' sWbs like :Wbs order by iItemOrden ');
        if ActividadesxOrden.FieldByName('sTipoActividad').AsString='Paquete' then
           QrAux.ParamByName('Wbs').AsString   := ActividadesxOrden.FieldValues['sWbs'] + '.%'
        else
           QrAux.ParamByName('Wbs').AsString   := ActividadesxOrden.FieldValues['sWbs'];
        QrAux.ParamByName('Contrato').AsString := ActividadesxOrden.FieldByName('sContrato').AsString;
        QrAux.ParamByName('Convenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
        QrAux.ParamByName('Orden').AsString    := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
        QrAux.Open;

        //Primero el Id de la Embarcacion principal... OSA 2013 ivan,,
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('select sIdEmbarcacion from embarcacion_vigencia '+
                             'where sContrato =:Contrato and dFechaInicio <= :Fecha and dFechaFinal >=:Fecha order by dFechaInicio');
        connection.QryBusca.ParamByName('Contrato').AsString := global_contrato_barco;
        connection.QryBusca.ParamByName('Fecha').AsDate      := actividadesxorden.FieldValues['dFechaFinal'];
        connection.QryBusca.Open;

        if connection.QryBusca.RecordCount > 0 then
            global_barco := connection.QryBusca.FieldValues['sIdEmbarcacion']
        else
           messageDLG('No existe una Vigencia de Embarcacion Principal', mtInformation, [mbOk], 0);

        ActualiaFactorGeneradorPER(global_barco, global_contrato, ordenesdetrabajo.FieldValues['sNumeroOrden']);
        ActualiaFactorGeneradorEQ(global_barco, global_contrato, ordenesdetrabajo.FieldValues['sNumeroOrden']);
        if Connection.Contrato.FieldByName('sIdResidencia').AsString = '03' then
          GenerarNotaPdasSinBarco(Excel,Hoja,QrConfiguracion,QrAux)
        else
          GenerarNotaPdas(Excel,Hoja,QrConfiguracion,QrAux);
        ConfigurarHoja(Excel,Hoja);

    finally
      Excel.ActiveWorkbook.SaveAs(SdgExcel.FileName);
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      if FileExists(tmpNombre) then
        DeleteFile(tmpNombre);

      if FileExists(tmpNombreC) then
        DeleteFile(tmpNombreC);
    end;
  end;

end;

Procedure TfrmActividades.PonerEncabezado(var Excel: Variant; var Hoja: Variant; var QrDatos: TZReadOnlyQuery);
var
  CadFecha:string;
  TempPath: array [0..MAX_PATH-1] of Char;
  imgAux:TImage;
  Pic : TJpegImage;
  fs: tStream;
  Altura, Margen:Extended;
  vPicture:Variant;
begin
  imgAux:=TImage.Create(nil);
  Excel.ActiveWindow.Zoom := 46;
  Excel.Columns['A:A'].ColumnWidth :=0.75;
  Excel.Columns['B:B'].ColumnWidth := 8.33;
  Excel.Columns['C:D'].ColumnWidth := 8.89;
  Excel.Columns['E:E'].ColumnWidth := 10.11;
  Excel.Columns['F:F'].ColumnWidth := 9.67;
  Excel.Columns['G:G'].ColumnWidth := 10.56;
  Excel.Columns['H:H'].ColumnWidth := 9.89;
  Excel.Columns['I:I'].ColumnWidth := 9.11;
  Excel.Columns['J:J'].ColumnWidth := 7.67;
  Excel.Columns['K:K'].ColumnWidth := 11.33;
  Excel.Columns['L:L'].ColumnWidth := 10.11;
  Excel.Columns['M:M'].ColumnWidth := 1.89;
  Excel.Columns['N:N'].ColumnWidth := 10;

  Excel.Rows['1:2'].RowHeight := 10.2;
  Excel.Rows['3:3'].RowHeight := 12.8;
  Excel.Rows['4:4'].RowHeight := 10.2;
  Excel.Rows['5:5'].RowHeight := 18;
  Excel.Rows['6:6'].RowHeight := 10.2;
  Excel.Rows['7:7'].RowHeight := 47.3;
  Excel.Rows['8:8'].RowHeight := 21.8;
  Excel.Rows['9:11'].RowHeight := 10.2;
  //Excel.Rows['10:10'].RowHeight := 32.62;
  application.ProcessMessages;
  // Colocar los encabezados del reporte

  Hoja.Range['B3:L3'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont+3;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  if QrDatos.RecordCount=1 then
    Excel.Selection.Value :='NOTA DE CAMPO';//QrDatos.FieldByName('mDescripcionBarco').AsString;

        (*'"ASEGURAMIENTO DE LA INTEGRIDAD Y CONFIABILIDAD' + #10 +
                                 'DEL SISTEMA DE TRANSPORTE DE HIDROCARBUROS POR DUCTOS, DE PEP, SISTEMA 1"' + #10 + #10 +
                                 'DESMANTELAMIENTO DE DUCTOS FUERA DE OPERACIÓN Y EMPACADOS';   *)

  Hoja.Range['B6:C6'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := 'CONTRATO:';
  Application.ProcessMessages;

  Hoja.Range['D6:F6'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := false;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := QrDatos.FieldByName('sContratoBarco').AsString;
  Application.ProcessMessages;

  Hoja.Range['G6:H6'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := 'FOLIO:';
  Application.ProcessMessages;

  Hoja.Range['I6:L6'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := QrDatos.FieldByName('sNumeroOrden').AsString;
  Application.ProcessMessages;

  Hoja.Range['B7:C8'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := 'DESCRIPCIÓN:';
  Application.ProcessMessages;

  Hoja.Range['D7:F8'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlJustify ;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := false;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := QrDatos.FieldByName('mDescripcionBarco').AsString;
  Application.ProcessMessages;

  Hoja.Range['G7:H7'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := 'OBRA:';
  Application.ProcessMessages;

  Hoja.Range['I7:L7'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlJustify;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := QrDatos.FieldByName('DescFolio').AsString;
  Application.ProcessMessages;

  Hoja.Range['G8:H8'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := 'LOCALIZACIÓN:';
  Application.ProcessMessages;

  Hoja.Range['I8:L8'].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Value := 'PLATAFORMA: '+ QrDatos.FieldByName('localizacion').AsString;
  //Excel.Rows('7:7').RowHeight := 65;
  Excel.Rows[IntToStr(7) + ':' + IntToStr(7)].RowHeight := 60;
  Application.ProcessMessages;


  // Obtener la imagen del cliente desde la base de datos

  if tmpNombreC='' then
  begin

    
    GetTempPath(SizeOf(TempPath), TempPath);
    tmpNombreC:=TempPath +'imgtempSln'+formatdatetime('dddddd hhnnss',now)+'.jpg';

    fs := QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
    If fs.Size > 1 Then
    Begin
      try
        Pic:=TJpegImage.Create;
        try
           Pic.LoadFromStream(fs);
           imgAux.Picture.Graphic := Pic;
        finally
           Pic.Free;
        end;
      finally
          fs.Free;
      End;
      imgAux.Picture.SaveToFile(tmpNombreC);
      
    End;
  end;

  Application.ProcessMessages;
  if FileExists(tmpNombreC) then
  begin
    // Agregar Imagen de la empresa
    Hoja.Cells[1,1].Select;
    //Excel.ActiveSheet.Pictures.Insert(tmpNombreC).Select;
    vPicture:=Excel.ActiveSheet.Pictures.Insert(tmpNombreC);
    // Determinar el tamaño real de la imagen
    Altura := (Excel.Rows[1].Height + Excel.Rows[2].Height + Excel.Rows[3].Height + Excel.Rows[4].Height + + (Excel.Rows[5].Height/2));   // * 0.7;
    Margen := Excel.Rows[1].Height ;//0;  //(Excel.Rows[1].Height + Excel.Rows[2].Height + Excel.Rows[3].Height + Excel.Rows[4].Height + Excel.Rows[5].Height - Altura) / 2;

    vPicture.ShapeRange.Left :=Excel.Columns['A:A'].Width;  //Margen;    //Excel.Columns['A:A'].Width + Margen;
    vPicture.ShapeRange.Top := 0;

    vPicture.ShapeRange.Width:=Excel.Columns['A:A'].Width + Excel.Columns['B:B'].Width+ (Excel.Columns['A:A'].Width/2);
    vPicture.ShapeRange.Height:=Altura;
    Excel.ActiveSheet.Shapes.AddPicture(tmpNombreC,false,True,vPicture.ShapeRange.Left,vPicture.ShapeRange.Top,vPicture.ShapeRange.Width,vPicture.ShapeRange.Height);
    vPicture.delete;
  end;
  Application.ProcessMessages;
  //if Color1<>RGB(150,150,150) then
  //Excel.Selection.ShapeRange.PictureFormat.ColorType := msoPictureGrayscale;
  if tmpNombre='' then
  begin
         // tmpNombre := GetTempFile('.~im');
    GetTempPath(SizeOf(TempPath), TempPath);
    tmpNombre:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
    fs := QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagen'), bmRead);
    If fs.Size > 1 Then
    Begin
      try
        Pic:=TJpegImage.Create;
        try
          Pic.LoadFromStream(fs);
          imgAux.Picture.Graphic := Pic;
        finally
          Pic.Free;
        end;
      finally
        fs.Free;
      End;
      imgAux.Picture.SaveToFile(tmpNombre);
    End;
  end;

  if FileExists(tmpNombre) then
  begin
    // Agregar Imagen Cliente a la hoja de excel
    Hoja.Cells[1,1].Select;
    Excel.ActiveSheet.Pictures.Insert(tmpNombre).select;
    Excel.Selection.Cut;

    Hoja.Cells[2,11].Select;
    Hoja.Paste;

     // Determinar el tamaño real de la imagen
    Altura := (Excel.Rows[1].Height +Excel.Rows[2].Height + Excel.Rows[3].Height + Excel.Rows[4].Height+ (Excel.Rows[5].Height)) ;   // * 0.7;
    Margen := 0;
    Excel.Selection.ShapeRange.Top := 0;
    Excel.Selection.ShapeRange.Width:=(Excel.Columns['K:K'].Width/2) + (Excel.Columns['L:L'].Width/2);
    Excel.Selection.ShapeRange.Height:=Altura;
    Excel.Selection.ShapeRange.IncrementLeft((Excel.Columns['K:K'].Width/2));

    Excel.ActiveSheet.Shapes.AddPicture(tmpNombre,false,True,Excel.Selection.ShapeRange.Left,Excel.Selection.ShapeRange.Top,Excel.Selection.ShapeRange.Width,Excel.Selection.ShapeRange.Height);
    Excel.Selection.ShapeRange.delete;
  end;

  Application.ProcessMessages;
  Hoja.Range['B6:L8'].Select;
  Excel.Selection.Borders[xlEdgeLeft].LineStyle         := xlContinuous;
  Excel.Selection.Borders[xlEdgeLeft].Weight            := xlThin;
  Excel.Selection.Borders[xlEdgeTop].LineStyle          := xlContinuous;
  Excel.Selection.Borders[xlEdgeTop].Weight             := xlThin;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle       := xlContinuous;
  Excel.Selection.Borders[xlEdgeBottom].Weight          := xlThin;
  Excel.Selection.Borders[xlEdgeRight].LineStyle        := xlContinuous;
  Excel.Selection.Borders[xlEdgeRight].Weight           := xlThin;
  Excel.Selection.Borders[xlInsideVertical].LineStyle   := xlContinuous ;
  Excel.Selection.Borders[xlInsideVertical].Weight      := xlThin;
  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideHorizontal].Weight    := xlThin;
  Excel.Selection.Borders[xlDiagonalDown].LineStyle := xlNone;
  Excel.Selection.Borders[xlDiagonalUp].LineStyle := xlNone;


end;

function TfrmActividades.lExisteMedida(sMedida: string): Boolean;
begin
  lExisteMedida := False;
  lExisteMedida := (strPos(pchar(connection.configuracion.FieldByName('txtMaterialAutomatico').AsString), pchar(sMedida + '|')) <> nil)
end;

procedure TfrmActividades.FormShow(Sender: TObject);
var
  i, x, y, z: integer;
begin
  isOpen:=False;
  sIdFrente:=global_FrenteTrabajo;
  sMenuP := stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cActividades', PopupPrincipal);
  UtGrid := TicdbGrid.create(grid_actividades);
    {Definimos colores...}
  colores[1, 1] := 3;
  colores[2, 1] := 5;
  colores[3, 1] := 6;
  colores[4, 1] := 8;
  colores[5, 1] := 10;
  colores[6, 1] := 11;
  colores[7, 1] := 12;
  colores[8, 1] := 13;
  colores[9, 1] := 14;
  colores[10, 1] := 15;

    {Colocamor color texto..}
  colores[1, 2] := 1;
  colores[2, 2] := 1;
  colores[3, 2] := 0;
  colores[4, 2] := 0;
  colores[5, 2] := 0;
  colores[6, 2] := 1;
  colores[7, 2] := 1;
  colores[8, 2] := 1;
  colores[9, 2] := 1;
  colores[10, 2] := 0;

    // ivan - > Llenado del array de las columnas del Excel..
  for x := 1 to 26 do
    columnas[x] := Chr(64 + x);

  i := 27;
  for x := 1 to 26 do
  begin
    for y := 1 to 26 do
    begin
      columnas[i] := Chr(64 + x) + Chr(64 + y);
      i := i + 1;
    end;
  end;

  for x := 1 to 1 do
  begin
    for y := 1 to 26 do
    begin
      for z := 1 to 26 do
      begin
        columnas[i] := Chr(64 + x) + Chr(64 + y) + Chr(64 + z);
        i := i + 1;
      end;
    end;
  end;


  iItemOrden := 0;
  sPaquete := '';
  OpcButton := '';

  OrdenesdeTrabajo.Active := False;
  OrdenesdeTrabajo.SQL.Clear;
  if (global_grupo = 'INTEL-CODE') then
    OrdenesdeTrabajo.SQL.Add('Select o.dFechaInicioT, o.dFechaSitioM, o.sDepSolicitante, o.dfiProgramado, o.dffProgramado, o.sNumeroOrden, o.sDescripcionCorta, ' +
      'o.sIdTipoOrden, o.sIdPlataforma, o.sEquipo, o.sPozo, o.dFechaElaboracion, o.sPuestoPEP, o.sFirmantePEP, o.sPuestoCia, o.sFirmantecia ' +
      'from ordenesdetrabajo o ' +
      'INNER JOIN ordenesxusuario ou On (ou.sContrato=o.sContrato ' +
      'And ou.sNumeroOrden=o.sNumeroOrden and ou.sIdUsuario =:Usuario ) ' +
      'where o.sContrato = :Contrato And o.cIdStatus =:Status Order by o.sNumeroOrden')
  else
    OrdenesdeTrabajo.SQL.Add('Select ot.dFechaInicioT, ot.dFechaSitioM, ot.sDepSolicitante, ot.dfiProgramado, ot.dffProgramado, ' +
      'ot.sNumeroOrden, ot.sDescripcionCorta, ot.sIdTipoOrden, ot.sIdPlataforma, ot.sEquipo, ot.sPozo, ' +
      'ot.dFechaElaboracion, ot.sPuestoPEP, ot.sFirmantePEP, ot.sPuestoCia, ot.sFirmantecia from ordenesdetrabajo ot ' +
      'INNER JOIN ordenesxusuario ou On (ou.sContrato=ot.sContrato ' +
      'And ou.sNumeroOrden=ot.sNumeroOrden) ' +
      'where ot.sContrato =:Contrato And ou.sDerechos<>"BLOQUEADO" ' +
      'And ou.sIdUsuario =:Usuario And ot.cIdStatus =:Status order by ot.sNumeroOrden');
  OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString;
  OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := Global_Contrato;
  OrdenesdeTrabajo.Params.ParamByName('status').DataType := ftString;
  OrdenesdeTrabajo.Params.ParamByName('status').Value := connection.configuracion.FieldValues['cStatusProceso'];
  OrdenesdeTrabajo.Params.ParamByName('Usuario').DataType := ftString;
  OrdenesdeTrabajo.Params.ParamByName('Usuario').Value := Global_Usuario;
  OrdenesdeTrabajo.Open;

  zqPlataformas.Active := False;
  zqPlataformas.Open;

  if OrdenesdeTrabajo.RecordCount > 0 then
  begin
    if sIdFrente='' then
    begin
      tsNumeroOrden.KeyValue := OrdenesdeTrabajo.FieldValues['sNumeroOrden'];
      sNumeroOrden := OrdenesdeTrabajo.FieldValues['sNumeroOrden'];
    end
    else
    begin
      tsNumeroOrden.KeyValue := sIdFrente;
      sNumeroOrden := sIdFrente;

    end;
    ConsultaReprogramacion;
    if zqReprogramacion.RecordCount > 0 then
       tsReprogramacion.KeyValue := zqReprogramacion.FieldByName('sIdConvenio').AsString;
    tsReprogramacion.OnExit(sender);
    Grid_Actividades.SetFocus
  end
  else
    tsNumeroOrden.SetFocus;

  zFasesProyecto.Active := False;
  zFasesProyecto.Open;

  ConsultaFolios;

  isOpen:=true;
  ActividadesxOrdenAfterScroll(ActividadesxOrden);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmActividades.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsReprogramacion.SetFocus
end;

procedure TfrmActividades.tsNumeroActividadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tmDescripcion.SetFocus
end;

procedure TfrmActividades.tdFechaInicioKeyPress(Sender: TObject;
  var Key: Char);
begin
//  if key = #13 then
//    tdDuracion.SetFocus
end;

procedure TfrmActividades.tdDuracionKeyPress(Sender: TObject;
  var Key: Char);
begin
//  if not keyFiltroTDBEdit(tdDuracion, key) then
//    key := #0;
  if key = #13 then
    tdFechaFinal.SetFocus
end;


procedure TfrmActividades.frmBarra1btnAddClick(Sender: TObject);
var
  sActividad: string;
begin
  activapop(frmActividades, popupprincipal);
  if tsNumeroOrden.Text <> '' then
  begin
    Insertar1.Enabled := False;
    Editar1.Enabled := False;
    Registrar1.Enabled := True;
    Can1.Enabled := True;
    Eliminar1.Enabled := False;
    Refresh1.Enabled := False;
    Salir1.Enabled := False;
    frmBarra1.btnAddClick(Sender);
    sOpcion := 'Nuevo';

    tdFechaInicio.Date := Date;
    tdFechaFinal.Date := Date;

    if ActividadesxOrden.RecordCount = 0 then
    begin
      sActividad := 'A';
      tsNumeroActividad.ReadOnly := True;
      try
          //Aquí insertamos el convenio 1 en automático en repogramaciones del contrato
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('insert into convenios (sContrato, sNumeroOrden, sIdConvenio, sDescripcion, dFecha, dFechaInicio, dFechaFinal) '+
                                      'values (:contrato, :orden, :convenio, :descripcion, :Fecha, :fechaI, :fechaF)');
          Connection.zCommand.Params.ParamByName('Contrato').AsString  := Global_Contrato;
          Connection.zCommand.Params.ParamByName('Convenio').AsString  := '1';
          Connection.zCommand.Params.ParamByName('Orden').AsString     := tsNumeroOrden.Text;
          Connection.zCommand.Params.ParamByName('descripcion').AsString := tsNumeroOrden.Text +' C-1';
          Connection.zCommand.Params.ParamByName('fecha').AsDate         := Now();
          Connection.zCommand.Params.ParamByName('fechaI').AsDate        := Now();
          Connection.zCommand.Params.ParamByName('fechaF').AsDate        := Now();
          connection.zCommand.ExecSQL();
          zqReprogramacion.Refresh;
      Except  
      end;
    end
    else
    begin
      sActividad := '';
      tsNumeroActividad.ReadOnly := false;
    end;
    if ActividadesxOrden.RecordCount > 0 then
      if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
        sPaquete := ActividadesxOrden.FieldValues['sWbs']
      else
        sPaquete := ActividadesxOrden.FieldValues['sWbsAnterior']
    else
      sPaquete := '';
    sPaqueteDesc := '';

    if tsPaquete.KeyValue <> null then
    begin
      tsPaquete.KeyValue := sPaquete;
      sPaqueteDesc := tsPaquete.Text;
    end;

    if Paquetes.RecordCount > 0 then
    begin
      OrdenPaqueteItem := Paquetes.FieldValues['iItemOrden'];
      OrdenPaqueteWbs := Paquetes.FieldValues['sWbs'];
      OrdenPaqueteNivel := Paquetes.FieldValues['iNivel'];
    end;

    ActividadesxOrden.Append;
    //ActividadesxOrden.Insert;
    ActividadesxOrden.FieldValues['sContrato'] := Global_Contrato;
    ActividadesxOrden.FieldValues['sIdConvenio'] := zqReprogramacion.FieldByName('sIdConvenio').AsString;
    ActividadesxOrden.FieldValues['sNumeroOrden'] := tsNumeroOrden.Text;
    ActividadesxOrden.FieldValues['sPaquete'] := '0';
    ActividadesxOrden.FieldValues['sNumeroActividad'] := sActividad;
    ActividadesxOrden.FieldValues['sTipoActividad'] := 'Actividad';
    ActividadesxOrden.FieldValues['sMedida'] := '';//ACTIVIDAD';
    ActividadesxOrden.FieldValues['lCalculo'] := 'Si';
    ActividadesxOrden.FieldValues['sHoraInicio'] := '00:00';
    ActividadesxOrden.FieldValues['sHoraFinal'] := '00:00';
    ActividadesxOrden.FieldValues['dVentaMN'] := 0;
    ActividadesxOrden.FieldValues['dVentaDLL'] := 0;
    ActividadesxOrden.FieldValues['dcostoMN'] := 0;
    ActividadesxOrden.FieldValues['dCostoDLL'] := 0;
    ActividadesxOrden.FieldValues['sWBSAnterior'] := sPaquete;
    ActividadesxOrden.FieldValues['dFechaInicio'] := Date;
    ActividadesxOrden.FieldValues['dFechaFinal'] := Date;
    ActividadesxOrden.FieldValues['dDuracion'] := 1;
    ActividadesxOrden.FieldValues['sWbsContrato'] := '';
    ActividadesxOrden.FieldValues['sWbsAnterior'] := sPaquete;
    ActividadesxOrden.FieldValues['sIdFase'] := '';
    if ActividadesxOrden.RecordCount = 0 then
       ActividadesxOrden.FieldValues['sMedida'] := ''
    else
       ActividadesxOrden.FieldValues['sMedida'] := 'Part.';
    ActividadesxOrden.FieldValues['lGerencial'] := 'No';
    ActividadesxOrden.FieldValues['sIdPlataforma'] := '';
    ActividadesxOrden.FieldValues['sIdPernocta'] := '';
    ActividadesxOrden.FieldValues['dPonderado'] := 0;
    ActividadesxOrden.FieldValues['dCargado'] := 0;
    ActividadesxOrden.FieldValues['dInstalado'] := 0;
    ActividadesxOrden.FieldValues['dExcedente'] := 0;
    ActividadesxOrden.FieldValues['mComentarios'] := '*';
    ActividadesxOrden.FieldValues['lGenerado'] := 'No';
    ActividadesxOrden.FieldValues['lCancelada'] := 'No';
    ActividadesxOrden.FieldValues['sAnexo'] := '';

    if ActividadesxOrden.RecordCount = 0 then
       tiColores.ItemIndex := 4
    else
       tiColores.ItemIndex := 0;

    case tiColores.ItemIndex of
      0: tiColor.ItemIndex := 0;
      1: tiColor.ItemIndex := 1;
      2: tiColor.ItemIndex := 2;
      3: tiColor.ItemIndex := 3;
      4: tiColor.ItemIndex := 4;
      5: tiColor.ItemIndex := 5;
      6: tiColor.ItemIndex := 6;
      7: tiColor.ItemIndex := 7;
      8: tiColor.ItemIndex := 8;
      9: tiColor.ItemIndex := 9;
      10: tiColor.ItemIndex := 10;
      11: tiColor.ItemIndex := 11;
      12: tiColor.ItemIndex := 12;
      13: tiColor.ItemIndex := 13;
      14: tiColor.ItemIndex := 14;
      15: tiColor.ItemIndex := 15;
    end;
    tsNumeroActividad.SetFocus;

    //PopUpNuevoRegistro;
  end;
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmActividades.frmBarra1btnEditClick(Sender: TObject);
begin
  activapop(frmActividades, popupprincipal);
  if ActividadesxOrden.RecordCount > 0 then
  begin
    Insertar1.Enabled := False;
    Editar1.Enabled := False;
    Registrar1.Enabled := True;
    Can1.Enabled := True;
    Eliminar1.Enabled := False;
    Refresh1.Enabled := False;
    grid_actividades.Enabled := false;
    sOpcion := 'Editar';

    Salir1.Enabled := False;
    frmBarra1.btnEditClick(Sender);
    try
      WbsAnt := ActividadesxOrden.FieldValues['sWbs'];
      ActividadAnt := ActividadesxOrden.FieldValues['sNumeroActividad'];
      DescripcionAnt := ActividadesxOrden.FieldValues['mDescripcion'];
      UnidadAnt := ActividadesxOrden.FieldValues['sMedida'];
      CostoAnt := ActividadesxOrden.FieldValues['dCostoMN'];
      VentaAnt := ActividadesxOrden.FieldValues['dVentaMN'];
      sWbsOrig := ActividadesxOrden.FieldValues['sWbs'];
      NivelAnt := ActividadesxOrden.FieldValues['iNivel'];
      ActividadesxOrden.Edit;

      tmDescripcion.Enabled := True;
      tdCostoMN.Enabled := true;
      tsUnidad.Enabled := true;
      tdVentaMN.Enabled := true;
    except
      on e: exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Conceptos/Partidas x Frente de Trabajo', 'Al editar registro', 0);
        frmBarra1.btnCancel.Click;
      end;
    end;
    tdCantidad.SetFocus
  end;
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmActividades.frmBarra1btnPostClick(Sender: TObject);
var
  lEdito, lContinua: Boolean;
  sItemOrdenAnterior,
    Wbs,
    ExtraePaquete,
    sItemOrdenAnterior_p,
    ExtraePaquete_p,
    NewWbs, sMensaje,
    sWbs_paquetes,
    DuracionEstablecida: string;
  iNivel_p, NivelAct: integer;
  Q_Item,
    Q_Paquete,
    Q_PaqueteCopy,
    Q_Actualiza: TZReadonlyQuery;
  nombres, cadenas: TStringList;
  PosReg:TBookMark;
  procedure ActualizaWbsMateriales(wAnterior,wActual,Contrat,folio,turno:String);
  var UptMat:TZQuery;
      MatxActivxFolio:TZReadOnlyQuery;
      OldCur:TCursor;
  begin
    //
    MatxActivxFolio:=TZReadOnlyQuery.Create(nil);
    try
      OldCur := Screen.Cursor;
      Screen.Cursor := crAppStart;
      MatxActivxFolio.Connection := connection.zConnection;
      UptMat := TZQuery.Create(nil);
      UptMat.Connection := connection.zConnection;
      try
//        MatxActivxFolio.Active := False;
//        MatxActivxFolio.SQL.Clear;
//        MatxActivxFolio.SQL.Add('select a.iiddiario,a.snumeroorden,m.* from bitacoradeactividades a '+'inner join bitacorademateriales m on (a.scontrato = m.scontrato and a.iiddiario = m.iiddiario and a.didfecha = m.didfecha) where a.snumeroorden = :snumeroorden and m.sContrato = :scontrato and a.sIdTurno = :sidturno and m.swbs = :swbs');
//        MatxActivxFolio.ParamByName('scontrato').AsString := contrat;
//        MatxActivxFolio.ParamByName('sidturno').AsString := Turno;
//        MatxActivxFolio.ParamByName('snumeroorden').AsString := folio;
//        MatxActivxFolio.ParamByName('swbs').AsString := wAnterior;
//        MatxActivxFolio.Open;
//
//        MatxActivxFolio.First;
//        while not MatxActivxFolio.Eof do
//        begin
//          //enlazar los materiales con la bitacora y por folio verificar el desastre
//          UptMat.Active := False;
//          UptMat.SQL.Clear;
//          UptMat.SQL.ADD( 'Update bitacorademateriales set swbs = :swbsn where scontrato = :contrato and swbs = :wbs and iiddiario = :iiddiario and didfecha = :didfecha');
//          UptMat.ParamByName('contrato').AsString := Contrat;
//          UptMat.ParamByName('iiddiario').asinteger := MatxActivxFolio.FieldByName('iiddiario').asinteger;
//          UptMat.ParamByName('didfecha').AsDateTime := MatxActivxFolio.FieldByName('didfecha').AsDateTime;
//          UptMat.ParamByName('swbsn').AsString := wActual;
//          UptMat.ParamByName('wbs').AsString :=wAnterior;
//
//          UptMat.ExecSQL;
//          MatxActivxFolio.Next;
//        end;

      finally
        UptMat.Free;
      end;
    finally
      Screen.Cursor := OldCur;
      MatxActivxFolio.free;
    end;
  end;


begin
  DuracionEstablecida := edt1.Text;
    {Validaciones de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('Descripcion'); //nombres.Add('Fecha Inio');
  cadenas.Add(tmDescripcion.Text); //cadenas.Add(tdFechaInicio.text);

  nombres.Add('Ponderado');
  cadenas.Add(tdPonderado.text);

  if not validaTexto(nombres, cadenas, 'Concepto/Part.', tsNumeroActividad.Text) then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;

  sMensaje := '';
  {Validacion si se est´cmabiando la actividad}
  if sOpcion = 'Editar' then
    if ActividadAnt <> tsNumeroActividad.Text then
      if MessageDlg('Desea modificar el Número de Partida?, Si modifica afectará contenido de Reportes Diarios.', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
         exit;

  {Verificamos que las cantidades de Frente no excedan a la Cantida de Anexo..}
  lContinua := SumaCantidades;
  if lContinua = False then
    exit;

  Q_Item := TZReadOnlyQuery.Create(self);
  Q_Item.Connection := connection.zConnection;

  Q_Paquete := TZReadOnlyQuery.Create(self);
  Q_Paquete.Connection := connection.zConnection;

  Q_PaqueteCopy := TZReadOnlyQuery.Create(self);
  Q_PaqueteCopy.Connection := connection.zConnection;

  Q_Actualiza := TZReadOnlyQuery.Create(self);
  Q_Actualiza.Connection := connection.zConnection;

  try

    if ActividadesxOrden.FieldValues['sWbs'] = 'A' then
      if ActividadesxOrden.RecordCount > 0 then
        if MessageDlg('Desea Cambiar las fechas de las actividades y tomar la de el paquete principal ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          Connection.zCommand.Active := False;
          Connection.zCommand.Filtered := False;
          Connection.zCommand.SQL.Clear;
          Connection.zCommand.SQL.Add('UPDATE actividadesxorden Set dFechaInicio = :FechaI, dFechaFinal = :FechaF, dDuracion = :duracion Where ' +
            'sContrato = :Contrato And sNumeroOrden = :Orden And sIdConvenio = :Convenio And ' +
            'sWBSAnterior Like :Paquete');
          Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          Connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
          Connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
          Connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
          Connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
          Connection.zCommand.Params.ParamByName('Convenio').value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Connection.zCommand.Params.ParamByName('Paquete').DataType := ftString;
          Connection.zCommand.Params.ParamByName('Paquete').value := Trim(ActividadesxOrden.FieldValues['sWBS']) + '%';
          Connection.zCommand.Params.ParamByName('FechaI').DataType := ftDate;
          Connection.zCommand.Params.ParamByName('FechaI').value := ActividadesxOrden.FieldValues['dFechaInicio'];
          Connection.zCommand.Params.ParamByName('FechaF').DataType := ftDate;
          Connection.zCommand.Params.ParamByName('FechaF').value := ActividadesxOrden.FieldValues['dFechaFinal'];
          Connection.zCommand.Params.ParamByName('duracion').DataType := ftInteger;
          Connection.zCommand.Params.ParamByName('duracion').value := DaysBetween(ActividadesxOrden.FieldValues['dFechaFinal'], ActividadesxOrden.FieldValues['dFechaInicio']) + 1;
          Connection.zCommand.ExecSQL;
          //Actualizar materiales*************************************************
        end;

    ActividadesxOrden.FieldValues['iColor'] := tiColor.Text;
    ActividadesxOrden.FieldByName('mDescripcion').AsString := tmDescripcion.Text;
        //Definimos si es paquete o es una actividad..
    if (ActividadesxOrden.FieldByName('sMedida').IsNull) or (ActividadesxOrden.FieldValues['sMedida'] = '') then
    begin
      ActividadesxOrden.FieldValues['sTipoActividad'] := 'Paquete';
      ActividadesxOrden.FieldValues['dVentaMN'] := 0;
      ActividadesxOrden.FieldValues['dVentaDLL'] := 0;
      ActividadesxOrden.FieldValues['dCostoMN'] := 0;
      ActividadesxOrden.FieldValues['dCostoDLL'] := 0;
      if (ActividadesxOrden.FieldByName('dCantidad').IsNull) or (ActividadesxOrden.FieldValues['dCantidad'] = 0) then
        ActividadesxOrden.FieldValues['dCantidad'] := 1;
      ActividadesxOrden.FieldValues['lGerencial'] := 'Si';
      ActividadesxOrden.FieldValues['sSimbolo'] := '+';
      sPaquete := ActividadesxOrden.Fieldbyname('sWBSAnterior').asstring;
    end
    else
    begin
      ActividadesxOrden.FieldValues['lGerencial'] := 'No';
      ActividadesxOrden.FieldValues['sTipoActividad'] := 'Actividad';
      ActividadesxOrden.FieldValues['sSimbolo'] := '';
      sPaquete := ActividadesxOrden.Fieldbyname('sWBSAnterior').asstring;
    end;

    // ActividadesxOrden.FieldValues['dDuracion'] := DaysBetween(ActividadesxOrden.FieldValues['dFechaFinal'], ActividadesxOrden.FieldValues['dFechaInicio']) + 1;

    if (ActividadesxOrden.FieldValues['dVentaMN'] > 0) then
    begin
      Connection.qryBusca.Active := False;
      Connection.qryBusca.Filtered := False;
      Connection.qryBusca.SQL.Clear;
      Connection.qryBusca.SQL.Add('Select sContrato from actividadesxorden Where sContrato = :contrato and sNumeroOrden = :orden And sIdConvenio = :convenio and sWBSAnterior = :wbs');
      Connection.qryBusca.Params.ParamByName('contrato').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('contrato').Value := global_contrato;
      Connection.qryBusca.Params.ParamByName('orden').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('orden').Value := tsNumeroOrden.Text;
      Connection.qryBusca.Params.ParamByName('convenio').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      Connection.qryBusca.Params.ParamByName('wbs').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('wbs').Value := ActividadesxOrden.FieldValues['sWbs'];
      Connection.qryBusca.Open;

      if Connection.qryBusca.RecordCount > 0 then
      begin
        ActividadesxOrden.FieldValues['sTipoActividad'] := 'Paquete';
        ActividadesxOrden.FieldValues['sMedida'] := '';
        ActividadesxOrden.FieldValues['dVentaMN'] := 0;
        ActividadesxOrden.FieldValues['dVentaDLL'] := 0;
        ActividadesxOrden.FieldValues['dCostoMN'] := 0;
        ActividadesxOrden.FieldValues['dCostoDLL'] := 0;
      end;
    end;

        {Antes de guardar el item buscamos a que paquete le pertenece..}
    Q_Item.Active := False;
    Q_Item.Filtered := False;
    Q_Item.SQL.Clear;
    Q_Item.SQL.Add('select iItemOrden, iNivel from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden and sWbs =:Wbs ');
    Q_Item.Params.ParamByName('Contrato').AsString := global_contrato;
    Q_Item.Params.ParamByName('Convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
    Q_Item.Params.ParamByName('Orden').AsString := tsNumeroOrden.text;
    if Paquetes.FieldValues['sWbs'] <> null then
      Q_Item.Params.ParamByName('Wbs').AsString := Paquetes.FieldValues['sWbs']
    else
      Q_Item.Params.ParamByName('Wbs').AsString := '';
    Q_Item.Open;

    if Q_Item.RecordCount > 0 then
      sItemOrdenAnterior := Q_Item.FieldValues['iItemOrden']
    else
      sItemOrdenAnterior := '';

    if Paquetes.FieldValues['sWbs'] <> null then
      ExtraePaquete := Paquetes.FieldValues['sWbs']
    else
      ExtraePaquete := '';

    if ActividadesxOrden.State = dsEdit then
    begin
        {Ahora hacemos la edicion del registro...}
        if chkItemOrden.Checked then
        begin
            ActividadesxOrden.FieldValues['iItemOrden'] := sItemOrdenAnterior + sFnBuscaItem(zqReprogramacion.FieldByName('sIdConvenio').AsString,ActividadesxOrden.FieldValues['sNumeroActividad'],
            ExtraePaquete,
            sItemOrdenAnterior,
            ActividadesxOrden.FieldValues['sTipoActividad'], tsNumeroOrden.Text, 'actividadesxorden',
            ActividadesxOrden.FieldValues['iNivel'] + 1);
        end;
        sItemOrdenAnterior_p := ActividadesxOrden.FieldValues['iItemOrden'];
        ExtraePaquete_p := ExtraePaquete;

        iNivel_p := ActividadesxOrden.FieldValues['iNivel'] + 1;

        if ActividadesxOrden.FieldValues['sAnexo'] = '' then
          ActividadesxOrden.FieldValues['sWbs'] := ExtraePaquete_p + '.' + ActividadesxOrden.FieldValues['sNumeroActividad']
        else
          ActividadesxOrden.FieldValues['sWbs'] := ExtraePaquete_p + '.' + ActividadesxOrden.FieldValues['sAnexo'] + '.' + ActividadesxOrden.FieldValues['sNumeroActividad'];

        ExtraePaquete_p := ActividadesxOrden.FieldValues['sWbs'];
        ActividadesxOrden.FieldByName('sWbs').OldValue;
        ActividadesxOrden.Post;

        if sOpcion = 'Editar' then
        if ActividadAnt <> tsNumeroActividad.Text then
        begin
            {Actualizamos Tablas con Wbs y Act...}
            {procedure BuscaElimina_datos(sParamTabla, sLlevaContrato, sLlevaFolio, sLlevaWbs, sLLevaAct, sParamContrato, sParamFolio, sParamWbs, sParamAct, sParamNuevoContrato, sParamNuevoFolio, sParamNuevaWbs, sParamNuevaAct : string; accion :string);}
            BuscaElimina_datos( 'actividadesxorden', 'sContrato', 'sNumeroOrden', 'sWbs', 'sNumeroActividad', global_contrato, tsNumeroOrden.KeyValue, WbsAnt, ActividadAnt, '', '', ActividadesxOrden.FieldValues['sWbs'], ActividadesxOrden.FieldValues['sNumeroActividad'], 'actualizar', False);

            {Actualizamos Tablas con Act...}
            {procedure BuscaElimina_datos(sParamTabla, sLlevaContrato, sLlevaFolio, sLlevaWbs, sLLevaAct, sParamContrato, sParamFolio, sParamWbs, sParamAct, sParamNuevoContrato, sParamNuevoFolio, sParamNuevaWbs, sParamNuevaAct : string; accion :string);}
            BuscaElimina_datos( 'actividadesxorden', 'sContrato', 'sNumeroOrden', '', 'sNumeroActividad', global_contrato, tsNumeroOrden.KeyValue, '', ActividadAnt, '', '', '', ActividadesxOrden.FieldValues['sNumeroActividad'], 'actualizar', True);
        end;

        {$region 'Edicion Paquetes'}
         //Ahora consultamos si es un paquete el dato editado..
        if (ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete') and (chkItemOrden.Checked)  then
        begin
          {consultamos los datos dependientes del paquete a editar..}
          Q_Paquete.Active := False;
          Q_Paquete.SQL.Clear;
          Q_Paquete.SQL.Add('select * from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden and sWbs LIKE :Wbs and iNivel >:Nivel order by iItemOrden, sNumeroActividad ');
          Q_Paquete.Params.ParamByName('Contrato').AsString := global_contrato;
          Q_Paquete.Params.ParamByName('Convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Q_Paquete.Params.ParamByName('Orden').AsString := tsNumeroOrden.Text;
          Q_Paquete.Params.ParamByName('Wbs').AsString := WbsAnt + '.%';
          Q_Paquete.Params.ParamByName('Nivel').AsInteger := NivelAnt;
          Q_Paquete.Open;

          {Esta es una copia de los elementos..}
          Q_PaqueteCopy.Active := False;
          Q_PaqueteCopy.SQL.Clear;
          Q_PaqueteCopy.SQL.Add('select * from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden and sWbs LIKE :Wbs and iNivel >:Nivel order by iItemOrden, sNumeroActividad ');
          Q_PaqueteCopy.Params.ParamByName('Contrato').AsString := global_contrato;
          Q_PaqueteCopy.Params.ParamByName('Convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Q_PaqueteCopy.Params.ParamByName('Orden').AsString := tsNumeroOrden.Text;
          Q_PaqueteCopy.Params.ParamByName('Wbs').AsString := WbsAnt + '.%';
          Q_PaqueteCopy.Params.ParamByName('Nivel').AsInteger := NivelAnt;
          Q_PaqueteCopy.Open;

          NivelAct := NivelAnt - (iNivel_p - 1);

          while not Q_Paquete.Eof do
          begin
                      {Antes de Hacer la actualizacion, se consultan por niveles para crear Wbs..}
            sWbs_paquetes := ExtraePaquete_p;
            Q_PaqueteCopy.First;
            while not Q_PaqueteCopy.Eof do
            begin
                          {Definimos las Wbs..}
              if WbsAnt = Q_PaqueteCopy.FieldValues['sWbsAnterior'] then
                sWbs_paquetes := ExtraePaquete_p;

              if Q_PaqueteCopy.FieldValues['sTipoActividad'] = 'Actividad' then
              begin
                if Q_PaqueteCopy.FieldValues['sAnexo'] = '' then
                  NewWbs := sWbs_paquetes + '.' + Q_PaqueteCopy.FieldValues['sNumeroActividad']
                else
                  NewWbs := sWbs_paquetes + '.' + Q_PaqueteCopy.FieldValues['sAnexo'] + '.' + Q_PaqueteCopy.FieldValues['sNumeroActividad']
              end
              else
                NewWbs := sWbs_paquetes + '.' + Q_PaqueteCopy.FieldValues['sNumeroActividad'];

              if Q_Paquete.FieldValues['sWbs'] = Q_PaqueteCopy.FieldValues['sWbs'] then
              begin
                Q_Actualiza.Active := False;
                Q_Actualiza.SQL.Clear;
                Q_Actualiza.SQL.Add('Update actividadesxorden set iItemOrden =:iOrden, sWbs =:Wbs, sWbsAnterior =:WbsAnt, iNivel =:Nivel where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden and sWbs =:WbsOld ');
                Q_Actualiza.Params.ParamByName('Contrato').AsString := global_contrato;
                Q_Actualiza.Params.ParamByName('Convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
                Q_Actualiza.Params.ParamByName('Orden').AsString := tsNumeroOrden.Text;
                Q_Actualiza.Params.ParamByName('Wbs').AsString := NewWbs;
                Q_Actualiza.Params.ParamByName('WbsOld').AsString := Q_Paquete.FieldValues['sWbs'];
                Q_Actualiza.Params.ParamByName('WbsAnt').AsString := sWbs_paquetes;
                if NivelAct > 0 then
                  Q_Actualiza.Params.ParamByName('Nivel').AsInteger := Q_PaqueteCopy.FieldValues['iNivel'] - NivelAct
                else
                  if NivelAct < 0 then
                    Q_Actualiza.Params.ParamByName('Nivel').AsInteger := Q_PaqueteCopy.FieldValues['iNivel'] + NivelAct
                  else
                    Q_Actualiza.Params.ParamByName('Nivel').AsInteger := Q_PaqueteCopy.FieldValues['iNivel'];

                  Q_Actualiza.Params.ParamByName('iOrden').AsString := sItemOrdenAnterior_P + sFnBuscaItem(zqReprogramacion.FieldByName('sIdConvenio').AsString,Q_Paquete.FieldValues['sNumeroActividad'],
                  ExtraePaquete_p,
                  sItemOrdenAnterior_p,
                  Q_Paquete.FieldValues['sTipoActividad'], tsNumeroOrden.Text, 'actividadesxanexo',
                  Q_Paquete.FieldValues['iNivel']);
                  Q_Actualiza.ExecSQL;

                Q_PaqueteCopy.Last;
              end;

              if Q_PaqueteCopy.FieldValues['sTipoActividad'] = 'Paquete' then
                sWbs_paquetes := NewWbs;

              Q_PaqueteCopy.Next;
            end;
            Q_Paquete.Next;
          end;
        end;
        {$endregion}
        lEdito := True;
    end
    else
    begin
      if tsPaquete.KeyValue = Null then
      begin
          ActividadesxOrden.FieldValues['iNivel'] := 0;
          ActividadesxOrden.FieldValues['iItemOrden'] := '';
      end
      else
      begin
          ActividadesxOrden.FieldValues['iNivel'] := Paquetes.FieldValues['iNivel'] + 1;
          ActividadesxOrden.FieldValues['iItemOrden'] := Paquetes.FieldValues['iItemOrden'];
      end;

      ActividadesxOrden.FieldValues['iItemOrden'] := sItemOrdenAnterior + sFnBuscaItem(zqReprogramacion.FieldByName('sIdConvenio').AsString, ActividadesxOrden.FieldValues['sNumeroActividad'],
        ExtraePaquete,
        sItemOrdenAnterior,
        ActividadesxOrden.FieldValues['sTipoActividad'], tsNumeroorden.Text, 'actividadesxorden',
        ActividadesxOrden.FieldValues['iNivel']);
      lEdito := False;
           {Se salvan los datos..}
      ActividadesxOrden.Post;
      posReg:=  ActividadesxOrden.GetBookmark;

    end;

    if ActividadesxOrden.State = dsEdit then
      lEdito := True
    else
    begin
      lEdito := False;
      Kardex('Conceptos Generales', 'Crea   Registro Programa de Trabajo', tsNumeroActividad.Text, Actividadesxorden.FieldValues['sTipoActividad'], tsNumeroOrden.Text, '', '','Tarifa Diaria','Actividades x folio');
    end;

    {$Region 'Calcula Fechas'}
    if lEdito then
    begin
      if (ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete') and (Connection.Configuracion.FieldValues['lCalculaFecha'] = 'Si') then
      begin
        try
          Connection.zCommand.Active := False;
          Connection.zCommand.Filtered := False;
          Connection.zCommand.SQL.Clear;
          Connection.zCommand.SQL.Add('UPDATE actividadesxorden Set dFechaInicio = :FechaI, dFechaFinal = :FechaF, dDuracion = :duracion Where ' +
            'sContrato = :Contrato And sNumeroOrden = :Orden And sIdConvenio = :Convenio And ' +
            'concat(sWBSAnterior , ".") Like :WbsAnterior');
          connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
          connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
          connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
          connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
          connection.zCommand.Params.ParamByName('Convenio').value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          connection.zCommand.Params.ParamByName('WbsAnterior').DataType := ftString;
          connection.zCommand.Params.ParamByName('WbsAnterior').value := Trim(ActividadesxOrden.FieldValues['sWBS']) + '.%';
          connection.zCommand.Params.ParamByName('FechaI').DataType := ftDate;
          connection.zCommand.Params.ParamByName('FechaI').value := ActividadesxOrden.FieldValues['dFechaInicio'];
          connection.zCommand.Params.ParamByName('FechaF').DataType := ftDate;
          connection.zCommand.Params.ParamByName('FechaF').value := ActividadesxOrden.FieldValues['dFechaFinal'];
          connection.zCommand.Params.ParamByName('duracion').DataType := ftInteger;
          connection.zCommand.Params.ParamByName('duracion').value := DaysBetween(ActividadesxOrden.FieldValues['dFechaFinal'], ActividadesxOrden.FieldValues['dFechaInicio']) + 1;
          connection.zCommand.ExecSQL;
        except
          MessageDlg('Ocurrio un error al actualizar el registro', mtWarning, [mbOk], 0);
        end;

        tsNumeroActividad.Enabled := True;
        tmDescripcion.Enabled := True;
        tdCostoMN.Enabled := True;
        tsUnidad.Enabled := True;
        tdVentaMN.Enabled := True;

      end;
    end;
    {$endregion}

    if (ActividadesxOrden.FieldByName('sTipoActividad').AsString = 'Paquete') then begin
      ActividadesxOrden.Edit;
      ActividadesxOrden.FieldByName('sDuracionHoras').AsString := DuracionEstablecida;
      ActividadesxOrden.Post;
    end;


    ActualizaDuracion;
    PosReg := ActividadesxOrden.GetBookmark;
    ActividadesxOrden.Refresh;

    try
      ActividadesxOrden.GotoBookmark(PosReg);
    except

     ActividadesxOrden.FreeBookmark(PosReg);
    end;


    Paquetes.Active := False;
    Paquetes.Open;

    tsPaquete.Enabled := True;
    Insertar1.Enabled := True;
    Editar1.Enabled := True;
    Registrar1.Enabled := False;
    Can1.Enabled := False;
    Eliminar1.Enabled := True;
    Refresh1.Enabled := True;
    Salir1.Enabled := True;

    frmBarra1.btnCancel.Click;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Conceptos/Partidas x Frente de Trabajo', 'Al salvar registro', 0);
      frmBarra1.btnCancel.Click;
    end;
  end;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmActividades.frmBarra1btnCancelClick(Sender: TObject);
begin
  desactivapop(popupprincipal);
  ActividadesxOrden.Cancel;
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  frmBarra1.btnCancelClick(Sender);
  tsPaquete.Enabled := True;

  tsNumeroActividad.Enabled := True;
  tmDescripcion.Enabled := True;
  tdCostoMN.Enabled := True;
  tsUnidad.Enabled := True;
  tdVentaMN.Enabled := True;
  grid_actividades.Enabled := True;

  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmActividades.frmBarra1btnDeleteClick(Sender: TObject);
var
  Actividad, TipoAct, cadena, tabla: string;
  total, i: integer;
  dPonderado: double;
begin
  if ActividadesxOrden.RecordCount > 0 then
    if MessageDlg('Desea eliminar el # Conepto / partida ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
          {Buscamos si la partida o paquete ya fue reportada..}
      Actividad := ActividadesxOrden.FieldValues['sNumeroActividad'];
      TipoAct := ActividadesxOrden.FieldValues['sTipoActividad'];
      dPonderado := ActividadesxOrden.FieldValues['dPonderado'];
      try
        if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Actividad' then
        begin
                   {Primero la partida..}
          cadena := AntesEliminarFrente(ActividadesxOrden.FieldValues['sNumeroActividad'], ActividadesxOrden.FieldValues['sWbs'], tsNumeroOrden.Text, TipoAct, zqReprogramacion.FieldByName('sIdConvenio').AsString);
          if cadena <> '' then
          begin
            MessageDlg('No se puede Eliminar!. La Partida ' + ActividadesxOrden.FieldValues['sNumeroActividad'] + ' se encuentra Registrada en:' + #13 + cadena, mtWarning, [mbOk], 0);
            exit;
          end;
        end
        else
        begin
                   {Ahora los paquetes..}
          cadena := AntesEliminarFrente(ActividadesxOrden.FieldValues['sNumeroActividad'], ActividadesxOrden.FieldValues['sWbs'] + '.%', tsNumeroOrden.Text, TipoAct, zqReprogramacion.FieldByName('sIdConvenio').AsString);
          if cadena <> '' then
          begin
            MessageDlg('No se puede Eliminar!. El Paquete ' + ActividadesxOrden.FieldValues['sNumeroActividad'] + ' contine Partidas registradas en: ' + #13 + cadena, mtWarning, [mbOk], 0);
            exit;
          end;
        end;

              //Se eliminan las distribuciones de los frentes,,
        DistribucionesFrente(tsNumeroOrden.Text, ActividadesxOrden.FieldValues['sWbs'], ActividadesxOrden.FieldValues['sTipoActividad'], ActividadesxOrden.FieldValues['iNivel']);

        if TipoAct = 'Actividad' then
        begin
          SavePlace := grid_actividades.DataSource.DataSet.GetBookmark;
          ActividadesxOrden.Delete;
        end
        else
        begin
          Connection.zCommand.Active := False;
          Connection.zCommand.Filtered := False;
          Connection.zCommand.SQL.Clear;
          Connection.zCommand.SQL.Add('delete from actividadesxorden Where sContrato = :contrato and sIdConvenio = :convenio and sNumeroOrden = :orden and sWbs LIKE :wbs and iNivel >=:Nivel ');
          connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          connection.zCommand.Params.ParamByName('Contrato').Value    := global_contrato;
          connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
          connection.zCommand.Params.ParamByName('Convenio').Value    := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          connection.zCommand.Params.ParamByName('Orden').DataType    := ftString;
          connection.zCommand.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text;
          connection.zCommand.Params.ParamByName('wbs').DataType      := ftString;
          connection.zCommand.Params.ParamByName('wbs').Value         := ActividadesxOrden.FieldValues['sWbs'] + '%';
          connection.zCommand.Params.ParamByName('Nivel').DataType    := ftInteger;
          connection.zCommand.Params.ParamByName('Nivel').Value       := ActividadesxOrden.FieldValues['iNivel'];
          connection.zCommand.ExecSQL;

          SavePlace := grid_actividades.DataSource.DataSet.GetBookmark;
          ActividadesxOrden.Delete;
        end;

        //Funcion elimina avances
        EliminaAvances(tsNumeroOrden.Text,zqReprogramacion.FieldByName('sIdConvenio').AsString );

        ActividadesxOrden.Refresh;
        try
          grid_actividades.DataSource.DataSet.GotoBookmark(SavePlace);
        except
          grid_actividades.DataSource.DataSet.FreeBookmark(SavePlace);
        end;
        Kardex('Conceptos Generales', 'Borra Registro Programa de Trabajo', Actividad, TipoAct, tsNumeroOrden.Text, '', '','Tarifa Diaria','ActividadesxFolio');
      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Conceptos/Partidas x Frente de Trabajo', 'Al borrar registro', 0);
        end;
      end;
      Paquetes.Active := False;
      Paquetes.Open;
    end;
end;

procedure TfrmActividades.frmBarra1btnRefreshClick(Sender: TObject);
var
  sSelect: string;
begin
  OrdenesdeTrabajo.Active := False;
  OrdenesdeTrabajo.Open;

  Paquetes.Active := False;
  Paquetes.Open;          
  connection.configuracion.Refresh;

  ActividadesxOrden.Filtered := false;
  ActividadesxOrden.Refresh;
  sFiltro := '';
  sPaquete := '';
end;

procedure TfrmActividades.frxReport1GetValue(const VarName: string;
  var Value: Variant);
var
  QryConsultaAvancesAcumulados:TzReadOnlyQuery;
begin
  if CompareText(VarName, 'AVANCE') = 0 then
  begin
    Value := '0.00 %';
    QryConsultaAvancesAcumulados:=TzReadOnlyQuery.create(nil);
    try
      QryConsultaAvancesAcumulados.Connection:=connection.zConnection;
      QryConsultaAvancesAcumulados.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', IncDay(now))+'", :Orden, :Folio), 4) AS dAvanceAcumulado;';
      QryConsultaAvancesAcumulados.ParamByName('Orden').AsString := Global_Contrato;
      QryConsultaAvancesAcumulados.ParamByName('Folio').AsString := tsNumeroOrden.KeyValue;
      QryConsultaAvancesAcumulados.Open;
      if QryConsultaAvancesAcumulados.RecordCount>0 then
        Value:=FormatFloat('0.00',QryConsultaAvancesAcumulados.FieldByName('dAvanceAcumulado').asfloat) + ' %';
    finally
      QryConsultaAvancesAcumulados.Destroy;
    end;
  end;
end;

procedure TfrmActividades.GenerarNotadeCampo1Click(Sender: TObject);
begin

  //////////////////////////////////////////////
 //   NotaCampoExcel
  TdConfiguracion(Global_Contrato,tsNumeroOrden.KeyValue,frxReport1);
  NotaCampo(Global_Contrato,tsNumeroOrden.KeyValue,frxReport1);
  frxReport1.LoadFromFile(global_files + global_Mireporte + '_TDNotaCampo.fr3') ;
  frxReport1.ShowReport();
  ReportePDF_ClearDataset(frxReport1);
end;

procedure TfrmActividades.frmBarra1btnExitClick(Sender: TObject);
begin
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  frmBarra1.btnExitClick(Sender);
  Close
end;

procedure TfrmActividades.FormClick(Sender: TObject);
begin
  if Assigned(frmGraficaGerencial) then
    if frmGraficaGerencial.Active = True then
      frmGraficaGerencial.Close;
end;

procedure TfrmActividades.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  utgrid.Destroy;
  ActividadesxOrden.Cancel;
  action := cafree;
  BotonPermiso.free;
end;

procedure TfrmActividades.FormCreate(Sender: TObject);
begin
  Paq := TstringList.Create;
end;

procedure TfrmActividades.tdDuracionExit(Sender: TObject);
begin
//  if (ActividadesxOrden.State = dsInsert) or (ActividadesxOrden.State = dsEdit) then
//    tdFechaFinal.Date := tdFechaInicio.Date + (ActividadesxOrden.FieldValues['dDuracion'] - 1);
//  tdDuracion.Color := global_color_salida
end;

procedure TfrmActividades.Insertar1Click(Sender: TObject);
begin
  frmBarra1.btnAdd.Click
end;

procedure TfrmActividades.Editar1Click(Sender: TObject);
begin
  frmBarra1.btnEdit.Click
end;

procedure TfrmActividades.Registrar1Click(Sender: TObject);
begin
  frmBarra1.btnPost.Click
end;

procedure TfrmActividades.ReprogramarFolio1Click(Sender: TObject);
var
   sConvenio : string;
begin
    if actividadesxorden.RecordCount = 0 then
    begin
        Application.CreateForm(TFrmPopUpReprogramacion, FrmPopUpReprogramacion);
        FrmPopUpReprogramacion.Left := trunc((Screen.Width)/2)-trunc((FrmPopUpReprogramacion.Width)/2);
        FrmPopUpReprogramacion.Top := trunc((screen.Height)/2)-trunc((FrmPopUpReprogramacion.Height)/2);
        if FrmPopUpReprogramacion.ShowModal = mrCancel then
        begin
            FrmPopUpReprogramacion.Free;
            exit;
        end
        else
        begin
            sConvenio := FrmPopUpReprogramacion.comboConvenios.Text;
            FrmPopUpReprogramacion.Free;
        end;

        zqCopiaReprogramacion.SQL.Clear;
        zqCopiaReprogramacion.SQL.Text := ObtenerSentencia( 'actividadesxorden', 'sql_reprograma_folio', ftInsert );
        zqCopiaReprogramacion.ParamByName('Orden').AsString         := global_contrato;
        zqCopiaReprogramacion.ParamByName('new_convenio').AsString  := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        zqCopiaReprogramacion.ParamByName('convenio_from').AsString := sConvenio;
        zqCopiaReprogramacion.ParamByName('folio').AsString         := tsNumeroOrden.Text;
        zqCopiaReprogramacion.ExecSQL;

        tsReprogramacion.OnExit(sender);
    end
    else
       messageDLG('No se puede copiar el Programa de una Reprogramación Anterior, existen Datos!', mtWarning, [mbOk], 0);
end;

procedure TfrmActividades.Can1Click(Sender: TObject);
begin
  frmBarra1.btnCancel.Click
end;

procedure TfrmActividades.Eliminar1Click(Sender: TObject);
begin
  frmBarra1.btnDelete.Click
end;

procedure TfrmActividades.ExportaaExcel1Click(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// GENERA PROGRAMA DE TRABAJO //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena: string;
      fs: tStream;
      Alto: Extended;
      Ren, nivel: integer;
      Progreso, TotalProgreso: real;
    begin
      Ren := 2;
  // Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 100;
//  if rAnexoC.Checked then
//  begin
      Excel.Columns['A:A'].ColumnWidth := 15;
      Excel.Columns['B:B'].ColumnWidth := 15;
      Excel.Columns['C:C'].ColumnWidth := 8;
      Excel.Columns['D:D'].ColumnWidth := 10;
      Excel.Columns['E:E'].ColumnWidth := 40;
      Excel.Columns['F:J'].ColumnWidth := 12;
      Excel.Columns['K:K'].ColumnWidth := 15;
      Excel.Columns['L:L'].ColumnWidth := 15;

      Hoja.Range['A1:A1'].Select;
      Excel.Selection.Value := 'Contrato';
      FormatoEncabezado;
      Hoja.Range['B1:B1'].Select;
      Excel.Selection.Value := 'Frente';
      FormatoEncabezado;
      Hoja.Range['C1:C1'].Select;
      Excel.Selection.Value := 'Nivel';
      FormatoEncabezado;
      Hoja.Range['D1:D1'].Select;
      Excel.Selection.NumberFormat := '@';
      Excel.Selection.Value := 'Actividad';
      FormatoEncabezado;
      Hoja.Range['E1:E1'].Select;
      Excel.Selection.Value := 'Descripcion';
      FormatoEncabezado;
      Hoja.Range['F1:F1'].Select;
      Excel.Selection.Value := 'Medida';
      FormatoEncabezado;
      Hoja.Range['G1:G1'].Select;
      Excel.Selection.NumberFormat := '@';
      Excel.Selection.Value := 'Cantidad';
      FormatoEncabezado;
      Hoja.Range['H1:H1'].Select;
      Excel.Selection.NumberFormat := '@';
      Excel.Selection.Value := 'Ponderado';
      FormatoEncabezado;
      Hoja.Range['I1:I1'].Select;
      Excel.Selection.NumberFormat := '@';
      Excel.Selection.Value := 'Fecha_Inicio';
      FormatoEncabezado;
      Excel.Selection.NumberFormat := '@';
      Hoja.Range['J1:J1'].Select;
      Excel.Selection.Value := 'Fecha_Final';
      FormatoEncabezado;
      Hoja.Range['K1:K1'].Select;
      Excel.Selection.Value := 'Tipo(PU,ADM)';
      FormatoEncabezado;
      Hoja.Range['L1:L1'].Select;
      Excel.Selection.Value := 'Id_Anexo';
      FormatoEncabezado;
      Hoja.Range['M1:M1'].Select;
      Excel.Selection.Value := 'Extraordinaria(Si/No)';
      FormatoEncabezado;

      connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select * from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden order by iitemorden');
      connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then
      begin
        while not connection.QryBusca.Eof do
        begin
                {Movimiento de la Barra..}
          Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
          TotalProgreso := TotalProgreso + Progreso;
          BarraEstado.Position := Trunc(TotalProgreso);

                {Escritura de Datos en el Archvio de Excel..}
          Hoja.Cells[Ren, 1].Select;
          Excel.Selection.Value := global_contrato;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 11;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Hoja.Cells[Ren, 2].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sNumeroOrden'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 3].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['iNivel'];

          Hoja.Cells[Ren, 4].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sNumeroActividad'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 5].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];
          Alto := Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight;
          Hoja.Cells[Ren, 5].Value := '';

          if Alto > 15 then
            Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := Alto
          else
            Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := 15;

          Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];

          Hoja.Cells[Ren, 6].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sMedida'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 7].Select;
          Excel.Selection.NumberFormat := '@';
          Excel.Selection.Value := connection.QryBusca.FieldValues['dCantidad'];
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 8].Select;
          Excel.Selection.NumberFormat := '@';
          Excel.Selection.Value := connection.QryBusca.FieldValues['dPonderado'];
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 9].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['dFechaInicio'];

          Hoja.Cells[Ren, 10].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['dFechaFinal'];

          Hoja.Cells[Ren, 11].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sTipoAnexo'];

          Hoja.Cells[Ren, 12].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sAnexo'];

          Hoja.Cells[Ren, 13].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['lExtraordinario'];

                {Colores de los paquetes..}
          if connection.QryBusca.FieldValues['sTipoActividad'] = 'Paquete' then
          begin
            nivel := connection.QryBusca.FieldValues['iNivel'];
            Hoja.Range['A' + IntToStr(Ren) + ':M' + IntToStr(Ren)].Select;
            if colores[nivel + 1, 2] = 1 then
              Excel.Selection.Font.Color := clWhite;
            Excel.Selection.Font.Bold := True;
            Excel.Selection.Interior.ColorIndex := colores[nivel + 1, 1];
          end;

          connection.QryBusca.Next;
          Inc(Ren);
        end;
      end;
      Hoja.Cells[2, 2].Select;


      Hoja.Range['A1:M1'].Select;

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
      try
        Hoja.Name := 'PROGRAMA DE TRABAJO ' + tsNumeroOrden.Text;
      except
        Hoja.Name := 'PROGRAMA DE TRABAJO ';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Programa de Trabajo:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
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

  if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    Excel.Visible := True;
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := True;
  end
  else
  begin
    Excel.Visible := True;
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := False;
  end;

  PanelProgress.Visible := True;
  Label15.Refresh;
  Label16.Refresh;
  Label17.Refresh;
  BarraEstado.Position := 0;

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
  begin
        // Grabar el archivo de excel con el nombre dado
    Excel.Visible := True;
    Excel.DisplayAlerts := True;
    Excel.ScreenUpdating := True;
    PanelProgress.Visible := False;
    messageDlg('El Archivo se generó Correctamente!', mtInformation, [mbOk], 0);
  end;

  Excel := '';

  if CadError <> '' then
    showmessage(CadError);
end;

procedure TfrmActividades.ExportaVolumenesExcel1Click(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// GENERACION DE PROGRAMA DE TRABAJO CON VOLUMENES //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena: string;
      fs: tStream;
      Alto: Extended;
      MiFechaI, MiFechaF, MiFecha: tDate;
      Ren, nivel, i, total: integer;
      Q_Partidas: TZReadOnlyQuery;
      dVolumen: double;
      Progreso, TotalProgreso: real;
    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      Ren := 2;
    // Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 100;

      Excel.Columns['A:A'].ColumnWidth := 20;
      Excel.Columns['B:B'].ColumnWidth := 18;
      Excel.Columns['C:C'].ColumnWidth := 0.58;
      Excel.Columns['D:D'].ColumnWidth := 10;
      Excel.Columns['E:E'].ColumnWidth := 40;
      Excel.Columns['F:G'].ColumnWidth := 12;
      Excel.Columns['H:J'].ColumnWidth := 0.58;

      // Colocar los encabezados de la plantilla...
      Hoja.Range['A1:A1'].Select;
      Excel.Selection.Value := 'Contrato';
      FormatoEncabezado;
      Hoja.Range['B1:B1'].Select;
      Excel.Selection.Value := 'Frente';
      FormatoEncabezado;
      Hoja.Range['C1:C1'].Select;
      Excel.Selection.Value := 'Nivel';
      FormatoEncabezado;
      Hoja.Range['D1:D1'].Select;
      Excel.Selection.Value := 'Actividad';
      FormatoEncabezado;
      Hoja.Range['E1:E1'].Select;
      Excel.Selection.Value := 'Descripcion';
      FormatoEncabezado;
      Hoja.Range['F1:F1'].Select;
      Excel.Selection.Value := 'Medida';
      FormatoEncabezado;
      Hoja.Range['G1:G1'].Select;
      Excel.Selection.Value := 'Cantidad';
      FormatoEncabezado;
      Hoja.Range['H1:H1'].Select;
      Excel.Selection.Value := 'Ponderado';
      FormatoEncabezado;
      Hoja.Range['I1:I1'].Select;
      Excel.Selection.Value := 'Fecha I.';
      FormatoEncabezado;
      Hoja.Range['J1:J1'].Select;
      Excel.Selection.Value := 'Fecha F.';
      FormatoEncabezado;

      //Consultamos las fechas del convenio modificatorio para impresion de las cantidades reportadas superiores al programa de trabajo.
      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Clear;
      connection.QryBusca2.SQL.Add('select max(dIdFecha) as dFechaFinal from reportediario where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden ');
      connection.QryBusca2.ParamByName('contrato').AsString := global_contrato;
      connection.QryBusca2.ParamByName('Convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      connection.QryBusca2.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
      connection.QryBusca2.Open;

      connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select * from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden order by iItemOrden');
      connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then
      begin
        MiFecha := connection.QryBusca.FieldByName('dFechaInicio').AsDateTime;
        MiFechaI := connection.QryBusca.FieldByName('dFechaInicio').AsDateTime;
        MiFechaF := connection.QryBusca2.FieldByName('dFechaFinal').AsDateTime;
        for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
        begin
          Hoja.Cells[Ren - 1, 10 + i].Select;
               {Formato de las fechas archivo Excel,, 24/07/2011..}
          Excel.Selection.NumberFormat := '@';
          Excel.Selection.Value := DateToStr(MiFecha);
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Interior.ColorIndex := 49;
          MiFecha := IncDay(MiFecha);
        end;
        total := i;

        Hoja.Cells[Ren - 1, 10 + i].Select;
        Excel.Selection.Value := 'Total';
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment := xlCenter;
        Excel.Selection.Font.Color := clWhite;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Interior.ColorIndex := 3;

        connection.QryBusca.First;
        while not connection.QryBusca.Eof do
        begin
                {Movimiento de la Barra..}
          Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
          TotalProgreso := TotalProgreso + Progreso;
          BarraEstado.Position := Trunc(TotalProgreso);

                {Escritura de Datos en el Archvio de Excel..}
          Hoja.Cells[Ren, 1].Select;
          Excel.Selection.Value := global_contrato;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Size := 11;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Hoja.Cells[Ren, 2].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sNumeroOrden'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 3].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['iNivel'];

          Hoja.Cells[Ren, 4].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sNumeroActividad'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 5].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];
          Alto := Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight;
          Hoja.Cells[Ren, 5].Value := '';

          if Alto > 15 then
            Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := Alto
          else
            Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := 15;

          Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];

          Hoja.Cells[Ren, 6].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sMedida'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 7].Select;
          Excel.Selection.NumberFormat := '@';
          Excel.Selection.Value := connection.QryBusca.FieldValues['dCantidad'];
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 8].Select;
          Excel.Selection.NumberFormat := '@';
          Excel.Selection.Value := connection.QryBusca.FieldValues['dPonderado'];
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;

          Hoja.Cells[Ren, 9].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['dFechaInicio'];

          Hoja.Cells[Ren, 10].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['dFechaFinal'];

                {Colores de los paquetes..}
          if connection.QryBusca.FieldValues['sTipoActividad'] = 'Paquete' then
          begin
            nivel := connection.QryBusca.FieldValues['iNivel'];
            Hoja.Range['A' + IntToStr(Ren) + ':J' + IntToStr(Ren)].Select;
            if colores[nivel + 1, 2] = 1 then
              Excel.Selection.Font.Color := clWhite;
            Excel.Selection.Font.Bold := True;
            Excel.Selection.Interior.ColorIndex := colores[nivel + 1, 1];
          end
          else
          begin
            MiFecha := MiFechaI;
                    {Consultamos si la partida esta reprotada..}
            Q_Partidas.Active := False;
            Q_Partidas.SQL.Clear;
            Q_Partidas.SQL.Add('Select b.sWbs,b.sNumeroActividad, sum(a.dCantidad) as dCantidad, a.dIdFecha, b.dCantidad as dVolumen ' +
              'From actividadesxorden b ' +
              'left JOIN bitacoradeactividades a ' +
              'ON (a.sContrato=b.sContrato And a.sWbs=b.sWbs And a.dIdFecha <=:Final and b.sNumeroOrden=a.sNumeroOrden) ' +
              'left JOIN tiposdemovimiento t ' +
              'ON (b.sContrato=t.sContrato And a.sIdTipoMovimiento=t.sIdTipoMovimiento And t.sClasificacion="Tiempo en Operacion") ' +
              'Where b.sContrato=:Contrato And b.sIdConvenio=:Convenio And b.sNumeroOrden =:Orden and a.sWbs =:Wbs ' +
              'Group By b.sWbs,a.dIdFecha Order By b.sNumeroActividad,b.iItemOrden,a.dIdFecha');
            Q_Partidas.Params.ParamByName('Contrato').DataType := ftString;
            Q_Partidas.Params.ParamByName('Contrato').Value := global_contrato;
            Q_Partidas.Params.ParamByName('Convenio').DataType := ftString;
            Q_Partidas.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            Q_Partidas.Params.ParamByName('Orden').DataType := ftString;
            Q_Partidas.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            Q_Partidas.Params.ParamByName('Final').DataType := ftDate;
            Q_Partidas.Params.ParamByName('Final').Value := MiFechaF;
            Q_Partidas.Params.ParamByName('Wbs').DataType := ftString;
            Q_Partidas.Params.ParamByName('Wbs').Value := connection.QryBusca.FieldByName('sWbs').AsString;
            Q_Partidas.Open;

            if Q_Partidas.RecordCount > 0 then
            begin
              dVolumen := 0;
              for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
              begin
                if MiFecha = Q_Partidas.FieldByName('dIdFecha').AsDateTime then
                begin
                  Hoja.Cells[Ren, 10 + i].Select;
                  Excel.Selection.Value := Q_Partidas.FieldByName('dCantidad').AsFloat;
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment := xlCenter;
                  Excel.Selection.Font.Bold := False;
                  Excel.Selection.Interior.ColorIndex := 41;
                  dVolumen := dVolumen + Q_Partidas.FieldByName('dCantidad').AsFloat;
                  Q_Partidas.Next;
                end;
                MiFecha := IncDay(MiFecha);
              end;

              Hoja.Cells[Ren, 10 + i].Select;
              Excel.Selection.Value := dVolumen;
              Excel.Selection.HorizontalAlignment := xlRight;
              Excel.Selection.VerticalAlignment := xlCenter;
              Excel.Selection.Font.Bold := True;
              if dVolumen = Q_Partidas.FieldByName('dVolumen').AsFloat then
                Excel.Selection.Font.Color := clBlue;
              if dVolumen > Q_Partidas.FieldByName('dVolumen').AsFloat then
                Excel.Selection.Font.Color := clRed
            end;
          end;
          connection.QryBusca.Next;
          Inc(Ren);
        end;
      end;
   {fORMATO DE LAS CELDAS..}
      Hoja.Range['K2:' + Columnas[total + 10] + IntToStr(Ren - 1)].Select;
      Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
      Excel.Selection.Borders[xlEdgeLeft].Weight := xlThin;
      Excel.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
      Excel.Selection.Borders[xlEdgeTop].Weight := xlThin;
      Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
      Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;
      Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
      Excel.Selection.Borders[xlEdgeRight].Weight := xlThin;
      Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous;
      Excel.Selection.Borders[xlInsideVertical].Weight := xlThin;
      Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous;
      Excel.Selection.Borders[xlInsideHorizontal].Weight := xlThin;

      Hoja.Range['A1:N1'].Select;
   // Formato general de encabezado de datos..
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment := xlCenter;
      Excel.Selection.Font.Color := clWhite;
      Excel.Selection.Interior.ColorIndex := 49;
      Excel.Selection.Interior.Pattern := xlSolid;

      Hoja.Cells[2, 2].Select;
    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'VOLUMENES REPORTADOS ' + tsNumeroOrden.Text;
      except
        Hoja.Name := 'VOLUMENES REPORTADOS ';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Programa de Trabajo:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
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
    showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
    Exit;
  end;

  if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    Excel.Visible := True;
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := True;
  end
  else
  begin
    Excel.Visible := True;
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := False;
  end;

  PanelProgress.Visible := True;
  Label15.Refresh;
  Label16.Refresh;
  Label17.Refresh;
  BarraEstado.Position := 0;

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
  begin
        // Grabar el archivo de excel con el nombre dado
    Excel.Visible := True;
    Excel.DisplayAlerts := True;
    Excel.ScreenUpdating := True;
    PanelProgress.Visible := False;
    messageDlg('El Archivo se generó Correctamente!', mtInformation, [mbOk], 0);
  end;

  Excel := '';

  if CadError <> '' then
    showmessage(CadError);
end;

procedure TfrmActividades.Refresh1Click(Sender: TObject);
begin
  frmBarra1.btnRefresh.Click
end;

procedure TfrmActividades.Salir1Click(Sender: TObject);
begin
  frmBarra1.btnExit.Click
end;

procedure TfrmActividades.tsNumeroOrdenExit(Sender: TObject);
begin
  frmBarra1.btnCancel.Click;
  sNumeroOrden := tsNumeroOrden.Text;
  IsOpen:=false;
  ConsultaFolios;
  ConsultaReprogramacion;
  tsReprogramacion.KeyValue := zqReprogramacion.FieldByName('sIdConvenio').AsString;
  tsNumeroOrden.Color := global_color_salida;
end;


procedure TfrmActividades.tdCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxDBCalcEdit(tdCantidad, key) then
    key := #0;
  if Key = #13 then
    tdventamn.SetFocus
end;

procedure TfrmActividades.tdCantidadAnexoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tdCantidad.SetFocus
end;

procedure TfrmActividades.tdCantidadChange(Sender: TObject);
begin
  TRxDBCalcEditChangef(tdCantidad, 'Cantidad');
end;

procedure TfrmActividades.tsNumeroActividadExit(Sender: TObject);
begin
  tsNumeroActividad.Color := global_color_salida;
  if (ActividadesxOrden.State = dsInsert) or (ActividadesxOrden.State = dsEdit) then
    if not ActividadesxOrden.FieldByName('sNumeroActividad').IsNull then
    begin {
        Connection.qryBusca.Active := False ;
        Connection.qryBusca.Filtered := False;
        Connection.qryBusca.SQL.Clear ;
        Connection.qryBusca.SQL.Add('Select sWbs, mDescripcion, dVentaMN, dVentaDLL, dCostoMN, dCostoDLL, dCantidadAnexo, ' +
                                    'sMedida, dFechaInicio, dFechaFinal from actividadesxanexo Where sContrato = :contrato And ' +
                                    'sNumeroActividad = :actividad And sWbs = :Wbs And sTipoActividad = "Actividad"') ;
        connection.qryBusca.Params.ParamByName('contrato').DataType := ftString ;
        connection.qryBusca.Params.ParamByName('contrato').Value := global_contrato ;
        connection.qryBusca.Params.ParamByName('actividad').DataType := ftString ;
        connection.qryBusca.Params.ParamByName('actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad'] ;
        connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString ;
        if ActividadesxOrden.FieldValues['sWbs']<>NULL then
          connection.qryBusca.Params.ParamByName('Wbs').Value := Trim(ActividadesxOrden.FieldValues['sWbs'])
        else
          connection.qryBusca.Params.ParamByName('Wbs').Value :='';

        connection.qryBusca.Open ;

        If connection.qryBusca.RecordCount > 0 then
        begin
            If (ActividadesxOrden.FieldValues['dCantidad'] = Null) Or (ActividadesxOrden.FieldValues['dCantidad'] = 0) Then
                ActividadesxOrden.FieldValues['dCantidad'] := connection.qryBusca.FieldValues['dCantidadAnexo'] ;
            ActividadesxOrden.FieldValues['sWbsContrato'] := connection.qryBusca.FieldValues['sWbs'] ;
            ActividadesxOrden.FieldValues['mDescripcion'] := connection.qryBusca.FieldValues['mDescripcion'] ;
            ActividadesxOrden.FieldValues['dFechaInicio'] := connection.qryBusca.FieldValues['dFechaInicio'] ;
            ActividadesxOrden.FieldValues['dFechaFinal'] := connection.qryBusca.FieldValues['dFechaFinal'] ;
            ActividadesxOrden.FieldValues['dDuracion'] := ( ActividadesxOrden.FieldValues['dFechaFinal'] - ActividadesxOrden.FieldValues['dFechaInicio'] ) + 1 ;
            ActividadesxOrden.FieldValues['sMedida'] := connection.qryBusca.FieldValues['sMedida'] ;
            If NOT connection.qryBusca.FieldByName('dVentaMN').IsNull Then
                ActividadesxOrden.FieldValues['dVentaMN'] := connection.qryBusca.FieldValues['dVentaMN'] ;
            If NOT connection.qryBusca.FieldByName('dVentaDLL').IsNull Then
                ActividadesxOrden.FieldValues['dVentaDLL'] := connection.qryBusca.FieldValues['dVentaDLL'] ;
            If NOT connection.qryBusca.FieldByName('dCostoMN').IsNull Then
                ActividadesxOrden.FieldValues['dCostoMN'] := connection.qryBusca.FieldValues['dCostoMN'] ;
            If NOT connection.qryBusca.FieldByName('dVentaDLL').IsNull Then
                ActividadesxOrden.FieldValues['dCostoDLL'] := connection.qryBusca.FieldValues['dCostoDLL'] ;
        end ;  }

      if tsPaquete.KeyValue = Null then
        ActividadesxOrden.FieldValues['sWbs'] := Trim(ActividadesxOrden.FieldValues['sNumeroActividad'])
      else
        ActividadesxOrden.FieldValues['sWbs'] := ActividadesxOrden.FieldValues['sWbsAnterior'] + '.' + Trim(ActividadesxOrden.FieldValues['sNumeroActividad']);
    end;
end;

procedure TfrmActividades.tdFechaInicioExit(Sender: TObject);
begin
  if frmBarra1.btnCancel.Enabled = True then
    tdFechaFinal.Date := tdFechaInicio.Date + (ActividadesxOrden.FieldValues['dDuracion'] - 1);
  tdFechaInicio.Color := global_color_salida
end;

procedure TfrmActividades.tsPaqueteKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    if tsNumeroActividad.Enabled = true then
      tsNumeroActividad.SetFocus
    else
      tdCantidad.SetFocus;
end;

procedure TfrmActividades.frmBarra1btnPrinterClick(Sender: TObject);
begin
  Global_OptGrafica := 'Frente';
  Application.CreateForm(TfrmGraficaGerencialDX, frmGraficaGerencialDx);
  frmGraficaGerencialDX.show;
end;

procedure TfrmActividades.tsNumeroOrdenEnter(Sender: TObject);
begin
  frmBarra1.btnCancel.Click;
  tsNumeroOrden.Color := global_color_entrada;
end;

procedure TfrmActividades.tsPaqueteEnter(Sender: TObject);
begin
  tsPaquete.Color := global_color_entrada
end;

procedure TfrmActividades.tsPaqueteExit(Sender: TObject);
begin
  tsPaquete.Color := global_color_salida;
  if frmBarra1.btnCancel.Enabled = True then
  begin
    ActividadesxOrden.FieldValues['dFechaInicio'] := Paquetes.FieldValues['dFechaInicio'];
    ActividadesxOrden.FieldValues['dFechaFinal'] := Paquetes.FieldValues['dFechaFinal'];
  end
end;

procedure TfrmActividades.tmDescripcionEnter(Sender: TObject);
begin
  tmDescripcion.Color := global_color_entrada
end;

procedure TfrmActividades.tmDescripcionExit(Sender: TObject);
begin
  tmDescripcion.Color := global_color_salida
end;

procedure TfrmActividades.tsActAnteriorEnter(Sender: TObject);
begin
  tsactanterior.Color := global_color_entrada;
end;

procedure TfrmActividades.tsActAnteriorExit(Sender: TObject);
begin
  tsactanterior.Color := global_color_salida
end;

procedure TfrmActividades.tsActAnteriorKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tmdescripcion.SetFocus
end;

procedure TfrmActividades.tsItemOrdenEnter(Sender: TObject);
begin
  tsItemOrden.Color := global_color_entrada
end;

procedure TfrmActividades.tsItemOrdenExit(Sender: TObject);
begin
  tsItemOrden.Color := global_color_salida
end;

procedure TfrmActividades.tsItemOrdenKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tmDescripcion.SetFocus
end;

procedure TfrmActividades.tsNumeroActividadEnter(Sender: TObject);
begin
  tsNumeroActividad.Color := global_color_entrada
end;

procedure TfrmActividades.tdFechaInicioEnter(Sender: TObject);
begin
  tdFechaInicio.Color := global_color_entrada
end;

procedure TfrmActividades.tdDuracionChange(Sender: TObject);
begin
//  TDBEditChangef(tdDuracion, 'Duración');
end;

procedure TfrmActividades.tdDuracionEnter(Sender: TObject);
begin
//  tdDuracion.Color := global_color_entrada
end;

procedure TfrmActividades.tdFechaFinalEnter(Sender: TObject);
begin
  tdFechaFinal.Color := global_color_entrada
end;

procedure TfrmActividades.tdFechaFinalExit(Sender: TObject);
begin
  if frmBarra1.btnCancel.Enabled = True then
    ActividadesxOrden.FieldValues['dDuracion'] := DaysBetween(tdFechaFinal.Date, tdFechaInicio.Date) + 1;
  tdFechaFinal.Color := global_color_salida
end;

procedure TfrmActividades.tdCantidadEnter(Sender: TObject);
begin
  tdCantidad.Color := global_color_entrada
end;

procedure TfrmActividades.tdCantidadExit(Sender: TObject);
begin
  tdCantidad.Color := global_color_salida
end;

procedure TfrmActividades.tiColorEnter(Sender: TObject);
begin
  tiColor.Color := global_color_entrada;
end;

procedure TfrmActividades.tiColorExit(Sender: TObject);
begin
  tiColor.Color := global_color_salida
end;

procedure TfrmActividades.tiColorKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsnumeroactividad.SetFocus;
end;

procedure TfrmActividades.tlCalculoEnter(Sender: TObject);
begin
  tlCalculo.Color := global_color_entrada
end;

procedure TfrmActividades.tlCalculoExit(Sender: TObject);
begin
  tlCalculo.Color := global_color_salida
end;

procedure TfrmActividades.tlCalculoKeyPress(Sender: TObject; var Key: Char);
begin
    if Key = #13 then
      tsnumeroactividad.SetFocus;
end;

procedure TfrmActividades.MenuItem9Click(Sender: TObject);
begin
  frmBarra1.btnPrinter.Click
end;

procedure TfrmActividades.PartidasBeforeDelete(DataSet: TDataSet);
begin
  Abort
end;

procedure TfrmActividades.ActividadesxOrdenAfterScroll(DataSet: TDataSet);
Var
  sSumaTotalHrsP : String;
begin
  sSumaTotalHrsP := '00:00';
  if isOpen then
  begin
    if ActividadesxOrden.RecordCount > 0 then
    begin
      if (ActividadesxOrden.State <> dsInsert) and (ActividadesxOrden.State <> dsEdit) then
      begin
        if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
        begin
          tlGerencial.Enabled := True;
          rxdbgrd1.Enabled:=false;
        end
        else
        begin
          tlGerencial.Enabled := False;
          rxdbgrd1.Enabled:=True;
        end;

        case tiColor.ItemIndex of
          0: tiColores.ItemIndex := 0;
          1: tiColores.ItemIndex := 1;
          2: tiColores.ItemIndex := 2;
          3: tiColores.ItemIndex := 3;
          4: tiColores.ItemIndex := 4;
          5: tiColores.ItemIndex := 5;
          6: tiColores.ItemIndex := 6;
          7: tiColores.ItemIndex := 7;
          8: tiColores.ItemIndex := 8;
          9: tiColores.ItemIndex := 9;
          10: tiColores.ItemIndex := 10;
          11: tiColores.ItemIndex := 11;
          12: tiColores.ItemIndex := 12;
          13: tiColores.ItemIndex := 13;
          14: tiColores.ItemIndex := 14;
          15: tiColores.ItemIndex := 15;
        end;
      end;
  //    CampoDuracion.Text := ActividadesxOrden.FieldValues['dDiferenciaDuracion'];
    end;
    zq_ProgramaDeActividad.Active := False;
    zq_ProgramaDeActividad.SQL.Text :=  '' +
                                        'SELECT ' +
                                        ' *, ' +
                                        '	@FechaInicial := (CONCAT(dFechaInicio, " ", cast(sHoraInicio AS Time))) AS Inicio, ' +
                                        '	@FechaFinal := (CONCAT(dFechaFinal, " ", cast(sHoraFinal AS Time))) AS Final, ' +
                                        '	CAST( ' +
                                        '		TIMEDIFF(@FechaFinal, @FechaInicial) AS CHAR ' +
                                        '	) AS dDiferenciaDuracion ' +
                                        'FROM ' +
                                        '	actividadesxorden_detalle ' +
                                        'WHERE ' +
                                        '	sContrato = :Contrato ' +
                                        '	AND sNumeroOrden = :Orden ' +
                                        '	AND sWbs = :Wbs ' +
                                        '	AND sIdConvenio = :Convenio ' +
                                        'ORDER BY ' +
                                        '	dFechaInicio, ' +
                                        '	Time(sHoraInicio) ' +
                                        '';
    zq_ProgramaDeActividad.Params.ParamByName('Contrato').AsString := ActividadesxOrden.FieldByName('sContrato').AsString;
    zq_ProgramaDeActividad.Params.ParamByName('Orden').AsString := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
    zq_ProgramaDeActividad.Params.ParamByName('Wbs').AsString := ActividadesxOrden.FieldByName('sWbs').AsString;
    zq_ProgramaDeActividad.Params.ParamByName('Convenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
    zq_ProgramaDeActividad.Open;
    while Not zq_ProgramaDeActividad.Eof do begin
      sSumaTotalHrsP := sfnSumaHoras(sSumaTotalHrsP, zq_ProgramaDeActividad.FieldByName('dDiferenciaDuracion').AsString);
      zq_ProgramaDeActividad.Next;
    end;
    zq_ProgramaDeActividad.First;
    if ActividadesxOrden.FieldByName('sTipoActividad').AsString = 'Paquete' then begin
      edt1.Text := ActividadesxOrden.FieldByName('sDuracionHoras').AsString;
    end else begin
      if sSumaTotalHrsP = '00:00' then begin
        edt1.Text := ActividadesxOrden.FieldByName('dDiferenciaDuracion').AsString;
      end else begin
        edt1.Text := sSumaTotalHrsP;
      end;
    end;
//  lbl5.BringToFront;
  end;
end;

procedure TfrmActividades.zProcCancelaInsert(DataSet: TDataSet);
begin
  abort
end;

procedure TfrmActividades.zq_ProgramaDeActividadAfterInsert(DataSet: TDataSet);
begin
  if ActividadesxOrden.RecordCount > 0 then begin
    if ActividadesxOrden.FieldByName('sTipoActividad').AsString = 'Actividad' then
    begin
      zq_ProgramaDeActividad.FieldByName('sContrato').AsString := ActividadesxOrden.FieldByName('sContrato').AsString;
      zq_ProgramaDeActividad.FieldByName('sIdConvenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
      zq_ProgramaDeActividad.FieldByName('sNumeroOrden').AsString := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
      zq_ProgramaDeActividad.FieldByName('sWbs').AsString := ActividadesxOrden.FieldByName('sWbs').AsString;
      zq_ProgramaDeActividad.FieldByName('sNumeroActividad').AsString := ActividadesxOrden.FieldByName('sNumeroActividad').AsString;
      zq_ProgramaDeActividad.FieldByName('dFechaInicio').AsDateTime:=now();
      zq_ProgramaDeActividad.FieldByName('dFechaFinal').AsDateTime:=now();
      zq_ProgramaDeActividad.FieldByName('sHoraInicio').AsString:='00:00';
      zq_ProgramaDeActividad.FieldByName('sHoraFinal').AsString:='00:00';
    end;
  end;
end;

procedure TfrmActividades.tdPonderadoChange(Sender: TObject);
begin
  TDBEditChangef(tdPonderado, 'Ponderado');
end;

procedure TfrmActividades.tdPonderadoEnter(Sender: TObject);
begin
  tdPonderado.Color := global_color_entrada;
end;

procedure TfrmActividades.tdPonderadoExit(Sender: TObject);
begin
  tdPonderado.Color := global_color_salida;
end;

procedure TfrmActividades.tdPonderadoKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTDBEdit(tdPonderado, key) then
    key := #0;
  if Key = #13 then
    tlcalculo.SetFocus;
end;

procedure TfrmActividades.grid_actividadesGetCellParams(
  Sender: TObject; Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
  try
    if (Sender as TrxDBGrid).DataSource.DataSet.State = dsBrowse then
      if ActividadesxOrden.RecordCount > 0 then
      begin
        AFont.Color := esColor(ActividadesxOrden.FieldValues['iColor']);

        if (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sTipoActividad').AsString = 'Paquete' then
        begin
          Afont.Style := [fsBold];
          Afont.Size := Afont.Size + 1;
          Background := EsBkColor[(Sender as TrxDBGrid).DataSource.DataSet.FieldByName('iNivel').AsInteger + 1];

            if ((Sender as TrxDBGrid).DataSource.DataSet.FieldByName('dPonderado').AsFloat > 100) then
            begin
                Afont.Style := [fsBold, fsItalic];
                AFont.Color := clRed;
            end
        end
        else
          if ((Sender as TrxDBGrid).DataSource.DataSet.FieldByName('dExcedente').AsFloat > 0) then
          begin
            Afont.Style := [fsBold, fsItalic];
            AFont.Color := clRed;
          end
      end
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Conceptos/Partidas x Frente de Trabajo', 'Al cambiar de registro de actividades', 0);
    end;
  end;
end;

procedure TfrmActividades.grid_actividadesMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x, y);
end;

procedure TfrmActividades.grid_actividadesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TfrmActividades.grid_actividadesTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmActividades.PartidasBeforeInsert(DataSet: TDataSet);
begin
  Abort
end;

procedure TfrmActividades.tiColoresEnter(Sender: TObject);
begin
  tiColores.Color := global_color_entrada
end;

procedure TfrmActividades.tiColoresExit(Sender: TObject);
begin
  tiColores.Color := global_color_salida
end;

procedure TfrmActividades.tiColoresKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tsunidad.SetFocus
end;

procedure TfrmActividades.tiColoresChange(Sender: TObject);
begin
  case tiColores.ItemIndex of
    0: tiColor.ItemIndex := 0;
    1: tiColor.ItemIndex := 1;
    2: tiColor.ItemIndex := 2;
    3: tiColor.ItemIndex := 3;
    4: tiColor.ItemIndex := 4;
    5: tiColor.ItemIndex := 5;
    6: tiColor.ItemIndex := 6;
    7: tiColor.ItemIndex := 7;
    8: tiColor.ItemIndex := 8;
    9: tiColor.ItemIndex := 9;
    10: tiColor.ItemIndex := 10;
    11: tiColor.ItemIndex := 11;
    12: tiColor.ItemIndex := 12;
    13: tiColor.ItemIndex := 13;
    14: tiColor.ItemIndex := 14;
    15: tiColor.ItemIndex := 15;
  end
end;

procedure TfrmActividades.PaquetesCalcFields(DataSet: TDataSet);
begin
  PaquetessDescripcion.Text := MidStr(Paquetes.FieldValues['mDescripcion'], 1, 70)
end;

procedure TfrmActividades.BtnImprimeClick(Sender: TObject);
begin
  if ChkAcumulativo.Checked then
    ImprimirAvxFPAcumulativo(DFechaInicio.Date,DFechaFin.date,CmbFolio,ChbPrevisualizar.Checked)
  else
    ImprimirAvxFP(DFechaInicio.Date,DFechaFin.date,CmbFolio,ChbPrevisualizar.Checked);
end;

procedure TfrmActividades.BtnSalirClick(Sender: TObject);
begin
  PnlRango.Visible := False;
end;

procedure TfrmActividades.ImprimirAvxFPAcumulativo(FInicio,FFin:TDateTime;CmbFl:TAdvComboBox;Ver:Boolean);
Const
  IdMEs =12;
  MesesDA  : array  [1..IdMEs] of string = ('ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE');
  FilaIni = 10;
  ColumnaIni = 5;

var
  zRoConsulta:TZReadOnlyQuery;
  ExcelAp,Libro,AHoja: Variant;
    AnchoTotal:Real;
  CFechas:TDateTime;
  ListaFolios,ListaPartidas:TStringList;
  EFila,EColumna,CFolios,Cpartidas,colinimes,Oldmes,ColIniAño,OldAño,ColFinal,OldAño2,OldMes2:Integer;
  FechaI,FechaF:TDateTime;
  AuxStr:string;
  BCreado:Boolean;
  Firmantes:Tfirmas;
  IndiceHoja,IndiceLibro,NHojas:Integer;

  RangoFi,RangoFf,Ffirmantei,Ffirmatef:TDateTime;

  x : integer;
  
  function ObtieneFirmantes(FIni,FFin:Tdatetime):Tfirmas;
  var ZCFirmas:Tzreadonlyquery;
  Vtemp:Tfirmas;
  begin
    Vtemp := Tfirmas.create;
    Zcfirmas := Tzreadonlyquery.create(nil);
    try
      Zcfirmas.connection := connection.zconnection;
      ZcFirmas.active := False;
      ZcFirmas.sql.text := 'Select f.* from firmas f where f.scontrato = :contrato and f.dIdFecha = (select max(didfecha) from firmas where didfecha <= :fecha and scontrato = f.sContrato) ';
      ZcFirmas.parambyname('Contrato').asstring := Global_contrato;
      ZcFirmas.parambyname('Fecha').asdatetime := FIni;
      ZcFirmas.open;
      if Zcfirmas.recordcount = 1 then
      begin
        Vtemp.Finicialcia := ZcFirmas.FieldByName('sfirmante1').AsString;
        Vtemp.FInicialPep := ZcFirmas.FieldByName('sfirmante2').AsString;
      end;

      ZcFirmas.active := False;
      ZcFirmas.parambyname('Fecha').asdatetime := Ffin;
      ZcFirmas.open;
      if Zcfirmas.recordcount = 1 then
      begin
        Vtemp.FFinalCia := ZcFirmas.FieldByName('sfirmante1').AsString;
        Vtemp.FFinalPep := ZcFirmas.FieldByName('sfirmante2').AsString;
      end;

    finally
      ZcFirmas.free;
      Result := Vtemp;
    end;

  end;


  procedure EncabezadoTexto(iFl,Icl,FcL:Integer);
    Procedure PFormatosExcel_H2(Var Excel: Variant; AutoFit: Integer = 0; Negritas: Boolean = False; SizeFont: Integer = 6; ColorExcel: Cardinal = clBlack; FontName: String = 'Arial'; FormatoCelda: String = '-');
    Var
      Rango: Variant;
    begin
      Excel.Selection.MergeCells := True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment := xlCenter;
      Excel.Selection.Font.Size := SizeFont;
      Excel.Selection.Font.Bold := Negritas;
      Excel.Selection.Font.Color := ColorExcel;
      Excel.Selection.Font.Name := FontName;
      Excel.Selection.WrapText := True;
      if FormatoCelda <> '-' then begin
        Excel.Selection.NumberFormat := FormatoCelda;
      end;
      Rango := Excel.Selection;
      if AutoFit > 1 then begin
        Excel.Selection.EntireRow.RowHeight := AutoFit;
      end;
    end;
  begin

    Fcl := 17;
    ExcelaP.Range[ColumnaNombre(Icl)+IntToStr(iFl)+':'+ColumnaNombre(FcL)+IntToStr(iFl)].Select;
    PFormatosExcel_H2(ExcelaP, 16, True, 12, clBlack, 'Arial');
    ExcelaP.Selection.HorizontalAlignment := xlCenter;
    ExcelAp.Selection.Value := 'AVANCES GLOBALES POR FOLIO / PARTIDAS';
    Inc (iFl);

    ExcelaP.Range[ColumnaNombre(Icl)+IntToStr(iFl)+':'+ColumnaNombre(FcL)+IntToStr(iFl)].Select;
    PFormatosExcel_H2(ExcelaP, 38, True, 8, clBlack, 'Arial');
    ExcelaP.Selection.HorizontalAlignment := xlCenter;
    ExcelaP.Selection.Value := global_contrato;
    ExcelaP.Selection.ReadingOrder := xlContext;
    ExcelaP.Selection.WrapText := True;
    //Hoja.PageSetup.PrintTitleRows := '$1:$5';
  end;

  procedure EncabezadoImagen(Izquierda,Derecha:boolean;FcL:Integer;Modo:integer = 1);
  VAR tMPNAME,TempPath :string;
  imgAux : TImage;
  Pic:TJpegImage;
  fs:TStream;
  begin
    Fcl := 15;
    Anchototal := 880;
    //Imagen Izquierda
    if Izquierda then
    begin
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then
        begin
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtemp1.jpg';
          fs := Connection.contrato.CreateBlobStream(Connection.contrato.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
          If fs.Size > 1 Then
          Begin
            try
              Pic:=TJpegImage.Create;
              try
                Pic.LoadFromStream(fs);
                imgAux.Picture.Graphic := Pic;
              finally
                Pic.Free;
              end;
            finally
              fs.Free;
            End;
            imgAux.Picture.SaveToFile(TmpName);
          End;
        end;
      Finally
        imgAux.Free;
      End;
      if FileExists(TmpName) then
      begin
        if Modo = 1 then
          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 1, 150, 85);

        if Modo = 2 then
          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 1, 135, 70);
      end;
    end;

    if derecha then
    begin
      //Imagen Derecha
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then
        begin
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtemp2.jpg';
          fs := Connection.configuracion.CreateBlobStream(Connection.configuracion.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
          If fs.Size > 1 Then Begin
            try
              Pic:=TJpegImage.Create;
              try
                Pic.LoadFromStream(fs);
                imgAux.Picture.Graphic := Pic;
              finally
                Pic.Free;
              end;
            finally
              fs.Free;
            End;
            imgAux.Picture.SaveToFile(TmpName);
          End;
        end;
      Finally
        imgAux.Free;
      End;
      if FileExists(TmpName) then
      begin
        if Modo = 1 then                                             //375

          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True,(Anchototal), 1, 90, 85);
        if Modo = 2 then
          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 490, 1, 75, 70);
      end;
    end;
  end;

  procedure Distintos(CpTipo:String);
  var
    OldValor:string;
  begin
    if (zRoConsulta.State <> dsBrowse) then
      raise Exception.Create('No se obtuvo ningún resultado de la consulta.');//esto dificilmente pasará
    OldValor := '';
    if CpTipo = 'Partida' then
    begin
      zRoConsulta.SortedFields := 'snumeroactividad';
      zRoConsulta.SortedFields := 'iitemorden';
    end;

    zRoConsulta.First;
    ListaPartidas.Clear;
    while not zRoConsulta.Eof do
    begin
      if CpTipo = 'Folio' then
      begin
        if ListaFolios.IndexOf( zRoConsulta.FieldByName('snumeroorden').asstring) = -1 then
          ListaFolios.Add(zRoConsulta.FieldByName('snumeroorden').asstring);
      end;
      if CpTipo = 'Partida' then
      begin
        if ListaPartidas.IndexOf(zRoConsulta.FieldByName('snumeroactividad').asstring) = -1 then
          ListaPartidas.Add(zRoConsulta.FieldByName('snumeroactividad').asstring);
      end;
      zRoConsulta.Next;
    end;
  end;

begin
  BCreado := False;
  anchototal := 0;
  if FFin < FInicio then
    raise Exception.Create('La fecha de fin no debe ser menor que la fecha de inicio.');

  zRoConsulta:=TZReadOnlyQuery.Create(nil);
  try
    PbExcel.Max := 100;
    PbExcel.Min := 0;
    PbExcel.Position := 0;
    PbExcel.Visible := True;
    zRoConsulta.Connection := connection.zConnection;
    zRoConsulta.Active := False;
    zRoConsulta.SQL.Text :=

    'SELECT ba.didfecha,ba.snumeroorden,ba.snumeroactividad ,ot.sidplataforma,ao.iitemorden, '+
    'ot.mdescripcion as descripcionfolio,ao.mdescripcion as descripcionpda, '+
    '(select (ifnull(sum(dcantidad),0)*100)  from bitacoradeactividades where dIdFecha <= ba.didfecha  and '+
    'sContrato = ot.sContrato and sNumeroOrden = ot.sNumeroOrden and snumeroactividad = ao.sNumeroActividad '+
    'and swbs = ao.swbs and  sIdTipoMovimiento = "ED" AND sidclasificacion IN ("TE","NP","FC") ) as davance, '+
    '( '+
    'SELECT sum(ba2.davance*ao2.dPonderado) as avpq FROM ordenesdetrabajo ot2 '+
    'INNER JOIN actividadesxorden ao2 '+
    'ON (ao2.scontrato = ot2.scontrato AND ao2.snumeroorden = ot2.snumeroorden) '+
    'INNER JOIN  '+
    'bitacoradeactividades ba2 '+
    'ON (ot2.sidfolio = ba2.sNumeroOrden AND ba2.swbs = ao2.swbs and ba2.snumeroactividad = ao2.sNumeroActividad) '+
    'INNER JOIN  '+
    'reportediario r2  '+
    'ON (ba2.didfecha = r2.dIdFecha AND r2.sorden = ot2.scontrato AND r2.sIdConvenio = ao2.sidconvenio) '+
    'WHERE ba2.sContrato = :contrato AND ao2.sIdConvenio = :convenio and ba2.dIdFecha <=ba.dIdFecha and ba2.sNumeroOrden = ba.sNumeroOrden '+
    ' AND ba2.sIdTipoMovimiento = "ED" AND ba2.sidclasificacion IN ("TE","NP","FC") '+
    'GROUP BY ba2.snumeroorden '+
    'ORDER BY ba2.dIdFecha,ba2.sNumeroOrden,  ba2.snumeroActividad '+
    ') as davancepaquete, '+
    '(select (ifnull(sum(dcantidad),0)*ao.dPonderado)  from bitacoradeactividades where dIdFecha < DATE_FORMAT(date(ba.didfecha),"%Y/%m/01")  and '+
    'sContrato = ot.sContrato and sNumeroOrden = ot.sNumeroOrden and snumeroactividad = ao.sNumeroActividad and swbs = ao.swbs and  sIdTipoMovimiento = "ED" AND sidclasificacion IN ("TE","NP","FC") ) as AvanceAnterior, '+
    '(select (ifnull(sum(dcantidad),0)*ao.dPonderado)   from bitacoradeactividades where (dIdFecha between  DATE_FORMAT(date(ba.didfecha),"%Y/%m/01") and LAST_DAY(ba.didfecha)) and '+
    'sContrato = ot.sContrato and sNumeroOrden = ot.sNumeroOrden and snumeroactividad = ao.sNumeroActividad and swbs = ao.swbs and  sIdTipoMovimiento = "ED" AND sidclasificacion IN ("TE","NP","FC") ) as AvanceActual, '+
    '(select (ifnull(sum(ba5.dcantidad*ao5.dPonderado),0))as no from actividadesxorden ao5 '+
    'inner join bitacoradeactividades ba5 on (ba5.scontrato = ao5.scontrato and ba5.sNumeroOrden = ao5.sNumeroOrden and ba5.snumeroactividad = ao5.sNumeroActividad and ao5.swbs = ba5.swbs) '+
    'where ao5.sContrato =  ot.sContrato and ao5.sNumeroOrden = ot.sNumeroOrden and ao5.sIdConvenio = :convenio and ba5.sIdTipoMovimiento = "ED" AND ba5.sidclasificacion IN ("TE","NP","FC") '+
    'and YEAR(ba5.didfecha) <= YEAR(ba.didfecha) and MONTH(ba5.didfecha) < MONTH(ba.didfecha) '+
    'group by ao5.sContrato,ao5.sNumeroOrden ) as AvanceAnteriorPaq, '+
    '(select (ifnull(sum(ba5.dcantidad*ao5.dPonderado),0))as no from actividadesxorden ao5 '+
    'inner join bitacoradeactividades ba5 on (ba5.scontrato = ao5.scontrato and ba5.sNumeroOrden = ao5.sNumeroOrden and ba5.snumeroactividad = ao5.sNumeroActividad and ao5.swbs = ba5.swbs) '+
    'where ao5.sContrato =  ot.sContrato and ao5.sNumeroOrden = ot.sNumeroOrden and ao5.sIdConvenio = :convenio and ba5.sIdTipoMovimiento = "ED" AND ba5.sidclasificacion IN ("TE","NP","FC") '+
    'and YEAR(ba5.didfecha) = YEAR(ba.didfecha) and MONTH(ba5.didfecha) = MONTH(ba.didfecha) '+
    'group by ao5.sContrato,ao5.sNumeroOrden  ) as AvanceActualPaq,  '+
    '(select min(didfecha)  from bitacoradeactividades where sContrato = ot.sContrato and sNumeroOrden = ot.sNumeroOrden '+
    'and snumeroactividad = ao.sNumeroActividad and swbs = ao.swbs and  sIdTipoMovimiento = "ED" AND sidclasificacion IN ("TE","NP","FC") ) as minfpartida, '+
    '(select max(didfecha)  from bitacoradeactividades where sContrato = ot.sContrato and sNumeroOrden = ot.sNumeroOrden '+
    'and snumeroactividad = ao.sNumeroActividad and swbs = ao.swbs and  sIdTipoMovimiento = "ED" AND sidclasificacion IN ("TE","NP","FC") ) as maxfpartida  '+
    'FROM ordenesdetrabajo ot '+
    'INNER JOIN actividadesxorden ao '+
    'ON (ao.scontrato = ot.scontrato AND ao.snumeroorden = ot.snumeroorden) '+
    'INNER JOIN '+
    'bitacoradeactividades ba '+
    'ON (ot.sidfolio = ba.sNumeroOrden AND ba.swbs = ao.swbs and ba.snumeroactividad = ao.sNumeroActividad) '+
    'INNER JOIN '+
    'reportediario r '+
    'ON (ba.didfecha = r.dIdFecha AND r.sorden = ot.scontrato AND r.sIdConvenio = ao.sidconvenio) '+
    'WHERE ba.sContrato = :contrato AND ao.sIdConvenio = :convenio and ba.dIdFecha BETWEEN :FechaI AND :FechaF  '+
    ' AND ba.sIdTipoMovimiento = "ED" AND ba.sidclasificacion IN ("TE","NP","FC") and (:Folio = -1 or (:Folio <> -1 and ba.snumeroorden = :Folio)) '+
    'GROUP BY ba.snumeroorden,ba.didfecha,ba.sNumeroActividad '+
    'ORDER BY ba.dIdFecha,ba.sNumeroOrden,  ao.iItemOrden';
    zRoConsulta.ParamByName('contrato').AsString := global_contrato;
    zRoConsulta.ParamByName('convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
    zRoConsulta.ParamByName('fechai').AsDate := FInicio;
    zRoConsulta.ParamByName('fechaf').AsDate := FFin;
    if CmbFl.ItemIndex = 0 then
      zRoConsulta.ParamByName('Folio').AsString := '-1'
    else
      zRoConsulta.ParamByName('Folio').AsString := CmbFl.Text;

    zRoConsulta.Open;
    
    if (zRoConsulta.RecordCount = 0) then
      raise Exception.Create('No se obtuvo ningún resultado de la consulta.');//esto dificilmente pasará

    if not GuardaExcel.Execute then
      raise Exception.Create('Proceso de generacion de archivo excel cancelado por el usuario.');

    ListaFolios:=TStringList.Create;
    ListaPartidas := Tstringlist.create;

    Distintos('Folio');
    PbExcel.Max := ListaFolios.Count-1;
    Try
      ExcelAp := CreateOleObject('Excel.Application');
    Except
      On E: Exception do
      begin
        raise Exception.Create('No se puede iniciar la aplicación excel, verifique que tenga instalado la paquetería office.');
      end;
    End;
    ExcelAp.Visible := Ver;
    ExcelAp.DisplayAlerts:= False;
    Libro := ExcelAp.Workbooks.Add;
    IndiceLibro := ExcelAp.Workbooks.count;

    //Generar encabezado con dias y folio
    zRoConsulta.First;
    FechaI := zRoConsulta.FieldByName('didfecha').AsDateTime;
    zRoConsulta.Last;
    FechaF := zRoConsulta.FieldByName('didfecha').AsDateTime;

    for x := 0 to 1 do
    begin
      CFechas := FechaI;
      Oldmes := 0;
      OldAño := 0;
      IndiceHoja := 1;

      while CFechas < FechaF  do
      begin
        if (YearOf(cfechas) <> OldAño) or (MonthOf(cfechas) <> Oldmes) then
        begin
          OldAño :=YearOf(cfechas);
          Oldmes :=MonthOf(cfechas);
          if x = 0 then
          begin
            if IndiceHoja > Libro.Sheets.count then
              AHoja := Libro.Sheets.Add;
          end
          else
          begin
            Libro.Sheets[IndiceHoja].Name := MesesDA[MonthOf(cfechas)] +FormatDateTime('-yyyy',CFechas);
          end;
          IndiceHoja := IndiceHoja+1;
        end;
        CFechas := IncDay(cfechas);
      end;
    end;
    Nhojas := IndiceHoja;

     RangoFF := FechaI -500;
     RangoFi := Fechai;
     oldaño2 := 0;
     oldMes2 := 0;
     while RangoFi < FechaF do
     begin
       if (yearof(RangoFi) <> yearof(RangoFF)) or (monthof(Rangofi) <> MonthOf(RangoFf)) then
       begin
         RangoFF := RangoFi;
         while DayOf(RangoFf) < DaysInMonth(RangoFi) do
         begin
           RangoFf := IncDay(Rangoff);
         end;
         if RangoFf > FechaF then
           RangoFf := FechaF;
         ExcelAp.Workbooks [indicelibro].Worksheets[MesesDA[MonthOf(rangoFi)] +FormatDateTime('-yyyy',rangoFi)].select;
         ExcelAp.activeWindow.DisplayGridlines := false;
         ExcelAp.Cells[1,1 ].Font.Name:='Arial';
         ExcelAp.Cells[1,1 ].Font.size:=8;
          //Fechas años, meses dias
          CFechas := RangoFI;
          Oldmes := MonthOf(CFechas);
          OldAño := YearOf(CFechas);
          EColumna := ColumnaIni;
          Efila := Filaini;
          while CFechas <= RangoFf do
          begin
            if CFechas = RangoFI then
            begin
              ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-3)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-3)].Select;
              ExcelAp.Selection.Value := OldAño;
              ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-2)].Select;
              ExcelAp.Selection.Value := MesesDA[oldmes];
              colinimes := ecolumna;
              ColIniAño := EColumna;
            end;
            if OldAño <> YearOf(CFechas) then
            begin
              OldAño := YearOf(CFechas);
              ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-3)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-3)].Select;
              ExcelAp.Selection.Value := Oldaño;
              ColIniAño := EColumna;
            end
            else
            begin
              ExcelAp.Range[ColumnaNombre(colinimes)+IntToStr(efila-3)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-3)].Select;
              ExcelAp.Selection.MergeCells := True;
            end;

            if Oldmes <> MonthOf(CFechas) then
            begin
              Oldmes := MonthOf(CFechas);
              ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-2)].Select;
              ExcelAp.Selection.Value := MesesDA[oldmes];
              colinimes := ecolumna;
            end
            else
            begin
              ExcelAp.Range[ColumnaNombre(colinimes)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-2)].Select;
              ExcelAp.Selection.MergeCells := True;
            end;
            ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-1)].Select;
            ExcelAp.Selection.Value := DayOf(CFechas);
            ExcelAp.Columns[ColumnaNombre(ecolumna)+':'+ColumnaNombre(ecolumna)].ColumnWidth := 8;
            CFechas := IncDay(CFechas,1);
            Anchototal :=  Anchototal + (8)*12;
            ColFinal := ecolumna;
            ecolumna := ecolumna+1;
          end;

          EncabezadoTexto(2,3,ColFinal);
          EncabezadoImagen(True,True,ColFinal,1);

          zRoConsulta.Filtered := False;
          zRoConsulta.Filter := ' didfecha >= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFi))+' and didfecha <= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFf));
          zRoConsulta.Filtered := True;

          Distintos('Folio');

          ExcelAp.Range[ColumnaNombre(ColumnaIni-4)+IntToStr(efila-1)+':'+ColumnaNombre(ColumnaIni-4)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'FOLIO';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni-3)+IntToStr(efila-1)+':'+ColumnaNombre(ColumnaIni-3)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'PARTIDA';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni-2)+IntToStr(efila-1)+':'+ColumnaNombre(ColumnaIni-2)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'PLATAFORMA';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni-1)+IntToStr(efila-1)+':'+ColumnaNombre(ColumnaIni-1)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'DESCRPCION';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;

          //Avances
          ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'AVANCE ANTERIOR';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ecolumna+1)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+1)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'AVANCE ACTUAL';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ecolumna+2)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+2)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'FECHA INICIO';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ecolumna+3)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+3)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'FECHA TERMINO';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;

          //Firmantes
          ExcelAp.Range[ColumnaNombre(ecolumna+4)+IntToStr(efila-3)+':'+ColumnaNombre(ecolumna+7)+IntToStr(efila-3)].Select;
          ExcelAp.Selection.MergeCells := True;
          ExcelAp.Selection.Value := 'FIRMANTES';

          ExcelAp.Range[ColumnaNombre(ecolumna+4)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna+5)+IntToStr(efila-2)].Select;
          ExcelAp.Selection.MergeCells := True;
          ExcelAp.Selection.Value := 'INCIALES';

          ExcelAp.Range[ColumnaNombre(ecolumna+6)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna+7)+IntToStr(efila-2)].Select;
          ExcelAp.Selection.MergeCells := True;
          ExcelAp.Selection.Value := 'FINALES';

          ExcelAp.Range[ColumnaNombre(ecolumna+4)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+4)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'CIA';

          ExcelAp.Range[ColumnaNombre(ecolumna+5)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+5)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'PEP';

          ExcelAp.Range[ColumnaNombre(ecolumna+6)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+6)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'CIA';

          ExcelAp.Range[ColumnaNombre(ecolumna+7)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+7)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'PEP';

          ExcelAp.Range[ColumnaNombre(ColumnaIni)+IntToStr(efila-3)+':'+ColumnaNombre(colfinal+8)+IntToStr(efila-3)].Select;
          ExcelAp.selection.interior.colorindex := 48;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni)+IntToStr(efila-2)+':'+ColumnaNombre(colfinal+8)+IntToStr(efila-2)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.selection.interior.colorindex := 15;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni-4)+IntToStr(efila-1)+':'+ColumnaNombre(colfinal+8)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.selection.interior.colorindex := 24;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;

          //Contenido
         //-----------------------------------------------------------------------------------------------

          EFila := FilaIni;

          for CFolios := 0 to ListaFolios.Count-1 do
          begin
            EColumna := ColumnaIni;
            //Paquete de folio por fechas
            zRoConsulta.Filtered := False;
            zRoConsulta.Filter := 'snumeroorden = '+QuotedStr(ListaFolios[CFolios])+' and didfecha >= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFi))+' and didfecha <= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFf));
            zRoConsulta.Filtered := True;
            Distintos('Partida');

            CFechas := RangoFi;

            EColumna := ColumnaIni;

            for Cpartidas := 0 to ListaPartidas.count-1 do
            begin
              if Cpartidas = 0 then
              begin
                ExcelAp.Range[ColumnaNombre(ecolumna-4)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-4)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := zRoConsulta.FieldByName('snumeroorden').AsString;
                ExcelAp.Selection.VerticalAlignment := xlCenter;
                ExcelAp.Columns[ColumnaNombre(ecolumna-4)+':'+ColumnaNombre(ecolumna-4)].ColumnWidth := 17;

                ExcelAp.Range[ColumnaNombre(ecolumna-2)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-2)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := zRoConsulta.FieldByName('sidplataforma').AsString;
                ExcelAp.Selection.VerticalAlignment := xlCenter;

                ExcelAp.Range[ColumnaNombre(ecolumna-1)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-1)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := zRoConsulta.FieldByName('descripcionfolio').AsString;
                ExcelAp.Columns[ColumnaNombre(ecolumna-1)+':'+ColumnaNombre(ecolumna-1)].ColumnWidth := 22.2;

                ExcelAp.Range[ColumnaNombre(ecolumna-4)+IntToStr(efila)+':'+ColumnaNombre(colfinal+8)+IntToStr(efila)].Select;
                ExcelAp.selection.interior.colorindex := 37;

                ExcelAp.Range[ColumnaNombre(ColFinal+1)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+1)+IntToStr(efila)].Select;
                if zRoConsulta.FieldByName('avanceanteriorpaq').asfloat <= 100 then
                  ExcelAp.Selection.Value := zRoConsulta.FieldByName('avanceanteriorpaq').asfloat
                else
                  ExcelAp.Selection.Value := 100;
                ExcelAp.Range[ColumnaNombre(ColFinal+2)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+2)+IntToStr(efila)].Select;
                if  zRoConsulta.FieldByName('avanceactualpaq').asfloat <= 100 then
                  ExcelAp.Selection.Value := zRoConsulta.FieldByName('avanceactualpaq').asfloat{-zRoConsulta.FieldByName('avanceanterior').asfloat}
                else
                  ExcelAp.Selection.Value := 100;
                zRoConsulta.First;
                Ffirmantei := zRoConsulta.FieldByName('minfpartida').AsDateTime;
                zRoConsulta.Last;
                Ffirmatef := zRoConsulta.FieldByName('maxfpartida').AsDateTime;
                zRoConsulta.First;

                ExcelAp.Range[ColumnaNombre(ColFinal+3)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+3)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := fFirmantei;
                ExcelAp.Selection.VerticalAlignment := xlCenter;
                ExcelAp.Range[ColumnaNombre(ColFinal+4)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+4)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Ffirmatef;
                ExcelAp.Selection.VerticalAlignment := xlCenter;

                if not assigned(Firmantes) then
                  Firmantes := Tfirmas.Create;
                Firmantes := obtienefirmantes(Ffirmantei,Ffirmatef);

                ExcelAp.Range[ColumnaNombre(ColFinal+5)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+5)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Firmantes.FInicialCia;
                ExcelAp.Range[ColumnaNombre(ColFinal+6)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+6)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Firmantes.FInicialPep;

                ExcelAp.Range[ColumnaNombre(ColFinal+7)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+7)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Firmantes.FFinalCia;
                ExcelAp.Range[ColumnaNombre(ColFinal+8)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+8)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Firmantes.FFinalPep;
                EFila := Efila+1;

                while CFechas <= RangoFf  do
                begin
                  if zRoConsulta.Locate('didfecha',CFechas,[]) then
                  begin
                    ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-1)].Select;
                    ExcelAp.Selection.Value := zRoConsulta.FieldByName('davancepaquete').AsFloat;
                  end
                  else
                  begin
                    ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-1)].Select;
                    if EColumna = ColumnaIni then
                      ExcelAp.Selection.Value := 0
                    else
                      ExcelAp.Selection.Value := ExcelAp.Workbooks [indicelibro].Worksheets[MesesDA[MonthOf(rangoFi)] +FormatDateTime('-yyyy',rangoFi)].Cells[efila-1,ecolumna-1].value;
                  end;

                  CFechas := IncDay(CFechas,1);
                  EColumna := EColumna+1;
                end;
                EColumna := ColumnaIni;

              end;

              zRoConsulta.Filtered := False;
              zRoConsulta.Filter := 'snumeroorden = '+QuotedStr(ListaFolios[CFolios]) + ' AND snumeroactividad = '+QuotedStr(Listapartidas[cPArtidas])+' and didfecha >= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFi))+' and didfecha <= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFf));
              zRoConsulta.Filtered := True;
              zRoConsulta.First;
              EColumna := columnaini;

              ExcelAp.Range[ColumnaNombre(ecolumna-3)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-3)+IntToStr(efila)].Select;
              ExcelAp.Selection.NumberFormat := '@';
              ExcelAp.Selection.Value := zRoConsulta.FieldByName('snumeroactividad').AsString;
              ExcelAp.Selection.VerticalAlignment := xlCenter;

              ExcelAp.Range[ColumnaNombre(ecolumna-1)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-1)+IntToStr(efila)].Select;
              ExcelAp.Selection.Value := zRoConsulta.FieldByName('descripcionpda').AsString;

              CFechas := RangoFi;
              while CFechas <= RangoFf  do
              begin
                if zRoConsulta.Locate('didfecha',CFechas,[]) then
                begin
                  ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila)+':'+ColumnaNombre(ecolumna)+IntToStr(efila)].Select;
                  if zRoConsulta.FieldByName('davance').AsFloat <= 100 then
                    ExcelAp.Selection.Value := zRoConsulta.FieldByName('davance').AsFloat
                  else
                    ExcelAp.Selection.Value := 100;
                  if zRoConsulta.FieldByName('davance').AsFloat >= 100 then
                    ExcelAp.selection.interior.colorindex := 35;

                end
                else
                begin
                  ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila)+':'+ColumnaNombre(ecolumna)+IntToStr(efila)].Select;
                  if EColumna = ColumnaIni then
                    ExcelAp.Selection.Value := 0
                  else
                    ExcelAp.Selection.Value := ExcelAp.Workbooks [indicelibro].Worksheets[MesesDA[MonthOf(rangoFi)] +FormatDateTime('-yyyy',rangoFi)].Cells[efila,ecolumna-1].value;
                end;

                //lo que va despues de las columnas finales
                if EColumna = ColFinal then
                begin
                  ExcelAp.Range[ColumnaNombre(ColFinal+1)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+1)+IntToStr(efila)].Select;
                  if zRoConsulta.FieldByName('avanceanterior').asfloat <= 100 then
                    ExcelAp.Selection.Value := zRoConsulta.FieldByName('avanceanterior').asfloat
                  else
                    ExcelAp.Selection.Value := 100;
                  ExcelAp.selection.interior.colorindex := 36;
                  ExcelAp.Range[ColumnaNombre(ColFinal+2)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+2)+IntToStr(efila)].Select;
                  ExcelAp.Selection.Value := zRoConsulta.FieldByName('avanceactual').asfloat;//-zRoConsulta.FieldByName('avanceanterior').asfloat;
                  ExcelAp.selection.interior.colorindex := 36;
                  ExcelAp.Range[ColumnaNombre(ColFinal+3)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+3)+IntToStr(efila)].Select;
                  ExcelAp.Selection.Value := zRoConsulta.FieldByName('minfpartida').AsDateTime;
                  ExcelAp.Range[ColumnaNombre(ColFinal+4)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+4)+IntToStr(efila)].Select;
                  ExcelAp.Selection.Value := zRoConsulta.FieldByName('maxfpartida').AsDateTime;
                  ExcelAp.selection.interior.colorindex := 36;
                end;

                CFechas := IncDay(CFechas,1);
                EColumna := EColumna+1;
              end;
              EFila := Efila+1;
            end;
            PbExcel.Position := CFolios;
          end;
          //Tamaño letra
          ExcelAp.Range[ColumnaNombre(ColumnaIni-4)+IntToStr(FilaIni-3)+':'+ColumnaNombre(colfinal+8)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.Font.size := 8;

          ExcelAp.Range[ColumnaNombre(ColumnaIni-4)+IntToStr(FilaIni)+':'+ColumnaNombre(ColumnaIni)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          //ExcelAp.Selection.NumberFormat := '@';
          ExcelAp.Selection.HorizontalAlignment := xlCenter;

          ExcelAp.Range[ColumnaNombre(ColFinal)+IntToStr(FilaIni)+':'+ColumnaNombre(ColFinal+8)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Selection.VerticalAlignment := xlCenter;

          //numerico dias
          ExcelAp.Range[ColumnaNombre(ColumnaIni)+IntToStr(FilaIni)+':'+ColumnaNombre(colfinal+2)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Selection.VerticalAlignment := xlCenter;
          ExcelAp.Selection.NumberFormat := '0.00';

          //Ancho fr columnas
          ExcelAp.Columns[ColumnaNombre(ColumnaIni-4)+':'+ColumnaNombre(ColumnaIni-4)].ColumnWidth := 17;
          ExcelAp.Columns[ColumnaNombre(ColumnaIni-3)+':'+ColumnaNombre(ColumnaIni-3)].ColumnWidth := 6;
          ExcelAp.Columns[ColumnaNombre(ColumnaIni-2)+':'+ColumnaNombre(ColumnaIni-2)].ColumnWidth := 12;
          ExcelAp.Columns[ColumnaNombre(ColumnaIni-1)+':'+ColumnaNombre(ColumnaIni-1)].ColumnWidth := 50;

          ExcelAp.Range[ColumnaNombre(ColumnaIni-1)+IntToStr(FilaIni)+':'+ColumnaNombre(ColumnaIni-1)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.WrapText := True;

          ExcelAp.Range[ColumnaNombre(ColumnaIni-1)+IntToStr(FilaIni)+':'+ColumnaNombre(ColumnaIni-1)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.WrapText := True;

          ExcelAp.Columns[ColumnaNombre(colfinal+1)+':'+ColumnaNombre(colfinal+1)].ColumnWidth := 12;
          ExcelAp.Columns[ColumnaNombre(colfinal+2)+':'+ColumnaNombre(colfinal+2)].ColumnWidth := 12;
          ExcelAp.Columns[ColumnaNombre(colfinal+3)+':'+ColumnaNombre(colfinal+3)].ColumnWidth := 10.5;
          ExcelAp.Columns[ColumnaNombre(colfinal+4)+':'+ColumnaNombre(colfinal+4)].ColumnWidth := 10.5;
          ExcelAp.Columns[ColumnaNombre(colfinal+5)+':'+ColumnaNombre(colfinal+5)].ColumnWidth := 32;
          ExcelAp.Columns[ColumnaNombre(colfinal+6)+':'+ColumnaNombre(colfinal+6)].ColumnWidth := 32;
          ExcelAp.Columns[ColumnaNombre(colfinal+7)+':'+ColumnaNombre(colfinal+7)].ColumnWidth := 32;
          ExcelAp.Columns[ColumnaNombre(colfinal+8)+':'+ColumnaNombre(colfinal+8)].ColumnWidth := 32;
         //-------------------------------------------------------------------------------------------------
          ExcelAp.ActiveWindow.View := 2;
          ExcelAp.ActiveSheet.PageSetup.LeftMargin := 0.7;
          ExcelAp.ActiveSheet.PageSetup.RightMargin := 0.7;
          ExcelAp.ActiveSheet.PageSetup.TopMargin := 0.75;
          ExcelAp.ActiveSheet.PageSetup.BottomMargin := 0.75;
          ExcelAp.ActiveSheet.PageSetup.HeaderMargin := 0.3 ;
          ExcelAp.ActiveSheet.PageSetup.FooterMargin := 0.3;
          ExcelAp.ActiveSheet.PageSetup.PrintHeadings := False;
          ExcelAp.ActiveSheet.PageSetup.PrintGridlines := False;
          ExcelAp.ActiveSheet.PageSetup.PrintQuality := 600;
          ExcelAp.ActiveSheet.PageSetup.CenterHorizontally := False;
          ExcelAp.ActiveSheet.PageSetup.CenterVertically := False;
          ExcelAp.ActiveSheet.PageSetup.Draft := False;
          ExcelAp.ActiveSheet.PageSetup.PaperSize := 1;
          ExcelAp.ActiveSheet.PageSetup.BlackAndWhite := False;
          ExcelAp.ActiveSheet.PageSetup.Zoom := False;
          ExcelAp.ActiveSheet.PageSetup.FitToPagesWide := 1;
          ExcelAp.ActiveSheet.PageSetup.FitToPagesTall := 1;
          ExcelAp.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
          ExcelAp.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := True;
          ExcelAp.ActiveWindow.Zoom := 100;
       end;
       rangoFi := incday(rangofi);
     end;

    BCreado := True;
    if Assigned(ListaFolios) then
      ListaFolios.free;
    if Assigned(ListaPartidas) then
      ListaPartidas.free;

  finally
    PbExcel.Visible := False;
    zRoConsulta.Free;
    if (BCreado)  then
    begin
      Libro.SaveAs(guardaexcel.FileName);
      Sleep(100);
      ExcelAp.quit;
      Sleep(200);
      ShellExecute(Handle,'open',pwidechar(guardaexcel.FileName), nil, nil,  SW_SHOWNORMAL);
    end
    else
      try
        ExcelAp.quit;
      Except
        ;
      end;
  end;
end;

procedure TfrmActividades.ImprimirAvxFP(FInicio,FFin:TDateTime;CmbFl:TAdvComboBox;Ver:Boolean);
Const
  IdMEs =12;
  MesesDA  : array  [1..IdMEs] of string = ('ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE');
  FilaIni = 10;
  ColumnaIni = 5;

var
  zRoConsulta:TZReadOnlyQuery;
  ExcelAp,Libro,AHoja: Variant;
    AnchoTotal:Real;
  CFechas:TDateTime;
  ListaFolios,ListaPartidas:TStringList;
  EFila,EColumna,CFolios,Cpartidas,colinimes,Oldmes,ColIniAño,OldAño,ColFinal,OldAño2,OldMes2:Integer;
  FechaI,FechaF:TDateTime;
  AuxStr:string;
  BCreado:Boolean;
  Firmantes:Tfirmas;
  IndiceHoja,IndiceLibro,NHojas:Integer;

  RangoFi,RangoFf,Ffirmantei,Ffirmatef:TDateTime;

  x : Integer;

  function ObtieneFirmantes(FIni,FFin:Tdatetime):Tfirmas;
  var ZCFirmas:Tzreadonlyquery;
  Vtemp:Tfirmas;
  begin
    Vtemp := Tfirmas.create;
    Zcfirmas := Tzreadonlyquery.create(nil);
    try
      Zcfirmas.connection := connection.zconnection;
      ZcFirmas.active := False;
      ZcFirmas.sql.text := 'Select f.* from firmas f where f.scontrato = :contrato and f.dIdFecha = (select max(didfecha) from firmas where didfecha <= :fecha and scontrato = f.sContrato) ';
      ZcFirmas.parambyname('Contrato').asstring := Global_contrato;
      ZcFirmas.parambyname('Fecha').asdatetime := FIni;
      ZcFirmas.open;
      if Zcfirmas.recordcount = 1 then
      begin
        Vtemp.Finicialcia := ZcFirmas.FieldByName('sfirmante1').AsString;
        Vtemp.FInicialPep := ZcFirmas.FieldByName('sfirmante2').AsString;
      end;

      ZcFirmas.active := False;
      ZcFirmas.parambyname('Fecha').asdatetime := Ffin;
      ZcFirmas.open;
      if Zcfirmas.recordcount = 1 then
      begin
        Vtemp.FFinalCia := ZcFirmas.FieldByName('sfirmante1').AsString;
        Vtemp.FFinalPep := ZcFirmas.FieldByName('sfirmante2').AsString;
      end;

    finally
      ZcFirmas.free;
      Result := Vtemp;
    end;

  end;


  procedure EncabezadoTexto(iFl,Icl,FcL:Integer);
    Procedure PFormatosExcel_H2(Var Excel: Variant; AutoFit: Integer = 0; Negritas: Boolean = False; SizeFont: Integer = 6; ColorExcel: Cardinal = clBlack; FontName: String = 'Arial'; FormatoCelda: String = '-');
    Var
      Rango: Variant;
    begin
      Excel.Selection.MergeCells := True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment := xlCenter;
      Excel.Selection.Font.Size := SizeFont;
      Excel.Selection.Font.Bold := Negritas;
      Excel.Selection.Font.Color := ColorExcel;
      Excel.Selection.Font.Name := FontName;
      Excel.Selection.WrapText := True;
      if FormatoCelda <> '-' then begin
        Excel.Selection.NumberFormat := FormatoCelda;
      end;
      Rango := Excel.Selection;
      if AutoFit > 1 then begin
        Excel.Selection.EntireRow.RowHeight := AutoFit;
      end;
    end;
  begin

    Fcl := 17;
    ExcelaP.Range[ColumnaNombre(Icl)+IntToStr(iFl)+':'+ColumnaNombre(FcL)+IntToStr(iFl)].Select;
    PFormatosExcel_H2(ExcelaP, 16, True, 12, clBlack, 'Arial');
    ExcelaP.Selection.HorizontalAlignment := xlCenter;
    ExcelAp.Selection.Value := 'AVANCES GLOBALES POR FOLIO / PARTIDAS';
    Inc (iFl);

    ExcelaP.Range[ColumnaNombre(Icl)+IntToStr(iFl)+':'+ColumnaNombre(FcL)+IntToStr(iFl)].Select;
    PFormatosExcel_H2(ExcelaP, 38, True, 8, clBlack, 'Arial');
    ExcelaP.Selection.HorizontalAlignment := xlCenter;
    ExcelaP.Selection.Value := global_contrato;
    ExcelaP.Selection.ReadingOrder := xlContext;
    ExcelaP.Selection.WrapText := True;
    //Hoja.PageSetup.PrintTitleRows := '$1:$5';
  end;

  procedure EncabezadoImagen(Izquierda,Derecha:boolean;FcL:Integer;Modo:integer = 1);
  VAR tMPNAME,TempPath :string;
  imgAux : TImage;
  Pic:TJpegImage;
  fs:TStream;
  begin

      Fcl := 15;
      Anchototal := 880;
    //Imagen Izquierda
    if Izquierda then
    begin
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then
        begin
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtemp1.jpg';
          fs := Connection.contrato.CreateBlobStream(Connection.contrato.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
          If fs.Size > 1 Then
          Begin
            try
              Pic:=TJpegImage.Create;
              try
                Pic.LoadFromStream(fs);
                imgAux.Picture.Graphic := Pic;
              finally
                Pic.Free;
              end;
            finally
              fs.Free;
            End;
            imgAux.Picture.SaveToFile(TmpName);
          End;
        end;
      Finally
        imgAux.Free;
      End;
      if FileExists(TmpName) then
      begin
        if Modo = 1 then
          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 1, 150, 85);

        if Modo = 2 then
          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 1, 135, 70);
      end;
    end;

    if derecha then
    begin
      //Imagen Derecha
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then
        begin
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtemp2.jpg';
          fs := Connection.configuracion.CreateBlobStream(Connection.configuracion.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
          If fs.Size > 1 Then Begin
            try
              Pic:=TJpegImage.Create;
              try
                Pic.LoadFromStream(fs);
                imgAux.Picture.Graphic := Pic;
              finally
                Pic.Free;
              end;
            finally
              fs.Free;
            End;
            imgAux.Picture.SaveToFile(TmpName);
          End;
        end;
      Finally
        imgAux.Free;
      End;
      if FileExists(TmpName) then
      begin
        if Modo = 1 then                                             //375

          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True,(Anchototal), 1, 90, 85);
        if Modo = 2 then
          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 490, 1, 75, 70);
      end;
    end;
  end;



  procedure Distintos(CpTipo:String);
  var
      OldValor:string;
  begin
    if (zRoConsulta.State <> dsBrowse) then
      raise Exception.Create('No se obtuvo ningún resultado de la consulta.');//esto dificilmente pasará
    OldValor := '';
    if CpTipo = 'Partida' then
    begin
      zRoConsulta.SortedFields := 'snumeroactividad';
      zRoConsulta.SortedFields := 'iitemorden';
    end;

    zRoConsulta.First;
    ListaPartidas.Clear;
    while not zRoConsulta.Eof do
    begin
      if CpTipo = 'Folio' then
      begin
        if ListaFolios.IndexOf( zRoConsulta.FieldByName('snumeroorden').asstring) = -1 then
          ListaFolios.Add(zRoConsulta.FieldByName('snumeroorden').asstring);
      end;
      if CpTipo = 'Partida' then
      begin
        if ListaPartidas.IndexOf(zRoConsulta.FieldByName('snumeroactividad').asstring) = -1 then
          ListaPartidas.Add(zRoConsulta.FieldByName('snumeroactividad').asstring);
      end;
      zRoConsulta.Next;
    end;
  end;

begin
  BCreado := False;
  anchototal := 0;
  if FFin < FInicio then
    raise Exception.Create('La fecha de fin no debe ser menor que la fecha de inicio.');

  zRoConsulta:=TZReadOnlyQuery.Create(nil);
  try
    PbExcel.Max := 100;
    PbExcel.Min := 0;
    PbExcel.Position := 0;
    PbExcel.Visible := True;
    zRoConsulta.Connection := connection.zConnection;
    zRoConsulta.Active := False;
    zRoConsulta.SQL.Text :=

    'SELECT ba.didfecha,ba.snumeroorden,ba.snumeroactividad ,ot.sidplataforma,ao.iitemorden, '+
    'ot.mdescripcion as descripcionfolio,ao.mdescripcion as descripcionpda,sum(ba.davance)*100 as davance, '+
    '( '+
    'SELECT sum(ba2.davance*ao2.dPonderado) as avpq FROM ordenesdetrabajo ot2 '+
    'INNER JOIN actividadesxorden ao2 '+
    'ON (ao2.scontrato = ot2.scontrato AND ao2.snumeroorden = ot2.snumeroorden) '+
    'INNER JOIN  '+
    'bitacoradeactividades ba2 '+
    'ON (ot2.sidfolio = ba2.sNumeroOrden AND ba2.swbs = ao2.swbs and ba2.snumeroactividad = ao2.sNumeroActividad) '+
    'INNER JOIN  '+
    'reportediario r2  '+
    'ON (ba2.didfecha = r2.dIdFecha AND r2.sorden = ot2.scontrato AND r2.sIdConvenio = ao2.sidconvenio) '+
    'WHERE ba2.sContrato = :contrato AND ao2.sIdConvenio = :convenio and ba2.dIdFecha =ba.dIdFecha and ba2.sNumeroOrden = ba.sNumeroOrden '+
    ' AND ba2.sIdTipoMovimiento = "ED" AND ba2.sidclasificacion IN ("TE","NP","FC") '+
    'GROUP BY ba2.snumeroorden,ba2.didfecha '+
    'ORDER BY ba2.dIdFecha,ba2.sNumeroOrden,  ba2.snumeroActividad '+
    ') as davancepaquete, '+
    '(select (ifnull(sum(dcantidad),0)*ao.dPonderado)  from bitacoradeactividades where dIdFecha < DATE_FORMAT(date(ba.didfecha),"%Y/%m/01")  and '+
    'sContrato = ot.sContrato and sNumeroOrden = ot.sNumeroOrden and snumeroactividad = ao.sNumeroActividad and swbs = ao.swbs and  sIdTipoMovimiento = "ED" AND sidclasificacion IN ("TE","NP","FC") ) as AvanceAnterior, '+
    '(select (ifnull(sum(dcantidad),0)*ao.dPonderado)   from bitacoradeactividades where (dIdFecha between  DATE_FORMAT(date(ba.didfecha),"%Y/%m/01") and LAST_DAY(ba.didfecha)) and '+
    'sContrato = ot.sContrato and sNumeroOrden = ot.sNumeroOrden and snumeroactividad = ao.sNumeroActividad and swbs = ao.swbs and  sIdTipoMovimiento = "ED" AND sidclasificacion IN ("TE","NP","FC") ) as AvanceActual, '+
    '(select (ifnull(sum(ba5.dcantidad*ao5.dPonderado),0))as no from actividadesxorden ao5 '+
    'inner join bitacoradeactividades ba5 on (ba5.scontrato = ao5.scontrato and ba5.sNumeroOrden = ao5.sNumeroOrden and ba5.snumeroactividad = ao5.sNumeroActividad and ao5.swbs = ba5.swbs) '+
    'where ao5.sContrato =  ot.sContrato and ao5.sNumeroOrden = ot.sNumeroOrden and ao5.sIdConvenio = :convenio and ba5.sIdTipoMovimiento = "ED" AND ba5.sidclasificacion IN ("TE","NP","FC") '+
    'and YEAR(ba5.didfecha) <= YEAR(ba.didfecha) and MONTH(ba5.didfecha) < MONTH(ba.didfecha) '+
    'group by ao5.sContrato,ao5.sNumeroOrden ) as AvanceAnteriorPaq, '+
    '(select (ifnull(sum(ba5.dcantidad*ao5.dPonderado),0))as no from actividadesxorden ao5 '+
    'inner join bitacoradeactividades ba5 on (ba5.scontrato = ao5.scontrato and ba5.sNumeroOrden = ao5.sNumeroOrden and ba5.snumeroactividad = ao5.sNumeroActividad and ao5.swbs = ba5.swbs) '+
    'where ao5.sContrato =  ot.sContrato and ao5.sNumeroOrden = ot.sNumeroOrden and ao5.sIdConvenio = :convenio and ba5.sIdTipoMovimiento = "ED" AND ba5.sidclasificacion IN ("TE","NP","FC") '+
    'and YEAR(ba5.didfecha) = YEAR(ba.didfecha) and MONTH(ba5.didfecha) = MONTH(ba.didfecha) '+
    'group by ao5.sContrato,ao5.sNumeroOrden  ) as AvanceActualPaq,  '+
    '(select min(didfecha)  from bitacoradeactividades where sContrato = ot.sContrato and sNumeroOrden = ot.sNumeroOrden '+
    'and snumeroactividad = ao.sNumeroActividad and swbs = ao.swbs and  sIdTipoMovimiento = "ED" AND sidclasificacion IN ("TE","NP","FC") ) as minfpartida, '+
    '(select max(didfecha)  from bitacoradeactividades where sContrato = ot.sContrato and sNumeroOrden = ot.sNumeroOrden '+
    'and snumeroactividad = ao.sNumeroActividad and swbs = ao.swbs and  sIdTipoMovimiento = "ED" AND sidclasificacion IN ("TE","NP","FC") ) as maxfpartida  '+
    'FROM ordenesdetrabajo ot '+
    'INNER JOIN actividadesxorden ao '+
    'ON (ao.scontrato = ot.scontrato AND ao.snumeroorden = ot.snumeroorden) '+
    'INNER JOIN '+
    'bitacoradeactividades ba '+
    'ON (ot.sidfolio = ba.sNumeroOrden AND ba.swbs = ao.swbs and ba.snumeroactividad = ao.sNumeroActividad) '+
    'INNER JOIN '+
    'reportediario r '+
    'ON (ba.didfecha = r.dIdFecha AND r.sorden = ot.scontrato AND r.sIdConvenio = ao.sidconvenio) '+
    'WHERE ba.sContrato = :contrato AND ao.sIdConvenio = :convenio and ba.dIdFecha BETWEEN :FechaI AND :FechaF  '+
    ' AND ba.sIdTipoMovimiento = "ED" AND ba.sidclasificacion IN ("TE","NP","FC") and (:Folio = -1 or (:Folio <> -1 and ba.snumeroorden = :Folio)) '+
    'GROUP BY ba.snumeroorden,ba.didfecha,ba.sNumeroActividad '+
    'ORDER BY ba.dIdFecha,ba.sNumeroOrden,  ao.iItemOrden';
    zRoConsulta.ParamByName('contrato').AsString := global_contrato;
    zRoConsulta.ParamByName('convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
    zRoConsulta.ParamByName('fechai').AsDate := FInicio;
    zRoConsulta.ParamByName('fechaf').AsDate := FFin;
    if CmbFl.ItemIndex = 0 then
      zRoConsulta.ParamByName('Folio').AsString := '-1'
    else
      zRoConsulta.ParamByName('Folio').AsString := CmbFl.Text;

    zRoConsulta.Open;
    
    if (zRoConsulta.RecordCount = 0) then
      raise Exception.Create('No se obtuvo ningún resultado de la consulta.');//esto dificilmente pasará

    if not GuardaExcel.Execute then
      raise Exception.Create('Proceso de generacion de archivo excel cancelado por el usuario.');

    ListaFolios:=TStringList.Create;
    ListaPartidas := Tstringlist.create;

    Distintos('Folio');
    PbExcel.Max := ListaFolios.Count-1;
    Try
      ExcelAp := CreateOleObject('Excel.Application');
    Except
      On E: Exception do
      begin
        raise Exception.Create('No se puede iniciar la aplicación excel, verifique que tenga instalado la paquetería office.');
      end;
    End;
    ExcelAp.Visible := Ver;
    ExcelAp.DisplayAlerts:= False;
    Libro := ExcelAp.Workbooks.Add;
    IndiceLibro := ExcelAp.Workbooks.count;

    //Generar encabezado con dias y folio
    zRoConsulta.First;
    FechaI := zRoConsulta.FieldByName('didfecha').AsDateTime;
    zRoConsulta.Last;
    FechaF := zRoConsulta.FieldByName('didfecha').AsDateTime;

    for x := 0 to 1 do
    begin
      CFechas := FechaI;
      Oldmes := 0;
      OldAño := 0;
      IndiceHoja := 1;

      while CFechas < FechaF  do
      begin
        if (YearOf(cfechas) <> OldAño) or (MonthOf(cfechas) <> Oldmes) then
        begin
          OldAño :=YearOf(cfechas);
          Oldmes :=MonthOf(cfechas);
          if x = 0 then
          begin
            if IndiceHoja > Libro.Sheets.count then
              AHoja := Libro.Sheets.Add;
          end
          else
          begin
            Libro.Sheets[IndiceHoja].Name := MesesDA[MonthOf(cfechas)] +FormatDateTime('-yyyy',CFechas);
          end;
          IndiceHoja := IndiceHoja+1;
        end;
        CFechas := IncDay(cfechas);
      end;
    end;
    Nhojas := IndiceHoja;

     RangoFF := FechaI -500;
     RangoFi := Fechai;
     oldaño2 := 0;
     oldMes2 := 0;
     while RangoFi < FechaF do
     begin
       if (yearof(RangoFi) <> yearof(RangoFF)) or (monthof(Rangofi) <> MonthOf(RangoFf)) then
       begin
         RangoFF := RangoFi;
         while DayOf(RangoFf) < DaysInMonth(RangoFi) do
         begin
           RangoFf := IncDay(Rangoff);
         end;
         if RangoFf > FechaF then
           RangoFf := FechaF;
         ExcelAp.Workbooks [indicelibro].Worksheets[MesesDA[MonthOf(rangoFi)] +FormatDateTime('-yyyy',rangoFi)].select;
         ExcelAp.activeWindow.DisplayGridlines := false;
         ExcelAp.Cells[1,1 ].Font.Name:='Arial';
         ExcelAp.Cells[1,1 ].Font.size:=8;
          //Fechas años, meses dias
          CFechas := RangoFI;
          Oldmes := MonthOf(CFechas);
          OldAño := YearOf(CFechas);
          EColumna := ColumnaIni;
          Efila := Filaini;
          while CFechas <= RangoFf do
          begin
            if CFechas = RangoFI then
            begin
              ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-3)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-3)].Select;
              ExcelAp.Selection.Value := OldAño;
              ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-2)].Select;
              ExcelAp.Selection.Value := MesesDA[oldmes];
              colinimes := ecolumna;
              ColIniAño := EColumna;
            end;
            if OldAño <> YearOf(CFechas) then
            begin
              OldAño := YearOf(CFechas);
              ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-3)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-3)].Select;
              ExcelAp.Selection.Value := Oldaño;
              ColIniAño := EColumna;
            end
            else
            begin
              ExcelAp.Range[ColumnaNombre(colinimes)+IntToStr(efila-3)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-3)].Select;
              ExcelAp.Selection.MergeCells := True;
            end;

            if Oldmes <> MonthOf(CFechas) then
            begin
              Oldmes := MonthOf(CFechas);
              ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-2)].Select;
              ExcelAp.Selection.Value := MesesDA[oldmes];
              colinimes := ecolumna;
            end
            else
            begin
              ExcelAp.Range[ColumnaNombre(colinimes)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-2)].Select;
              ExcelAp.Selection.MergeCells := True;
            end;
            ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-1)].Select;
            ExcelAp.Selection.Value := DayOf(CFechas);
            ExcelAp.Columns[ColumnaNombre(ecolumna)+':'+ColumnaNombre(ecolumna)].ColumnWidth := 8;
            CFechas := IncDay(CFechas,1);
            Anchototal :=  Anchototal + (8)*12;
            ColFinal := ecolumna;
            ecolumna := ecolumna+1;
          end;

          EncabezadoTexto(2,3,ColFinal);
          EncabezadoImagen(True,True,ColFinal,1);

          zRoConsulta.Filtered := False;
          zRoConsulta.Filter := ' didfecha >= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFi))+' and didfecha <= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFf));
          zRoConsulta.Filtered := True;

          Distintos('Folio');

          ExcelAp.Range[ColumnaNombre(ColumnaIni-4)+IntToStr(efila-1)+':'+ColumnaNombre(ColumnaIni-4)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'FOLIO';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni-3)+IntToStr(efila-1)+':'+ColumnaNombre(ColumnaIni-3)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'PARTIDA';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni-2)+IntToStr(efila-1)+':'+ColumnaNombre(ColumnaIni-2)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'PLATAFORMA';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni-1)+IntToStr(efila-1)+':'+ColumnaNombre(ColumnaIni-1)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'DESCRPCION';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;

          //Avances
          ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'AVANCE ANTERIOR';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ecolumna+1)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+1)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'AVANCE ACTUAL';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ecolumna+2)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+2)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'FECHA INICIO';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ecolumna+3)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+3)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'FECHA TERMINO';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;

          //Firmantes
          ExcelAp.Range[ColumnaNombre(ecolumna+4)+IntToStr(efila-3)+':'+ColumnaNombre(ecolumna+7)+IntToStr(efila-3)].Select;
          ExcelAp.Selection.MergeCells := True;
          ExcelAp.Selection.Value := 'FIRMANTES';

          ExcelAp.Range[ColumnaNombre(ecolumna+4)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna+5)+IntToStr(efila-2)].Select;
          ExcelAp.Selection.MergeCells := True;
          ExcelAp.Selection.Value := 'INCIALES';

          ExcelAp.Range[ColumnaNombre(ecolumna+6)+IntToStr(efila-2)+':'+ColumnaNombre(ecolumna+7)+IntToStr(efila-2)].Select;
          ExcelAp.Selection.MergeCells := True;
          ExcelAp.Selection.Value := 'FINALES';

          ExcelAp.Range[ColumnaNombre(ecolumna+4)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+4)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'CIA';

          ExcelAp.Range[ColumnaNombre(ecolumna+5)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+5)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'PEP';

          ExcelAp.Range[ColumnaNombre(ecolumna+6)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+6)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'CIA';

          ExcelAp.Range[ColumnaNombre(ecolumna+7)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna+7)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Value := 'PEP';

          ExcelAp.Range[ColumnaNombre(ColumnaIni)+IntToStr(efila-3)+':'+ColumnaNombre(colfinal+8)+IntToStr(efila-3)].Select;
          ExcelAp.selection.interior.colorindex := 48;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni)+IntToStr(efila-2)+':'+ColumnaNombre(colfinal+8)+IntToStr(efila-2)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.selection.interior.colorindex := 15;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Range[ColumnaNombre(ColumnaIni-4)+IntToStr(efila-1)+':'+ColumnaNombre(colfinal+8)+IntToStr(efila-1)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.selection.interior.colorindex := 24;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;

          //Contenido
         //-----------------------------------------------------------------------------------------------

          EFila := FilaIni;

          for CFolios := 0 to ListaFolios.Count-1 do
          begin
            EColumna := ColumnaIni;
            //Paquete de folio por fechas
            zRoConsulta.Filtered := False;
            zRoConsulta.Filter := 'snumeroorden = '+QuotedStr(ListaFolios[CFolios])+' and didfecha >= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFi))+' and didfecha <= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFf));
            zRoConsulta.Filtered := True;
            Distintos('Partida');

            CFechas := RangoFi;

            EColumna := ColumnaIni;

            for Cpartidas := 0 to ListaPartidas.count-1 do
            begin
              if Cpartidas = 0 then
              begin
                ExcelAp.Range[ColumnaNombre(ecolumna-4)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-4)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := zRoConsulta.FieldByName('snumeroorden').AsString;
                ExcelAp.Selection.VerticalAlignment := xlCenter;
                ExcelAp.Columns[ColumnaNombre(ecolumna-4)+':'+ColumnaNombre(ecolumna-4)].ColumnWidth := 17;

                ExcelAp.Range[ColumnaNombre(ecolumna-2)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-2)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := zRoConsulta.FieldByName('sidplataforma').AsString;
                ExcelAp.Selection.VerticalAlignment := xlCenter;

                ExcelAp.Range[ColumnaNombre(ecolumna-1)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-1)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := zRoConsulta.FieldByName('descripcionfolio').AsString;
                ExcelAp.Columns[ColumnaNombre(ecolumna-1)+':'+ColumnaNombre(ecolumna-1)].ColumnWidth := 22.2;

                ExcelAp.Range[ColumnaNombre(ecolumna-4)+IntToStr(efila)+':'+ColumnaNombre(colfinal+8)+IntToStr(efila)].Select;
                ExcelAp.selection.interior.colorindex := 37;

                ExcelAp.Range[ColumnaNombre(ColFinal+1)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+1)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := zRoConsulta.FieldByName('avanceanteriorpaq').asfloat;

                ExcelAp.Range[ColumnaNombre(ColFinal+2)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+2)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := zRoConsulta.FieldByName('avanceactualpaq').asfloat{-zRoConsulta.FieldByName('avanceanterior').asfloat};


                zRoConsulta.First;
                Ffirmantei := zRoConsulta.FieldByName('minfpartida').AsDateTime;
                zRoConsulta.Last;
                Ffirmatef := zRoConsulta.FieldByName('maxfpartida').AsDateTime;
                zRoConsulta.First;

                ExcelAp.Range[ColumnaNombre(ColFinal+3)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+3)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := fFirmantei;
                ExcelAp.Selection.VerticalAlignment := xlCenter;
                ExcelAp.Range[ColumnaNombre(ColFinal+4)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+4)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Ffirmatef;
                ExcelAp.Selection.VerticalAlignment := xlCenter;

                if not assigned(Firmantes) then
                  Firmantes := Tfirmas.Create;
                Firmantes := obtienefirmantes(Ffirmantei,Ffirmatef);

                ExcelAp.Range[ColumnaNombre(ColFinal+5)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+5)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Firmantes.FInicialCia;
                ExcelAp.Range[ColumnaNombre(ColFinal+6)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+6)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Firmantes.FInicialPep;
                //EFila := Efila+1;
                ExcelAp.Range[ColumnaNombre(ColFinal+7)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+7)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Firmantes.FFinalCia;
                ExcelAp.Range[ColumnaNombre(ColFinal+8)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+8)+IntToStr(efila)].Select;
                ExcelAp.Selection.Value := Firmantes.FFinalPep;
                EFila := Efila+1;

                while CFechas <= RangoFf  do
                begin
                  if zRoConsulta.Locate('didfecha',CFechas,[]) then
                  begin
                    ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila-1)+':'+ColumnaNombre(ecolumna)+IntToStr(efila-1)].Select;
                    if zRoConsulta.FieldByName('davancepaquete').AsFloat <= 100 then
                      ExcelAp.Selection.Value := zRoConsulta.FieldByName('davancepaquete').AsFloat
                    else
                      ExcelAp.Selection.Value := 100;
                  end;

                  CFechas := IncDay(CFechas,1);
                  EColumna := EColumna+1;
                end;
                EColumna := ColumnaIni;

              end;

              zRoConsulta.Filtered := False;
              zRoConsulta.Filter := 'snumeroorden = '+QuotedStr(ListaFolios[CFolios]) + ' AND snumeroactividad = '+QuotedStr(Listapartidas[cPArtidas])+' and didfecha >= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFi))+' and didfecha <= '+quotedstr(FormatDateTime('YYYY/MM/DD',RangoFf));
              zRoConsulta.Filtered := True;
              zRoConsulta.First;
              EColumna := columnaini;

              ExcelAp.Range[ColumnaNombre(ecolumna-3)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-3)+IntToStr(efila)].Select;
              ExcelAp.Selection.NumberFormat := '@';
              ExcelAp.Selection.Value := zRoConsulta.FieldByName('snumeroactividad').AsString;
              ExcelAp.Selection.VerticalAlignment := xlCenter;

              ExcelAp.Range[ColumnaNombre(ecolumna-1)+IntToStr(efila)+':'+ColumnaNombre(ecolumna-1)+IntToStr(efila)].Select;
              ExcelAp.Selection.Value := zRoConsulta.FieldByName('descripcionpda').AsString;

              CFechas := RangoFi;
              while CFechas <= RangoFf  do
              begin
                if zRoConsulta.Locate('didfecha',CFechas,[]) then
                begin
                  ExcelAp.Range[ColumnaNombre(ecolumna)+IntToStr(efila)+':'+ColumnaNombre(ecolumna)+IntToStr(efila)].Select;
                  if zRoConsulta.FieldByName('davance').AsFloat <= 100 then
                    ExcelAp.Selection.Value := zRoConsulta.FieldByName('davance').AsFloat
                  else
                    ExcelAp.Selection.Value := 100;
                  if zRoConsulta.FieldByName('davance').AsFloat >= 100 then
                    ExcelAp.selection.interior.colorindex := 35;
                  
                end;

                //lo que va despues de las columnas finales
                if EColumna = ColFinal then
                begin
                  ExcelAp.Range[ColumnaNombre(ColFinal+1)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+1)+IntToStr(efila)].Select;
                  if zRoConsulta.FieldByName('avanceanterior').asfloat <= 100 then
                    ExcelAp.Selection.Value := zRoConsulta.FieldByName('avanceanterior').asfloat
                  else
                    ExcelAp.Selection.Value := 100;
                  ExcelAp.selection.interior.colorindex := 36;
                  ExcelAp.Range[ColumnaNombre(ColFinal+2)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+2)+IntToStr(efila)].Select;
                  if zRoConsulta.FieldByName('avanceactual').asfloat <= 100 then
                    ExcelAp.Selection.Value := zRoConsulta.FieldByName('avanceactual').asfloat//-zRoConsulta.FieldByName('avanceanterior').asfloat;
                  else
                    ExcelAp.Selection.Value := 100;
                  ExcelAp.selection.interior.colorindex := 36;
                  ExcelAp.Range[ColumnaNombre(ColFinal+3)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+3)+IntToStr(efila)].Select;
                  ExcelAp.Selection.Value := zRoConsulta.FieldByName('minfpartida').AsDateTime;
                  ExcelAp.Range[ColumnaNombre(ColFinal+4)+IntToStr(efila)+':'+ColumnaNombre(ColFinal+4)+IntToStr(efila)].Select;
                  ExcelAp.Selection.Value := zRoConsulta.FieldByName('maxfpartida').AsDateTime;
                  ExcelAp.selection.interior.colorindex := 36;
                end;

                CFechas := IncDay(CFechas,1);
                EColumna := EColumna+1;
              end;
              EFila := Efila+1;
            end;
            PbExcel.Position := CFolios;
          end;
          //Tamaño letra
          ExcelAp.Range[ColumnaNombre(ColumnaIni-4)+IntToStr(FilaIni-3)+':'+ColumnaNombre(colfinal+8)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.Font.size := 8;

          ExcelAp.Range[ColumnaNombre(ColumnaIni-4)+IntToStr(FilaIni)+':'+ColumnaNombre(ColumnaIni)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          //ExcelAp.Selection.NumberFormat := '@';
          ExcelAp.Selection.HorizontalAlignment := xlCenter;

          ExcelAp.Range[ColumnaNombre(ColFinal)+IntToStr(FilaIni)+':'+ColumnaNombre(ColFinal+8)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Selection.VerticalAlignment := xlCenter;

          //numerico dias
          ExcelAp.Range[ColumnaNombre(ColumnaIni)+IntToStr(FilaIni)+':'+ColumnaNombre(colfinal+2)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
          ExcelAp.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.Selection.VerticalAlignment := xlCenter;
          ExcelAp.Selection.NumberFormat := '0.00';

          //Ancho fr columnas
          ExcelAp.Columns[ColumnaNombre(ColumnaIni-4)+':'+ColumnaNombre(ColumnaIni-4)].ColumnWidth := 17;
          ExcelAp.Columns[ColumnaNombre(ColumnaIni-3)+':'+ColumnaNombre(ColumnaIni-3)].ColumnWidth := 6;
          ExcelAp.Columns[ColumnaNombre(ColumnaIni-2)+':'+ColumnaNombre(ColumnaIni-2)].ColumnWidth := 12;
          ExcelAp.Columns[ColumnaNombre(ColumnaIni-1)+':'+ColumnaNombre(ColumnaIni-1)].ColumnWidth := 50;

          ExcelAp.Range[ColumnaNombre(ColumnaIni-1)+IntToStr(FilaIni)+':'+ColumnaNombre(ColumnaIni-1)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.WrapText := True;

          ExcelAp.Range[ColumnaNombre(ColumnaIni-1)+IntToStr(FilaIni)+':'+ColumnaNombre(ColumnaIni-1)+IntToStr(Efila-1)].Select;
          ExcelAp.Selection.WrapText := True;

          ExcelAp.Columns[ColumnaNombre(colfinal+1)+':'+ColumnaNombre(colfinal+1)].ColumnWidth := 12;
          ExcelAp.Columns[ColumnaNombre(colfinal+2)+':'+ColumnaNombre(colfinal+2)].ColumnWidth := 12;
          ExcelAp.Columns[ColumnaNombre(colfinal+3)+':'+ColumnaNombre(colfinal+3)].ColumnWidth := 10.5;
          ExcelAp.Columns[ColumnaNombre(colfinal+4)+':'+ColumnaNombre(colfinal+4)].ColumnWidth := 10.5;
          ExcelAp.Columns[ColumnaNombre(colfinal+5)+':'+ColumnaNombre(colfinal+5)].ColumnWidth := 32;
          ExcelAp.Columns[ColumnaNombre(colfinal+6)+':'+ColumnaNombre(colfinal+6)].ColumnWidth := 32;
          ExcelAp.Columns[ColumnaNombre(colfinal+7)+':'+ColumnaNombre(colfinal+7)].ColumnWidth := 32;
          ExcelAp.Columns[ColumnaNombre(colfinal+8)+':'+ColumnaNombre(colfinal+8)].ColumnWidth := 32;
         //-------------------------------------------------------------------------------------------------
          ExcelAp.ActiveWindow.View := 2;
          ExcelAp.ActiveSheet.PageSetup.LeftMargin := 0.7;
          ExcelAp.ActiveSheet.PageSetup.RightMargin := 0.7;
          ExcelAp.ActiveSheet.PageSetup.TopMargin := 0.75;
          ExcelAp.ActiveSheet.PageSetup.BottomMargin := 0.75;
          ExcelAp.ActiveSheet.PageSetup.HeaderMargin := 0.3 ;
          ExcelAp.ActiveSheet.PageSetup.FooterMargin := 0.3;
          ExcelAp.ActiveSheet.PageSetup.PrintHeadings := False;
          ExcelAp.ActiveSheet.PageSetup.PrintGridlines := False;
          ExcelAp.ActiveSheet.PageSetup.PrintQuality := 600;
          ExcelAp.ActiveSheet.PageSetup.CenterHorizontally := False;
          ExcelAp.ActiveSheet.PageSetup.CenterVertically := False;
          ExcelAp.ActiveSheet.PageSetup.Draft := False;
          ExcelAp.ActiveSheet.PageSetup.PaperSize := 1;
          ExcelAp.ActiveSheet.PageSetup.BlackAndWhite := False;
          ExcelAp.ActiveSheet.PageSetup.Zoom := False;
          ExcelAp.ActiveSheet.PageSetup.FitToPagesWide := 1;
          //if tpo = 1 then  //Todo en una pagina
            ExcelAp.ActiveSheet.PageSetup.FitToPagesTall := 1;
          //if tpo = 2 then
            //ExcelAp.ActiveSheet.PageSetup.FitToPagesTall := 2;
          ExcelAp.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
          ExcelAp.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := True;
          ExcelAp.ActiveWindow.Zoom := 100;


       end;
       rangoFi := incday(rangofi);
     end;

    BCreado := True;
    if Assigned(ListaFolios) then
      ListaFolios.free;
    if Assigned(ListaPartidas) then
      ListaPartidas.free;

  finally
    PbExcel.Visible := False;
    zRoConsulta.Free;
    if (BCreado)  then
    begin
      Libro.SaveAs(guardaexcel.FileName);
      Sleep(100);
      ExcelAp.quit;
      Sleep(200);
      ShellExecute(Handle,'open',pwidechar(guardaexcel.FileName), nil, nil,  SW_SHOWNORMAL);
    end
    else
      try
        ExcelAp.quit;
      Except
        ;
      end;
  end;
end;

procedure TfrmActividades.BuscarPartida1Click(Sender: TObject);
var
  sNumeroPartida: string;
begin
  if ActividadesxOrden.RecordCount > 0 then
  begin
    sNumeroPartida := UPPERCASE(InputBox('Inteligent', 'Inserte la partida a buscar?', ActividadesxOrden.FieldValues['sNumeroActividad']));
    ActividadesxOrden.Locate('sWBS', sNumeroPartida, [loCaseInsensitive])
  end
end;

procedure TfrmActividades.ActividadesxOrdenAfterInsert(
  DataSet: TDataSet);
begin
  ActividadesxOrden.FieldValues['iNivel'] := 0;
  ActividadesxOrden.FieldValues['dCantidad'] := 1;
  ActividadesxOrden.FieldValues['dPonderado'] := 0;
  ActividadesxOrden.FieldValues['dDuracion'] := 0;
  ActividadesxOrden.FieldValues['sIdPlataforma'] := '';
  ActividadesxOrden.FieldValues['sIdPernocta'] := '';
end;

procedure TfrmActividades.tdVentaMNChange(Sender: TObject);
begin
  TRxDBCalcEditChangef(tdVentaMN, '$Moneda Nacional');
end;

procedure TfrmActividades.tdVentaMNEnter(Sender: TObject);
begin
  tdVentaMN.Color := global_color_entrada
end;

procedure TfrmActividades.tdVentaMNExit(Sender: TObject);
begin
  tdVentaMN.Color := global_color_salida
end;

procedure TfrmActividades.tdVentaMNKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxDBCalcEdit(tdVentaMN, key) then
    key := #0;
  if Key = #13 then
    tdCostoMN.SetFocus
end;

procedure TfrmActividades.tdFechaFinalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tmComentarios.SetFocus
end;

procedure TfrmActividades.tmComentariosEnter(Sender: TObject);
begin
  tmComentarios.Color := global_color_entrada
end;

procedure TfrmActividades.tmComentariosExit(Sender: TObject);
begin
  tmComentarios.Color := global_color_salida
end;

procedure TfrmActividades.ActividadesxOrdenCalcFields(
  DataSet: TDataSet);
begin
 //   If not ActividadesxOrden.State in [dsinsert,dsedit] then
//    Begin
  ActividadesxOrdendMontoMN.Value := ActividadesxOrdendCantidad.Value * ActividadesxOrdendVentaMN.Value;
  ActividadesxOrdendMontoDLL.Value := ActividadesxOrdendCantidad.Value * ActividadesxOrdendVentaDLL.Value;

  ActividadesxOrdendMontoCostoMN.Value := ActividadesxOrdendCantidad.Value * ActividadesxOrdendCostoMN.Value;
  ActividadesxOrdendMontoCostoDLL.Value := ActividadesxOrdendCantidad.Value * ActividadesxOrdendCostoDLL.Value;

  if ActividadesxOrden.FieldValues['sWbs'] <> Null then
    ActividadesxOrden.FieldByName('sWbsSpace').AsString := espaces(ActividadesxOrden.FieldValues['iNivel']) + ActividadesxOrden.FieldValues['sNumeroActividad']
    {ActividadesxOrdensWbsSpace.Text}
//   End
end;

procedure TfrmActividades.ActividadesxOrdendDuracionChange(Sender: TField);
begin
  if ActividadesxOrden.State in [DsInsert,DsEdit] then
  begin
    if Not Ciclar then
    begin
      Ciclar:=true;
      ActividadesxOrden.FieldByName('dFechaFinal').AsDateTime:=IncDay(tdFechaInicio.DateTime,strToIntDef(Sender.AsString,0));
      Ciclar:=false;
    end;
  end;
end;

procedure TfrmActividades.ActividadesxOrdendFechaFinalChange(Sender: TField);
begin
 if ActividadesxOrden.State in [DsInsert,DsEdit] then
  begin
    if Not Ciclar then
    begin
      Ciclar:=true;
      ActividadesxOrden.FieldByName('dDuracion').AsInteger:=daysBetween(tdFechaInicio.DateTime,Sender.AsDateTime);
      Ciclar:=false;
    end;
  end;
end;

procedure TfrmActividades.ActividadesxOrdendFechaInicioChange(Sender: TField);
begin
  if ActividadesxOrden.State in [DsInsert,DsEdit] then
  begin
    if Not Ciclar then
    begin
      Ciclar:=true;
      ActividadesxOrden.FieldByName('dDuracion').AsInteger:=daysBetween(Sender.AsDateTime,tdFechaFinal.DateTime);
      Ciclar:=false;
    end;
  end;
end;

procedure TfrmActividades.ActividadesxOrdendPonderadoSetText(Sender: TField;
  const Text: string);
begin
  sender.Value := abs(StrToFloatDef(text, 0));
end;

procedure TfrmActividades.ActividadesxOrdenNewSimbolGetText(Sender: TField;
  var Text: string; DisplayText: Boolean);
begin
  text := '';
  if (ActividadesxOrdenstipoactividad.asstring = 'Paquete') then
    if (Paq.IndexOf(ActividadesxOrdenswbs.asstring) <> -1) then
      text := '+'
    else
      text := '-';
end;

procedure TfrmActividades.ActividadesxOrdensWbsAnteriorChange(
  Sender: TField);
begin
  if (ActividadesxOrden.State = dsInsert) or (ActividadesxOrden.State = dsEdit) then
  begin
    if (ActividadesxOrden.FieldByName('sWBSAnterior').AsString <> '') and
      (ActividadesxOrden.FieldByName('sWBSAnterior').AsString <> '0') then
    begin
      if Paquetes.RecordCount > 0 then
      begin
        ActividadesxOrden.FieldValues['dFechaInicio'] := Paquetes.FieldValues['dFechaInicio'];
        ActividadesxOrden.FieldValues['dDuracion'] := Paquetes.FieldValues['dDuracion'];
        ActividadesxOrden.FieldValues['dFechaFinal'] := Paquetes.FieldValues['dFechaFinal'];
        ActividadesxOrden.FieldValues['iNivel'] := Paquetes.FieldValues['iNivel'] + 1;
        if ActividadesxOrden.FieldValues['iNivel'] > 0 then
          if not ActividadesxOrden.FieldByName('sNumeroActividad').IsNull then
            ActividadesxOrden.FieldValues['sWbs'] := ActividadesxOrden.FieldValues['sWBSAnterior'] + '.' + ActividadesxOrden.FieldValues['sAnexo'] + '.' + Trim(ActividadesxOrden.FieldValues['sNumeroActividad'])
          else
            ActividadesxOrden.FieldValues['sWbs'] := ActividadesxOrden.FieldValues['sWBSAnterior'] + '.' + ActividadesxOrden.FieldValues['sAnexo'] + '.Sin Partida'
        else
          ActividadesxOrden.FieldValues['sWbs'] := Trim(ActividadesxOrden.FieldValues['sNumeroActividad']);
      end
    end
    else
    begin
      ActividadesxOrden.FieldValues['iNivel'] := 0;
      ActividadesxOrden.FieldValues['iItemOrden'] := 0;
      if not ActividadesxOrden.FieldByName('sNumeroActividad').IsNull then
        ActividadesxOrden.FieldValues['sWbs'] := Trim(ActividadesxOrden.FieldValues['sNumeroActividad'])
      else
        ActividadesxOrden.FieldValues['sWbs'] := 'Sin Partida'
    end
  end
end;


procedure TfrmActividades.tsUnidadEnter(Sender: TObject);
begin
  tsUnidad.Color := global_color_entrada
end;

procedure TfrmActividades.tsUnidadExit(Sender: TObject);
begin
  tsUnidad.Color := global_color_salida;
end;

procedure TfrmActividades.tsUnidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tdPonderado.SetFocus;
end;

procedure TfrmActividades.acumularDiferencia(suma, sMensaje: string);
begin
  RxMDValida.Append;
  RxMDValida.FieldByName('sNumeroActividad').Value := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
  RxMDValida.FieldByName('sWbs').Value := connection.QryBusca.FieldByName('sWbs').AsString;
  RxMDValida.FieldByName('dCantidad').Value := connection.QryBusca.FieldByName('dCantidad').AsString;
  RxMDValida.FieldByName('suma').Value := suma;
  RxMDValida.FieldByName('aMN').Value := connection.QryBusca.FieldByName('aMN').AsString;
  RxMDValida.FieldByName('aDLL').Value := connection.QryBusca.FieldByName('aDLL').AsString;
  RxMDValida.FieldByName('dCantidadAnexo').Value := connection.QryBusca.FieldByName('dCantidadAnexo').AsString;
  RxMDValida.FieldByName('bMN').Value := connection.QryBusca.FieldByName('bMN').AsString;
  RxMDValida.FieldByName('bDLL').Value := connection.QryBusca.FieldByName('bDLL').AsString;
  RxMDValida.FieldByName('descripcion').Value := connection.QryBusca.FieldByName('descripcion').AsString;
  RxMDValida.FieldByName('mensaje').Value := sMensaje;
  RxMDValida.FieldByName('sNumeroOrden').Value := connection.QryBusca.FieldByName('sNumeroOrden').AsString;
  RxMDValida.FieldByName('sWbs2').Value := connection.QryBusca.FieldByName('wbs2').AsString;
  RxMDValida.Post;
end;

function TfrmActividades.cantidadesDiferentes(sActividad: string): string;
var
  sSQL: string;
begin
  result := '';

  sSQL := 'SELECT ' +
    'sum(a.dCantidad) as suma ' +
    'FROM actividadesxorden a ' +
    'INNER JOIN  actividadesxanexo b ' +
    'ON a.sContrato = b.sContrato ' +
    'AND a.sIdConvenio = b.sIdConvenio ' +
    'AND a.sNumeroActividad = b.sNumeroActividad ' +
    'AND a.sTipoActividad = "Actividad" ' +
    'WHERE b.sContrato = :contrato ' +
    'AND b.sIdConvenio = :convenio ' +
    'AND b.sNumeroActividad = :actividad ' +
    'AND b.sTipoActividad = "Actividad"';

  connection.QryBusca.Active := false;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add(sSQL);
  connection.QryBusca.ParamByName('actividad').Value := sActividad;
  connection.QryBusca.ParamByName('contrato').Value := global_contrato;
  connection.QryBusca.ParamByName('convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
    result := connection.QryBusca.FieldByName('suma').AsString
end;

procedure TfrmActividades.ventasDiferentes(sActividad, suma: string);
var
  sSQL: string;
  lError1, lError2: boolean;
begin
  sSQL := 'SELECT ' +
    'b.sNumeroActividad, b.sWbs, a.dCantidad, substr(b.mDescripcion,1,255) as descripcion, ' +
    'a.dVentaMN as aMN, a.dVentaDLL as aDLL, a.sTipoActividad, a.sNumeroOrden, a.sWbs as wbs2, ' +
    'b.dCantidadAnexo,  b.dVentaMN as bMN, b.dVentaDLL as bDLL  ' +
    'FROM actividadesxorden a ' +
    'INNER JOIN  actividadesxanexo b ' +
    'ON a.sContrato = b.sContrato ' +
    'AND a.sIdConvenio = b.sIdConvenio ' +
    'AND a.sNumeroActividad = b.sNumeroActividad ' +
    'AND a.sTipoActividad = "Actividad" ' +
    'WHERE b.sContrato = :contrato ' +
    'AND b.sIdConvenio = :convenio ' +
    'AND b.sNumeroActividad = :actividad ' +
    'AND b.sTipoActividad = "Actividad" ' +
    'ORDER BY b.sNumeroActividad';

  connection.QryBusca.Active := false;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add(sSQL);
  connection.QryBusca.ParamByName('actividad').Value := sActividad;
  connection.QryBusca.ParamByName('contrato').Value := global_contrato;
  connection.QryBusca.ParamByName('convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
  connection.QryBusca.Open;

  lError1 := false;
  lError2 := false;
  while not connection.QryBusca.Eof do begin
    if (connection.QryBusca.FieldByName('aMN').Value <>
      connection.QryBusca.FieldByName('bMN').Value)
      or (connection.QryBusca.FieldByName('aDLL').Value <>
      connection.QryBusca.FieldByName('bDLL').Value) then begin
      acumularDiferencia(suma, 'Existe diferencia entre los valores de ventas');
      lError1 := true;
    end
    else begin
      if (not lError1) and (not lError2) then begin
        if (connection.QryBusca.FieldByName('dCantidadAnexo').Value <> suma) then
          lError2 := true;
      end;
    end;
    connection.QryBusca.Next;
  end;
  if (not lError1) and (lError2) then begin
    connection.QryBusca.First;
    acumularDiferencia(suma, 'Existe diferencia entre la suma total de las cantidades y la cantidad del anexo');
  end;
end;

procedure TfrmActividades.CalcDiferenciasOT(lista: TStringList);
var
  ii: integer;
begin
  RxMDValida.Active := True;
  if RxMDValida.RecordCount > 0 then
    RxMDValida.EmptyTable;
  for ii := 0 to Lista.Count - 1 do begin
    ventasDiferentes(Lista.Strings[ii], cantidadesDiferentes(Lista.Strings[ii]));
  end;
end;

procedure TfrmActividades.grid_actividadesEnter(Sender: TObject);
begin
  if ActividadesxOrden.active then
  begin
    if ActividadesxOrden.state in [dsinsert, dsedit] then
    begin
    //If ActividadesxOrden.FieldByName('sNumeroActividad').IsNull  then
    //begin
      ActividadesxOrden.Cancel;
      frmBarra1.btnCancel.Click
    //end ;
    end;
  end;
end;

procedure TfrmActividades.Imprimir1Click(Sender: TObject);
begin
  frmBarra1.btnPrinter.Click
end;

procedure TfrmActividades.tsPlataformaEnter(Sender: TObject);
begin
    tsPlataforma.Color := global_color_entrada;
end;

procedure TfrmActividades.tsPlataformaExit(Sender: TObject);
begin
    tsPlataforma.Color := global_color_salida;
end;

procedure TfrmActividades.tsPlataformaKeyPress(Sender: TObject; var Key: Char);
begin
  if key =#13 then
     tmDescripcion.SetFocus;
end;

procedure TfrmActividades.tsReprogramacionEnter(Sender: TObject);
begin
    tsReprogramacion.Color := global_color_entrada;
end;

procedure TfrmActividades.tsReprogramacionExit(Sender: TObject);
begin
//    if zqReprogramacion.RecordCount > 0 then
//    begin
//        sNumeroOrden := tsNumeroOrden.Text;
//        IsOpen:=false;
//        ConsultaFolios;
//    end;
    tsReprogramacion.Color := global_color_salida
end;

procedure TfrmActividades.tsReprogramacionKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key = #13 then
       grid_actividades.SetFocus
end;

procedure TfrmActividades.tsTipoAnexoKeyPress(Sender: TObject; var Key: Char);
begin
   if KEY=#13 then
      tlCalculo.SetFocus;
end;

procedure TfrmActividades.grid_actividadesCellClick(Column: TColumn);
begin
  if actividadesxorden.RecordCount > 0 then
    grid_actividades.Hint := actividadesxorden.FieldValues['sWbs'];
end;

procedure TfrmActividades.grid_actividadesDblClick(
  Sender: TObject);
var
  sCondicion: string;
  sSelect: string;
  inicio, reg: Integer;
  Lugar: Tbookmark;
begin

  if (ActividadesxOrden.FieldValues['sWbs'] <> NULL) then
  begin
    sCondicion := 'sWbs not Like ' + quotedstr(Trim(ActividadesxOrden.FieldValues['sWbs']) + '.*');

    if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
    begin

      lugar := ActividadesxOrden.GetBookmark;

      reg := Paq.indexof(ActividadesxOrdenswbs.asstring);

      if reg = -1 then
      begin
        if Pos(sCondicion, sFiltro) = 0 then
          if sFiltro <> '' then
            sFiltro := sFiltro + ' and ' + scondicion
          else
            sFiltro := sCondicion;

        Paq.Add(ActividadesxOrdenswbs.asstring);

      end
      else
      begin
        inicio := Pos(sCondicion, sFiltro);

        Paq.Delete(reg);
        if (inicio > 0) then
        begin
          if inicio = 1 then
            sFiltro := MidStr(sFiltro, Length(scondicion) + 6, Length(sFiltro))
          else
            sFiltro := MidStr(sFiltro, 1, inicio - 6) + MidStr(sFiltro, inicio + Length(scondicion), length(sfiltro));

        end;

      end;

      ActividadesxOrden.Filtered := false;
      ActividadesxOrden.Filter := sfiltro;
      ActividadesxOrden.Filtered := true;
      try
        ActividadesxOrden.GotoBookmark(lugar);
      finally
        ActividadesxOrden.FreeBookmark(lugar);
      end;


    end;
  end;
end;

procedure TfrmActividades.btnImportarClick(Sender: TObject);
var
  registro: Integer;
  sParametro: string;
  lContinua: Boolean;
begin
  try
    if GridProgConstExist.DataSource.DataSet.FieldValues['sNumeroOrden'] <> tsNumeroOrden.Text then
    begin
      if ActividadesxOrden.RecordCount > 0 then
      begin
            // Verifico que no existan partidas reportadas, si ya se ha reportado algo del programa, se cancela toda la operacion
        Connection.qryBusca.Active := False;
        Connection.qryBusca.Filtered := False;
        Connection.qryBusca.SQL.Clear;
        Connection.qryBusca.SQL.Add('Select a2.sContrato From bitacoradeactividades b ' +
          'INNER JOIN actividadesxorden a2 ON (a2.sContrato = b.sContrato And a2.sNumeroOrden = b.sNumeroOrden And a2.sWbs = b.sWbs And ' +
          'a2.sNumeroActividad = b.sNumeroActividad And a2.sTipoActividad = "Actividad" ' +
          'Where a2.sContrato = :Contrato And a2.sNumeroOrden = :Orden');
        Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
        Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
        Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
        Connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
        Connection.qryBusca.Open;
        if Connection.QryBusca.RecordCount > 0 then
        begin
          MessageDlg('Existen partidas dentro del programa de trabajo seleccionado, no se puede insertar el programa de trabajo seleccionado.', mtInformation, [mbOk], 0);
          lContinua := False
        end
        else
          lContinua := True
      end
      else
        lContinua := True;

      if lContinua then
      begin
        Connection.qryBusca.Active := False;
        Connection.qryBusca.Filtered := False;
        Connection.qryBusca.SQL.Clear;
        Connection.qryBusca.SQL.Add('Select * from actividadesxorden Where sContrato = :Contrato And sIdConvenio = :convenio And sNumeroOrden = :orden');
        Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
        Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
        Connection.qryBusca.Params.ParamByName('convenio').DataType := ftString;
        Connection.qryBusca.Params.ParamByName('convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        Connection.qryBusca.Params.ParamByName('orden').DataType := ftString;
        Connection.qryBusca.Params.ParamByName('orden').Value := GridProgConstExist.DataSource.DataSet.FieldValues['sNumeroOrden'];
        Connection.qryBusca.Open;
        if Connection.qryBusca.RecordCount > 0 then
        begin
          Connection.zCommand.Active := False;
          Connection.zCommand.Filtered := False;
          Connection.zCommand.SQL.Clear;
          Connection.zCommand.SQL.Add('DELETE FROM actividadesxorden Where sContrato = :contrato and sNumeroOrden = :orden and sIdConvenio = :convenio');
          connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
          connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
          connection.zCommand.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          connection.zCommand.Params.ParamByName('orden').DataType := ftString;
          connection.zCommand.Params.ParamByName('orden').Value := tsNumeroOrden.Text;
          connection.zCommand.ExecSQL();
                // Empiezo exportando los anexos ...
          Connection.qryBusca.First;
          while not Connection.qryBusca.Eof do
          begin
            Connection.zCommand.Active := False;
            Connection.zCommand.Filtered := False;
            Connection.zCommand.SQL.Clear;
            Connection.zCommand.SQL.Add(funcsql(Connection.qryBusca, 'actividadesxorden'));
            for registro := 0 to Connection.qryBusca.fieldcount - 1 do
            begin
              sparametro := 'param' + trim(inttostr(registro + 1));
              connection.zCommand.Params.parambyname(sparametro).datatype := Connection.qryBusca.fields[registro].datatype;
              if Connection.qryBusca.fields[registro].DisplayName = 'sNumeroOrden' then
                connection.zCommand.Params.parambyname(sparametro).value := tsNumeroOrden.Text
              else
                connection.zCommand.Params.parambyname(sparametro).value := Connection.qryBusca.fields[registro].value;
            end;
            connection.zCommand.ExecSQL;
            Connection.qryBusca.Next
          end;
          frmBarra1.btnRefresh.Click
        end;
      end
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Conceptos/Partidas x Frente de Trabajo', 'Al importar', 0);
    end;
  end;
end;

procedure TfrmActividades.CmbFolioChange(Sender: TObject);
var zFls:TZReadOnlyQuery;
begin
  try
    zFls:=TZReadOnlyQuery.Create(nil);
    try
      zFls.Connection := connection.zConnection;
      zFls.Active := False;
      zFls.SQL.Text := 'select min(didfecha) as minimo, max(didfecha) as maximo from bitacoradeactividades where sContrato = :Contrato and (:folio = -1 or (:folio <> -1 and sNumeroOrden = :folio))';
      zFls.ParamByName('Contrato').AsString := global_contrato;
      if CmbFolio.ItemIndex = 0 then
        zFls.ParamByName('Folio').AsString := '-1'
      else
        zFls.ParamByName('Folio').AsString := CmbFolio.Text;
      zFls.Open;
      DFechaInicio.Date := zFls.FieldByName('minimo').AsDateTime;
      DFechaFin.Date := zFls.FieldByName('maximo').AsDateTime;
    finally
      zFls.Free;
    end;
  except
    ;
  end;
end;

procedure TfrmActividades.Com1Click(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// GENERACION DE PROGRAMA DE TRABAJO CON VOLUMENES //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena: string;
      fs: tStream;
      Alto: Extended;
      MiFechaI, MiFechaF, MiFecha: tDate;
      Ren, nivel, i, total: integer;
      Q_Partidas: TZReadOnlyQuery;
      dVolumen, dReal, dCantidad: double;
      Progreso, TotalProgreso: real;
      sWbs: string;
      dFechaInicial1, dFechaFinal1, dFechaInicial2, dFechaFinal2: TDate;
    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      Ren := 2;
    // Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 100;

      Excel.Columns['A:A'].ColumnWidth := 20;
      Excel.Columns['B:B'].ColumnWidth := 18;
      Excel.Columns['C:C'].ColumnWidth := 18;
      Excel.Columns['D:D'].ColumnWidth := 10;
      Excel.Columns['E:E'].ColumnWidth := 40;
      Excel.Columns['F:G'].ColumnWidth := 12;
      Excel.Columns['H:J'].ColumnWidth := 18;

      // Colocar los encabezados de la plantilla...
      Hoja.Range['A1:A1'].Select;
      Excel.Selection.Value := 'Contrato';
      FormatoEncabezado;
      Hoja.Range['B1:B1'].Select;
      Excel.Selection.Value := 'Frente';
      FormatoEncabezado;
      Hoja.Range['C1:C1'].Select;
      Excel.Selection.Value := 'Nivel';
      FormatoEncabezado;
      Hoja.Range['D1:D1'].Select;
      Excel.Selection.Value := 'Actividad';
      FormatoEncabezado;
      Hoja.Range['E1:E1'].Select;
      Excel.Selection.Value := 'Descripcion';
      FormatoEncabezado;
      Hoja.Range['F1:F1'].Select;
      Excel.Selection.Value := 'Medida';
      FormatoEncabezado;
      Hoja.Range['G1:G1'].Select;
      Excel.Selection.Value := 'Cantidad';
      FormatoEncabezado;
      Hoja.Range['H1:H1'].Select;
      Excel.Selection.Value := 'Ponderado';
      FormatoEncabezado;
      Hoja.Range['I1:I1'].Select;
      Excel.Selection.Value := 'Fecha I.';
      FormatoEncabezado;
      Hoja.Range['J1:J1'].Select;
      Excel.Selection.Value := 'Fecha F.';
      FormatoEncabezado;
      with Connection do
      begin
      {obtener las fechas iniciales y finales programas por parte de la cia}
        QryBusca.Active := false;
        QryBusca.SQL.Clear;
        QryBusca.SQL.Add('select min(dIdFecha) as dFechaInicialProg from distribuciondeactividadescia b ' +
          ' inner join  actividadesxorden  a ' +
          '   on b.sContrato=a.sContrato and b.sNumeroOrden=a.sNumeroOrden and b.sWbs=a.sWbs and b.sNumeroActividad=a.sNumeroActividad  ' +
          ' where a.sContrato=:contrato and a.sNumeroOrden=:orden and a.sIdConvenio=:convenio  ');
        QryBusca.ParamByName('contrato').AsString := global_contrato;
        QryBusca.ParamByName('orden').AsString := tsNumeroOrden.text;
        QryBusca.ParamByName('convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        QryBusca.Open;
        if QryBusca.RecordCount > 0 then
        begin
          dFechaInicial1 := QryBusca.FieldByName('dFechaInicialProg').AsDateTime;
        end;
        QryBusca.Active := false;
        QryBusca.SQL.Clear;
        QryBusca.SQL.Add('select max(dIdFecha) as dFechaFinalProg from distribuciondeactividadescia b ' +
          ' inner join  actividadesxorden  a  ' +
          '   on b.sContrato=a.sContrato and b.sNumeroOrden=a.sNumeroOrden and b.sWbs=a.sWbs and b.sNumeroActividad=a.sNumeroActividad  ' +
          ' where a.sContrato=:contrato and a.sNumeroOrden=:orden and a.sIdConvenio=:convenio  ');
        QryBusca.ParamByName('contrato').AsString := global_contrato;
        QryBusca.ParamByName('orden').AsString := tsNumeroOrden.text;
        QryBusca.ParamByName('convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        QryBusca.Open;
        if QryBusca.RecordCount > 0 then
        begin
          dFechaFinal1 := QryBusca.FieldByName('dFechaFinalProg').AsDateTime;
        end;
       {Obtener las fechas iniciales y finales de las partidas que estan programadas(cia) y reportadas }
        QryBusca.Active := false;
        QryBusca.SQL.Clear;
        QryBusca.SQL.Add('select min(a.dIdFecha) as dFechaInicialProg from distribuciondeactividadescia b ' +
          ' inner join  bitacoradeactividades  a ' +
          '   on b.sContrato=a.sContrato and b.sNumeroOrden=a.sNumeroOrden and b.sWbs=a.sWbs and b.sNumeroActividad=a.sNumeroActividad  ' +
          ' where b.sContrato=:contrato and b.sNumeroOrden=:orden and b.sIdConvenio=:convenio  ');
        QryBusca.ParamByName('contrato').AsString := global_contrato;
        QryBusca.ParamByName('orden').AsString := tsNumeroOrden.text;
        QryBusca.ParamByName('convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        QryBusca.Open;
        if QryBusca.RecordCount > 0 then
        begin
          dFechaInicial2 := QryBusca.FieldByName('dFechaInicialProg').AsDateTime;
        end;
        QryBusca.Active := false;
        QryBusca.SQL.Clear;
        QryBusca.SQL.Add('select max(a.dIdFecha) as dFechaFinalProg from distribuciondeactividadescia b ' +
          ' inner join  bitacoradeactividades  a ' +
          '   on b.sContrato=a.sContrato and b.sNumeroOrden=a.sNumeroOrden and b.sWbs=a.sWbs and b.sNumeroActividad=a.sNumeroActividad  ' +
          ' where b.sContrato=:contrato and b.sNumeroOrden=:orden and b.sIdConvenio=:convenio  ');
        QryBusca.ParamByName('contrato').AsString := global_contrato;
        QryBusca.ParamByName('orden').AsString := tsNumeroOrden.text;
        QryBusca.ParamByName('convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        QryBusca.Open;
        if QryBusca.RecordCount > 0 then
        begin
          dFechaFinal2 := QryBusca.FieldByName('dFechaFinalProg').AsDateTime;
        end;

        {Compara scual de las fechas iniciales es mas vieja..}
        if dFechaInicial1 > dFechaInicial2 then dFechaInicial1 := dFechaInicial2;
        {Compara scual de las fechas finales es mas reciente..}
        if dFechaFinal1 < dFechaFinal2 then dFechaFinal1 := dFechaFinal2;
        
        {Ahora obtener la fecha de cuando se genera}
        QryBusca.Active := false;
        QryBusca.SQL.Clear;
        QryBusca.SQL.Add(' select max(a.dFecha) as dFechaFinalProg from distribuciondeactividadescia b ' +
          ' inner join  estimacionxpartida  a ' +
          '   on b.sContrato=a.sContrato and b.sNumeroOrden=a.sNumeroOrden and b.sWbs=a.sWbs and b.sNumeroActividad=a.sNumeroActividad  ' +
          ' where b.sContrato=:contrato and b.sNumeroOrden=:orden and b.sIdConvenio=:convenio  ');
        QryBusca.ParamByName('contrato').AsString := global_contrato;
        QryBusca.ParamByName('orden').AsString := tsNumeroOrden.text;
        QryBusca.ParamByName('convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        QryBusca.Open;
        if QryBusca.RecordCount > 0 then
        begin
          dFechaFinal2 := QryBusca.FieldByName('dFechaFinalProg').AsDateTime;
        end;
        {Compara scual de las fechas finales es mas reciente..}
        if dFechaFinal1 < dFechaFinal2 then dFechaFinal1 := dFechaFinal2;
                     
        {Crear los encabezados de las fechas en base a los fechas inicial y final obtenidos}
        i := 1;
        dFechaInicial2 := dFechaInicial1;
        while (dFechaInicial2 <= dFechaFinal1) do
        begin
          Hoja.Cells[Ren - 1, 10 + i].Select;
                 {Formato de las fechas archivo Excel,, 24/07/2011..}
          Excel.Selection.NumberFormat := '@';
          Excel.Selection.Value := DateToStr(dFechaInicial2);
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Interior.ColorIndex := 49;
          dFechaInicial2 := IncDay(dFechaInicial2);
          Inc(i);
        end;
        Hoja.Cells[Ren - 1, 10 + i].Select;
        Excel.Selection.Value := 'TOTAL';
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.VerticalAlignment := xlCenter;
        Excel.Selection.Font.Color := clWhite;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Interior.ColorIndex := 11;

        {Ahora leer todas las actividades programadas y en base a estas, obtener la informacion requerida (reportadas, programadas)}
        QryBusca.Active := false;
        QryBusca.SQL.Clear;
        QryBusca.SQL.Add(' select  a.sWbs,a.iNivel,a.sMedida,a.sNumeroActividad, a.mDescripcion ,' +
          '     b.dIdFecha,a.sNumeroOrden, a.dCantidad,a.sTipoActividad,a.dPonderado,a.dFechaInicio,dFechaFinal, a.dCantidad as dReal,' +
          '    b.dCantidad as dProgramado ' +
          ' from distribuciondeactividadescia b ' +
          ' inner join  actividadesxorden  a  ' +
          '   on b.sContrato=a.sContrato and b.sNumeroOrden=a.sNumeroOrden and b.sWbs=a.sWbs and b.sNumeroActividad=a.sNumeroActividad  ' +
          ' where a.sContrato=:contrato and a.sNumeroOrden=:orden and a.sIdConvenio=:convenio  ' +
          ' group by a.sWbs  ' +
          ' order by a.iItemOrden;');
        QryBusca.ParamByName('contrato').AsString := global_contrato;
        QryBusca.ParamByName('orden').AsString := tsNumeroOrden.Text;
        QryBusca.ParamByName('convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        QryBusca.Open;

        sWbs := '';
        Inc(Ren);
        while not QryBusca.EOF do
        begin
          {Descripcion del concepto}
          if sWbs <> QryBusca.FieldValues['sWbs'] then
          begin
            sWbs := QryBusca.FieldValues['sWbs'];
                {Escritura de Datos en el Archvio de Excel..}
            Hoja.Cells[Ren, 1].Select;
            Excel.Selection.Value := global_contrato;
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment := xlCenter;
            Excel.Selection.Font.Size := 11;
            Excel.Selection.Font.Bold := False;
            Excel.Selection.Font.Name := 'Calibri';

            Hoja.Cells[Ren, 2].Select;
            Excel.Selection.Value := QryBusca.FieldValues['sNumeroOrden'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment := xlCenter;

            Hoja.Cells[Ren, 3].Select;
            Excel.Selection.Value := QryBusca.FieldValues['iNivel'];

            Hoja.Cells[Ren, 4].Select;
            Excel.Selection.Value := QryBusca.FieldValues['sNumeroActividad'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment := xlCenter;

            Hoja.Cells[Ren, 5].Select;
            Excel.Selection.Value := QryBusca.FieldValues['mDescripcion'];
            Alto := Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight;
            Hoja.Cells[Ren, 5].Value := '';

            if Alto > 15 then
              Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := Alto
            else
              Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := 15;

            Excel.Selection.Value := QryBusca.FieldValues['mDescripcion'];

            Hoja.Cells[Ren, 6].Select;
            Excel.Selection.Value := QryBusca.FieldValues['sMedida'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment := xlCenter;

            Hoja.Cells[Ren, 7].Select;
            Excel.Selection.NumberFormat := '@';
            Excel.Selection.Value := QryBusca.FieldValues['dCantidad'];
            Excel.Selection.HorizontalAlignment := xlRight;
            Excel.Selection.VerticalAlignment := xlCenter;

            Hoja.Cells[Ren, 8].Select;
            Excel.Selection.NumberFormat := '@';
            Excel.Selection.Value := QryBusca.FieldValues['dPonderado'];
            Excel.Selection.HorizontalAlignment := xlRight;
            Excel.Selection.VerticalAlignment := xlCenter;

            Hoja.Cells[Ren, 9].Select;
            Excel.Selection.Value := QryBusca.FieldValues['dFechaInicio'];

            Hoja.Cells[Ren, 10].Select;
            Excel.Selection.Value := QryBusca.FieldValues['dFechaFinal'];
            Inc(Ren);
          end;
          {Programado}
          Inc(Ren);
          i := 1;

          Hoja.Cells[Ren - 1, 9 + i].Select;
          Excel.Selection.Value := 'PROGRAMADO';
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Color := clBlack;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Interior.ColorIndex := 2;
                    
          dFechaInicial2 := dFechaInicial1;
          dCantidad := 0;
          while (dFechaInicial2 <= dFechaFinal1) do
          begin

            QryBusca2.Active := false;
            QryBusca2.SQL.Clear;
            QryBusca2.SQL.Add('select if(dCantidad is null,0,dCantidad) as dCantidad from distribuciondeactividadescia where ' +
              ' sContrato=:contrato and sIdConvenio =:convenio and sNumeroOrden = :orden ' +
              ' and sNumeroActividad=:actividad and sWbs =:wbs and dIdFecha=:fecha ');
            QryBusca2.ParamByName('contrato').AsString := global_contrato;
            QryBusca2.ParamByName('convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            QryBusca2.ParamByName('orden').AsString := QryBusca.FieldByName('sNumeroOrden').AsString;
            QryBusca2.ParamByName('actividad').AsString := QryBusca.FieldByName('sNumeroActividad').AsString;
            QryBusca2.ParamByName('wbs').AsString := QryBusca.FieldByName('sWbs').AsString;
            QryBusca2.ParamByName('fecha').AsDate := dFechaInicial2;
            QryBusca2.Open;

            Hoja.Cells[Ren - 1, 10 + i].Select;
            Excel.Selection.Value := QryBusca2.FieldByName('dCantidad').AsFloat;
            Excel.Selection.HorizontalAlignment := xlRight;
            Excel.Selection.VerticalAlignment := xlCenter;
            Excel.Selection.Font.Color := clWhite;
            Excel.Selection.Font.Bold := True;
            Excel.Selection.Interior.ColorIndex := 49;
            dCantidad := dCantidad + QryBusca2.FieldByName('dCantidad').AsFloat;
            dFechaInicial2 := IncDay(dFechaInicial2);
            Inc(i);
          end;
          Hoja.Cells[Ren - 1, 10 + i].Select;
          Excel.Selection.Value := dCantidad;
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Interior.ColorIndex := 11;
          {Reportado}
          dCantidad := 0;
          Inc(Ren);
          i := 1;

          Hoja.Cells[Ren - 1, 9 + i].Select;
          Excel.Selection.Value := 'REPORTADO';
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Color := clBlack;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Interior.ColorIndex := 2;


          dFechaInicial2 := dFechaInicial1;
          while (dFechaInicial2 <= dFechaFinal1) do
          begin

            QryBusca2.Active := false;
            QryBusca2.SQL.Clear;
            QryBusca2.SQL.Add('select if(sum(dCantidad) is null,0,sum(dCantidad)) as dCantidad from bitacoradeactividades where ' +
              ' sContrato=:contrato and sNumeroOrden = :orden ' +
              ' and sNumeroActividad=:actividad and sWbs =:wbs and dIdFecha=:fecha ' +
              ' group by dIdFecha ');
            QryBusca2.ParamByName('contrato').AsString := global_contrato;
            QryBusca2.ParamByName('orden').AsString := tsNumeroOrden.Text;
            QryBusca2.ParamByName('actividad').AsString := QryBusca.FieldByName('sNumeroActividad').AsString;
            QryBusca2.ParamByName('wbs').AsString := QryBusca.FieldByName('sWbs').AsString;
            QryBusca2.ParamByName('fecha').AsDate := dFechaInicial2;
            QryBusca2.Open;

            Hoja.Cells[Ren - 1, 10 + i].Select;
            Excel.Selection.Value := QryBusca2.FieldByName('dCantidad').AsFloat;
            Excel.Selection.HorizontalAlignment := xlRight;
            Excel.Selection.VerticalAlignment := xlCenter;
            Excel.Selection.Font.Color := clWhite;
            Excel.Selection.Font.Bold := True;
            Excel.Selection.Interior.ColorIndex := 29;
            dCantidad := dCantidad + QryBusca2.FieldByName('dCantidad').AsFloat;
            dFechaInicial2 := IncDay(dFechaInicial2);
            Inc(i);
          end;
          Hoja.Cells[Ren - 1, 10 + i].Select;
          Excel.Selection.Value := dCantidad;
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Interior.ColorIndex := 11;
          {GENERADO}
          dCantidad := 0;
          Inc(Ren);
          i := 1;

          Hoja.Cells[Ren - 1, 9 + i].Select;
          Excel.Selection.Value := 'GENERADO';
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Color := clBlack;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Interior.ColorIndex := 2;


          dFechaInicial2 := dFechaInicial1;
          while (dFechaInicial2 <= dFechaFinal1) do
          begin

            QryBusca2.Active := false;
            QryBusca2.SQL.Clear;
            QryBusca2.SQL.Add(' select if(sum(dCantidad) is null,0,sum(dCantidad)) as dCantidad from estimacionxpartida where ' +
              ' sContrato=:contrato and sNumeroOrden = :orden ' +
              ' and sNumeroActividad=:actividad and sWbs =:wbs and dFecha=:fecha ' +
              ' group by dFecha ');
            QryBusca2.ParamByName('contrato').AsString := global_contrato;
            QryBusca2.ParamByName('orden').AsString := tsNumeroOrden.Text;
            QryBusca2.ParamByName('actividad').AsString := QryBusca.FieldByName('sNumeroActividad').AsString;
            QryBusca2.ParamByName('wbs').AsString := QryBusca.FieldByName('sWbs').AsString;
            QryBusca2.ParamByName('fecha').AsDate := dFechaInicial2;
            QryBusca2.Open;

            Hoja.Cells[Ren - 1, 10 + i].Select;
            Excel.Selection.Value := QryBusca2.FieldByName('dCantidad').AsFloat;
            Excel.Selection.HorizontalAlignment := xlRight;
            Excel.Selection.VerticalAlignment := xlCenter;
            Excel.Selection.Font.Color := clWhite;
            Excel.Selection.Font.Bold := True;
            Excel.Selection.Interior.ColorIndex := 52;
            dCantidad := dCantidad + QryBusca2.FieldByName('dCantidad').AsFloat;
            dFechaInicial2 := IncDay(dFechaInicial2);
            Inc(i);
          end;
          Hoja.Cells[Ren - 1, 10 + i].Select;
          Excel.Selection.Value := dCantidad;
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Interior.ColorIndex := 11;          
          QryBusca.Next();
        end;
      end;
     Hoja.Cells[2, 2].Select;
    end;


  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'ProgVsRealVsGenOrden'; // + tsNumeroOrden.Text;
      except
        Hoja.Name := 'ProgVsRealVsGenOrden';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Programa de Trabajo:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
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
    showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
    Exit;
  end;

  if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    Excel.Visible := True;
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := True;
  end
  else
  begin
    Excel.Visible := True;
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := False;
  end;

  PanelProgress.Visible := True;
  Label15.Refresh;
  Label16.Refresh;
  Label17.Refresh;
  BarraEstado.Position := 0;

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
  begin
        // Grabar el archivo de excel con el nombre dado
    Excel.Visible := True;
    Excel.DisplayAlerts := True;
    Excel.ScreenUpdating := True;
    PanelProgress.Visible := False;
    messageDlg('El Archivo se generó Correctamente!', mtInformation, [mbOk], 0);
  end;

  Excel := '';

  if CadError <> '' then
    showmessage(CadError);

end;

procedure TfrmActividades.PonderarConceptosClick(Sender: TObject);
begin
    PonderadoAnterior;
end;

procedure TfrmActividades.AgregarObservaciones1Click(Sender: TObject);
begin
  application.CreateForm(TFrmNotaCampoObservaciones,FrmNotaCampoObservaciones);
  try
    FrmNotaCampoObservaciones.ParamContrato:=Global_Contrato;
    FrmNotaCampoObservaciones.ParamFolio:=tsNumeroOrden.KeyValue;
    FrmNotaCampoObservaciones.ShowModal;
  finally

    FrmNotaCampoObservaciones.Destroy;
  end;
end;

procedure TfrmActividades.AVANCESGLOBALESXFOLIOPARTIDAS1Click(Sender: TObject);
var zFls:TZReadOnlyQuery;
begin
  CmbFolio.Items.Clear;
  CmbFolio.Items.Add('Todos');
  zFls:=TZReadOnlyQuery.Create(nil);
  try
    zFls.Connection := connection.zConnection;
    zFls.Active := False;
    zFls.SQL.Text := 'select distinct(snumeroorden) from ordenesdetrabajo where scontrato = :Contrato order by snumeroorden';
    zFls.ParamByName('Contrato').AsString := global_contrato;
    zFls.Open;
    zFls.First;
    while not zFls.Eof do
    begin
      CmbFolio.Items.Add(zFls.FieldByName('snumeroorden').AsString);
      zFls.Next;
    end;
  finally
    zFls.Free;
  end;
  CmbFolio.ItemIndex := 0;
  CmbFolioChange(CmbFolio);
  PnlRango.Visible := True;
end;

procedure TfrmActividades.tdCostoMNChange(Sender: TObject);
begin
  TRxDBCalcEditChangef(tdCostoMN, '$ Costo M.N.');
end;

procedure TfrmActividades.tdCostoMNEnter(Sender: TObject);
begin
  tdCostoMN.Color := global_color_entrada
end;

procedure TfrmActividades.tdCostoMNExit(Sender: TObject);
begin
  tdCostoMN.Color := global_color_salida
end;

procedure TfrmActividades.tdCostoMNKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxDBCalcEdit(tdCostoMN, key) then
    key := #0;
  if Key = #13 then
    tsTipoAnexo.SetFocus
end;

procedure TfrmActividades.ImportaXLS(Sender: TObject);
var
  sArchivo: string;
  flcid, Fila: Integer;
  zExcel: tExcelApplication;
  zExcelLibro: tExcelWorkbook;
  zExcelHoja: tExcelWorksheet;

  sValue,
    ImpsNumeroActividad,
    ImpdCantidad,
    ImpdFechaInicio,
    ImpdFechaFinal: string;
begin
  try
    with tOpenDialog.Create(Self) do
    begin
      Title := 'Inserta Archivo de Consulta';
      if Execute then
        sArchivo := FileName
    end;

    if sArchivo <> '' then
    begin
      if ActividadesxOrden.RecordCount > 0 then
      begin
        if MessageDlg('Desea Introducir las partidas seleccionadas dentro del paquete de la Orden de Trabajo?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
          begin
            sWbsAnterior := ActividadesxOrden.FieldValues['sWbs'];
            sItemOrden := MidStr(ActividadesxOrden.FieldValues['iItemOrden'], 1, (ActividadesxOrden.FieldValues['iNivel'] + 1) * LongNivel);
            iNivel := ActividadesxOrden.FieldValues['iNivel'] + 1
          end
          else
          begin
            sWbsAnterior := ActividadesxOrden.FieldValues['sWbsAnterior'];
            sItemOrden := MidStr(ActividadesxOrden.FieldValues['iItemOrden'], 1, ActividadesxOrden.FieldValues['iNivel'] * LongNivel);
            iNivel := ActividadesxOrden.FieldValues['iNivel']
          end
        end
        else
        begin
          sWbsAnterior := '';
          sItemOrden := '';
        end
      end
      else
      begin
        sWbsAnterior := '';
        sItemOrden := '';
      end;

      flcid := GetUserDefaultLCID;
      zExcel := tExcelApplication.Create(Self);
      zExcel.Connect;
      zExcel.Visible[flcid] := true;
      zExcel.UserControl := true;
      try
        zExcelLibro := tExcelWorkbook.Create(Self);
        zExcelLibro.ConnectTo(zExcel.Workbooks.Open(sArchivo, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam,
          emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, flcid));

        zExcelHoja := tExcelWorkSheet.Create(Self);
        zExcelHoja.ConnectTo(zExcelLibro.Sheets.Item[1] as ExcelWorkSheet);
      finally
        Fila := 2;
        sValue := zExcelHoja.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
        while (sValue <> '') do
        begin
                // Verifico si el Contrato y la Orden Son Iguales ...
          if (zExcelHoja.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2 = global_contrato) and
          (zExcelHoja.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2 = tsNumeroOrden.Text) then
          begin
            ImpdFechaInicio := '';
            ImpdFechaFinal := '';
            ImpsNumeroActividad := zExcelHoja.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
            ImpdCantidad := zExcelHoja.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;
            if zExcelHoja.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2 <> '' then
              ImpdFechaInicio := DateToStr(zExcelHoja.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2);
            if zExcelHoja.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2 <> '' then
              ImpdFechaFinal := DateToStr(zExcelHoja.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2);

                    //Busco la partida en el anexo del contrato ...

            connection.QryBusca.Active := False;
            Connection.qryBusca.Filtered := False;
            connection.QryBusca.SQL.Clear;
            connection.QryBusca.SQL.Add('select sWbs, sNumeroActividad, sTipoActividad, sActividadAnterior from actividadesxanexo ' +
              'Where sContrato = :Contrato and sIdConvenio=:convenio and sNumeroActividad = :Actividad And sTipoActividad = "Actividad"');
            connection.QryBusca.Params.ParamByName('contrato').DataType := ftString;
            connection.QryBusca.Params.ParamByName('contrato').Value := global_contrato;
            connection.QryBusca.Params.ParamByName('convenio').DataType := ftString;
            connection.QryBusca.Params.ParamByName('convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.QryBusca.Params.ParamByName('Actividad').DataType := ftString;
            connection.QryBusca.Params.ParamByName('Actividad').Value := ReplaceStr(Trim(ImpsNumeroActividad), ' ', '');
            connection.QryBusca.Open;
                    // Si es ACtividad de Anexo se Registra la Partida ...
            if Connection.qryBusca.RecordCount > 0 then
              ProcIntroduceRegistro(sWbsAnterior, sItemOrden, global_contrato, zqReprogramacion.FieldByName('sIdConvenio').AsString, tsNumeroOrden.Text, connection.QryBusca.FieldValues['sWbs'], connection.QryBusca.FieldValues['sNumeroActividad'], connection.QryBusca.FieldValues['sTipoActividad'], iNivel, ImpdCantidad, ImpdFechaInicio, ImpdFechaFinal, connection.QryBusca.FieldValues['sActividadAnterior'], '')
          end;
          Fila := Fila + 1;
          sValue := zExcelHoja.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
        end;

            // Modifico la fecha de los paquetes superiores ..
        if (Connection.Configuracion.FieldValues['lCalculaFecha'] = 'Si') then
          ProcRegeneraMontos(global_contrato, zqReprogramacion.FieldByName('sIdConvenio').AsString, tsNumeroOrden.Text, sWbsAnterior);
        MessageDlg('Proceso Terminado.', mtInformation, [mbOk], 0);
      end;

      SavePlace := grid_actividades.DataSource.DataSet.GetBookmark;
      ActividadesxOrden.Refresh;
      try
        grid_actividades.DataSource.DataSet.GotoBookmark(SavePlace);
      except
        grid_actividades.DataSource.DataSet.FreeBookmark(SavePlace);
      end;

        // Cierro Todo
      zExcel.Quit;
      zExcel.Disconnect;
      zExcel.Destroy;
      zExcelLibro.Destroy;
      zExcelHoja.Destroy;
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Conceptos/Partidas x Frente de Trabajo', 'Al Importar XLS', 0);
    end;
  end;
end;

procedure TfrmActividades.InsertaActividad(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: Integer;
begin
  try
    if ActividadesxOrden.RecordCount > 0 then
    begin
      if not lYaPregunto then
        if MessageDlg('Desea Introducir las partidas seleccionadas dentro del paquete de la Orden de Trabajo?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          lYaPregunto := True;
          if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
          begin
            sWbsAnterior := ActividadesxOrden.FieldValues['sWbs'];
            sItemOrden := MidStr(ActividadesxOrden.FieldValues['iItemOrden'], 1, (ActividadesxOrden.FieldValues['iNivel'] + 1) * LongNivel);
            iNivel := ActividadesxOrden.FieldValues['iNivel'] + 1
          end
          else
          begin
            sWbsAnterior := ActividadesxOrden.FieldValues['sWbsAnterior'];
            sItemOrden := MidStr(ActividadesxOrden.FieldValues['iItemOrden'], 1, ActividadesxOrden.FieldValues['iNivel'] * LongNivel);
            iNivel := ActividadesxOrden.FieldValues['iNivel']
          end
        end
        else
        begin
          sWbsAnterior := '';
          sItemOrden := '';
        end
    end
    else
    begin
      sWbsAnterior := '';
      sItemOrden := '';
    end;
    SavePlace := GridActividadesxAnexo.DataSource.DataSet.GetBookmark;
    with GridActividadesxAnexo.DataSource.DataSet do
    begin
      for iGrid := 0 to GridActividadesxAnexo.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(GridActividadesxAnexo.SelectedRows.Items[iGrid]));
        if sWbsAnterior = '' then
          ProcIntroduceRegistro(sWbsAnterior, sItemOrden, global_contrato, zqReprogramacion.FieldByName('sIdConvenio').AsString, tsNumeroOrden.Text, FieldValues['sWbs'], FieldValues['sNumeroActividad'], FieldValues['sTipoActividad'], iNivel, '', '', '', FieldValues['sActividadAnterior'], '')
        else
          if FieldValues['sTipoActividad'] = 'Actividad' then
            ProcIntroduceRegistro(sWbsAnterior, sItemOrden, global_contrato, zqReprogramacion.FieldByName('sIdConvenio').AsString, tsNumeroOrden.Text, FieldValues['sWbs'], FieldValues['sNumeroActividad'], FieldValues['sTipoActividad'], iNivel, '', '', '', FieldValues['sActividadAnterior'], '')
      end
    end;

    // Modifico la fecha de los paquetes superiores ..
    if (Connection.Configuracion.FieldValues['lCalculaFecha'] = 'Si') then
      ProcRegeneraMontos(global_contrato, zqReprogramacion.FieldByName('sIdConvenio').AsString, tsNumeroOrden.Text, sWbsAnterior);
    SavePlace := grid_actividades.DataSource.DataSet.GetBookmark;
    ActividadesxOrden.Refresh;
    try
      grid_actividades.DataSource.DataSet.GotoBookmark(SavePlace);
    except
      grid_actividades.DataSource.DataSet.FreeBookmark(SavePlace);
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Conceptos/Partidas x Frente de Trabajo', 'Al insertar actividad', 0);
    end;
  end;
end;

procedure TfrmActividades.procBuscaPartida(Sender: TObject);
var
  sNumeroPartida: string;
begin
  if zActividadesxAnexo.RecordCount > 0 then
  begin
    sNumeroPartida := (Sender as tEdit).Text;
    zActividadesxAnexo.Locate('sNumeroActividad', sNumeroPartida, [loCaseInsensitive])
  end;
end;

procedure TfrmActividades.ProgramaDiariodelConceptodelaCia1Click(
  Sender: TObject);
begin
  ProgramarActividad();
end;

procedure TfrmActividades.ProgramarActividad();
var
  FechaInicial, FechaFinal: TDateTime;
  iNumeroDias: Integer;
  dCantidad: Double;
  dAjuste: Double;
  lProgramarDias: Boolean;
  lReprogramarDias: Boolean;
begin
  FechaInicial := 0;
  FechaFinal := 0;
  Connection.QryBusca.Active := False;
  Connection.QryBusca.SQL.Clear;
  Connection.QryBusca.SQL.Add('select min(dIdFecha)  as fecha from distribuciondeactividadescia  ' +
    ' where sContrato =:contrato and sNumeroOrden = :orden and sIdConvenio=:convenio ' +
    ' and sWbs=:wbs and sNumeroActividad=:actividad ');
  Connection.QryBusca.ParamByName('contrato').AsString := global_contrato;
  Connection.QryBusca.ParamByName('orden').AsString := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
  Connection.QryBusca.ParamByName('convenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
  Connection.QryBusca.ParamByName('wbs').AsString := ActividadesxOrden.FieldByName('sWbs').AsString;
  Connection.QryBusca.ParamByName('actividad').AsString := ActividadesxOrden.FieldByName('sNumeroActividad').AsString;
  Connection.QryBusca.Open;
  if Connection.QryBusca.RecordCount > 0 then
  begin
    FechaInicial := Connection.QryBusca.FieldByName('fecha').AsDateTime;
  end;
  Connection.QryBusca.Active := False;
  Connection.QryBusca.SQL.Clear;
  Connection.QryBusca.SQL.Add('select max(dIdFecha) as fecha  from distribuciondeactividadescia  ' +
    ' where sContrato =:contrato and sNumeroOrden = :orden and sIdConvenio=:convenio ' +
    ' and sWbs=:wbs and sNumeroActividad=:actividad    ');
  Connection.QryBusca.ParamByName('contrato').AsString := global_contrato;
  Connection.QryBusca.ParamByName('orden').AsString := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
  Connection.QryBusca.ParamByName('convenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
  Connection.QryBusca.ParamByName('wbs').AsString := ActividadesxOrden.FieldByName('sWbs').AsString;
  Connection.QryBusca.ParamByName('actividad').AsString := ActividadesxOrden.FieldByName('sNumeroActividad').AsString;
  Connection.QryBusca.Open;
  if Connection.QryBusca.RecordCount > 0 then
  begin
    FechaFinal := Connection.QryBusca.FieldByName('fecha').AsDateTime;
  end;

  lProgramarDias := False;
  lReprogramarDias := False;
//si no hay programacion diaria crearla
  if (FechaInicial = 0) or (FechaFinal = 0) then
  begin
    if MessageDlg('No existe programacion diaria para esta actividad, desea crearla ahora?', mtInformation, [mbYes, mbNo], 0) = mrYes then
    begin
      FechaInicial := Actividadesxorden.FieldByName('dFechaInicio').AsDateTime;
      FechaFinal := Actividadesxorden.FieldByName('dFechaFinal').AsDateTime;
      iNumeroDias := DaysBetween(FechaFinal, FechaInicial) + 1;
      lProgramarDias := True;
      if Actividadesxorden.FieldByName('dCantidad').AsFloat > 0 then
        dCantidad := Actividadesxorden.FieldByName('dCantidad').AsFloat / iNumeroDias
      else
        dCantidad := 0;
    end;
  end;

//si hay programacion , pero el rango de fechas difiere del que tiene la partida actualmente
  if
    (not lProgramarDias) and
    (
    (FechaInicial <> Actividadesxorden.FieldByName('dFechaInicio').AsDateTime)
    or
    (FechaFinal <> Actividadesxorden.FieldByName('dFechaFinal').AsDateTime)
    )
    then
  begin
    if MessageDlg('La programacion actual difiere de las fechas de inicio y termino de la actividad, desea reprogramar?' + chr(10) +
      '[Las nuevas cantidades se anexan en ceros,las que estan fuera de ese rango de fechas se borran]', mtInformation, [mbYes, mbNo], 0) = mrYes then
    begin
      FechaInicial := Actividadesxorden.FieldByName('dFechaInicio').AsDateTime;
      FechaFinal := Actividadesxorden.FieldByName('dFechaFinal').AsDateTime;
      iNumeroDias := DaysBetween(FechaFinal, FechaInicial);
      lReprogramarDias := True;
      dCantidad := 0;
      //borrar la programacion diaria que esta arriba de la fecha maxima actual
      Connection.zCommand.Active := False;
      Connection.zCommand.SQL.Clear;
      Connection.zCommand.SQL.Add('delete from distribuciondeactividadescia  ' +
        ' where sContrato =:contrato and sNumeroOrden = :orden and sIdConvenio=:convenio ' +
        ' and sWbs=:wbs and sNumeroActividad=:actividad and dIdFecha >:fechamaxima');
      Connection.zCommand.ParamByName('contrato').AsString := global_contrato;
      Connection.zCommand.ParamByName('orden').AsString := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
      Connection.zCommand.ParamByName('convenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
      Connection.zCommand.ParamByName('wbs').AsString := ActividadesxOrden.FieldByName('sWbs').AsString;
      Connection.zCommand.ParamByName('actividad').AsString := ActividadesxOrden.FieldByName('sNumeroActividad').AsString;
      Connection.zCommand.ParamByName('fechamaxima').AsDateTime := FechaFinal;
      Connection.zCommand.ExecSQL;

      //borrar la programacion diaria que esta arriba de la fecha minima actual
      Connection.zCommand.Active := False;
      Connection.zCommand.SQL.Clear;
      Connection.zCommand.SQL.Add('delete from distribuciondeactividadescia  ' +
        ' where sContrato =:contrato and sNumeroOrden = :orden and sIdConvenio=:convenio ' +
        ' and sWbs=:wbs and sNumeroActividad=:actividad and dIdFecha <:fechamaxima');
      Connection.zCommand.ParamByName('contrato').AsString := global_contrato;
      Connection.zCommand.ParamByName('orden').AsString := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
      Connection.zCommand.ParamByName('convenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
      Connection.zCommand.ParamByName('wbs').AsString := ActividadesxOrden.FieldByName('sWbs').AsString;
      Connection.zCommand.ParamByName('actividad').AsString := ActividadesxOrden.FieldByName('sNumeroActividad').AsString;
      Connection.zCommand.ParamByName('fechamaxima').AsDateTime := FechaInicial;
      Connection.zCommand.ExecSQL;
    end;
  end;


  if (lProgramarDias) or (lReprogramarDias) then
  begin
    dAjuste := 0;
    while (FechaInicial <= FechaFinal) do
    begin
      if (FechaInicial = FechaFinal) and lProgramarDias then
      begin
        Connection.QryBusca.Active := False;
        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Add('select sum(dCantidad) as dSuma from distribuciondeactividadescia ' +
          ' where sContrato =:contrato and sNumeroOrden = :orden and sIdConvenio=:convenio ' +
          ' and sWbs=:wbs and sNumeroActividad=:actividad ');
        Connection.QryBusca.ParamByName('contrato').AsString := global_contrato;
        Connection.QryBusca.ParamByName('orden').AsString := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
        Connection.QryBusca.ParamByName('convenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
        Connection.QryBusca.ParamByName('wbs').AsString := ActividadesxOrden.FieldByName('sWbs').AsString;
        Connection.QryBusca.ParamByName('actividad').AsString := ActividadesxOrden.FieldByName('sNumeroActividad').AsString;
        Connection.QryBusca.Open;
        dAjuste := Connection.QryBusca.FieldByName('dSuma').AsFloat;

        dCantidad := Actividadesxorden.FieldByName('dCantidad').AsFloat - dAjuste;

        if dCantidad < 0 then dCantidad := dCantidad * (-1);
      end;
      Connection.zCommand.Active := False;
      Connection.zCommand.SQL.Clear;
      Connection.zCommand.SQL.Add('insert into distribuciondeactividadescia set ' +
        '  sContrato = :contrato ,' +
        '  sIdConvenio = :convenio,' +
        '  sNumeroOrden =:orden ,' +
//        '  sWbsContrato =:wbscontrato ,' +
        '  sWbs =:wbs,' +
//        '  sPaquete =:paquete,' +
        '  sNumeroActividad=:actividad,' +
//        '  sTipoActividad =:tipoactividad,' +
//        '  sIdFase =:fase,' +
        '  dIdFecha =:fecha ,' +
        '  dCantidad = :Cantidad on duplicate key update dIdFecha=:fecha');
      Connection.zCommand.ParamByName('contrato').AsString := Actividadesxorden.FieldByName('sContrato').AsString;
      Connection.zCommand.ParamByName('convenio').AsString := Actividadesxorden.FieldByName('sIdConvenio').AsString;
      Connection.zCommand.ParamByName('orden').AsString := Actividadesxorden.FieldByName('sNumeroOrden').AsString;
//      Connection.zCommand.ParamByName('wbscontrato').AsString := Actividadesxorden.FieldByName('sWbsContrato').AsString;
      Connection.zCommand.ParamByName('wbs').AsString := Actividadesxorden.FieldByName('sWbs').AsString;
//      Connection.zCommand.ParamByName('paquete').AsString := Actividadesxorden.FieldByName('sPaquete').AsString;
      Connection.zCommand.ParamByName('actividad').AsString := Actividadesxorden.FieldByName('sNumeroActividad').AsString;
//      Connection.zCommand.ParamByName('fase').AsString := Actividadesxorden.FieldByName('sIdFase').AsString;
//      Connection.zCommand.ParamByName('tipoactividad').AsString := Actividadesxorden.FieldByName('sTipoActividad').AsString;
      Connection.zCommand.ParamByName('fecha').AsDateTime := FechaInicial;
      Connection.zCommand.ParamByName('Cantidad').AsFloat := dCantidad;
      Connection.zCommand.ExecSql;
      FechaInicial := incday(FechaInicial);
    end

  end;

//  application.CreateForm(TfrmProgramacionActividadxOrden, frmProgramacionActividadxOrden);
//
//  frmProgramacionActividadxOrden.QryDetalle.Active := False;
//  frmProgramacionActividadxOrden.QryDetalle.ParamByName('contrato').AsString := Actividadesxorden.FieldByName('sContrato').AsString;
//  frmProgramacionActividadxOrden.QryDetalle.ParamByName('convenio').AsString := Actividadesxorden.FieldByName('sIdConvenio').AsString;
//  frmProgramacionActividadxOrden.QryDetalle.ParamByName('orden').AsString := Actividadesxorden.FieldByName('sNumeroOrden').AsString;
//  frmProgramacionActividadxOrden.QryDetalle.ParamByName('wbs').AsString := Actividadesxorden.FieldByName('sWbs').AsString;
//  frmProgramacionActividadxOrden.QryDetalle.ParamByName('actividad').AsString := Actividadesxorden.FieldByName('sNumeroActividad').AsString;
//  frmProgramacionActividadxOrden.QryDetalle.ParamByName('isometrico').AsString := Actividadesxorden.FieldByName('sIsometrico').AsString;
//  frmProgramacionActividadxOrden.QryDetalle.ParamByName('empleado').AsString := Actividadesxorden.FieldByName('sIdEmpleado').AsString;
//  frmProgramacionActividadxOrden.QryDetalle.Open;
//
//  frmProgramacionActividadxOrden.lblDescripcion.Caption := Actividadesxorden.FieldByName('mDescripcion').AsString;
//  frmProgramacionActividadxOrden.FechaInicial.Caption := datetostr(Actividadesxorden.FieldByName('dFechaInicio').AsDateTime);
//  frmProgramacionActividadxOrden.FechaFinal.Caption := datetostr(Actividadesxorden.FieldByName('dFechaFinal').AsDateTime);
//  frmProgramacionActividadxOrden.lblPartida.Caption := Actividadesxorden.FieldByName('sNumeroActividad').AsString;
//
//
//  frmProgramacionActividadxOrden.Visible := False;
//  frmProgramacionActividadxOrden.ShowModal;

  application.CreateForm(TfrmProgramacionPartidasCia, frmProgramacionPartidasCia);

  frmProgramacionPartidasCia.QryDetalle.Active := False;
  frmProgramacionPartidasCia.QryDetalle.SQL.Clear;
  frmProgramacionPartidasCia.QryDetalle.SQL.Add('select * from distribuciondeactividadescia where ' +
    ' sContrato=:contrato ' +
    ' and sNumeroOrden=:orden ' +
    ' and sIdConvenio=:convenio ' +
    ' and sWbs=:wbs ' +
    ' and sNumeroActividad=:actividad ' +
    ' order by dIdFecha ');

  frmProgramacionPartidasCia.QryDetalle.ParamByName('contrato').AsString := Actividadesxorden.FieldByName('sContrato').AsString;
  frmProgramacionPartidasCia.QryDetalle.ParamByName('convenio').AsString := Actividadesxorden.FieldByName('sIdConvenio').AsString;
  frmProgramacionPartidasCia.QryDetalle.ParamByName('orden').AsString := Actividadesxorden.FieldByName('sNumeroOrden').AsString;
  frmProgramacionPartidasCia.QryDetalle.ParamByName('wbs').AsString := Actividadesxorden.FieldByName('sWbs').AsString;
  frmProgramacionPartidasCia.QryDetalle.ParamByName('actividad').AsString := Actividadesxorden.FieldByName('sNumeroActividad').AsString;
  frmProgramacionPartidasCia.QryDetalle.Open;

  frmProgramacionPartidasCia.lblDescripcion.Caption := Actividadesxorden.FieldByName('mDescripcion').AsString;
  frmProgramacionPartidasCia.FechaInicial.Caption := datetostr(Actividadesxorden.FieldByName('dFechaInicio').AsDateTime);
  frmProgramacionPartidasCia.FechaFinal.Caption := datetostr(Actividadesxorden.FieldByName('dFechaFinal').AsDateTime);
  frmProgramacionPartidasCia.lblPartida.Caption := Actividadesxorden.FieldByName('sNumeroActividad').AsString;


  frmProgramacionPartidasCia.Visible := False;
  frmProgramacionPartidasCia.ShowModal;

  frmProgramacionPartidasCia.Destroy;

end;

procedure TfrmActividades.InsertarConceptosClick(Sender: TObject);
var
  myForm: TForm;
  zDSActividadesxAnexo: tDataSource;
  sPaquete: string;
begin
  if tsNumeroOrden.Text <> '' then
    if (ActividadesxOrden.State <> dsInsert) or (ActividadesxOrden.State <> dsEdit) then
    begin
      lYaPregunto := False;
      if ActividadesxOrden.RecordCount > 0 then
        if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
          sPaquete := ActividadesxOrden.FieldValues['sWbs']
        else
          sPaquete := ActividadesxOrden.FieldValues['sWbsAnterior']
      else
        sPaquete := '';

      myForm := TForm.Create(Self);
      try
        myForm.Position := poDesktopCenter;
        myForm.Caption := 'Insertar Conceptos del Contrato Principal en el Paquete No. ' + sPaquete + ' "' + tsPaquete.Text + '"';
        MyForm.BorderIcons := [];
        MyForm.Width := 1200;
        MyForm.Height := 480;
        MyForm.BorderStyle := bsSizeable;
        MyForm.Color := $00FEC6BA;

        zActividadesxAnexo := TZReadOnlyQuery.Create(nil);
        zActividadesxAnexo.Connection := connection.zConnection;
        zActividadesxAnexo.Active := False;
        zActividadesxAnexo.Sql.Clear;
        zActividadesxAnexo.Sql.Add('Select *, sNumeroActividad as sSpacesNumeroActividad, SubStr(mDescripcion, 1, 255) as sDescripcion From actividadesxanexo ' +
          'where sContrato = :contrato And sIdConvenio = :Convenio Order By iItemOrden');
        zActividadesxAnexo.Params.ParamByName('Contrato').DataType := ftString;
        zActividadesxAnexo.Params.ParamByName('Contrato').Value    := global_contrato;
        zActividadesxAnexo.Params.ParamByName('Convenio').DataType := ftString;
        zActividadesxAnexo.Params.ParamByName('Convenio').Value    := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        zActividadesxAnexo.Open;
        zDSActividadesxAnexo := tDataSource.Create(nil);
        zDSActividadesxAnexo.DataSet := zActividadesxAnexo;

        GridActividadesxAnexo := TRxDBGrid.Create(MyForm);
        with GridActividadesxAnexo do
        begin
          Parent := myForm;
          Visible := True;
          Align := alCustom;
          Options := [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgAlwaysShowSelection, dgCancelOnExit, dgMultiSelect];
          TitleButtons := True;
          DataSource := zDSActividadesxAnexo;
          Width := 1200;
          Height := 430;
          Anchors := [akLeft, akTop, akRight, akBottom];
          ParentColor := True;
          Ctl3D := False;

          Columns.Clear;
          Columns.Add;
          Columns[0].FieldName := 'sSpacesNumeroActividad';
          Columns[0].Width := 70;
          Columns[0].Title.Caption := 'Actividad';
          Columns[0].ReadOnly := True;
          Columns[0].Font.Style := [fsBold];
          Columns[0].Font.Color := clBlue;
          Columns.Add;
          Columns[1].FieldName := 'sActividadAnterior';
          Columns[1].Width := 70;
          Columns[1].Title.Caption := 'Act. Anterior';
          Columns[1].ReadOnly := True;
          Columns[1].Font.Style := [fsBold];
          Columns.Add;
          Columns[2].FieldName := 'sDescripcion';
          Columns[2].Width := 680;
          Columns[2].Title.Caption := 'Descripcion';
          Columns[2].ReadOnly := True;
          Columns[2].Font.Style := [];
          Columns.Add;
          Columns[3].FieldName := 'dFechaInicio';
          Columns[3].Width := 60;
          Columns[3].Title.Caption := 'F. Inicio';
          Columns[3].Font.Style := [];
          Columns.Add;
          Columns[4].FieldName := 'dFechaFinal';
          Columns[4].Width := 60;
          Columns[4].Title.Caption := 'F. Final';
          Columns[4].Title.Alignment := taRightJustify;
          Columns[4].Font.Style := [];
          Columns.Add;
          Columns[5].FieldName := 'dCantidadAnexo';
          Columns[5].Width := 70;
          Columns[5].Title.Caption := 'Cant. a Inst.';
          Columns[5].Title.Alignment := taRightJustify;
          Columns[5].Font.Style := [];
          Columns.Add;
          Columns[6].FieldName := 'sMedida';
          Columns[6].Width := 60;
          Columns[6].Title.Caption := 'U. Medida';
          Columns[6].Title.Alignment := taRightJustify;
          Columns[6].Font.Style := [];
          Columns.Add;
          Columns[7].FieldName := 'sAnexo';
          Columns[7].Width := 60;
          Columns[7].Title.Caption := 'Anexo';
          Columns[7].Title.Alignment := taRightJustify;
          Columns[7].Font.Style := [];
          Columns.Add;
          Columns[8].FieldName := 'dVentaMN';
          Columns[8].Width := 70;
          Columns[8].Title.Caption := '$ Precio MN';
          Columns[8].Title.Alignment := taRightJustify;
          Columns[8].Font.Style := [];
          Columns.Add;
          Columns[9].FieldName := 'dPonderado';
          Columns[9].Width := 70;
          Columns[9].Title.Caption := '% Ponderado';
          Columns[9].Title.Alignment := taRightJustify;
          Columns[9].Font.Style := [];
        end;

        with TButton.Create(Self) do
        begin
          Left := 10;
          Top := 440;
          Width := 120;
          Height := 35;
          Default := True;
          Parent := MyForm;
          Caption := 'Importar File XLS';
          OnClick := ImportaXLS;
          Anchors := [akLeft, akBottom];
        end;

        with TButton.Create(Self) do
        begin
          Left := 140;
          Top := 440;
          Width := 120;
          Height := 35;
          Default := True;
          Parent := MyForm;
          Caption := 'Insertar Partidas';
          OnClick := InsertaActividad;
          Anchors := [akLeft, akBottom];
        end;

        with TButton.Create(Self) do
        begin
          Left := 270;
          Top := 440;
          Width := 120;
          Height := 35;
          ModalResult := mrCancel;
          Cancel := True;
          Parent := MyForm;
          Caption := 'Cancelar Inserccion';
          Anchors := [akLeft, akBottom];
        end;

        with TLabel.Create(Self) do
        begin
          Left := 1000;
          Top := 455;
          Width := 120;
          Height := 35;
          Parent := MyForm;
          Caption := 'Buscar ...:';
          Anchors := [akRight, akBottom];
        end;
        with TEdit.Create(Self) do
        begin
          Left := 1060;
          Top := 450;
          Width := 130;
          Height := 35;
          Parent := MyForm;
          Anchors := [akRight, akBottom];
          OnChange := procBuscaPartida;
        end;
        myForm.ShowModal;
      finally
        zActividadesxAnexo.Destroy;
        zDSActividadesxAnexo.Destroy;
        GridActividadesxAnexo.Destroy;
        Paquetes.Refresh;
        myForm.Free;
      end;
    end;
end;

procedure TfrmActividades.ActividadesxOrdensTipoActividadChange(
  Sender: TField);
begin
  if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
    tlGerencial.Enabled := True
  else
    tlGerencial.Enabled := False
end;

procedure TfrmActividades.formatoEncabezado;
begin
  Excel.Selection.MergeCells := False;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size := 12;
  Excel.Selection.Font.Bold := False;
  Excel.Selection.Font.Name := 'Calibri';
end;

//*************************************BRITO 25-03-11***************************

procedure TfrmActividades.PopUpNuevoRegistro;
var
  myForm: TForm;
  zDSActividadesxAnexo: tDataSource;
  sPaquete: string;
begin
  try
    if tsNumeroOrden.Text <> '' then
      if (ActividadesxOrden.State <> dsInsert) or (ActividadesxOrden.State <> dsEdit) then
      begin
                //lYaPregunto := False ;
        if ActividadesxOrden.RecordCount > 0 then
          if ActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
            sPaquete := ActividadesxOrden.FieldValues['sWbs']
          else
            sPaquete := ActividadesxOrden.FieldValues['sWbsAnterior']
        else
          sPaquete := '';

        myForm := TForm.Create(Self);
        try
          myForm.Position := poDesktopCenter;
          myForm.Caption := 'Insertar Concepto del Contrato Principal en el Paquete No. ' + sPaquete + ' "' + sPaqueteDesc + '"';
          MyForm.BorderIcons := [];
          MyForm.Width := 1200;
          MyForm.Height := 480;
          MyForm.BorderStyle := bsSizeable;
          MyForm.Color := $00FEC6BA;

          zActividadesxAnexo := TZReadOnlyQuery.Create(nil);
          zActividadesxAnexo.Connection := connection.zConnection;
          zActividadesxAnexo.Active := False;
          zActividadesxAnexo.Sql.Clear;
          zActividadesxAnexo.Sql.Add('Select *, sNumeroActividad as sSpacesNumeroActividad, SubStr(mDescripcion, 1, 255) as sDescripcion From actividadesxanexo ' +
            'where sContrato = :contrato And sIdConvenio = :Convenio And sTipoActividad = "Actividad" Order By iItemOrden');
          zActividadesxAnexo.Params.ParamByName('Contrato').DataType := ftString;
          zActividadesxAnexo.Params.ParamByName('Contrato').Value := global_contrato;
          zActividadesxAnexo.Params.ParamByName('Convenio').DataType := ftString;
          zActividadesxAnexo.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          zActividadesxAnexo.Open;
          zDSActividadesxAnexo := tDataSource.Create(nil);
          zDSActividadesxAnexo.DataSet := zActividadesxAnexo;

          GridActividadesxAnexo := TRxDBGrid.Create(MyForm);
          with GridActividadesxAnexo do
          begin
            Parent := myForm;
            Visible := True;
            Align := alCustom;
            Options := [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgRowSelect, dgAlwaysShowSelection, dgCancelOnExit];
            TitleButtons := True;
            DataSource := zDSActividadesxAnexo;
            Width := 1200;
            Height := 430;
            Anchors := [akLeft, akTop, akRight, akBottom];
            ParentColor := True;
            Ctl3D := False;

            Columns.Clear;
            Columns.Add;
            Columns[0].FieldName := 'sAnexo';
            Columns[0].Width := 60;
            Columns[0].Title.Caption := 'Anexo';
            Columns[0].Title.Alignment := taRightJustify;
            Columns[0].Font.Style := [];
            Columns.Add;
            Columns[1].FieldName := 'sSpacesNumeroActividad';
            Columns[1].Width := 70;
            Columns[1].Title.Caption := 'Actividad';
            Columns[1].ReadOnly := True;
            Columns[1].Font.Style := [fsBold];
            Columns[1].Font.Color := clBlue;
            Columns.Add;
            Columns[2].FieldName := 'sDescripcion';
            Columns[2].Width := 600;
            Columns[2].Title.Caption := 'Descripcion';
            Columns[2].ReadOnly := True;
            Columns[2].Font.Style := [];
            Columns.Add;
            Columns[3].FieldName := 'dFechaInicio';
            Columns[3].Width := 60;
            Columns[3].Title.Caption := 'F. Inicio';
            Columns[3].Font.Style := [];
            Columns.Add;
            Columns[4].FieldName := 'dFechaFinal';
            Columns[4].Width := 60;
            Columns[4].Title.Caption := 'F. Final';
            Columns[4].Title.Alignment := taRightJustify;
            Columns[4].Font.Style := [];
            Columns.Add;
            Columns[5].FieldName := 'dCantidadAnexo';
            Columns[5].Width := 70;
            Columns[5].Title.Caption := 'Cant. a Inst.';
            Columns[5].Title.Alignment := taRightJustify;
            Columns[5].Font.Style := [];
            Columns.Add;
            Columns[6].FieldName := 'sMedida';
            Columns[6].Width := 60;
            Columns[6].Title.Caption := 'U. Medida';
            Columns[6].Title.Alignment := taRightJustify;
            Columns[6].Font.Style := [];
            Columns.Add;
            Columns[7].FieldName := 'dVentaMN';
            Columns[7].Width := 70;
            Columns[7].Title.Caption := '$ Precio MN';
            Columns[7].Title.Alignment := taRightJustify;
            Columns[7].Font.Style := [];
            Columns.Add;
            Columns[8].FieldName := 'dPonderado';
            Columns[8].Width := 70;
            Columns[8].Title.Caption := '% Ponderado';
            Columns[8].Title.Alignment := taRightJustify;
            Columns[8].Font.Style := [];
          end;

          with TButton.Create(Self) do
          begin
            Left := 10;
            Top := 440;
            Width := 120;
            Height := 30;
            ModalResult := mrOk;
            Default := True;
            Parent := MyForm;
            Caption := 'Nuevo Paquete';
            OnClick := NuevoPaquete;
            Anchors := [akLeft, akBottom];
          end;

          with TButton.Create(Self) do
          begin
            Left := 140;
            Top := 440;
            Width := 120;
            Height := 30;
            ModalResult := mrOk;
            Default := True;
            Parent := MyForm;
            Caption := 'Seleccionar Partida';
            OnClick := SeleccionarNuevaActividad;
            Anchors := [akLeft, akBottom];
          end;

          with TButton.Create(Self) do
          begin
            Left := 270;
            Top := 440;
            Width := 120;
            Height := 30;
            ModalResult := mrCancel;
            Cancel := True;
            Parent := MyForm;
            Caption := 'Cancelar Seleccion';
            Anchors := [akLeft, akBottom];
          end;

          with TLabel.Create(Self) do
          begin
            Left := 1000;
            Top := 455;
            Width := 120;
            Height := 30;
            Parent := MyForm;
            Caption := 'Buscar ...:';
            Anchors := [akRight, akBottom];
          end;
          with TEdit.Create(Self) do
          begin
            Left := 1060;
            Top := 450;
            Width := 130;
            Height := 30;
            Parent := MyForm;
            Anchors := [akRight, akBottom];
            OnChange := procBuscaPartida;
          end;
          if myForm.ShowModal = mrOk then
          begin
            tsPaquete.SetFocus;
          end
          else begin
            frmBarra1.btnCancel.Click
          end;
        finally
          zActividadesxAnexo.Destroy;
          zDSActividadesxAnexo.Destroy;
          GridActividadesxAnexo.Destroy;
          Paquetes.Refresh;
          myForm.Free;
        end;
      end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Conceptos/Partidas x Frente de Trabajo', 'Al mostrar ventana de nuevo registro', 0);
    end;
  end;
end;
//*************************************BRITO 25-03-11***************************

//*************************************BRITO 25-03-11***************************

procedure TfrmActividades.SeleccionarNuevaActividad(Sender: TObject);
begin
  with GridActividadesxAnexo.DataSource.DataSet do
  begin
    ActividadesxOrden.FieldValues['sWbsContrato'] := FieldValues['sWbs'];
    ActividadesxOrden.FieldValues['sNumeroActividad'] := FieldValues['sNumeroActividad'];
    ActividadesxOrden.FieldValues['sWbsAnterior'] := sPaquete;
    ActividadesxOrden.FieldValues['mDescripcion'] := FieldValues['mDescripcion'];
    ActividadesxOrden.FieldValues['sMedida'] := FieldValues['sMedida'];
    ActividadesxOrden.FieldValues['sActividadAnterior'] := FieldValues['sActividadAnterior'];
    ActividadesxOrden.FieldValues['dVentaMN'] := FieldValues['dVentaMN'];
    ActividadesxOrden.FieldValues['dCostoMN'] := FieldValues['dCostoMN'];
    ActividadesxOrden.FieldValues['sTipoAnexo'] := FieldValues['sTipoAnexo'];
    ActividadesxOrden.FieldValues['sAnexo'] := FieldValues['sAnexo'];
    ActividadesxOrden.FieldValues['sTipoAnexo'] := FieldValues['sTipoAnexo'];
       {Ahora creamos el itemorden...}
    ActividadesxOrden.FieldValues['iItemOrden'] := OrdenPaqueteItem + sFnBuscaItem(zqReprogramacion.FieldByName('sIdConvenio').AsString,FieldValues['sNumeroActividad'],
      sPaquete,
      OrdenPaqueteItem,
      FieldValues['sTipoActividad'], tsNumeroOrden.Text, 'actividadesxorden',
      OrdenPaqueteNivel + 1);
    if FieldValues['sTipoActividad'] = 'Paquete' then
      ActividadesxOrden.FieldValues['sWbs'] := sPaquete + '.' + FieldValues['sNumeroActividad']
    else
    begin
      if FieldValues['sAnexo'] <> '' then
        ActividadesxOrden.FieldValues['sWbs'] := sPaquete + '.' + FieldValues['sAnexo'] + '.' + FieldValues['sNumeroActividad']
      else
        ActividadesxOrden.FieldValues['sWbs'] := sPaquete + '.' + FieldValues['sNumeroActividad'];
    end;

  end;
  ActividadesxOrden.FieldValues['sTipoActividad'] := 'Actividad';
  tsPaquete.Enabled := True;
  tmDescripcion.Enabled := False;
  tsNumeroActividad.Enabled := False;
  tdCostoMN.Enabled := true;
  tdVentaMN.Enabled := true;
  tsUnidad.Enabled := False;
end;

//*************************************BRITO 25-03-11***************************

procedure TfrmActividades.NuevoPaquete(Sender: TObject);
begin
  ActividadesxOrden.FieldValues['sTipoActividad'] := 'Paquete';
  ActividadesxOrden.FieldValues['sWbsContrato'] := SwbsPrincipal(global_contrato, zqReprogramacion.FieldByName('sIdConvenio').AsString, 'Paquete', '', connection.zConnection); //'*';
  ActividadesxOrden.FieldValues['sWbsAnterior'] := sPaquete;

  tsUnidad.Enabled := False;
  tdVentaMn.Enabled := False;
  tdCostoMn.Enabled := False;
end;

function TfrmActividades.SumaCantidades(): boolean;
var
  dCantidad: double;
begin
  SumaCantidades := True;
     {Primero sumamos las cantidades de Cada Frente de trabajo..}
  connection.qryBusca.Active := False;
  connection.qryBusca.SQL.Clear;
  connection.qryBusca.SQL.Add('Select sum(dCantidad) as dCantidad from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio and sWbsContrato =:WbsContrato and sNumeroActividad =:Actividad and sTipoActividad = "Actividad"  ' +
    'and (sWbs <> :Wbs and sNumeroOrden =:Orden) group by sContrato ');
  connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
  connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
  connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
  connection.qryBusca.Params.ParamByName('WbsContrato').DataType := ftString;
  connection.qryBusca.Params.ParamByName('WbsContrato').Value := actividadesxorden.FieldValues['sWbsContrato'];
  connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString;
  connection.qryBusca.Params.ParamByName('Wbs').Value := actividadesxorden.FieldValues['sWbs'];
  connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
  connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
  connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
  connection.qryBusca.Params.ParamByName('Actividad').Value := actividadesxorden.FieldValues['sNumeroActividad'];
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
    dCantidad := connection.qryBusca.FieldValues['dCantidad']
  else
    dCantidad := 0;

  dCantidad := dCantidad + tdCantidad.Value;

     {Primero buscamos la cantidad de Anexo de la partida..}
  connection.qryBusca2.Active := False;
  connection.qryBusca2.SQL.Clear;
  connection.qryBusca2.SQL.Add('Select dCantidadAnexo from actividadesxanexo where sContrato =:Contrato and sIdConvenio =:Convenio and sWbs =:Wbs and sNumeroActividad =:Actividad and sTipoActividad = "Actividad" ');
  connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString;
  connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato;
  connection.qryBusca2.Params.ParamByName('Convenio').DataType := ftString;
  connection.qryBusca2.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
  connection.qryBusca2.Params.ParamByName('Wbs').DataType := ftString;
  connection.qryBusca2.Params.ParamByName('Wbs').Value := actividadesxorden.FieldValues['sWbsContrato'];
  connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString;
  connection.qryBusca2.Params.ParamByName('Actividad').Value := actividadesxorden.FieldValues['sNumeroActividad'];
  connection.qryBusca2.Open;

  if connection.qryBusca2.RecordCount > 0 then
    if dCantidad > connection.QryBusca2.FieldValues['dCantidadAnexo'] then
    begin
      messageDLG('La cantidad de Anexo para el Concepto / Partida ' + actividadesxorden.FieldValues['sNumeroActividad'] +
        ' es menor a la suma de los Frentes. Cant. Anexo = ' + FloatToStr(connection.QryBusca2.FieldValues['dCantidadAnexo']) +
        ' Suma Frentes = ' + FloatToStr(dCantidad), mtInformation, [mbOk], 0);
      SumaCantidades := False;
    end;

end;


procedure TfrmActividades.ActualiaFactorGeneradorPER(sParamEmbarcacion: string; sParamOrden: string; sParamFolio: string);
var
   zqFactoresPersonal,
   zqFolios,
   zqMovFolios,
   zqActualizaFolios : TZQuery;
   TotalFolios,
   CantPersonalFolio : Double;
   sFecha : string;
   Progreso, TotalProgreso: real;
begin
    //Funcion par aactualizar los factores de los folios antes de imprimir generador de barco JJF by ivan 2 Nov 2013
    zqFolios:=TZQuery.Create (Self);
    zqFolios.connection:= connection.zConnection;

    zqActualizaFolios:=TZQuery.Create (Self);
    zqActualizaFolios.connection:= connection.zConnection;

    zqFactoresPersonal := TZQuery.Create(Self);
    zqFactoresPersonal.Connection := Connection.zConnection;
    zqFactoresPersonal.Active := False ;
    zqFactoresPersonal.SQL.Clear ;
    zqFactoresPersonal.SQL.Add('select bp.sContrato, bp.sNumeroOrden, bp.sIdPersonal as IdRecurso, '+
                                'bp.dIdFecha, ROUND(SUM(bp.dCanthh), 2) as sFactor, SUM(bp.dCanthh) as sFactorTotal, (ROUND(SUM(bp.dCanthh), 2) - SUM(bp.dCanthh)) as dDiferencia from bitacoradepersonal bp '+
                                'Inner Join personal p on (bp.sContrato = p.sContrato and  bp.sIdPersonal = p.sIdPersonal and p.lCobro ="Si") '+
                                'Inner Join contratos c on (bp.sContrato = c.sContrato) '+
                                'where bp.sIdPernocta not like :Embarcacion and bp.sContrato =:Orden  and bp.sNumeroOrden =:Folio and bp.dCantidad > 0 and bp.sTipoObra = "PU" '+
                                'Group By bp.sContrato, bp.sNumeroOrden, p.sIdPersonal, bp.dIdFecha order By bp.sContrato, bp.dIdFecha  asc ');
    zqFactoresPersonal.Params.ParamByName('Embarcacion').DataType := ftString ;
    zqFactoresPersonal.Params.ParamByName('Embarcacion').Value    := sParamEmbarcacion ;
    zqFactoresPersonal.Params.ParamByName('Orden').DataType       := ftString ;
    zqFactoresPersonal.Params.ParamByName('Orden').Value          := sParamOrden ;
    zqFactoresPersonal.Params.ParamByName('Folio').DataType       := ftString ;
    zqFactoresPersonal.Params.ParamByName('folio').Value          := sParamFolio ;
    zqFactoresPersonal.Open;

    sFecha := '';
    PanelProgress.Visible := True;
    zqFactoresPersonal.First;
    while not zqFactoresPersonal.Eof do
    begin
        Progreso := (1 / (zqFactoresPersonal.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
        TotalProgreso := TotalProgreso + Progreso;
        Label15.Caption := 'Procesando Personal...';
        Label15.Refresh;
        BarraEstado.Position := Trunc(TotalProgreso);

        {Actualizamos los valores de dCantHH a dCantHHGenerador}
        zqFolios.Active := False;
        zqFolios.SQL.Clear;
        zqFolios.SQL.Add('Update bitacoradepersonal set dCantHHGenerador = dCantHH '+
                         'where dIdFecha =:Fecha and sIdPernocta not like :Embarcacion and sContrato =:Orden '+
                         'and sNumeroOrden =:Folio and sTipoObra = "PU" and sIdPersonal =:Id ');
        zqFolios.ParamByName('Fecha').AsDateTime     := zqFactoresPersonal.FieldByName('dIdFecha').AsDateTime;
        zqFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
        zqFolios.ParamByName('Orden').AsString       := zqFactoresPersonal.FieldByName('sContrato').AsString;
        zqFolios.ParamByName('Folio').AsString       := zqFactoresPersonal.FieldByName('sNumeroOrden').AsString;
        zqFolios.ParamByName('Id').AsString          := zqFactoresPersonal.FieldByName('idRecurso').AsString;
        zqFolios.ExecSQL;

        {Consultamos el mayor factor de personal para aplicar ajuste}
        zqFolios.Active := False;
        zqFolios.SQL.Clear;
        zqFolios.SQL.Add('select * from bitacoradepersonal bp '+
                         'where bp.dIdFecha =:Fecha and bp.sIdPernocta not like :Embarcacion and bp.sContrato =:Orden '+
                         'and bp.sNumeroOrden =:Folio and bp.sTipoObra = "PU" and bp.sIdPersonal =:Id '+
                         'order By bp.sFactor DESC limit 1 ');
        zqFolios.ParamByName('Fecha').AsDateTime     := zqFactoresPersonal.FieldByName('dIdFecha').AsDateTime;
        zqFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
        zqFolios.ParamByName('Orden').AsString       := zqFactoresPersonal.FieldByName('sContrato').AsString;
        zqFolios.ParamByName('Folio').AsString       := zqFactoresPersonal.FieldByName('sNumeroOrden').AsString;
        zqFolios.ParamByName('Id').AsString          := zqFactoresPersonal.FieldByName('idRecurso').AsString;
        zqFolios.Open;

        if zqFolios.RecordCount > 0 then
        begin
            zqActualizaFolios.Active := False;
            zqActualizaFolios.SQL.Clear;
            zqActualizaFolios.SQL.Add('Update bitacoradepersonal set dCantHHGenerador = :Cantidad '+
                             'where dIdFecha =:Fecha and sIdPernocta not like :Embarcacion and sContrato =:Orden '+
                             'and sNumeroOrden =:Folio and sTipoObra = "PU" and sIdPersonal =:Id and sIdPlataforma =:Plataforma and sHoraInicio =:HoraI and sHoraFinal =:HoraF and sDescripcion =:Descripcion ');
            zqActualizaFolios.ParamByName('Fecha').AsDateTime     := zqFolios.FieldByName('dIdFecha').AsDateTime;
            zqActualizaFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
            zqActualizaFolios.ParamByName('Orden').AsString       := zqFolios.FieldByName('sContrato').AsString;
            zqActualizaFolios.ParamByName('Folio').AsString       := zqFolios.FieldByName('sNumeroOrden').AsString;
            zqActualizaFolios.ParamByName('Id').AsString          := zqFolios.FieldByName('sIdPersonal').AsString;
            zqActualizaFolios.ParamByName('Plataforma').AsString  := zqFolios.FieldByName('sIdPlataforma').AsString;
            zqActualizaFolios.ParamByName('HoraI').AsString       := zqFolios.FieldByName('sHoraInicio').AsString;
            zqActualizaFolios.ParamByName('HoraF').AsString       := zqFolios.FieldByName('sHoraFinal').AsString;
            zqActualizaFolios.ParamByName('Cantidad').AsFloat     := zqFolios.FieldByName('dCantHHGenerador').AsFloat + zqFactoresPersonal.FieldByName('dDiferencia').AsFloat;
            zqActualizaFolios.ParamByName('Descripcion').AsString := zqFolios.FieldByName('sDescripcion').AsString ;
            zqActualizaFolios.ExecSQL;
        end;

        zqFactoresPersonal.Next;
    end;

end;

procedure TfrmActividades.ActualiaFactorGeneradorEQ(sParamEmbarcacion: string; sParamOrden: string; sParamFolio: string);
var
   zqFactoresEquipo,
   zqFolios,
   zqMovFolios,
   zqActualizaFolios : TZQuery;
   TotalFolios,
   CantPersonalFolio : Double;
   sFecha : string;
   Progreso, TotalProgreso: real;
begin
    //Funcion par aactualizar los factores de los folios antes de imprimir generador de barco JJF by ivan 2 Nov 2013
    zqFolios:=TZQuery.Create (Self);
    zqFolios.connection:= connection.zConnection;

    zqActualizaFolios:=TZQuery.Create (Self);
    zqActualizaFolios.connection:= connection.zConnection;

    zqFactoresEquipo := TZQuery.Create(Self);
    zqFactoresEquipo.Connection := Connection.zConnection;
    zqFactoresEquipo.Active := False ;
    zqFactoresEquipo.SQL.Clear ;
    zqFactoresEquipo.SQL.Add('select  bp.sContrato, bp.sNumeroOrden, bp.sIdEquipo as IdRecurso, '+
                             'bp.dIdFecha, ROUNd(sum(bp.dCantHH),2) as sFactor, SUM(bp.dCanthh) as sFactorTotal, (ROUND(SUM(bp.dCanthh), 2) - SUM(bp.dCanthh)) as dDiferencia '+
                             'from bitacoradeequipos bp '+
                             'Inner Join equipos p on (bp.sContrato = p.sContrato and  bp.sIdEquipo = p.sIdEquipo and p.lCobro ="Si") '+
                             'Inner Join contratos c on (bp.sContrato = c.sContrato) '+
                             'where bp.sIdPernocta not like :Embarcacion and bp.sContrato =:Orden and bp.sNumeroOrden =:Folio and bp.dCantidad > 0 and bp.sTipoObra = "PU" '+
                             'Group By bp.sContrato, bp.sNumeroOrden, p.sIdEquipo, bp.dIdFecha order By bp.sContrato, bp.dIdFecha  asc ');
    zqFactoresEquipo.Params.ParamByName('Embarcacion').DataType := ftString ;
    zqFactoresEquipo.Params.ParamByName('Embarcacion').Value    := sParamEmbarcacion ;
    zqFactoresEquipo.Params.ParamByName('Orden').DataType       := ftString ;
    zqFactoresEquipo.Params.ParamByName('Orden').Value          := sParamOrden ;
    zqFactoresEquipo.Params.ParamByName('Folio').DataType       := ftString ;
    zqFactoresEquipo.Params.ParamByName('folio').Value          := sParamFolio ;
    zqFactoresEquipo.Open;

    sFecha := '';
    PanelProgress.Visible := True;
    zqFactoresEquipo.First;
    while not zqFactoresEquipo.Eof do
    begin
        Progreso := (1 / (zqFactoresEquipo.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
        TotalProgreso := TotalProgreso + Progreso;
        Label15.Caption := 'Procesando Equipo...';
        Label15.Refresh;
        BarraEstado.Position := Trunc(TotalProgreso);

        {Actualizamos los valores de dCantHH a dCantHHGenerador}
        zqFolios.Active := False;
        zqFolios.SQL.Clear;
        zqFolios.SQL.Add('Update bitacoradeequipos set dCantHHGenerador = dCantHH '+
                         'where dIdFecha =:Fecha and sIdPernocta not like :Embarcacion and sContrato =:Orden '+
                         'and sNumeroOrden =:Folio and sTipoObra = "PU" and sIdEquipo =:Id ');
        zqFolios.ParamByName('Fecha').AsDateTime     := zqFactoresEquipo.FieldByName('dIdFecha').AsDateTime;
        zqFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
        zqFolios.ParamByName('Orden').AsString       := zqFactoresEquipo.FieldByName('sContrato').AsString;
        zqFolios.ParamByName('Folio').AsString       := zqFactoresEquipo.FieldByName('sNumeroOrden').AsString;
        zqFolios.ParamByName('Id').AsString          := zqFactoresEquipo.FieldByName('idRecurso').AsString;
        zqFolios.ExecSQL;

        {Consultamos el mayor factor de personal para aplicar ajuste}
        zqFolios.Active := False;
        zqFolios.SQL.Clear;
        zqFolios.SQL.Add('select * from bitacoradeequipos bp '+
                         'where bp.dIdFecha =:Fecha and bp.sIdPernocta not like :Embarcacion and bp.sContrato =:Orden '+
                         'and bp.sNumeroOrden =:Folio and bp.sTipoObra = "PU" and bp.sIdEquipo =:Id '+
                         'order By bp.sFactor DESC limit 1 ');
        zqFolios.ParamByName('Fecha').AsDateTime     := zqFactoresEquipo.FieldByName('dIdFecha').AsDateTime;
        zqFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
        zqFolios.ParamByName('Orden').AsString       := zqFactoresEquipo.FieldByName('sContrato').AsString;
        zqFolios.ParamByName('Folio').AsString       := zqFactoresEquipo.FieldByName('sNumeroOrden').AsString;
        zqFolios.ParamByName('Id').AsString          := zqFactoresEquipo.FieldByName('idRecurso').AsString;
        zqFolios.Open;

        if zqFolios.RecordCount > 0 then
        begin
            zqActualizaFolios.Active := False;
            zqActualizaFolios.SQL.Clear;
            zqActualizaFolios.SQL.Add('Update bitacoradeequipos set dCantHHGenerador = :Cantidad '+
                             'where dIdFecha =:Fecha and sIdPernocta not like :Embarcacion and sContrato =:Orden '+
                             'and sNumeroOrden =:Folio and sTipoObra = "PU" and sIdEquipo =:Id and sHoraInicio =:HoraI and sHoraFinal =:HoraF and sDescripcion =:Descripcion ');
            zqActualizaFolios.ParamByName('Fecha').AsDateTime     := zqFolios.FieldByName('dIdFecha').AsDateTime;
            zqActualizaFolios.ParamByName('Embarcacion').AsString := sParamEmbarcacion;
            zqActualizaFolios.ParamByName('Orden').AsString       := zqFolios.FieldByName('sContrato').AsString;
            zqActualizaFolios.ParamByName('Folio').AsString       := zqFolios.FieldByName('sNumeroOrden').AsString;
            zqActualizaFolios.ParamByName('Id').AsString          := zqFolios.FieldByName('sIdEquipo').AsString;
            zqActualizaFolios.ParamByName('HoraI').AsString       := zqFolios.FieldByName('sHoraInicio').AsString;
            zqActualizaFolios.ParamByName('HoraF').AsString       := zqFolios.FieldByName('sHoraFinal').AsString;
            zqActualizaFolios.ParamByName('Cantidad').AsFloat     := zqFolios.FieldByName('dCantHHGenerador').AsFloat + zqFactoresEquipo.FieldByName('dDiferencia').AsFloat;
            zqActualizaFolios.ParamByName('Descripcion').AsString := zqFolios.FieldByName('sDescripcion').AsString ;
            zqActualizaFolios.ExecSQL;
        end;

        zqFactoresEquipo.Next;
    end;

end;


procedure TfrmActividades.PonderadoAnterior;
var
  dMontoContratoMN: Currency;
  dMontoContratoDLL: Currency;
  dPonderadoAjuste,
    Difer, Ponderado,
    decPonderado, Suma: Extended;
  scalcula: string;
begin
  scalcula := 'Si';
    //Inicia proceso de estructura del proyecto ...
  if Actividadesxorden.RecordCount > 0 then
    if MessageDlg('Desea Ponderar los Conceptos del Contrato Seleccionado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
        // Que ponderados se calcularan ?
        // Sumo todos las partidas anexo que tengan en lCalculo <> Si

      Connection.zCommand.Active := False;
      connection.zCommand.Filtered := False;
      Connection.zCommand.SQL.Clear;
      Connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = 0 ' +
        'Where sContrato = :Contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden');
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
      connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
      connection.zCommand.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      connection.zCommand.Params.ParamByName('orden').DataType := ftString;
      connection.zCommand.Params.ParamByName('orden').Value := tsNumeroOrden.Text;
      connection.zCommand.ExecSQL;

      dPonderadoAjuste := 100;

        // Actualizacion de ponderados ....
      dMontoContrato := 0;
      Connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('Select sum(dCantidad * dVentaMN) as dMontoMN From actividadesxorden ' +
        'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
        'and lcalculo=:calculo group by sContrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      Connection.QryBusca.Params.ParamByName('calculo').AsString := sCalcula;
      Connection.QryBusca.Open;
      if Connection.QryBusca.RecordCount > 0 then
        dMontoContratoMN := Connection.QryBusca.FieldValues['dMontoMN'];

      if connection.configuracion.FieldValues['lCalculoPonderado'] = 'Financiero' then
      begin
        if dMontoContratoMN > 0 then
        begin
          connection.zCommand.Active := False;
          connection.zCommand.Filtered := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = (((dCantidad * dVentaMN) / :montocontrato) * :miMaximoPonderado) ' +
            'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" and dCantidad <> 0 ' +
            'and lcalculo=:calculo');
          connection.zcommand.params.ParamByName('Contrato').DataType := ftString;
          connection.zcommand.params.ParamByName('Contrato').Value := global_contrato;
          connection.zcommand.params.ParamByName('Convenio').DataType := ftString;
          connection.zcommand.params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          connection.zcommand.params.ParamByName('Orden').DataType := ftString;
          connection.zcommand.params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          connection.zcommand.params.ParamByName('montocontrato').DataType := ftFloat;
          connection.zcommand.params.ParamByName('montocontrato').Value := dMontoContratoMN;
          connection.zcommand.params.ParamByName('miMaximoPonderado').DataType := ftFloat;
          connection.zcommand.params.ParamByName('miMaximoPonderado').Value := dPonderadoAjuste;
          connection.zcommand.params.ParamByName('calculo').AsString := scalcula;
          connection.zCommand.ExecSQL;
        end;
      end
      else
        if connection.configuracion.FieldValues['lCalculoPonderado'] = 'Duracion' then
        begin
                //Calculo el monto del programa ...
          Connection.QryBusca.Active := False;
          Connection.QryBusca.Filtered := False;
          Connection.QryBusca.SQL.Clear;
          Connection.QryBusca.SQL.Add('Select sum(dDuracion) as dDuracionTotal From actividadesxorden ' +
            'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
            'and lcalculo=:calculo group by sContrato');
          Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
          Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
          Connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
          Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          Connection.QryBusca.Params.ParamByName('calculo').AsString := sCalcula;
          Connection.QryBusca.Open;
          if connection.QryBusca.RecordCount > 0 then
          begin
            connection.zCommand.Active := False;
            connection.zCommand.Filtered := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Text := 'select * from actividadesxorden where sContrato = :Contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden And sTipoActividad = "Actividad" and dDuracion <> 0 ' +
              'and lcalculo=:calculo order by iItemOrden';
            connection.zcommand.params.ParamByName('Contrato').DataType := ftString;
            connection.zcommand.params.ParamByName('Contrato').Value := global_contrato;
            connection.zcommand.params.ParamByName('Convenio').DataType := ftString;
            connection.zcommand.params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.zcommand.params.ParamByName('Orden').DataType := ftString;
            connection.zcommand.params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.zcommand.params.ParamByName('calculo').AsString := sCalcula;
            connection.zCommand.Open;

            Difer := 0; // Diferencia para ajuste entre partidas
            Ponderado := 0; // Almacenamiento de ponderado total calculado
            Suma := 0;
            while not Connection.zCommand.Eof do
            begin
              Ponderado := (Connection.zCommand.FieldByName('dDuracion').AsFloat / Connection.QryBusca.FieldValues['dDuracionTotal']);
              Ponderado := Ponderado + Difer; // Sumar la diferencia anterior para ajuste automático
              decPonderado := Trunc(Ponderado * 1000000) / 1000000;
              Difer := Ponderado - decPonderado;
              decPonderado := decPonderado * dPonderadoAjuste;

              Suma := Suma + decPonderado;

              if (Connection.zCommand.RecNo = Connection.zCommand.RecordCount) and (Suma <> dPonderadoAjuste) then
                decPonderado := decPonderado + (dPonderadoAjuste - Suma);

              Connection.zCommand.Edit;
              Connection.zCommand.FieldByName('dPonderado').AsFloat := decPonderado;
              Connection.zCommand.Post;

              Connection.zCommand.Next;
            end;
          end;
        end
        else
        begin
                // Primero el Financiero MN
          dMontoContrato := 0;
          Connection.qryBusca.Active := False;
          Connection.QryBusca.Filtered := False;
          Connection.qryBusca.SQL.Clear;
          Connection.qryBusca.SQL.Add('Select sum(dCantidad * dVentaMN) as dMontoMN From actividadesxorden ' +
            'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
            'and lcalculo=:calculo group by sContrato');
          Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
          Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          Connection.qryBusca.Params.ParamByName('calculo').AsString := sCalcula;
          Connection.qryBusca.Open;
          if Connection.qryBusca.RecordCount > 0 then
            dMontoContratoMN := Connection.qryBusca.FieldValues['dMontoMN'];

          if dMontoContratoMN > 0 then
          begin
            Connection.QryBusca2.Active := False;
            connection.QryBusca2.Filtered := False;
            Connection.QryBusca2.SQL.Clear;
            Connection.QryBusca2.SQL.Add('select dCantidad, dVentaMN, sWbs from actividadesxorden ' +
              'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" And dCantidad <> 0 ' +
              'and lcalculo=:calculo');
            connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Contrato').Value := global_contrato;
            connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.QryBusca2.Params.ParamByName('Calculo').AsString := sCalcula;
            connection.QryBusca2.Open;

            while not connection.QryBusca2.Eof do
            begin
              Connection.zCommand.Active := False;
              connection.zCommand.Filtered := False;
              Connection.zCommand.SQL.Clear;
              Connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = :ponderado ' +
                'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sWbs =:Wbs and sTipoActividad = "Actividad" ' +
                'and lcalculo=:calculo');
              connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              connection.zCommand.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
              connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
              connection.zCommand.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
              connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
              connection.zCommand.Params.ParamByName('Wbs').Value := connection.QryBusca2.FieldValues['sWbs'];
              connection.zCommand.Params.ParamByName('ponderado').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('ponderado').Value := (((connection.QryBusca2.FieldByName('dCantidad').AsFloat * connection.QryBusca2.FieldByName('dVentaMN').AsFloat) / dMontoContratoMN) * dPonderadoAjuste);
              connection.zCommand.Params.ParamByName('calculo').AsString := sCalcula;
              connection.zCommand.ExecSQL;
              connection.QryBusca2.Next;
            end;
          end;

                // Primero el Financiero DLL
          dMontoContrato := 0;
          Connection.qryBusca.Active := False;
          Connection.QryBusca.Filtered := False;
          Connection.qryBusca.SQL.Clear;
          Connection.qryBusca.SQL.Add('Select sum(dCantidad * dVentaDLL) as dMontoDLL From actividadesxorden ' +
            'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
            'and lcalculo=:calculo group by sContrato');
          Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
          Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          Connection.qryBusca.Params.ParamByName('calculo').AsString := sCalcula;
          Connection.qryBusca.Open;
          if Connection.qryBusca.RecordCount > 0 then
            dMontoContratoDLL := Connection.qryBusca.FieldValues['dMontoDLL'];

          if dMontoContratoDLL > 0 then
          begin
            Connection.QryBusca2.Active := False;
            connection.QryBusca2.Filtered := False;
            Connection.QryBusca2.SQL.Clear;
            Connection.QryBusca2.SQL.Add('select dCantidad, dVentaDLL, sWbs, dPonderado from actividadesxorden ' +
              'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" And dCantidad <> 0 ' +
              'and lcalculo=:calculo');
            connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Contrato').Value := global_contrato;
            connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.QryBusca2.Params.ParamByName('calculo').AsString := sCalcula;
            connection.QryBusca2.Open;

            while not connection.QryBusca2.Eof do
            begin
              Connection.zCommand.Active := False;
              connection.zCommand.Filtered := False;
              Connection.zCommand.SQL.Clear;
              Connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = :ponderado ' +
                'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sWbs =:Wbs and sTipoActividad = "Actividad" ' +
                'and lcalculo=:calculo');
              connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              connection.zCommand.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
              connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
              connection.zCommand.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
              connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
              connection.zCommand.Params.ParamByName('Wbs').Value := connection.QryBusca2.FieldValues['sWbs'];
              connection.zCommand.Params.ParamByName('ponderado').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('ponderado').Value := (connection.QryBusca2.FieldByName('dPonderado').AsFloat +
                (((connection.QryBusca2.FieldByName('dCantidad').AsFloat * connection.QryBusca2.FieldByName('dVentaDLL').AsFloat)
                / dMontoContratoMN) * dPonderadoAjuste)) / 2;
              connection.zCommand.Params.ParamByName('calculo').AsString := sCalcula;
              connection.zCommand.ExecSQL;
              connection.QryBusca2.Next;
            end;
          end;

                // Fisico en Moneda Nacional
                //Calculo el monto del programa ...
          Connection.qryBusca.Active := False;
          Connection.QryBusca.Filtered := False;
          Connection.qryBusca.SQL.Clear;
          Connection.qryBusca.SQL.Add('Select sum(dDuracion) as dDuracionTotal From actividadesxorden ' +
            'Where sContrato = :Contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden And sTipoActividad = "Actividad" group by sContrato ' +
            'and lcalculo=:calculo');
          Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
          Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          Connection.qryBusca.Params.ParamByName('Calculo').AsString := sCalcula;
                //Connection.qryBusca.Open ;
          Connection.qryBusca.Open;
          if connection.qryBusca.RecordCount > 0 then
          begin
            Connection.zCommand.Active := False;
            connection.zCommand.Filtered := False;
            Connection.zCommand.SQL.Clear;
            Connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = (dPonderado + (((dDuracion / :duracioncontrato) * :miMaximoPonderado)) / 2) ' +
              'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" and dDuracion <> 0 ' +
              'and lcalculo=:calculo');
            connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
            connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
            connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
            connection.zCommand.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
            connection.zCommand.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.zCommand.Params.ParamByName('duracioncontrato').DataType := ftFloat;
            connection.zCommand.Params.ParamByName('duracioncontrato').Value := Connection.qryBusca.FieldValues['dDuracionTotal'];
            connection.zcommand.params.ParamByName('miMaximoPonderado').DataType := ftFloat;
            connection.zcommand.params.ParamByName('miMaximoPonderado').Value := dPonderadoAjuste;
            connection.zcommand.params.ParamByName('Calculo').AsString := sCalcula;
            connection.zCommand.ExecSQL;
          end
        end;


      Connection.QryBusca2.Active := False;
      Connection.QryBusca2.SQL.Clear;
      Connection.QryBusca2.SQL.Add('Select Distinct sWBS From actividadesxorden ' +
        'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Paquete" Order By iNivel DESC');
      Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      Connection.QryBusca2.Open;
      while not Connection.QryBusca2.Eof do
      begin
        Connection.QryBusca.Active := False;
        Connection.QryBusca.Filtered := False;
        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Add('Select Min(dFechaInicio) as dFechaInicio, Max(dFechaFinal) as dFechaFinal, sum(dPonderado) as dPonderado, ' +
          'sum(dCantidad * dVentaMN) as dMontoMN, sum(dCantidad * dVentaDLL) as dMontoDLL From actividadesxorden ' +
          'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sWBSAnterior = :Paquete ' +
          'and lcalculo=:calculo Group By sWBSAnterior');
        Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
        Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
        Connection.QryBusca.Params.ParamByName('Paquete').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Paquete').Value := Connection.QryBusca2.FieldValues['sWBS'];
        Connection.QryBusca.Params.ParamByName('calculo').AsString := sCalcula;
        Connection.QryBusca.Open;
        if Connection.QryBusca.RecordCount > 0 then
          if (not Connection.QryBusca.FieldByName('dFechaInicio').IsNull) and (not Connection.QryBusca.FieldByName('dFechaFinal').IsNull) then
          begin
            connection.zCommand.Active := False;
            connection.zCommand.Filtered := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Update actividadesxorden SET dFechaInicio = :Inicio, dFechaFinal = :Final, dPonderado = :Ponderado, dVentaMN = :MontoMN, dVentaDLL = :MontoDLL ' +
              ',dDuracion=:Duracion Where sContrato = :Contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden And ' +
              'sWBS = :Paquete And sTipoActividad = "Paquete"');
            connection.zcommand.params.ParamByName('Contrato').DataType := ftString;
            connection.zcommand.params.ParamByName('Contrato').Value := global_contrato;
            connection.zcommand.params.ParamByName('Convenio').DataType := ftString;
            connection.zcommand.params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.zcommand.params.ParamByName('Orden').DataType := ftString;
            connection.zcommand.params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.zcommand.params.ParamByName('Paquete').DataType := ftString;
            connection.zcommand.params.ParamByName('Paquete').Value := Connection.QryBusca2.FieldValues['sWBS'];
            connection.zcommand.params.ParamByName('Inicio').DataType := ftDate;
            connection.zcommand.params.ParamByName('Inicio').Value := Connection.QryBusca.FieldValues['dFechaInicio'];
            connection.zcommand.params.ParamByName('Final').DataType := ftDate;
            connection.zcommand.params.ParamByName('Final').Value := Connection.QryBusca.FieldValues['dFechaFinal'];
            connection.zcommand.ParamByName('Duracion').AsInteger:=DaysBetween(Connection.QryBusca.FieldByName('dFechaFinal').AsDateTime,Connection.QryBusca.FieldByName('dFechaInicio').AsDateTime);
            connection.zcommand.params.ParamByName('Ponderado').DataType := ftFloat;
            if roundTo(Connection.QryBusca.FieldValues['dPonderado'], -2) >= 100 then
              connection.zcommand.params.ParamByName('Ponderado').Value := 100
            else
              connection.zcommand.params.ParamByName('Ponderado').Value := Connection.QryBusca.FieldValues['dPonderado'];
            connection.zcommand.params.ParamByName('MontoMN').DataType := ftFloat;
            connection.zcommand.params.ParamByName('MontoMN').Value := Connection.QryBusca.FieldValues['dMontoMN'];
            connection.zcommand.params.ParamByName('MontoDLL').DataType := ftFloat;
            connection.zcommand.params.ParamByName('MontoDLL').Value := Connection.QryBusca.FieldValues['dMontoDLL'];
            Connection.zCommand.ExecSQL;
          end;
        Connection.QryBusca2.Next
      end;


      dMontoContratoDLL := 0;
      Connection.qryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      Connection.qryBusca.SQL.Clear;
      Connection.qryBusca.SQL.Add('Select sum(dCantidad * dVentaDLL) as dMontoDLL From actividadesxorden ' +
        'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
        'and lcalculo=:calculo group by sContrato');
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      Connection.qryBusca.Params.ParamByName('calculo').AsString := sCalcula;
      Connection.qryBusca.Open;
      if Connection.qryBusca.RecordCount > 0 then
        dMontoContratoDLL := Connection.qryBusca.FieldValues['dMontoDLL'];

      Connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('Select dFechaInicio, dFechaFinal From actividadesxorden ' +
        'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And iNivel = 0');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      Connection.QryBusca.Open;

      frmBarra1.btnRefresh.Click
    end;
  Actividadesxorden.Refresh;
end;

procedure TfrmActividades.PonderadoHorarios;
var
  dMontoContratoMN: Currency;
  dMontoContratoDLL: Currency;
  dPonderadoAjuste,
    Difer, Ponderado,
    decPonderado, Suma: Extended;
  scalcula: string;
  sSumaTotalHrsP: String;
begin
  scalcula := 'Si';
    //Inicia proceso de estructura del proyecto ...
  if Actividadesxorden.RecordCount > 0 then
    if MessageDlg('Desea Ponderar los Conceptos del Contrato Seleccionado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
        // Que ponderados se calcularan ?
        // Sumo todos las partidas anexo que tengan en lCalculo <> Si

      Connection.zCommand.Active := False;
      connection.zCommand.Filtered := False;
      Connection.zCommand.SQL.Clear;
      Connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = 0 ' +
        'Where sContrato = :Contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden');
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
      connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
      connection.zCommand.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      connection.zCommand.Params.ParamByName('orden').DataType := ftString;
      connection.zCommand.Params.ParamByName('orden').Value := tsNumeroOrden.Text;
      connection.zCommand.ExecSQL;
        // and lCalculo = "Si"


      dPonderadoAjuste := 100;
     //   If (Connection.QryBusca.RecordCount > 0) And (Connection.QryBusca.FieldValues['TotalPonderado'] > 0 ) Then
     //         dPonderadoAjuste :=  Connection.QryBusca.FieldValues ['TotalPonderado'] ;

        // Actualizacion de ponderados ....
      dMontoContrato := 0;
      Connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('Select sum(dCantidad * dVentaMN) as dMontoMN From actividadesxorden ' +
        'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
        'and lcalculo=:calculo group by sContrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      Connection.QryBusca.Params.ParamByName('calculo').AsString := sCalcula;
      Connection.QryBusca.Open;


      if Connection.QryBusca.RecordCount > 0 then
        dMontoContratoMN := Connection.QryBusca.FieldValues['dMontoMN'];

      if connection.configuracion.FieldValues['lCalculoPonderado'] = 'Financiero' then
      begin
        if dMontoContratoMN > 0 then
        begin
          connection.zCommand.Active := False;
          connection.zCommand.Filtered := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = (((dCantidad * dVentaMN) / :montocontrato) * :miMaximoPonderado) ' +
            'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" and dCantidad <> 0 ' +
            'and lcalculo=:calculo');
          connection.zcommand.params.ParamByName('Contrato').DataType := ftString;
          connection.zcommand.params.ParamByName('Contrato').Value := global_contrato;
          connection.zcommand.params.ParamByName('Convenio').DataType := ftString;
          connection.zcommand.params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          connection.zcommand.params.ParamByName('Orden').DataType := ftString;
          connection.zcommand.params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          connection.zcommand.params.ParamByName('montocontrato').DataType := ftFloat;
          connection.zcommand.params.ParamByName('montocontrato').Value := dMontoContratoMN;
          connection.zcommand.params.ParamByName('miMaximoPonderado').DataType := ftFloat;
          connection.zcommand.params.ParamByName('miMaximoPonderado').Value := dPonderadoAjuste;
          connection.zcommand.params.ParamByName('calculo').AsString := scalcula;
          connection.zCommand.ExecSQL;
        end;
      end
      else
        if connection.configuracion.FieldValues['lCalculoPonderado'] = 'Duracion' then
        begin
          //Calculo la duración total del programa en horas.
          ActividadesxOrden.First;
          sSumaTotalHrsP := '00:00';
          while Not ActividadesxOrden.Eof do begin
            if (ActividadesxOrden.FieldByName('sTipoActividad').AsString = 'Actividad') AND (ActividadesxOrden.FieldByName('lCalculo').AsString = 'Si') then begin
              Connection.QryBusca.Active := False;
              Connection.QryBusca.SQL.Text :=  '' +
                                                  'SELECT ' +
                                                  '	dFechaInicio, ' +
                                                  '	sHoraInicio, ' +
                                                  '	dFechaFinal, ' +
                                                  '	sHoraFinal, ' +
                                                  '	@FechaInicial := (CONCAT(dFechaInicio, " ", cast(sHoraInicio AS Time))) AS Inicio, ' +
                                                  '	@FechaFinal := (CONCAT(dFechaFinal, " ", cast(sHoraFinal AS Time))) AS Final, ' +
                                                  '	CAST( ' +
                                                  '		TIMEDIFF(@FechaFinal, @FechaInicial) AS CHAR ' +
                                                  '	) AS dDiferenciaDuracion ' +
                                                  'FROM ' +
                                                  '	actividadesxorden_detalle ' +
                                                  'WHERE ' +
                                                  '	sContrato = :Contrato ' +
                                                  '	AND sNumeroOrden = :Orden ' +
                                                  '	AND sWbs = :Wbs ' +
                                                  '	AND sIdConvenio = :Convenio ' +
                                                  'ORDER BY ' +
                                                  '	dFechaInicio, ' +
                                                  '	Time(sHoraInicio) ' +
                                                  '';
              Connection.QryBusca.Params.ParamByName('Contrato').AsString := ActividadesxOrden.FieldByName('sContrato').AsString;
              Connection.QryBusca.Params.ParamByName('Orden').AsString := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
              Connection.QryBusca.Params.ParamByName('Wbs').AsString := ActividadesxOrden.FieldByName('sWbs').AsString;
              Connection.QryBusca.Params.ParamByName('Convenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
              Connection.QryBusca.Open;
              while Not Connection.QryBusca.Eof do begin
                sSumaTotalHrsP := sfnSumaHoras(sSumaTotalHrsP, Connection.QryBusca.FieldByName('dDiferenciaDuracion').AsString);
                Connection.QryBusca.Next;
              end;
            end;
            ActividadesxOrden.Next;
          end;

//          ShowMessage(sSumaTotalHrsP);
//          ShowMessage(FloatToStr(StrToFloat(sSumaTotalHrsP)));

          //Calculo el monto del programa ...
          Connection.QryBusca.Active := False;
          Connection.QryBusca.Filtered := False;
          Connection.QryBusca.SQL.Clear;
          Connection.QryBusca.SQL.Add('Select sum(dDuracion) as dDAnt, (CAST("'+sSumaTotalHrsP+'" AS TIME) + "0") / 10000 AS dDuracionTotal From actividadesxorden ' +
            'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
            'and lcalculo=:calculo group by sContrato');
          Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
          Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
          Connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
          Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          Connection.QryBusca.Params.ParamByName('calculo').AsString := sCalcula;
          Connection.QryBusca.Open;
//          ShowMessage(FloatToStr(Connection.zCommand.FieldByName('dDuracion').AsFloat));
          if connection.QryBusca.RecordCount > 0 then
          begin
            connection.zCommand.Active := False;
            connection.zCommand.Filtered := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Text := 'select * from actividadesxorden where sContrato = :Contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden And sTipoActividad = "Actividad" ' + //and dDuracion <> 0 
              'and lcalculo=:calculo order by iItemOrden';
            connection.zcommand.params.ParamByName('Contrato').DataType := ftString;
            connection.zcommand.params.ParamByName('Contrato').Value := global_contrato;
            connection.zcommand.params.ParamByName('Convenio').DataType := ftString;
            connection.zcommand.params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.zcommand.params.ParamByName('Orden').DataType := ftString;
            connection.zcommand.params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.zcommand.params.ParamByName('calculo').AsString := sCalcula;
            connection.zCommand.Open;

            Difer := 0; // Diferencia para ajuste entre partidas
            Ponderado := 0; // Almacenamiento de ponderado total calculado
            Suma := 0;
            while not Connection.zCommand.Eof do
            begin
              sSumaTotalHrsP := '00:00';
              Connection.QryBusca2.Active := False;
              Connection.QryBusca2.SQL.Text :=  '' +
                                                  'SELECT ' +
                                                  '	dFechaInicio, ' +
                                                  '	sHoraInicio, ' +
                                                  '	dFechaFinal, ' +
                                                  '	sHoraFinal, ' +
                                                  '	@FechaInicial := (CONCAT(dFechaInicio, " ", cast(sHoraInicio AS Time))) AS Inicio, ' +
                                                  '	@FechaFinal := (CONCAT(dFechaFinal, " ", cast(sHoraFinal AS Time))) AS Final, ' +
                                                  '	CAST( ' +
                                                  '		TIMEDIFF(@FechaFinal, @FechaInicial) AS CHAR ' +
                                                  '	) AS dDiferenciaDuracion ' +
                                                  'FROM ' +
                                                  '	actividadesxorden_detalle ' +
                                                  'WHERE ' +
                                                  '	sContrato = :Contrato ' +
                                                  '	AND sNumeroOrden = :Orden ' +
                                                  '	AND sWbs = :Wbs ' +
                                                  '	AND sIdConvenio = :Convenio ' +
                                                  'ORDER BY ' +
                                                  '	dFechaInicio, ' +
                                                  '	Time(sHoraInicio) ' +
                                                  '';
              Connection.QryBusca2.Params.ParamByName('Contrato').AsString := ActividadesxOrden.FieldByName('sContrato').AsString;
              Connection.QryBusca2.Params.ParamByName('Orden').AsString := ActividadesxOrden.FieldByName('sNumeroOrden').AsString;
              Connection.QryBusca2.Params.ParamByName('Wbs').AsString := Connection.zCommand.FieldByName('sWbs').AsString;
              Connection.QryBusca2.Params.ParamByName('Convenio').AsString := ActividadesxOrden.FieldByName('sIdConvenio').AsString;
              Connection.QryBusca2.Open;
              while Not Connection.QryBusca2.Eof do begin
                sSumaTotalHrsP := sfnSumaHoras(sSumaTotalHrsP, Connection.QryBusca2.FieldByName('dDiferenciaDuracion').AsString);
                Connection.QryBusca2.Next;
              end;

              Connection.QryBusca2.SQL.Text := 'SELECT (CAST("'+sSumaTotalHrsP+'" AS TIME) + "0") / 10000 AS dDuracionPartida;';
              Connection.QryBusca2.Open;



              Ponderado := (Connection.QryBusca2.FieldByName('dDuracionPartida').AsFloat / Connection.QryBusca.FieldValues['dDuracionTotal']);
              Ponderado := Ponderado + Difer; // Sumar la diferencia anterior para ajuste automático
              decPonderado := Trunc(Ponderado * 1000000) / 1000000;
              Difer := Ponderado - decPonderado;
              decPonderado := decPonderado * dPonderadoAjuste;

              Suma := Suma + decPonderado;

              if (Connection.zCommand.RecNo = Connection.zCommand.RecordCount) and (Suma <> dPonderadoAjuste) then
                decPonderado := decPonderado + (dPonderadoAjuste - Suma);

              Connection.zCommand.Edit;
              Connection.zCommand.FieldByName('dPonderado').AsFloat := decPonderado;
              Connection.zCommand.Post;

              Connection.zCommand.Next;
            end;
          end;
        end
        else
        begin
                // Primero el Financiero MN
          dMontoContrato := 0;
          Connection.qryBusca.Active := False;
          Connection.QryBusca.Filtered := False;
          Connection.qryBusca.SQL.Clear;
          Connection.qryBusca.SQL.Add('Select sum(dCantidad * dVentaMN) as dMontoMN From actividadesxorden ' +
            'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
            'and lcalculo=:calculo group by sContrato');
          Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
          Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          Connection.qryBusca.Params.ParamByName('calculo').AsString := sCalcula;
          Connection.qryBusca.Open;
          if Connection.qryBusca.RecordCount > 0 then
            dMontoContratoMN := Connection.qryBusca.FieldValues['dMontoMN'];

          if dMontoContratoMN > 0 then
          begin
            Connection.QryBusca2.Active := False;
            connection.QryBusca2.Filtered := False;
            Connection.QryBusca2.SQL.Clear;
            Connection.QryBusca2.SQL.Add('select dCantidad, dVentaMN, sWbs from actividadesxorden ' +
              'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" And dCantidad <> 0 ' +
              'and lcalculo=:calculo');
            connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Contrato').Value := global_contrato;
            connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.QryBusca2.Params.ParamByName('Calculo').AsString := sCalcula;
            connection.QryBusca2.Open;

            while not connection.QryBusca2.Eof do
            begin
              Connection.zCommand.Active := False;
              connection.zCommand.Filtered := False;
              Connection.zCommand.SQL.Clear;
              Connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = :ponderado ' +
                'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sWbs =:Wbs and sTipoActividad = "Actividad" ' +
                'and lcalculo=:calculo');
              connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              connection.zCommand.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
              connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
              connection.zCommand.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
              connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
              connection.zCommand.Params.ParamByName('Wbs').Value := connection.QryBusca2.FieldValues['sWbs'];
              connection.zCommand.Params.ParamByName('ponderado').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('ponderado').Value := (((connection.QryBusca2.FieldByName('dCantidad').AsFloat * connection.QryBusca2.FieldByName('dVentaMN').AsFloat) / dMontoContratoMN) * dPonderadoAjuste);
              connection.zCommand.Params.ParamByName('calculo').AsString := sCalcula;
              connection.zCommand.ExecSQL;
              connection.QryBusca2.Next;
            end;
          end;

                // Primero el Financiero DLL
          dMontoContrato := 0;
          Connection.qryBusca.Active := False;
          Connection.QryBusca.Filtered := False;
          Connection.qryBusca.SQL.Clear;
          Connection.qryBusca.SQL.Add('Select sum(dCantidad * dVentaDLL) as dMontoDLL From actividadesxorden ' +
            'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
            'and lcalculo=:calculo group by sContrato');
          Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
          Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          Connection.qryBusca.Params.ParamByName('calculo').AsString := sCalcula;
          Connection.qryBusca.Open;
          if Connection.qryBusca.RecordCount > 0 then
            dMontoContratoDLL := Connection.qryBusca.FieldValues['dMontoDLL'];

          if dMontoContratoDLL > 0 then
          begin
            Connection.QryBusca2.Active := False;
            connection.QryBusca2.Filtered := False;
            Connection.QryBusca2.SQL.Clear;
            Connection.QryBusca2.SQL.Add('select dCantidad, dVentaDLL, sWbs, dPonderado from actividadesxorden ' +
              'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" And dCantidad <> 0 ' +
              'and lcalculo=:calculo');
            connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Contrato').Value := global_contrato;
            connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
            connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.QryBusca2.Params.ParamByName('calculo').AsString := sCalcula;
            connection.QryBusca2.Open;

            while not connection.QryBusca2.Eof do
            begin
              Connection.zCommand.Active := False;
              connection.zCommand.Filtered := False;
              Connection.zCommand.SQL.Clear;
              Connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = :ponderado ' +
                'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sWbs =:Wbs and sTipoActividad = "Actividad" ' +
                'and lcalculo=:calculo');
              connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
              connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
              connection.zCommand.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
              connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
              connection.zCommand.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
              connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
              connection.zCommand.Params.ParamByName('Wbs').Value := connection.QryBusca2.FieldValues['sWbs'];
              connection.zCommand.Params.ParamByName('ponderado').DataType := ftFloat;
              connection.zCommand.Params.ParamByName('ponderado').Value := (connection.QryBusca2.FieldByName('dPonderado').AsFloat +
                (((connection.QryBusca2.FieldByName('dCantidad').AsFloat * connection.QryBusca2.FieldByName('dVentaDLL').AsFloat)
                / dMontoContratoMN) * dPonderadoAjuste)) / 2;
              connection.zCommand.Params.ParamByName('calculo').AsString := sCalcula;
              connection.zCommand.ExecSQL;
              connection.QryBusca2.Next;
            end;
          end;

                // Fisico en Moneda Nacional
                //Calculo el monto del programa ...
          Connection.qryBusca.Active := False;
          Connection.QryBusca.Filtered := False;
          Connection.qryBusca.SQL.Clear;
          Connection.qryBusca.SQL.Add('Select sum(dDuracion) as dDuracionTotal From actividadesxorden ' +
            'Where sContrato = :Contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden And sTipoActividad = "Actividad" group by sContrato ' +
            'and lcalculo=:calculo');
          Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
          Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
          Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
          Connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
          Connection.qryBusca.Params.ParamByName('Calculo').AsString := sCalcula;
                //Connection.qryBusca.Open ;
          Connection.qryBusca.Open;
          if connection.qryBusca.RecordCount > 0 then
          begin
            Connection.zCommand.Active := False;
            connection.zCommand.Filtered := False;
            Connection.zCommand.SQL.Clear;
            Connection.zCommand.SQL.Add('update actividadesxorden SET dPonderado = (dPonderado + (((dDuracion / :duracioncontrato) * :miMaximoPonderado)) / 2) ' +
              'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" and dDuracion <> 0 ' +
              'and lcalculo=:calculo');
            connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
            connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato;
            connection.zCommand.Params.ParamByName('Convenio').DataType := ftString;
            connection.zCommand.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
            connection.zCommand.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.zCommand.Params.ParamByName('duracioncontrato').DataType := ftFloat;
            connection.zCommand.Params.ParamByName('duracioncontrato').Value := Connection.qryBusca.FieldValues['dDuracionTotal'];
            connection.zcommand.params.ParamByName('miMaximoPonderado').DataType := ftFloat;
            connection.zcommand.params.ParamByName('miMaximoPonderado').Value := dPonderadoAjuste;
            connection.zcommand.params.ParamByName('Calculo').AsString := sCalcula;
            connection.zCommand.ExecSQL;
          end
        end;


      Connection.QryBusca2.Active := False;
      Connection.QryBusca2.SQL.Clear;
      Connection.QryBusca2.SQL.Add('Select Distinct sWBS From actividadesxorden ' +
        'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Paquete" Order By iNivel DESC');
      Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      Connection.QryBusca2.Open;
      while not Connection.QryBusca2.Eof do
      begin
        Connection.QryBusca.Active := False;
        Connection.QryBusca.Filtered := False;
        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Add('Select Min(dFechaInicio) as dFechaInicio, Max(dFechaFinal) as dFechaFinal, sum(dPonderado) as dPonderado, ' +
          'sum(dCantidad * dVentaMN) as dMontoMN, sum(dCantidad * dVentaDLL) as dMontoDLL From actividadesxorden ' +
          'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sWBSAnterior = :Paquete ' +
          'and lcalculo=:calculo Group By sWBSAnterior');
        Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
        Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
        Connection.QryBusca.Params.ParamByName('Paquete').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Paquete').Value := Connection.QryBusca2.FieldValues['sWBS'];
        Connection.QryBusca.Params.ParamByName('calculo').AsString := sCalcula;
        Connection.QryBusca.Open;
        if Connection.QryBusca.RecordCount > 0 then
          if (not Connection.QryBusca.FieldByName('dFechaInicio').IsNull) and (not Connection.QryBusca.FieldByName('dFechaFinal').IsNull) then
          begin
            connection.zCommand.Active := False;
            connection.zCommand.Filtered := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Update actividadesxorden SET dFechaInicio = :Inicio, dFechaFinal = :Final, dPonderado = :Ponderado, dVentaMN = :MontoMN, dVentaDLL = :MontoDLL ' +
              'Where sContrato = :Contrato And sIdConvenio = :Convenio and sNumeroOrden =:Orden And ' +
              'sWBS = :Paquete And sTipoActividad = "Paquete"');
            connection.zcommand.params.ParamByName('Contrato').DataType := ftString;
            connection.zcommand.params.ParamByName('Contrato').Value := global_contrato;
            connection.zcommand.params.ParamByName('Convenio').DataType := ftString;
            connection.zcommand.params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
            connection.zcommand.params.ParamByName('Orden').DataType := ftString;
            connection.zcommand.params.ParamByName('Orden').Value := tsNumeroOrden.Text;
            connection.zcommand.params.ParamByName('Paquete').DataType := ftString;
            connection.zcommand.params.ParamByName('Paquete').Value := Connection.QryBusca2.FieldValues['sWBS'];
            connection.zcommand.params.ParamByName('Inicio').DataType := ftDate;
            connection.zcommand.params.ParamByName('Inicio').Value := Connection.QryBusca.FieldValues['dFechaInicio'];
            connection.zcommand.params.ParamByName('Final').DataType := ftDate;
            connection.zcommand.params.ParamByName('Final').Value := Connection.QryBusca.FieldValues['dFechaFinal'];
            connection.zcommand.params.ParamByName('Ponderado').DataType := ftFloat;
            if roundTo(Connection.QryBusca.FieldValues['dPonderado'], -2) >= 100 then
              connection.zcommand.params.ParamByName('Ponderado').Value := 100
            else
              connection.zcommand.params.ParamByName('Ponderado').Value := Connection.QryBusca.FieldValues['dPonderado'];
            connection.zcommand.params.ParamByName('MontoMN').DataType := ftFloat;
            connection.zcommand.params.ParamByName('MontoMN').Value := Connection.QryBusca.FieldValues['dMontoMN'];
            connection.zcommand.params.ParamByName('MontoDLL').DataType := ftFloat;
            connection.zcommand.params.ParamByName('MontoDLL').Value := Connection.QryBusca.FieldValues['dMontoDLL'];
            Connection.zCommand.ExecSQL;
          end;
        Connection.QryBusca2.Next
      end;


      dMontoContratoDLL := 0;
      Connection.qryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      Connection.qryBusca.SQL.Clear;
      Connection.qryBusca.SQL.Add('Select sum(dCantidad * dVentaDLL) as dMontoDLL From actividadesxorden ' +
        'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And sTipoActividad = "Actividad" ' +
        'and lcalculo=:calculo group by sContrato');
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      Connection.qryBusca.Params.ParamByName('calculo').AsString := sCalcula;
      Connection.qryBusca.Open;
      if Connection.qryBusca.RecordCount > 0 then
        dMontoContratoDLL := Connection.qryBusca.FieldValues['dMontoDLL'];

      Connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('Select dFechaInicio, dFechaFinal From actividadesxorden ' +
        'Where sContrato = :Contrato and sNumeroOrden =:Orden And sIdConvenio = :Convenio And iNivel = 0');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
      Connection.QryBusca.Open;


      if Connection.QryBusca.RecordCount > 0 then
      begin
            //Actualizo el convenio
        connection.zCommand.Active := False;
        connection.zCommand.Filtered := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Update convenios SET dFechaInicio = :Inicio, dFechaFinal = :Final, dMontoMN = :MontoMN, dMontoDLL = :MontoDLL ' +
          'Where sContrato = :Contrato And sIdConvenio = :Convenio');
        connection.zcommand.params.ParamByName('Contrato').DataType := ftString;
        connection.zcommand.params.ParamByName('Contrato').Value := global_contrato;
        connection.zcommand.params.ParamByName('Convenio').DataType := ftString;
        connection.zcommand.params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        connection.zcommand.params.ParamByName('Inicio').DataType := ftDate;
        connection.zcommand.params.ParamByName('Inicio').Value := Connection.QryBusca.FieldValues['dFechaInicio'];
        connection.zcommand.params.ParamByName('Final').DataType := ftDate;
        connection.zcommand.params.ParamByName('Final').Value := Connection.QryBusca.FieldValues['dFechaFinal'];
        connection.zcommand.params.ParamByName('MontoMN').DataType := ftFloat;
        connection.zcommand.params.ParamByName('MontoMN').Value := dMontoContratoMN;
        connection.zcommand.params.ParamByName('MontoDLL').DataType := ftFloat;
        connection.zcommand.params.ParamByName('MontoDLL').Value := dMontoContratoDLL;
        Connection.zCommand.ExecSQL;
      end;
      frmBarra1.btnRefresh.Click
    end;
  Actividadesxorden.Refresh;

end;

procedure TfrmActividades.ConsultaFolios;
begin
    ActividadesxOrden.Active := False;
    ActividadesxOrden.Params.ParamByName('Contrato').DataType := ftString;
    ActividadesxOrden.Params.ParamByName('Contrato').Value    := Global_Contrato;
    ActividadesxOrden.Params.ParamByName('Orden').DataType    := ftString;
    ActividadesxOrden.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text;
    ActividadesxOrden.Open;

    isOpen:=true;
    ActividadesxOrdenAfterScroll(ActividadesxOrden);

    Paquetes.Active := False;
    Paquetes.Params.ParamByName('contrato').DataType := ftString;
    Paquetes.Params.ParamByName('contrato').Value    := global_contrato;
    Paquetes.Params.ParamByName('orden').DataType    := ftString;
    Paquetes.Params.ParamByName('orden').Value       := tsNumeroOrden.Text;
    Paquetes.Open;
end;

procedure TfrmActividades.ConsultaReprogramacion;
begin
    zqReprogramacion.Active := False;
    zqReprogramacion.Params.ParamByName('Contrato').DataType := ftString;
    zqReprogramacion.Params.ParamByName('Contrato').Value    := Global_Contrato;
    zqReprogramacion.Params.ParamByName('Folio').DataType    := ftString;
    zqReprogramacion.Params.ParamByName('folio').Value       := tsNumeroOrden.Text;
    zqReprogramacion.Open;

    if zqReprogramacion.FieldByName('sIdConvenio').AsString <> '1' then
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('update convenios set sIdConvenio = "1" where sContrato =:Contrato and sNumeroOrden =:folio ');
        connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
        connection.zCommand.ParamByName('Folio').AsString    := tsNumeroOrden.Text;
        connection.zCommand.ExecSQL;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('update actividadesxorden set sIdConvenio = "1" where sContrato =:Contrato and sIdConvenio =:convenio and sNumeroOrden =:folio ');
        connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
        connection.zCommand.ParamByName('Convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        connection.zCommand.ParamByName('Folio').AsString    := tsNumeroOrden.Text;
        connection.zCommand.ExecSQL;

        zqReprogramacion.Active := False;
        zqReprogramacion.Params.ParamByName('Contrato').DataType := ftString;
        zqReprogramacion.Params.ParamByName('Contrato').Value    := Global_Contrato;
        zqReprogramacion.Params.ParamByName('Folio').DataType    := ftString;
        zqReprogramacion.Params.ParamByName('folio').Value       := tsNumeroOrden.Text;
        zqReprogramacion.Open;
    end;


end;


end.

