unit frm_inteligent;

interface

uses                                                 
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, global, ComCtrls, frm_connection, DB, ExtCtrls,
  ImgList, StdCtrls, 
  Sockets, 
  Mask, StoHtmlHelp, 
  AdvToolBar, AdvToolBarStylers, frxpngimage, 
  UnitExcepciones, frm_SintesisGerencial,
  AdvMenus,
  JvBackgrounds, Buttons, frm_catalogoerrores, iniFiles, ZConnection, unitmanejofondo, 
  DBCtrls, DBGrids, rxToolEdit, rxCurrEdit, RXDBCtrl,
  JvComponentBase, CalcEvent, AppEvnts,frm_sincinformes,
  JvAppStorage, frxClass, jpeg, dxSkinsCore, dxSkinBlack, dxSkinBlue,
  dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide,
  dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle, dxSkinFoggy,
  dxSkinGlassOceans, dxSkinHighContrast, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMetropolis,
  dxSkinMetropolisDark, dxSkinMoneyTwins, dxSkinOffice2007Black,
  dxSkinOffice2007Blue, dxSkinOffice2007Green, dxSkinOffice2007Pink,
  dxSkinOffice2007Silver, dxSkinOffice2010Black, dxSkinOffice2010Blue,
  dxSkinOffice2010Silver, dxSkinOffice2013DarkGray, dxSkinOffice2013LightGray,
  dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic,
  dxSkinSharp, dxSkinSharpPlus, dxSkinSilver, dxSkinSpringTime, dxSkinStardust,
  dxSkinSummer2008, dxSkinTheAsphaltWorld, dxSkinsDefaultPainters,
  dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint, dxSkinXmas2008Blue,
  dxSkinsdxBarPainter, dxBar, cxClasses, cxGraphics, Frm_ImportaProject;
       ///gamael
type
    TfrmInteligent = class(TForm)
    Status: TStatusBar;
    mInteligent: TMainMenu;
    mnConfiguracion: TMenuItem;
    cEmbarcaciones: TMenuItem;
    mnPersonal: TMenuItem;
    cPersonal: TMenuItem;
    cPaquetesPer: TMenuItem;
    opProgramacion: TMenuItem;
    cPernoctan: TMenuItem;
    mnEquipos: TMenuItem;
    cEquipos: TMenuItem;
    cPaquetesEq: TMenuItem;
    N1: TMenuItem;
    cPlataformas: TMenuItem;
    cCuentas: TMenuItem;
    mnObra: TMenuItem;
    opSQLAnexo: TMenuItem;
    N37: TMenuItem;
    cAnexo: TMenuItem;
    cVerifica: TMenuItem;
    N36: TMenuItem;
    cOrdenes: TMenuItem;
    cActividades: TMenuItem;
    mnTarifaDiaria: TMenuItem;
    Firmantes1: TMenuItem;
    N19: TMenuItem;
    JornadasEspeciales1: TMenuItem;
    MovtosdeEmbarcacin1: TMenuItem;
    FasesxOrden1: TMenuItem;
    N21: TMenuItem;
    BitacoradeActividades1: TMenuItem;
    N10: TMenuItem;
    ReporteDiario2: TMenuItem;
    N3: TMenuItem;
    Estimaciones5: TMenuItem;
    GeneradoresdeObra1: TMenuItem;
    Estimaciones2: TMenuItem;
    NotasdeCampo1: TMenuItem;
    mnPrecioU: TMenuItem;
    opPermisos: TMenuItem;
    N26: TMenuItem;
    mnOperaciones: TMenuItem;
    adDistPrograma: TMenuItem;
    N23: TMenuItem;
    adReg02: TMenuItem;
    adReg03: TMenuItem;
    N30: TMenuItem;
    cPaquetesAct: TMenuItem;
    mnAvances: TMenuItem;
    cAvContrato: TMenuItem;
    cAvOrden: TMenuItem;
    mnConsultas: TMenuItem;
    cConsulta1: TMenuItem;
    cConsulta3: TMenuItem;
    N16: TMenuItem;
    mnDiario: TMenuItem;
    opFirmas: TMenuItem;
    N42: TMenuItem;
    rInstalado: TMenuItem;
    N14: TMenuItem;
    rDiario: TMenuItem;
    qComentarios: TMenuItem;
    N13: TMenuItem;
    opValida: TMenuItem;
    opAbre: TMenuItem;
    mnEstimaciones: TMenuItem;
    opProyeccion: TMenuItem;
    N15: TMenuItem;
    opEstimaciones: TMenuItem;
    opTiemposM: TMenuItem;
    N24: TMenuItem;
    N41: TMenuItem;
    opComparativo1: TMenuItem;
    rComparativo: TMenuItem;
    mnAlmacen: TMenuItem;
    opAvisodeEmb: TMenuItem;
    mnHerramientas: TMenuItem;
    adSql: TMenuItem;
    adImportar: TMenuItem;
    adExportar: TMenuItem;
    adImportarOk: TMenuItem;
    mnPanel: TMenuItem;
    mnAdmon: TMenuItem;
    adDeptos: TMenuItem;
    N49: TMenuItem;
    adProgramas: TMenuItem;
    adUsuariosC: TMenuItem;
    N50: TMenuItem;
    adUsuarios: TMenuItem;
    adGrupos: TMenuItem;
    N51: TMenuItem;
    adGrupoP: TMenuItem;
    N7: TMenuItem;
    adContratos: TMenuItem;
    adConfiguracion: TMenuItem;
    adTurnos: TMenuItem;
    adFestivos: TMenuItem;
    N44: TMenuItem;
    mnAdmonPEP: TMenuItem;
    adResidencias: TMenuItem;
    N46: TMenuItem;
    cFactorCosto: TMenuItem;
    N45: TMenuItem;
    adTiposMov: TMenuItem;
    N17: TMenuItem;
    cProgPlaticas: TMenuItem;
    mnGerencial: TMenuItem;
    geAvFiFin: TMenuItem;
    N48: TMenuItem;
    gePersonalProg: TMenuItem;
    N53: TMenuItem;
    grPenasRet: TMenuItem;
    mnSistema: TMenuItem;
    sAcerca: TMenuItem;
    N6: TMenuItem;
    sLogin: TMenuItem;
    sSeleccion: TMenuItem;
    sSalir: TMenuItem;
    ImInteligent: TImageList;
    AccesoObra: TMenuItem;
    adReg04: TMenuItem;
    opValidaEst: TMenuItem;
    opOrdenCam: TMenuItem;
    N54: TMenuItem;
    cProveedores: TMenuItem;
    cConsulta5: TMenuItem;
    N27: TMenuItem;
    cConsulta4: TMenuItem;
    cConsulta6: TMenuItem;
    adExportar2: TMenuItem;
    imgKardex: TImage;
    sCambiaP: TMenuItem;
    adCancelacion: TMenuItem;
    rGerencial: TMenuItem;
    opRequisiciones: TMenuItem;
    opPedidos: TMenuItem;
    opComparativo5: TMenuItem;
    rSintesis: TMenuItem;
    sTips: TMenuItem;
    sWarning: TMenuItem;
    adAvisos: TMenuItem;
    adActivos: TMenuItem;
    opFiltro: TMenuItem;
    N5: TMenuItem;
    opInventario: TMenuItem;
    opadmonCatalogo: TMenuItem;
    N2: TMenuItem;
    opGeneradorSub: TMenuItem;
    N8: TMenuItem;
    adSubContratos: TMenuItem;
    opAyuda: TMenuItem;
    mnConsumibles: TMenuItem;
    ProyecciondeAct: TMenuItem;
    ControlPers: TMenuItem;
    AsignaciondeOrd: TMenuItem;
    Tripulacion: TMenuItem;
    reporteBarco: TMenuItem;
    OpCrearGenerado: TMenuItem;
    opGeneradores: TMenuItem;
    opHistorico: TMenuItem;
    rPartidasIsom: TMenuItem;
    mnuConversiones: TMenuItem;
    cFases: TMenuItem;
    oficmodif: TMenuItem;
    MovPerofic: TMenuItem;
    mnuKardex: TMenuItem;
    segprograma: TMenuItem;
    CargaPrograma: TMenuItem;
    mnuAgrupacionP: TMenuItem;
    Tiempo: TTimer;
    Graficador: TMenuItem;
    ChartPro: TMenuItem;
    MnuEntAlmace: TMenuItem;
    MnuSalAlmacen: TMenuItem;
    MnuCatalogodeMo: TMenuItem;
    N11: TMenuItem;
    MnuFamiliadePro: TMenuItem;
    MnuAlmacen: TMenuItem;
    imgIteliBarra: TImageList;
    AdvToolBarFantasyStyler1: TAdvToolBarFantasyStyler;
    mnCopiarParametro: TMenuItem;
    menuEstimacion: TMenuItem;
    optEstimaciones: TMenuItem;
    optValida: TMenuItem;
    optDesautoriza: TMenuItem;
    SubAdministrador: TMenuItem;
    CatalogoErrores: TMenuItem;
    OpenDialog1: TOpenDialog;
    MnuImpAvCont: TMenuItem;
    mnuPersonal2: TMenuItem;
    subMateriales: TMenuItem;
    N9: TMenuItem;
    N22: TMenuItem;
    InformedeSincronizado1: TMenuItem;
    GerencialBarco1: TMenuItem;
    ReportedeProduccion: TMenuItem;
    N28: TMenuItem;
    JvAppStorage1: TJvAppStorage;
    frxReport1: TfrxReport;
    MOE1: TMenuItem;
    AgrupadordePersonal1: TMenuItem;
    NombresdeFirmantes1: TMenuItem;
    Unificadordeequipos1: TMenuItem;
    ContenidoNotaCampo1: TMenuItem;
    Cargadeformatos1: TMenuItem;
    ListadodePersonal1: TMenuItem;
    mniListadoPersonal: TMenuItem;
    iconosPop: TcxImageList;
    iconos2: TcxImageList;
    dxBarManager1: TdxBarManager;
    dxBarManager1Bar1: TdxBar;
    dxBarLargeButton2: TdxBarLargeButton;
    dxBarLargeButton3: TdxBarLargeButton;
    dxBarLargeButton4: TdxBarLargeButton;
    dxBarLargeButton5: TdxBarLargeButton;
    dxBarLargeButton6: TdxBarLargeButton;
    dxBarLargeButton7: TdxBarLargeButton;
    dxBarLargeButton8: TdxBarLargeButton;
    dxBarLargeButton9: TdxBarLargeButton;
    dxBarLargeButton10: TdxBarLargeButton;
    dxBarLargeButton11: TdxBarLargeButton;
    dxBarLargeButton12: TdxBarLargeButton;
    dxBarLargeButton13: TdxBarLargeButton;
    dxBarLargeButton14: TdxBarLargeButton;
    dxBarLargeButton15: TdxBarLargeButton;
    dxBarLargeButton16: TdxBarLargeButton;
    dxBarLargeButton17: TdxBarLargeButton;
    dxBarLargeButton18: TdxBarLargeButton;
    dxBarLargeButton19: TdxBarLargeButton;
    dxBarLargeButton20: TdxBarLargeButton;
    dxBarLargeButton21: TdxBarLargeButton;
    dxBarLargeButton22: TdxBarLargeButton;
    dxBarLargeButton23: TdxBarLargeButton;
    dxBarLargeButton24: TdxBarLargeButton;
    dxBarLargeButton25: TdxBarLargeButton;
    dxBarLargeButton26: TdxBarLargeButton;
    dxBarLargeButton27: TdxBarLargeButton;
    dxBarLargeButton1: TdxBarLargeButton;
    dxBarLargeButton28: TdxBarLargeButton;
    JvBackground1: TJvBackground;
    inteligentpop: TPopupMenu;
    Cambiarimagendefondo1: TMenuItem;
    cambiarmododefondo1: TMenuItem;
    estirado1: TMenuItem;
    centrado1: TMenuItem;
    mosaico1: TMenuItem;
    Ventanasen1: TMenuItem;
    Cascada1: TMenuItem;
    MosaicoVertical1: TMenuItem;
    MosaicoHorizontal1: TMenuItem;
    Irareportesdiarios1: TMenuItem;
    Iraestimaciones1: TMenuItem;
    Irageneradores1: TMenuItem;
    Irageneradoresdeinformes1: TMenuItem;
    Button1: TButton;
    RecursosPT1: TMenuItem;
    MovimientosdeBarcoPT1: TMenuItem;
    MovimientosdeBarcoPT2: TMenuItem;
    EquipoPT1: TMenuItem;
    PernoctasPT1: TMenuItem;
    dxBarButton1: TdxBarButton;
    dxBarButton2: TdxBarButton;
    dxBarLargeButton29: TdxBarLargeButton;
    ControldeCalidad1: TMenuItem;
    ControldeCalidadRIR1: TMenuItem;
    ReprogramacionesXFolios1: TMenuItem;
    HorariosGerenciales1: TMenuItem;
    cProgramados: TMenuItem;
    ReportesProduccionSabanas1: TMenuItem;
    ResumendePersonal1: TMenuItem;
    mniGeneradores: TMenuItem;
    mniRendimiento: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure adConfiguracionClick(Sender: TObject);
    procedure adContratosClick(Sender: TObject);
    procedure adDeptosClick(Sender: TObject);
    procedure adUsuariosClick(Sender: TObject);
    procedure cPersonalClick(Sender: TObject);
    procedure cEquiposClick(Sender: TObject);
    procedure cEmbarcacionesClick(Sender: TObject);
    procedure cPlataformasClick(Sender: TObject);
    procedure cCuentasClick(Sender: TObject);
    procedure cPernoctanClick(Sender: TObject);
    procedure cOrdenesClick(Sender: TObject);
    procedure cActividadesClick(Sender: TObject);
    procedure Firmantes1Click(Sender: TObject);
    procedure adTiposMovClick(Sender: TObject);
    procedure opMuertoClick(Sender: TObject);
    procedure sSeleccionClick(Sender: TObject);
    procedure sLoginClick(Sender: TObject);
    procedure sSalirClick(Sender: TObject);
    procedure sAcercaClick(Sender: TObject);
    procedure CatalogodeProveedores1Click(Sender: TObject);
    procedure opPermisosClick(Sender: TObject);
    procedure rDiarioClick(Sender: TObject);
    procedure opValidaClick(Sender: TObject);
    procedure opAbreClick(Sender: TObject);
    procedure cPaquetesEqClick(Sender: TObject);
    procedure cPaquetesPerClick(Sender: TObject);
    procedure cConsulta1Click(Sender: TObject);
    procedure adReg02Click(Sender: TObject);
    procedure opFirmasClick(Sender: TObject);
    procedure cAnexoClick(Sender: TObject);
    procedure cVerificaClick(Sender: TObject);
    procedure adFestivosClick(Sender: TObject);
    procedure cProgPlaticasClick(Sender: TObject);
    procedure opComparativo1Click(Sender: TObject);
    procedure adTurnosClick(Sender: TObject);
    procedure adReg03Click(Sender: TObject);
    procedure opEstimacionesClick(Sender: TObject);
    procedure cPaquetesActClick(Sender: TObject);
    procedure adDistProgramaClick(Sender: TObject);
    procedure rComparativoClick(Sender: TObject);
    procedure cAvContratoClick(Sender: TObject);
    procedure Firmantes3Click(Sender: TObject);
    procedure opTiemposMClick(Sender: TObject);
    procedure qComentariosClick(Sender: TObject);
    procedure JornadasEspeciales1Click(Sender: TObject);
    procedure SpeedItem2Click(Sender: TObject);
    procedure SpeedItem3Click(Sender: TObject);
    procedure SpeedItem5Click(Sender: TObject);
    procedure SpeedItem6Click(Sender: TObject);
    procedure SpeedItem7Click(Sender: TObject);
    procedure SpeedItem8Click(Sender: TObject);
    procedure SpeedItem9Click(Sender: TObject);
    procedure SpeedItem12Click(Sender: TObject);
    procedure SpeedItem11Click(Sender: TObject);
    procedure SpeedItem14Click(Sender: TObject);
    procedure SpeedItem10Click(Sender: TObject);
    procedure opProyeccionClick(Sender: TObject);
    procedure adSqlClick(Sender: TObject);
    procedure opSQLAnexoClick(Sender: TObject);
    procedure cFactorCostoClick(Sender: TObject);
    procedure geAvFiFinClick(Sender: TObject);
    procedure adImportarClick(Sender: TObject);
    procedure adExportarClick(Sender: TObject);
    procedure adResidenciasClick(Sender: TObject);
    procedure adProgramasClick(Sender: TObject);
    procedure adGruposClick(Sender: TObject);
    procedure adGrupoPClick(Sender: TObject);
    procedure opProgramacionClick(Sender: TObject);
    procedure gePersonalProgClick(Sender: TObject);
    procedure adUsuariosCClick(Sender: TObject);
    procedure grPenasRetClick(Sender: TObject);
    procedure cConsulta4Click(Sender: TObject);
    procedure adImportarOkClick(Sender: TObject);
    procedure SpeedItem4Click0(Sender: TObject);
    procedure SpeedItem15Click(Sender: TObject);
    procedure SpeedItem16Click(Sender: TObject);
    procedure SpeedItem17Click(Sender: TObject);
    procedure SpeedItem16Click0(Sender: TObject);
    procedure AccesoObraClick(Sender: TObject);
    procedure adReg04Click(Sender: TObject);
    procedure opValidaEstClick(Sender: TObject);
    procedure opOrdenCamClick(Sender: TObject);
    procedure opAvisodeEmbClick(Sender: TObject);
    procedure cProveedoresClick(Sender: TObject);
    procedure SubidadePersonal1Click(Sender: TObject);
    procedure SpeedItem18Click(Sender: TObject);
    procedure SpeedItem19Click(Sender: TObject);
    procedure SpeedItem20Click(Sender: TObject);
    procedure cConsulta5Click(Sender: TObject);
    procedure cConsulta6Click(Sender: TObject);
    procedure SpeedItem13Click(Sender: TObject);
    procedure SpeedItem10Click0(Sender: TObject);
    procedure SpeedItem22Click(Sender: TObject);
    procedure cConsulta3Click(Sender: TObject);
    procedure adExportar2Click(Sender: TObject);
    procedure cConsulta7Click(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure imgKardexClick(Sender: TObject);
    procedure sCambiaPClick(Sender: TObject);
    procedure adCancelacionClick(Sender: TObject);
    procedure rGerencialClick(Sender: TObject);
    procedure opRequisicionesClick(Sender: TObject);
    procedure opPedidosClick(Sender: TObject);
    procedure opComparativo5Click(Sender: TObject);
    procedure SpeedItem23Click(Sender: TObject);
    procedure rSintesisClick(Sender: TObject);
    procedure sTipsClick(Sender: TObject);
    procedure sWarningClick(Sender: TObject);
    procedure adAvisosClick(Sender: TObject);
    procedure adActivosClick(Sender: TObject);
    procedure opFiltroClick(Sender: TObject);
    procedure cTrinomiosClick(Sender: TObject);
    procedure opadmonCatalogoClick(Sender: TObject);
    procedure opGeneradorSubClick(Sender: TObject);
    procedure adSubContratosClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure opAyudaClick(Sender: TObject);
    procedure mnuPozosdePerfoClick(Sender: TObject);
    procedure rInstaladoClick(Sender: TObject);
    procedure ProyecciondeActClick(Sender: TObject);
    procedure AsignaciondeOrdClick(Sender: TObject);
    procedure RepBarcoClick(Sender: TObject);
    procedure sTripulacionClick(Sender: TObject);
    procedure opInventarioClick(Sender: TObject);
    procedure opGeneradoresClick(Sender: TObject);
    procedure rPartidasIsomClick(Sender: TObject);
    procedure mnuConversionesClick(Sender: TObject);
    procedure cFasesClick(Sender: TObject);
    procedure tripulacionClick(Sender: TObject);
    procedure mnuKardexClick(Sender: TObject);
    procedure oficmodifClick(Sender: TObject);
    procedure MovPeroficClick(Sender: TObject);
    procedure CargaProgramaClick(Sender: TObject);
    procedure mnuAgrupacionPClick(Sender: TObject);
    procedure TiempoTimer(Sender: TObject);
    procedure GraficadorClick(Sender: TObject);
    procedure ChartProClick(Sender: TObject);
    procedure MnuCatalogodeMoClick(Sender: TObject);
    procedure MnuFamiliadeProClick(Sender: TObject);
    procedure MnuAlmacenClick(Sender: TObject);
    procedure tbbSetupClick(Sender: TObject);
    procedure tbbConsult3Click(Sender: TObject);
    procedure tbbPaquetePerClick(Sender: TObject);
    procedure tbbConsult2Click(Sender: TObject);
    procedure tbbConsult4Click(Sender: TObject);
    procedure tbbConsult5Click(Sender: TObject);
    procedure tbbConsult6Click(Sender: TObject);
    procedure tbbPaqEqClick(Sender: TObject);
    procedure tbbRepDiarioClick(Sender: TObject);
    procedure tbbRepBarcoClick(Sender: TObject);
    procedure tbbAvisoEmbClick(Sender: TObject);
    procedure tbbFotosClick(Sender: TObject);
    procedure tbbOrdenCambioClick(Sender: TObject);
    procedure tbbFirmantesClick(Sender: TObject);
    procedure tbbAutorizaClick(Sender: TObject);
    procedure tbbDesautorizaClick(Sender: TObject);
    procedure tbbEstimaClick(Sender: TObject);
    procedure tbbGeneraClick(Sender: TObject);
    procedure tbbInformesClick(Sender: TObject);
    procedure AdvToolBarButton24Click(Sender: TObject);
    procedure MnuEntAlmaceClick(Sender: TObject);
    procedure MnuSalAlmacenClick(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure ToolButton2Click(Sender: TObject);
    procedure ToolButton3Click(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure ToolButton10Click(Sender: TObject);
    procedure ToolButton9Click(Sender: TObject);
    procedure ToolButton6Click(Sender: TObject);
    procedure ToolButton7Click(Sender: TObject);
    procedure ToolButton8Click(Sender: TObject);
    procedure ToolButton11Click(Sender: TObject);
    procedure ToolButton13Click(Sender: TObject);
    procedure ToolButton14Click(Sender: TObject);
    procedure ToolButton15Click(Sender: TObject);
    procedure ToolButton17Click(Sender: TObject);
    procedure ToolButton18Click(Sender: TObject);
    procedure ToolButton21Click(Sender: TObject);
    procedure ToolButton22Click(Sender: TObject);
    procedure ToolButton23Click(Sender: TObject);
    procedure ToolButton24Click(Sender: TObject);
    procedure ToolButton25Click(Sender: TObject);
    procedure ToolButton27Click(Sender: TObject);
    procedure AdvToolBarMenuButton1Click(Sender: TObject);
    procedure tbbCambiaContratoClick(Sender: TObject);
    procedure tbbConsult1Click(Sender: TObject);
    procedure mnCopiarParametroClick(Sender: TObject);
    procedure optEstimacionesClick(Sender: TObject);
    procedure optValidaClick(Sender: TObject);
    procedure optDesautorizaClick(Sender: TObject);
    procedure SubAdministradorClick(Sender: TObject);
    procedure CatalogoErroresClick(Sender: TObject);
    procedure Generaciondeinformes2Click(Sender: TObject);
    procedure Generadores2Click(Sender: TObject);
    procedure Estimaciones3Click(Sender: TObject);
    procedure Reportesdiarios1Click(Sender: TObject);
    procedure inteligentpopPopup(Sender: TObject);
    procedure Irareportesdiarios1Click(Sender: TObject);
    procedure Iraestimaciones1Click(Sender: TObject);
    procedure Irageneradores1Click(Sender: TObject);
    procedure Irageneradoresdeinformes1Click(Sender: TObject);
    procedure Cambiarimagendefondo1Click(Sender: TObject);
    procedure FormMouseEnter(Sender: TObject);
    procedure FormMouseLeave(Sender: TObject);
    procedure MnuImpAvContClick(Sender: TObject);
    procedure mnuPersonal2Click(Sender: TObject);
    procedure reporteBarcoClick(Sender: TObject);
    procedure estirado1Click(Sender: TObject);
    procedure centrado1Click(Sender: TObject);
    procedure mosaico1Click(Sender: TObject);
    procedure rt1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);

    procedure Cascada1Click(Sender: TObject);
    procedure MosaicoVertical1Click(Sender: TObject);
    procedure MosaicoHorizontal1Click(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure JvAppEvents1ActiveControlChange(Sender: TObject);
    procedure subMaterialesClick(Sender: TObject);
    procedure InformedeSincronizado1Click(Sender: TObject);
    procedure GerencialBarco1Click(Sender: TObject);
    procedure ReportedeProduccionClick(Sender: TObject);
    procedure AgrupadordePersonal1Click(Sender: TObject);
    procedure NombresdeFirmantes1Click(Sender: TObject);
    procedure Panel2CanResize(Sender: TObject; var NewWidth, NewHeight: Integer;
      var Resize: Boolean);
    procedure Unificadordeequipos1Click(Sender: TObject);
    procedure ContenidoNotaCampo1Click(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure tbbProcGeneradorClick(Sender: TObject);
    procedure Cargadeformatos1Click(Sender: TObject);
    procedure ListadodePersonal1Click(Sender: TObject);
    procedure mniListadoPersonalClick(Sender: TObject);
    procedure MOE1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure MovimientosdeBarcoPT1Click(Sender: TObject);
    procedure MovimientosdeBarcoPT2Click(Sender: TObject);
    procedure EquipoPT1Click(Sender: TObject);
    procedure PernoctasPT1Click(Sender: TObject);
    procedure dxBarLargeButton29Click(Sender: TObject);
    procedure ControldeCalidadRIR1Click(Sender: TObject);
    procedure ReprogramacionesXFolios1Click(Sender: TObject);
    procedure HorariosGerenciales1Click(Sender: TObject);
    procedure cProgramadosClick(Sender: TObject);
    procedure ReportesProduccionSabanas1Click(Sender: TObject);
    procedure ResumendePersonal1Click(Sender: TObject);
    procedure mniGeneradoresClick(Sender: TObject);
    procedure mniRendimientoClick(Sender: TObject);

//    procedure BitacoradeActividades1Click(Sender: TObject);
//    procedure Estimaciones5Click(Sender: TObject);
//    procedure MovimientosdeEmbarcacin1Click(Sender: TObject);
//    procedure N3Click(Sender: TObject);
//    procedure ReporteDiario2Click(Sender: TObject);

//    procedure MovtosdeEmbarcacin1Click(Sender: TObject);
//    procedure ReporteBarco1Click(Sender: TObject);
//    procedure oooClick(Sender: TObject);
//    procedure hClick(Sender: TObject);
//    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }

    function MostrarFormChild(sForm: string): boolean;

  public
    { Public declarations }
    adentro: boolean;
    procedure permisosUsuarios;
    procedure AppMessage(var Msg: TMsg; var Handled: Boolean);
  end;

const
  WM_UPDATESTATUS = WM_USER + 2;

var
  frmInteligent: TfrmInteligent;
  detectar: string;
  Letra: char;

  function  GetAppVersion:string;

implementation

uses frm_contratos, frm_deptos, frm_usuarios,
  frm_personal, frm_equipos, frm_embarcaciones,
  frm_plataformas,
  frm_Cuentas, frm_Pernoctan,
  frm_ordenes, frm_actividades,
  frm_tiposdeMovimiento,
  frm_acceso,
  frm_acerca, frm_proveedores, frm_TipoMovto,
  frm_Almacenes, frm_tramitedepermisos,
  frm_valida,
  frm_abrereporte,
  frm_paquetesdepersonal,
  frm_ConsultadeActividades,
  frm_CalculoAvancesxPartida, frm_ConsultaxDescripcion,
  frm_ActividadesAnexo,
  frm_diasfestivos,
  frm_platicas, frm_turnos,
  frm_ReporteDiarioTurno,
  frm_compara2,
  frm_comparativo,
  frm_ConsultadeActividades2,
  frm_BusquedadeNotas,
  frm_Reprogramacion,

  frm_SqlManager,
  frm_importaciondedatos,
  frm_factordecosto, frm_AvancesFinancieros,
  frm_SqlImportar,
  frm_SqlExportar, frm_residencias,
  frm_programas,
  frm_GruposUsuarios, frm_gruposxprograma, frm_personalprogramado,
  frm_personalconsolidado, frm_contratosxusuario,
  frm_AjustaAnexo, frm_retecionesypenas,
  frm_PendientesNew, frm_ImportarDiarios,
  frm_ControlDirecto, frm_CalculoAvancesPaquetes,
  frm_ValidaEstimacion, frm_OrdendeCambio,
  frm_ConsultadeActividades3,
  frm_ConsultadeActividades4, frm_ExportaGeneral, frm_Kardex, frm_setup,
  frm_jornadasdiarias, frm_cambiapassword,
  frm_ReportePeriodo, frmReporteDiarioGerencial,
  frm_FichaTecnica,
  frm_ConsultadeActividades5,
  frm_firmantes,
  frm_ActividadesAnexo2, frm_generado, frm_estimaciones,
  frm_actividadesxgrupo, frm_EstimaInstalado, frm_compara, frm_tipsdia,
  frm_warningdia, frm_AvisosAlertas,
  frm_activos, frm_seleccion2, frmFiltroInteligent,
  frm_trinomios, frm_ProcRegAvFisico,
  
  frm_DistribucionPrograma, frm_admonCatalogos,
  frm_EstimaProveedor, frm_AjustaOrden, frm_SubContratos,
  frm_CatNomFirmantes,
  frm_Pedidos, frm_ordenesPerf, frm_EqPozos,
  frm_Consumibles, frm_RequisicionPerf,
  frm_detalledeinstalacion, frm_Proyeccion2, 
  frm_contratosxordenes, 
  frm_tripulacion, frm_ReporteDiario_Barco,
  frm_partidasxisometrico, frm_Conversiones, frm_Fases,
  frm_ordenesGral, FrmMovtoPersonalxoficio,
  frm_GruposPersonal,
  frm_Graficador, frm_IntelChart, frm_grupofamilias, 
  frm_EntradaAlmacen, frm_SalidaAlmacen,
  frm_EstimacionGeneral, frm_ValidaEstimacionGral,
  frm_AperturaEstimacionGral, 
  frm_AdministrarBd, frm_cancelacion,
  frm_ActualizaAvancesRemotos, frm_entradaanex, 
  
  
  frm_OpcionesGerencial, frm_OpcionesReporteProduccion,
  Frm_Moe, frm_gruposdepersonal, frm_GruposDeEquipo,frm_ProcesaGenerador,
  frm_ModuloReporteGerencial,frm_ResumenPersonal,frm_unificadorequipos,frm_contenidoNotaCampo,
  frm_Formatos, frm_Listado_Personal, frm_lista_personalV2, frm_cuadre, UTFrmMoeBordo,
  frm_Recursos_equipo, frm_Recursos_movimientos, frm_Recursos_pernocta,
  frm_Recursos_personal, frm_Calidad_Rir,
  frm_ReprogramacionFolio, frm_HorariosGerenciales , Frm_NotaCampo,
  Frm_generadores, frm_ReporteDiarioTurnoTierra;
{$R *.dfm}

function  GetAppVersion:string;
var
  Size, Size2: DWord;
  Pt, Pt2: Pointer;
begin
  Size := GetFileVersionInfoSize(PChar (ParamStr (0)), Size2);
  if Size > 0 then begin
    GetMem (Pt, Size);
    try
      GetFileVersionInfo (PChar (ParamStr (0)), 0, Size, Pt);
      VerQueryValue (Pt, '\', Pt2, Size2);
      with TVSFixedFileInfo (Pt2^) do begin
      Result:= ' Ver '+
               IntToStr (HiWord (dwFileVersionMS)) + '.' +
               IntToStr (LoWord (dwFileVersionMS)) + '.' +
               IntToStr (HiWord (dwFileVersionLS)) + '.' +
               IntToStr (LoWord (dwFileVersionLS));
      end;
    finally
      FreeMem (Pt);
    end;
  end;
end;

procedure TfrmInteligent.MosaicoHorizontal1Click(Sender: TObject);
begin
  FRMINTELIGENT.TileModE := tbHorizontal;
  FRMINTELIGENT.Tile;
end;

procedure TfrmInteligent.MosaicoVertical1Click(Sender: TObject);
begin
  FRMINTELIGENT.TileModE := tbVertical;
  FRMINTELIGENT.Tile;
end;


procedure TfrmInteligent.AppMessage(var Msg: TMsg; var Handled: Boolean); // TMSg
var
  actual: TWincontrol;
begin


  if Msg.Message = WM_KEYDOWN then
  begin
    //Msg.

  (*
  // esto es para controlar que con la flecha abajo se desplegen las
  // listas
    if Msg.wParam = VK_DOWN then
    begin
      Actual := Screen.ActiveControl;
      if Actual is TDBLookupComboBox then
            if not TDBLookupComboBox(Screen.ActiveForm.ActiveControl).listVisible then
               TDBLookupComboBox(Screen.ActiveForm.ActiveControl).DropDown;

      // si utilizas las RX
      if Actual is TRxDBLookupCombo then
            if not TRxDBLookupCombo(Screen.ActiveForm.ActiveControl).listVisible then
               TRxDBLookupCombo(Screen.ActiveForm.ActiveControl).DropDown;


      if Actual is TComboBox then
         SendMessage(TComboBox(Screen.ActiveForm.ActiveControl).Handle,CB_SHOWDROPDOWN,-1,0);

    // se pueden añadir mas controles....

    end;  *)

  // y aqui es el control del intro.
  (*if Msg.wParam = VK_RETURN THEN
  begin
     Actual := Screen.ActiveControl;
     // para los edit.
     if (Actual is TCustomEdit) and
        not (Actual is TCustomMemo) then
        Msg.wParam := VK_TAB;

     // para los memo. Con ctrl + enter se salta de linea dentro del
     // memo, si solo es intro salta al siguiente control
     if (Actual is TCustomMemo) and
        (HiWord(GetKeyState(VK_CONTROL)) = 0)then
        Msg.wParam := VK_TAB;

     // para los lookup
     if Actual is TDBLookupComboBox then
        if not TDBLookupComboBox(Actual).listVisible then
           Msg.wParam := VK_TAB;

     // Si utilizas las RX
     if Actual is TRxDBLookupCombo then
        if not TRxDBLookupCombo(Actual).listVisible then
           Msg.wParam := VK_TAB;

     // Los combobox
     if Actual is TCustomComboBox then
        if not TCustomComboBox(Actual).DroppedDown then
           Msg.wParam := VK_TAB;

     // el radiobutton
     if Actual is TRadioButton then
        Msg.wParam := VK_TAB;

     // en el grid
     if Actual is TDBGrid then
        if (TDBGrid(Actual).ReadOnly) or
           (HiWord(GetKeyState(VK_CONTROL)) = 1)then
           Msg.wParam := VK_TAB;

     // o esto otro Con Ctrl + intro siguiente celda con
     // ctrl + shift + intro celda anterior
     // es complicado pero es por no liar con el intro solo
     // que siempre pasa de un control a otro.
     if Actual is TDBGrid then
        begin
        if (TDBGrid(Actual).ReadOnly) or
           (HiWord(GetKeyState(VK_CONTROL)) = 0) then
           Msg.wParam := VK_TAB
        else
           if not(HiWord(GetKeyState(VK_SHIFT)) = 0) then
              begin
              if TDBGrid(Actual).selectedindex > 0 then             { increment the field }
                 TDBGrid(Actual).selectedindex := TDBGrid(Actual).selectedindex -1
              else
                 TDBGrid(Actual).selectedindex := TDBGrid(Actual).fieldcount -1;
              end
           else
              begin
              if TDBGrid(Actual).selectedindex < (TDBGrid(Actual).fieldcount -1) then             { increment the field }
                 TDBGrid(Actual).selectedindex := TDBGrid(Actual).selectedindex +1
              else
                 TDBGrid(Actual).selectedindex := 0;
              end;
        end;


     // ListBox
     if Actual is TCustomListBox then
        Msg.wParam := VK_TAB;

     // TabControl
     if Actual is TCustomTabControl then
        Msg.wParam := VK_TAB;

     // Aqui se pondrían todos los controles que usamos en la
     // aplicación y necesitamos cambiar el tab por intro.
     // de esta forma te despreocupas de cuantos añades y es más
     // limpio el código

     end;
  end;            *)
    Actual := Screen.ActiveControl;
    if (Msg.wParam = 189) or (Msg.wParam = 109) then
    begin
      if (Actual is TCustomEdit) and
        not (Actual is TCustomMemo) then
        if TCustomEdit(Actual).Tag = 2123 then
          Msg.wParam := VK_CANCEL;

      if Actual is TDBGrid then
         // if (TDBGrid(Actual).ReadOnly) or
         //    (HiWord(GetKeyState(VK_CONTROL)) = 1)then
        if (TDBGrid(Actual).Tag = 2123) then
        begin
          if ((pos('CANTIDAD', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('AVANCE', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('PRECIO', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('DURACION', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('VENTA M.N.', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('VENTA DLL', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('GRUPO', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('RENGLON', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('ORDEN', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('NO. ESTIMACION', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('NO. REPROG.', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('FASE', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('COSTO MN', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('COSTO DLL', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('ID', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('ANEXO', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('PARTIDA', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0)
          or (pos('ID PAGO', uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.Caption)) > 0))
          then
            if uppercase(TDBGrid(Actual).Columns[TDBGrid(Actual).selectedindex].Title.caption) <> 'TRAZABILIDAD' then //EL APARTADO DE MATERIALES EN BITACORA2 NO PERMITE -
              Msg.wParam := VK_CANCEL;
        end;

    end;

    if Msg.wParam = VK_RETURN then
    begin
      if (Actual is TRxDBCalcEdit) then
        TRxDBCalcEdit(Actual).value := abs(TRxDBCalcEdit(Actual).value);

    end;
  end;
  //WM_ACTIVATE
          //WM_DISPLAYCHANGE
          //WM_CAPTURECHANGED
          //EN_CHANGE
          //WM_USERCHANGED
          //EM_GETMODIFY
 // if Msg.message = WM_COMMAND then
//  begin
   // if TMESSAGE(Msg.Message).     =EN_CHANGE then
  //    showmessage('Cange');
      //if (Actual is TRxDBCalcEdit) then
        //TRxDBCalcEdit(Actual).value:=abs(TRxDBCalcEdit(Actual).value);
 // end;


end;


function TfrmInteligent.MostrarFormChild(sForm: string): boolean;
var
  i: integer;
begin
  result := false;


  if MDIChildCount <> 0 then
  begin
    for I := 0 to MDIChildCount - 1 do
      if uppercase(MDIChildren[i].Name) = uppercase(sform) then
      begin
        MDIChildren[i].Show;
        result := true;
        break;
      end;
  end;

end;

procedure TfrmInteligent.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if Licencia = 'Si' then
    if MessageDlg('Esta seguro que desea salir completamente de la aplicación?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      CanClose := true
    else
      CanClose := False;
end;

procedure TfrmInteligent.FormCreate(Sender: TObject);
begin
  Application.OnMessage := AppMessage;
end;

procedure TfrmInteligent.FormKeyPress(Sender: TObject; var Key: Char);
begin
  //inherited;

  //If Key = #13 Then
  //begin
    //if Ord(key) in [0,1,2,3,4,5,6,7,8,9] then
      //showmessage('Numero');

   // SelectNext(ActiveControl as tWinControl, True, True );
   // Key := #0;
  //end;

end;

procedure TfrmInteligent.FormMouseEnter(Sender: TObject);
begin
  inteligentpop.AutoPopup := true;
end;

procedure TfrmInteligent.FormMouseLeave(Sender: TObject);
begin
  inteligentpop.AutoPopup := FALSE;
end;

procedure TfrmInteligent.FormShow(Sender: TObject);
var
  WinDir: array[0..MAX_PATH - 1] of char;
  sTipo: string;
  iPos: Integer;
  InfoSize, H, RsltLen: Cardinal;
  VersionBlock: Pointer;
  Rslt: PVSFixedFileInfo;

  StringList: TStrings;
  S: wideString;
  F: TextFile;
var i: integer;
  ini: tinifile;
  validarpath: string;
  bueno: boolean;
  bueno2: boolean;
  pathimagen: string;

  rutaaux: string;
begin

//*************************************
  adentro := False;
  InfoSize := GetFileVersionInfoSize(PChar(Application.ExeName), H);
  VersionBlock := AllocMem(InfoSize);
  try
    GetFileVersionInfo(PChar(Application.ExeName), H, InfoSize, VersionBlock);
    VerQueryValue(VersionBlock, '\', Pointer(Rslt), RsltLen);
    Caption := global_version; //'Sistema Inteligente para la Administración de Contratos Obra Publica Versión 12 de Marzo de 2012';
        //Format('%d.%d.%d.%d', [
        //Rslt.dwProductVersionMS div 65536,
        //Rslt.dwProductVersionMS mod 65536,
        //Rslt.dwProductVersionLS div 65536,
        //Rslt.dwProductVersionLS mod 65536]) +
  finally
    FreeMem(VersionBlock);
  end;

 // 'Inteligent 2011 VC 1.1'


//  SetString(global_ruta, WinDir, GetWindowsDirectory(WinDir, MAX_PATH));
//  if global_ruta = '' then
//  begin
//    raise Exception.Create(SysErrorMessage(GetLastError));
//    global_ruta := 'c:\inteligent\';
//  end;
//
//  global_ruta := MidStr(global_ruta, 1, 3) + 'inteligent\';
//
//  ///// AQUI ESTUVO CARMEN /////
//  detectar := global_ruta + 'image.ini';
//
//  if leeini(detectar) <> 'no' then
//    muestrafondo(JvBackground1, unitmanejofondo.imapatglobal, unitmanejofondo.estadoglobal)
//  else
//    escribeinidefault(detectar, 'bmCenter');
//
//  //// TEMRINA LECTURA DE LA IMAGEN..

  frmAcceso.ShowModal;
  if frmAcceso.salir then
  begin
    tiempo.enabled := True;

    abort;
  end;

  if (global_usuario <> '') and (global_usuario <> 'INTEL-CODE') then
  begin
    global_activo := 'S';
    frmSeleccion2.showModal;
  end
  else
    if global_usuario <> 'INTEL-CODE' then
    begin
      if global_grupo = 'INTEL-CODE' then
      begin
        frmSeleccion2.showModal;
        connection.contrato.Active := False;
        connection.contrato.Params.ParamByName('Contrato').DataType := ftString;
        connection.contrato.Params.ParamByName('Contrato').Value := global_contrato;
        connection.contrato.Open;

        connection.configuracion.Active := False;
        connection.configuracion.Params.ParamByName('Contrato').Value := global_contrato;
        connection.configuracion.Params.ParamByName('Contrato').DataType := ftString;
        connection.configuracion.Open;
        global_convenio := 'C';
        if connection.configuracion.RecordCount = 0 then
          application.MessageBox('Precaución: No se encontro el archivo principal de configuración, notifique al Administrador del Sistema', 'Inteligent', 0)
        else
          Global_Convenio := connection.configuracion.FieldValues['sIdConvenio']
      end
      else
        application.Terminate;
    end;

  Licencia := 'Si';

  status.Panels.Items[1].Text := global_nombre;
  status.Panels.Items[3].Text := global_server;
  status.Panels.Items[5].Text := global_db;
  status.Panels.Items[7].Text := global_contrato;

  ImgKardex.Top := frmInteligent.Height - 132;
  ImgKardex.Left := frmInteligent.Width - 63;

  if global_contrato <> '' then
  begin
    try
      stMenu := '';
      if sender is TMenuItem then
        stMenu := (sender as TMenuItem).Name;
      Application.CreateForm(TfrmPendientesNew, frmPendientesNew);
      frmPendientesNew.show;

      Connection.qryBusca.Active := False;
      Connection.qryBusca.SQL.Clear;
      Connection.qryBusca.SQL.Add('Select dFechaInicio, dFechaFinal From convenios Where sContrato = :Contrato And sIdConvenio = :Convenio');
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
      Connection.qryBusca.Open;
    except
      on e: exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Mis Pendientes', 0);
      end;
    end;

    if Connection.qryBusca.RecordCount > 0 then
    begin
      frmPendientesNew.tdFechaInicio.Text := Connection.qryBusca.FieldValues['dFechaInicio'];
      if VarIsNull(frmPendientesNew.tdFechaFinal.Text) then
        MessageDlg('No hay fecha final de Convenio!!', mtError, [mbOk], 0)
      else
        frmPendientesNew.tdFechaFinal.Text := Connection.qryBusca.FieldValues['dFechaFinal'];
      frmPendientesNew.tdLaborado.Value := (Date - Connection.qryBusca.FieldValues['dFechaInicio']) + 1;
      frmPendientesNew.tdTranscurrido.Value := Connection.qryBusca.FieldValues['dFechaFinal'] - Date;

      if Date <= Connection.qryBusca.FieldValues['dFechaFinal'] then
      begin
        frmPendientesNew.avProyecto.Value := (Connection.qryBusca.FieldValues['dFechaFinal'] - Connection.qryBusca.FieldValues['dFechaInicio']) + 1;
        frmPendientesNew.avProyecto.Value := (frmPendientesNew.tdLaborado.Value / frmPendientesNew.avProyecto.Value) * 100;
        frmPendientesNew.avPendiente.Value := 100 - frmPendientesNew.avProyecto.Value
      end
      else
      begin
        frmPendientesNew.avProyecto.Value := 100;
        frmPendientesNew.avPendiente.Value := 0;
      end

    end
    else
    begin
      frmPendientesNew.tdFechaInicio.Text := DateToStr(Date);
      if VarIsNull(frmPendientesNew.tdFechaFinal.Text) then
        MessageDlg('NO HAY FECHA FINAL  !!', mtError, [mbOk], 0)
      else
        frmPendientesNew.tdFechaFinal.Text := DateToStr(Date);
      frmPendientesNew.tdLaborado.Value := 0;
      frmPendientesNew.tdTranscurrido.Value := 0;
      frmPendientesNew.avProyecto.Value := 0;
      frmPendientesNew.avPendiente.Value := 0;

    end;

    if frmPendientesNew.tdTranscurrido.Value <= 10 then
    begin
      frmPendientesNew.tdTranscurrido.Font.Style := [fsBold];
      frmPendientesNew.tdTranscurrido.Font.Color := clRed;
      frmPendientesNew.tdTranscurrido.Font.Size := 9;
    end
    else
    begin
      frmPendientesNew.tdTranscurrido.Font.Style := [];
      frmPendientesNew.tdTranscurrido.Font.Color := clWindowText;
      frmPendientesNew.tdTranscurrido.Font.Size := 8;
    end;


    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select dAvancePonderadoGlobal From avancesglobales Where sContrato = :Contrato And sIdConvenio = :Convenio And dIdFecha = :Fecha And sNumeroOrden = ""');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := Date;
    Connection.qryBusca.Open;
    frmPendientesNew.avProgramado.Value := 0;
    if Connection.qryBusca.RecordCount > 0 then
      frmPendientesNew.avProgramado.Value := Connection.qryBusca.FieldValues['dAvancePonderadoGlobal'];

    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select Sum(dAvance)  as dAvance From avancesglobalesxorden Where sContrato = :Contrato And sIdConvenio = :Convenio And dIdFecha <= :Fecha And sNumeroOrden = "" Group By sContrato');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := Date;
    Connection.qryBusca.Open;
    frmPendientesNew.avReal.Value := 0;
    if Connection.qryBusca.RecordCount > 0 then
      frmPendientesNew.avReal.Value := Connection.qryBusca.FieldValues['dAvance'];
    if Connection.configuracion.FieldValues['sTipsInicial'] = 'Si' then
    begin
      Application.CreateForm(TfrmTipsDia, frmTipsDia);
      frmTipsDia.show;
    end;

      // Inicial proceso de creación de warning ...
    {connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "No" Where sTipo = "Warning" And iMessage <= 10');
    connection.zCommand.ExecSQL;}

      // Warning 1 y 2
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select dFechaFinal From convenios Where sContrato = :Contrato And sIdConvenio = :Convenio');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
      if (Connection.qryBusca.FieldValues['dFechaFinal'] - Date) <= 20 then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "Si" Where sTipo = "Warning" And iMessage = 2');
        connection.zCommand.ExecSQL;
      end
      else
        if (Connection.qryBusca.FieldValues['dFechaFinal'] - Date) <= 30 then
        begin
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "Si" Where sTipo = "Warning" And iMessage = 1');
          connection.zCommand.ExecSQL;
        end;

      // Warning 3
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select dIdFecha From reportediario Where sContrato = :Contrato And dIdFecha = :Fecha');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := Date - 1;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount = 0 then
    begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "Si" Where sTipo = "Warning" And iMessage = 3');
      connection.zCommand.ExecSQL;
    end;

      // Warning 4
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select Count(dIdFecha) as iReportes From reportediario Where sContrato = :Contrato And lStatus <> "Autorizado"');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount = 0 then
      if Connection.qryBusca.FieldValues['iReportes'] > 0 then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "Si" Where sTipo = "Warning" And iMessage = 4');
        connection.zCommand.ExecSQL;
      end;
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select sAutor From inteligent_message where sTipo = "Warning" And lVisible = "Si"');
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
    begin
      Application.CreateForm(TfrmWarningDia, frmWarningDia);
      frmWarningDia.Show
    end
  end;

  if (global_usuario <> 'INTEL-CODE') and (global_usuario <> 'ADMIN') then
    permisosUsuarios();

//  frmInteligent.WindowState := wsNormal ;
//  frmInteligent.Show ;
//  frmInteligent.WindowState := wsMaximized ;
//  frmInteligent.Show ;

end;

procedure TfrmInteligent.adConfiguracionClick(Sender: TObject);
begin
  frmSetup.ShowModal;

  status.Panels.Items[9].Text := global_convenio;
  status.Panels.Items[13].Text := global_afectacion;
end;


procedure TfrmInteligent.adContratosClick(Sender: TObject);
begin
  if not MostrarFormChild('frmcontratos') then
  begin
    Application.CreateForm(TfrmContratos, frmContratos);
    frmContratos.show;
  end;

end;

procedure TfrmInteligent.adDeptosClick(Sender: TObject);
begin
  if not MostrarFormChild('frmdeptos') then
  begin
    Application.CreateForm(TfrmDeptos, frmDeptos);
    frmDeptos.show
  end;
end;

procedure TfrmInteligent.adUsuariosClick(Sender: TObject);
begin
  if not MostrarFormChild('frmusuarios') then
  begin
    Application.CreateForm(TfrmUsuarios, frmUsuarios);
    frmUsuarios.show
  end;
end;

procedure TfrmInteligent.cPersonalClick(Sender: TObject);
begin
//<ROJAS>
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmpersonal') then
    begin
      Application.CreateForm(TfrmPersonal, frmPersonal);
      frmPersonal.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Catalogo de Especialidades de Personal', 0);
    end;
  end;
//
end;

procedure TfrmInteligent.cEquiposClick(Sender: TObject);
begin
//<ROJAS>
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmequipos') then
    begin
      Application.CreateForm(TfrmEquipos, frmEquipos);
      frmEquipos.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Reportes Diarios', 0);
    end;
  end;
//

end;

procedure TfrmInteligent.cEmbarcacionesClick(Sender: TObject);
begin
  stMenu := '';
  if Sender is TMenuItem then
    stMenu := (Sender as TMenuItem).Name;
  if not MostrarFormChild('frmembarcaciones') then
  begin
    Application.CreateForm(TfrmEmbarcaciones, frmEmbarcaciones);
    frmEmbarcaciones.FormStyle := fsMDIForm;
    inteligentpop.AutoPopup := false;
    frmembarcaciones.Visible := False;
    frmembarcaciones.showModal;
  end;
end;

procedure TfrmInteligent.centrado1Click(Sender: TObject);
begin
  detectar := ExtractFilePath(Application.Exename) + 'image.ini'; //extraepath
  if leeini(detectar) <> 'no' then
    modofondo(JvBackground1, 'bmCenter', detectar)
  else
    escribeinidefault(detectar, 'bmCenter');
end;

procedure TfrmInteligent.cPlataformasClick(Sender: TObject);
begin
  stMenu := '';
  if Sender is TMenuItem then
    stMenu := (Sender as TMenuItem).Name;
  if not MostrarFormChild('frmplataformas') then
  begin
    Application.CreateForm(TfrmPlataformas, frmPlataformas);
    frmPlataformas.show
  end;
end;


procedure TfrmInteligent.cCuentasClick(Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmcuentas') then
    begin
      Application.CreateForm(TfrmCuentas, frmCuentas);
      frmCuentas.show
    end;

  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Reportes Diarios', 0);
    end;

  end;

end;

procedure TfrmInteligent.cPernoctanClick(Sender: TObject);
begin
  if not MostrarFormChild('frmpernoctan') then
  begin
    Application.CreateForm(TfrmPernoctan, frmPernoctan);
    frmPernoctan.show
  end;
end;

procedure TfrmInteligent.cOrdenesClick(Sender: TObject);
begin
  if Connection.configuracion.FieldValues['sCampPerf'] = 'No' then
  begin
    if not MostrarFormChild('frmordenes') then
    begin
      Application.CreateForm(TfrmOrdenes, frmOrdenes);
      frmOrdenes.show;
    end
  end
  else
  begin
    if not MostrarFormChild('frmordenesperf') then
    begin
      Application.CreateForm(TfrmOrdenesPerf, frmOrdenesPerf);
      frmOrdenesPerf.Show;
    end
  end;
end;

procedure TfrmInteligent.cActividadesClick(
  Sender: TObject);
begin
  stMenu := '';
  if Sender is TMenuItem then
    stMenu := (Sender as TMenuItem).Name;
  if not MostrarFormChild('frmactividades') then
  begin
    Application.CreateForm(TfrmActividades, frmActividades);
    frmActividades.show
  end
end;

procedure TfrmInteligent.Cambiarimagendefondo1Click(Sender: TObject);
var pathimagen: string;
  Ini: TiniFile;
begin
  detectar := ExtractFilePath(Application.Exename) + 'image.ini';
  escribeini(detectar, OpenDialog1);
  if leeini(detectar) <> 'no' then
    muestrafondo(JvBackground1, unitmanejofondo.imapatglobal, unitmanejofondo.estadoglobal)

{opendialog1.Filter:='Fotografia|*.jpg|Imagen|*.bmp';
  if openDialog1.Execute then
   begin
       pathimagen:=(openDialog1.FileName);
       jvbackground1.Image.Picture.LoadFromFile(pathimagen);
       ini := TIniFile.Create(global_ruta+'img.ini');
       ini.WriteString ('Configuración', 'wallpaper', pathimagen);
       ini.free;
    end
   else ShowMessage('Abrir archivo se a cancelado');}
end;





procedure TfrmInteligent.cFasesClick(Sender: TObject);
begin
  if not MostrarFormChild('frmfases') then
  begin
    Application.CreateForm(TfrmFases, frmFases);
    frmFases.show
  end;
end;

procedure TfrmInteligent.Firmantes1Click(Sender: TObject);
begin
  global_orden := '';
  if not MostrarFormChild('frmfirmas') then
  begin
    Application.CreateForm(TfrmFirmas, frmFirmas);
    frmfirmas.show
  end;
end;

procedure TfrmInteligent.adTiposMovClick(Sender: TObject);
begin
  if not MostrarFormChild('frmtiposdemovimiento') then
  begin
    Application.CreateForm(TfrmTiposdeMovimiento, frmTiposdeMovimiento);
    frmTiposdeMovimiento.show
  end;
end;

procedure TfrmInteligent.opMuertoClick(Sender: TObject);
begin
  global_orden := '';
  Application.CreateForm(TfrmJornadasDiarias, frmJornadasDiarias);
  frmJornadasDiarias.show
end;

procedure TfrmInteligent.sSeleccionClick(Sender: TObject);
begin
  frmSeleccion2.showModal;
  if Assigned(frmPendientesNew) then
  begin
    try
      if global_PendientesOculto = False then
      begin
        global_PendientesOculto := False;
        frmPendientesNew.Close;
      end;
      Application.CreateForm(TfrmPendientesNew, frmPendientesNew);
      frmPendientesNew.show;

      frmPendientesNew.tmDescripcion.Text := Connection.contrato.FieldValues['mDescripcion'];
      Connection.qryBusca.Active := False;
      Connection.qryBusca.SQL.Clear;
      Connection.qryBusca.SQL.Add('Select dFechaInicio, dFechaFinal From convenios Where sContrato = :Contrato And sIdConvenio = :Convenio');
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
      Connection.qryBusca.Open;
      if Connection.qryBusca.RecordCount > 0 then
      begin
        frmPendientesNew.tdFechaInicio.Text := Connection.qryBusca.FieldValues['dFechaInicio'];
        frmPendientesNew.tdFechaFinal.Text := Connection.qryBusca.FieldValues['dFechaFinal'];
        frmPendientesNew.tdLaborado.Value := (Date - Connection.qryBusca.FieldValues['dFechaInicio']) + 1;
        frmPendientesNew.tdTranscurrido.Value := Connection.qryBusca.FieldValues['dFechaFinal'] - Date;

        if Date <= Connection.qryBusca.FieldValues['dFechaFinal'] then
        begin
          frmPendientesNew.avProyecto.Value := (Connection.qryBusca.FieldValues['dFechaFinal'] - Connection.qryBusca.FieldValues['dFechaInicio']) + 1;
          frmPendientesNew.avProyecto.Value := (frmPendientesNew.tdLaborado.Value / frmPendientesNew.avProyecto.Value) * 100;
          frmPendientesNew.avPendiente.Value := 100 - frmPendientesNew.avProyecto.Value
        end
        else
        begin
          frmPendientesNew.avProyecto.Value := 100;
          frmPendientesNew.avPendiente.Value := 0;
        end

      end
      else
      begin
        frmPendientesNew.tdFechaInicio.Text := DateToStr(Date);
        frmPendientesNew.tdFechaFinal.Text := DateToStr(Date);
        frmPendientesNew.tdLaborado.Value := 0;
        frmPendientesNew.tdTranscurrido.Value := 0;
        frmPendientesNew.avProyecto.Value := 0;
        frmPendientesNew.avPendiente.Value := 0;

      end;

      if frmPendientesNew.tdTranscurrido.Value <= 10 then
      begin
        frmPendientesNew.tdTranscurrido.Font.Style := [fsBold];
        frmPendientesNew.tdTranscurrido.Font.Color := clRed;
        frmPendientesNew.tdTranscurrido.Font.Size := 9;
      end
      else
      begin
        frmPendientesNew.tdTranscurrido.Font.Style := [];
        frmPendientesNew.tdTranscurrido.Font.Color := clWindowText;
        frmPendientesNew.tdTranscurrido.Font.Size := 8;
      end;


      Connection.qryBusca.Active := False;
      Connection.qryBusca.SQL.Clear;
      Connection.qryBusca.SQL.Add('Select dAvancePonderadoGlobal From avancesglobales Where sContrato = :Contrato And sIdConvenio = :Convenio And dIdFecha = :Fecha And sNumeroOrden = ""');
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
      Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
      Connection.qryBusca.Params.ParamByName('Fecha').Value := Date;
      Connection.qryBusca.Open;
      frmPendientesNew.avProgramado.Value := 0;
      if Connection.qryBusca.RecordCount > 0 then
        frmPendientesNew.avProgramado.Value := Connection.qryBusca.FieldValues['dAvancePonderadoGlobal'];

      Connection.qryBusca.Active := False;
      Connection.qryBusca.SQL.Clear;
      Connection.qryBusca.SQL.Add('Select Sum(dAvance) as dAvance From avancesglobalesxorden Where sContrato = :Contrato And sIdConvenio = :Convenio And dIdFecha <= :Fecha And sNumeroOrden = "" Group By sContrato');
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
      Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
      Connection.qryBusca.Params.ParamByName('Fecha').Value := Date;
      Connection.qryBusca.Open;
      frmPendientesNew.avReal.Value := 0;
      if Connection.qryBusca.RecordCount > 0 then
        frmPendientesNew.avReal.Value := Connection.qryBusca.FieldValues['dAvance'];
      status.Panels.Items[1].Text := global_nombre;
      status.Panels.Items[3].Text := global_server;
      status.Panels.Items[5].Text := global_db;
      status.Panels.Items[7].Text := global_contrato;
      status.Panels.Items[9].Text := global_convenio;
      status.Panels.Items[11].Text := global_sturno;
      status.Panels.Items[13].Text := global_afectacion;
    except

    end;
  end;
end;

procedure TfrmInteligent.sLoginClick(Sender: TObject);
var
  sPrograma: string;
  iElemento: Integer;
  Component: tMenuItem;

  StringList: TStrings;
  S: wideString;
  F: TextFile;


begin
  adentro := True;
  try
    frmTipsDia.Close;
  except
  end;
  global_contrato := '';
  global_usuario := '';
  connection.zConnection.Disconnect;
  frmAcceso.ShowModal;
  if frmacceso.salir then
  begin
    tiempo.Enabled := True;
    abort;
  end;

  if global_usuario <> '' then
  begin
    global_activo := 'S';
    if global_contrato <> '' then
    begin
           // El usuario pertenece a un contrato ...
           // Se inicializan los Querys al contrato seleccionado ...
      connection.configuracion.Active := False;
      connection.configuracion.SQL.Clear;
      connection.configuracion.SQL.Add('select * from configuracion where sContrato = :contrato');
      connection.configuracion.Params.ParamByName('Contrato').Value := global_contrato;
      connection.configuracion.Params.ParamByName('Contrato').DataType := ftString;
      connection.configuracion.Open;
      global_convenio := 'C';
      if connection.configuracion.RecordCount = 0 then
        application.MessageBox('Precaución: No se encontro el archivo principal de configuración, notifique al Administrador del Sistema', 'Inteligent', 0)
      else
        Global_Convenio := connection.configuracion.FieldValues['sIdConvenio']
    end
    else
      frmInteligent.permisosUsuarios;
    frmSeleccion2.showModal
  end
  else
    if global_grupo = 'INTEL-CODE' then
      frmSeleccion2.showModal
    else
      application.Terminate;

  status.Panels.Items[1].Text := global_nombre;
  status.Panels.Items[3].Text := global_server;
  status.Panels.Items[5].Text := global_db;
  status.Panels.Items[7].Text := global_contrato;
  status.Panels.Items[9].Text := global_convenio;
  status.Panels.Items[11].Text := global_turno;
  status.Panels.Items[13].Text := global_afectacion;


  if global_PendientesOculto = False then
  begin
    global_PendientesOculto := False;
    frmPendientesNew.Close;
  end;
  if global_contrato <> '' then
  begin
    Application.CreateForm(TfrmPendientesNew, frmPendientesNew);
    frmPendientesNew.show;


    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select dFechaInicio, dFechaFinal From convenios Where sContrato = :Contrato And sIdConvenio = :Convenio');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
    begin
      frmPendientesNew.tdFechaInicio.Text := Connection.qryBusca.FieldValues['dFechaInicio'];
      frmPendientesNew.tdFechaFinal.Text := Connection.qryBusca.FieldValues['dFechaFinal'];
      frmPendientesNew.tdLaborado.Value := (Date - Connection.qryBusca.FieldValues['dFechaInicio']) + 1;
      frmPendientesNew.tdTranscurrido.Value := Connection.qryBusca.FieldValues['dFechaFinal'] - Date;

      if Date <= Connection.qryBusca.FieldValues['dFechaFinal'] then
      begin
        frmPendientesNew.avProyecto.Value := (Connection.qryBusca.FieldValues['dFechaFinal'] - Connection.qryBusca.FieldValues['dFechaInicio']) + 1;
        frmPendientesNew.avProyecto.Value := (frmPendientesNew.tdLaborado.Value / frmPendientesNew.avProyecto.Value) * 100;
        frmPendientesNew.avPendiente.Value := 100 - frmPendientesNew.avProyecto.Value
      end
      else
      begin
        frmPendientesNew.avProyecto.Value := 100;
        frmPendientesNew.avPendiente.Value := 0;
      end

    end
    else
    begin
      frmPendientesNew.tdFechaInicio.Text := DateToStr(Date);
      frmPendientesNew.tdFechaFinal.Text := DateToStr(Date);
      frmPendientesNew.tdLaborado.Value := 0;
      frmPendientesNew.tdTranscurrido.Value := 0;
      frmPendientesNew.avProyecto.Value := 0;
      frmPendientesNew.avPendiente.Value := 0;

    end;

    if frmPendientesNew.tdTranscurrido.Value <= 10 then
    begin
      frmPendientesNew.tdTranscurrido.Font.Style := [fsBold];
      frmPendientesNew.tdTranscurrido.Font.Color := clRed;
      frmPendientesNew.tdTranscurrido.Font.Size := 9;
    end
    else
    begin
      frmPendientesNew.tdTranscurrido.Font.Style := [];
      frmPendientesNew.tdTranscurrido.Font.Color := clWindowText;
      frmPendientesNew.tdTranscurrido.Font.Size := 8;
    end;


    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select dAvancePonderadoGlobal From avancesglobales Where sContrato = :Contrato And sIdConvenio = :Convenio And dIdFecha = :Fecha And sNumeroOrden = ""');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := Date;
    Connection.qryBusca.Open;
    frmPendientesNew.avProgramado.Value := 0;
    if Connection.qryBusca.RecordCount > 0 then
      frmPendientesNew.avProgramado.Value := Connection.qryBusca.FieldValues['dAvancePonderadoGlobal'];

    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select Sum(dAvance)  as dAvance From avancesglobalesxorden Where sContrato = :Contrato And sIdConvenio = :Convenio And dIdFecha <= :Fecha And sNumeroOrden = "" Group By sContrato');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := Date;
    Connection.qryBusca.Open;
    frmPendientesNew.avReal.Value := 0;
    if Connection.qryBusca.RecordCount > 0 then
      frmPendientesNew.avReal.Value := Connection.qryBusca.FieldValues['dAvance'];
    if Connection.configuracion.FieldValues['sTipsInicial'] = 'Si' then
    begin
      Application.CreateForm(TfrmTipsDia, frmTipsDia);
      frmTipsDia.show;
    end;

      // Inicial proceso de creación de warning ...
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "No" Where sTipo = "Warning" And iMessage <= 10');
    connection.zCommand.ExecSQL;

      // Warning 1 y 2
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select dFechaFinal From convenios Where sContrato = :Contrato And sIdConvenio = :Convenio');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
      if (Connection.qryBusca.FieldValues['dFechaFinal'] - Date) <= 20 then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "Si" Where sTipo = "Warning" And iMessage = 2');
        connection.zCommand.ExecSQL;
      end
      else
        if (Connection.qryBusca.FieldValues['dFechaFinal'] - Date) <= 30 then
        begin
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "Si" Where sTipo = "Warning" And iMessage = 1');
          connection.zCommand.ExecSQL;
        end;

      // Warning 3
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select dIdFecha From reportediario Where sContrato = :Contrato And dIdFecha = :Fecha');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := Date - 1;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount = 0 then
    begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "Si" Where sTipo = "Warning" And iMessage = 3');
      connection.zCommand.ExecSQL;
    end;

      // Warning 4
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select Count(dIdFecha) as iReportes From reportediario Where sContrato = :Contrato And lStatus <> "Autorizado"');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount = 0 then
      if Connection.qryBusca.FieldValues['iReportes'] > 0 then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE inteligent_message SET lVisible = "Si" Where sTipo = "Warning" And iMessage = 4');
        connection.zCommand.ExecSQL;
      end;
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select sAutor From inteligent_message where sTipo = "Warning" And lVisible = "Si"');
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
    begin
      Application.CreateForm(TfrmWarningDia, frmWarningDia);
      frmWarningDia.Show
    end
  end;

  permisosUsuarios();

end;

procedure TfrmInteligent.sSalirClick(Sender: TObject);
begin
  close;
end;

procedure TfrmInteligent.sAcercaClick(Sender: TObject);
begin
  frmAcerca.showModal
end;

procedure TfrmInteligent.Cargadeformatos1Click(Sender: TObject);
begin
  Application.CreateForm(TfrmFormatos, frmFormatos);
  frmFormatos.ShowModal;
end;

procedure TfrmInteligent.CargaProgramaClick(Sender: TObject);
begin
  {if not MostrarFormChild('frmCargaPrograma') then
  begin
    Application.CreateForm(TfrmcargaPrograma, frmcargaPrograma);
    frmcargaPrograma.Show;
  end;}
end;

procedure TfrmInteligent.Cascada1Click(Sender: TObject);
begin
  FRMINTELIGENT.Cascade;
end;

procedure TfrmInteligent.MnuCatalogodeMoClick(Sender: TObject);
begin
  if not MostrarFormChild('frmmovtos') then
  begin
    Application.CreateForm(TfrmMovtos, frmMovtos);
    frmMovtos.show
  end;
end;


procedure TfrmInteligent.subMaterialesClick(Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmconsumibles') then
    begin
      Application.CreateForm(TfrmConsumibles, frmConsumibles);
      frmConsumibles.Show;
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Reportes Diarios', 0);
    end;
  end;
      //
end;

procedure TfrmInteligent.CatalogodeProveedores1Click(Sender: TObject);
begin
  Application.CreateForm(TfrmProveedores, frmProveedores);
  frmProveedores.show
end;

procedure TfrmInteligent.CatalogoErroresClick(Sender: TObject);
begin
  if not MostrarFormChild('frmcatalogoerrores') then
  begin
    Application.CreateForm(TfrmCatalogoErrores, frmCatalogoErrores);
    frmCatalogoErrores.show
  end;
end;

procedure TfrmInteligent.opPermisosClick(Sender: TObject);
begin
    global_orden := '';
    Application.CreateForm(TfrmTramitedePermisos, frmTramitedePermisos);
    frmTramitedePermisos.ShowModal;
end;

procedure TfrmInteligent.rDiarioClick(Sender: TObject);
begin
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;


    if connection.contrato.FieldByName('eLugarOt').AsString='Tierra' then
    begin
      if not MostrarFormChild('frmDiarioTurnoTierra') then
      begin
          Application.CreateForm(TfrmDiarioTurnoTierra,frmDiarioTurnoTierra);
          frmDiarioTurnoTierra.show;
      end;


    end
    else
    begin

      if not MostrarFormChild('frmdiarioturno') then
      begin
          Application.CreateForm(TfrmDiarioTurno, frmDiarioTurno);
          frmDiarioTurno.show;
      end;
    end;
end;


procedure TfrmInteligent.opValidaClick(Sender: TObject);
begin
  if not MostrarFormChild('frmvalida') then
  begin
    Application.CreateForm(TfrmValida, frmValida);
    frmValida.show
  end;
end;

procedure TfrmInteligent.oficmodifClick(Sender: TObject);
begin
  if not MostrarFormChild('frmordenesgeneral') then
  begin
    Application.CreateForm(TfrmOrdenesGeneral, frmOrdenesGeneral);
    frmOrdenesGeneral.Show;
  end;
end;

procedure TfrmInteligent.opAbreClick(Sender: TObject);
begin
  if not MostrarFormChild('frmabrereporte') then
  begin
    Application.CreateForm(TfrmAbreReporte, frmAbreReporte);
    frmAbreReporte.show
  end;
end;

procedure TfrmInteligent.cPaquetesEqClick(Sender: TObject);
begin
  if not MostrarFormChild('frmgruposdeequipo') then
  begin
    Application.CreateForm(TfrmGruposDeEquipo, frmGruposdeEquipo);
    frmGruposdeEquipo.show
  end;
end;

procedure TfrmInteligent.cPaquetesPerClick(Sender: TObject);
begin
  if not MostrarFormChild('frmpaquetepersonal') then
  begin
    Application.CreateForm(TfrmPaquetePersonal, frmPaquetePersonal);
    frmPaquetePersonal.show
  end;
end;

procedure TfrmInteligent.cConsulta1Click(Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;
    if not MostrarFormChild('frmconsultaactividad2') then
    begin
      Application.CreateForm(TfrmConsultaActividad2, frmConsultaActividad2);
      frmConsultaActividad2.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Reportes Diarios', 0);
    end;
  end;

end;

procedure TfrmInteligent.adReg02Click(Sender: TObject);
begin
  if not MostrarFormChild('frmcalculoavancesxpartida') then
  begin
    Application.CreateForm(TfrmCalculoAvancesxPartida, frmCalculoAvancesxPartida);
    frmCalculoAvancesxPartida.show
  end;
end;

procedure TfrmInteligent.opFirmasClick(Sender: TObject);
begin
  global_orden := '';
  Application.CreateForm(TfrmFirmas, frmFirmas);
  frmfirmas.ShowModal;
end;

procedure TfrmInteligent.cAnexoClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;

  if connection.configuracion.FieldValues['sTipoContrato'] = 'Precio Unitario' then
  begin
    if not MostrarFormChild('frmactividadesanexo2') then
    begin
      Application.CreateForm(TfrmActividadesAnexo2, frmActividadesAnexo2);
      frmActividadesAnexo2.show
    end
  end
  else
    if connection.configuracion.FieldValues['sTipoContrato'] = 'Precio Unitario x OS' then
    begin
      if not MostrarFormChild('frmactividadesanexo') then
      begin
        Application.CreateForm(TfrmActividadesAnexo, frmActividadesAnexo);
        frmActividadesAnexo.show
      end
    end
end;

procedure TfrmInteligent.cVerificaClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmajustaanexo') then
  begin
    Application.CreateForm(TfrmAjustaAnexo, frmAjustaAnexo);
    frmAjustaAnexo.show
  end;
end;

procedure TfrmInteligent.dxBarLargeButton29Click(Sender: TObject);
begin
    opTiemposM.Click
end;

procedure TfrmInteligent.optDesautorizaClick(Sender: TObject);
begin
  if not MostrarFormChild('frmaperturaestimaciongral') then
  begin
    Application.CreateForm(TfrmAperturaEstimacionGral, frmAperturaEstimacionGral);
    frmAperturaEstimacionGral.show
  end;
end;

procedure TfrmInteligent.optEstimacionesClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmestimaciongeneral') then
  begin
    Application.CreateForm(TfrmEstimacionGeneral, frmEstimacionGeneral);
    frmEstimacionGeneral.show
  end;
end;

procedure TfrmInteligent.EquipoPT1Click(Sender: TObject);
begin
  if not MostrarFormChild('frmrecursosequipo') then
  begin
    Application.CreateForm(TfrmRecursosEquipo, frmRecursosEquipo);
    FrmRecursosEquipo.show;
  end;
end;

procedure TfrmInteligent.Estimaciones3Click(Sender: TObject);
begin
  opEstimaciones.Click
end;

procedure TfrmInteligent.estirado1Click(Sender: TObject);
begin
  detectar := ExtractFilePath(Application.Exename) + 'image.ini'; //extraepath
  if leeini(detectar) <> 'no' then
    modofondo(JvBackground1, 'bmStretch', detectar)
  else
    escribeinidefault(detectar, 'bmStretch');
end;

procedure TfrmInteligent.adFestivosClick(Sender: TObject);
begin
  if not MostrarFormChild('frmdiasfestivos') then
  begin
    Application.CreateForm(TfrmDiasFestivos, frmDiasFestivos);
    frmDiasFestivos.show
  end;
end;

procedure TfrmInteligent.cProgPlaticasClick(
  Sender: TObject);
begin
  if not MostrarFormChild('frmplaticas') then
  begin
    Application.CreateForm(TfrmPlaticas, frmPlaticas);
    frmPlaticas.show
  end;
end;

procedure TfrmInteligent.cProgramadosClick(Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmcomparativo2') then
    begin
      Application.CreateForm(TfrmComparativo2, frmComparativo2);
      frmComparativo2.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Comparativo de Avance', 0);
    end;
  end;
end;

procedure TfrmInteligent.opComparativo1Click(
  Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmgenerado') then
    begin
      Application.CreateForm(TfrmGenerado, frmGenerado);
      frmGenerado.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Comparativo Cantidad a Instalar Vs Generadores Vs Reportes Diarios', 0);
    end;
  end;

end;

procedure TfrmInteligent.adTurnosClick(Sender: TObject);
begin
  if not MostrarFormChild('frmturnos') then
  begin
    Application.CreateForm(TfrmTurnos, frmTurnos);
    frmTurnos.show
  end;
end;

procedure TfrmInteligent.adReg03Click(Sender: TObject);
begin
  if not MostrarFormChild('frmprocregavfisico') then
  begin
    Application.CreateForm(TfrmProcRegAvFisico, frmProcRegAvFisico);
    frmProcRegAvFisico.show
  end;
end;

procedure TfrmInteligent.opEstimacionesClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmestimaciones') then
  begin
    Application.CreateForm(TfrmEstimaciones, frmEstimaciones);
    frmEstimaciones.show
  end;
end;

procedure TfrmInteligent.cPaquetesActClick(Sender: TObject);
begin
  if not MostrarFormChild('frmactividadesxgrupo') then
  begin
    Application.CreateForm(TfrmActividadesxGrupo, frmActividadesxGrupo);
    frmActividadesxGrupo.show
  end;
end;

procedure TfrmInteligent.adDistProgramaClick(Sender: TObject);
begin
  if not MostrarFormChild('frmdistribucionprograma') then
  begin
    Application.CreateForm(TfrmDistribucionPrograma, frmDistribucionPrograma);
    frmDistribucionPrograma.show
  end;
end;

procedure TfrmInteligent.rComparativoClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmcompara') then
  begin
    Application.CreateForm(TfrmCompara, frmCompara);
    frmCompara.show
  end;
end;

procedure TfrmInteligent.cAvContratoClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmcomparativo') then
  begin
    Application.CreateForm(TfrmComparativo, frmComparativo);
    frmComparativo.show
  end;
end;

procedure TfrmInteligent.ChartProClick(Sender: TObject);
begin
  if not MostrarFormChild('fIntelChart') then
  begin
    Application.CreateForm(TIntelChart, FIntelChart);
    FIntelChart.show;
  end;
end;

procedure TfrmInteligent.ContenidoNotaCampo1Click(Sender: TObject);
begin
  Application.CreateForm(TFrmNotaCampo,FrmNotaCampo);
  FrmNotaCampo.show;
end;

procedure TfrmInteligent.ControldeCalidadRIR1Click(Sender: TObject);
begin
  Application.CreateForm( TfrmCalidad_Rir, frmCalidad_Rir );
  frmCalidad_Rir.ShowModal;
end;

procedure TfrmInteligent.Firmantes3Click(Sender: TObject);
begin
  Application.CreateForm(TfrmFirmas, frmFirmas);
  frmfirmas.show
end;

procedure TfrmInteligent.opTiemposMClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmreporteperiodo') then
  begin
    Application.CreateForm(TfrmReportePeriodo, frmReportePeriodo);
    frmReportePeriodo.show
  end;
end;

procedure TfrmInteligent.qComentariosClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmbuscacomentarios') then
  begin
    Application.CreateForm(TfrmBuscaComentarios, frmBuscaComentarios);
    frmBuscaComentarios.show
  end;
end;

procedure TfrmInteligent.JornadasEspeciales1Click(Sender: TObject);
begin
//     Application.CreateForm(TfrmJornadasEspeciales, frmJornadasEspeciales);
//     frmJornadasEspeciales.show
end;

procedure TfrmInteligent.JvAppEvents1ActiveControlChange(Sender: TObject);
begin
  self.Caption := sender.ClassName;
end;

procedure TfrmInteligent.ListadodePersonal1Click(Sender: TObject);
begin
  if not MostrarFormChild('frmListado_Personal') then
  begin
    Application.CreateForm(TfrmListado_Personal, frmListado_Personal);
    frmListado_Personal.Show;
  end;
end;

procedure TfrmInteligent.SpeedItem2Click(Sender: TObject);
begin
  adConfiguracion.Click
end;

procedure TfrmInteligent.SpeedItem3Click(Sender: TObject);
begin
  adentro := True;
  sSeleccion.Click;
end;

procedure TfrmInteligent.SpeedItem5Click(Sender: TObject);
begin
  qComentarios.Click
end;

procedure TfrmInteligent.SpeedItem6Click(Sender: TObject);
begin
  cConsulta4.Click
end;

procedure TfrmInteligent.SpeedItem7Click(Sender: TObject);
begin
  opEstimaciones.Click
end;

procedure TfrmInteligent.SpeedItem8Click(Sender: TObject);
begin
  cPaquetesPer.Click
end;

procedure TfrmInteligent.SpeedItem9Click(Sender: TObject);
begin
  cPaquetesEQ.Click
end;

procedure TfrmInteligent.SpeedItem12Click(Sender: TObject);
begin
  opOrdenCam.Click
end;

procedure TfrmInteligent.SpeedItem11Click(Sender: TObject);
begin
  rDiario.Click
end;

procedure TfrmInteligent.SpeedItem14Click(Sender: TObject);
begin
  opGeneradores.Click
end;

procedure TfrmInteligent.SpeedItem10Click(Sender: TObject);
begin
  opComparativo1.Click
end;


procedure TfrmInteligent.opProyeccionClick(Sender: TObject);
begin {
  if not MostrarFormChild('frmproyeccion') then
  begin
  Application.CreateForm(TfrmProyeccion, frmProyeccion);
  frmProyeccion.show
  end;}
  if not MostrarFormChild('frmProyeccion2') then
  begin
    Application.CreateForm(TfrmProyeccion2, frmProyeccion2);
    frmProyeccion2.show;
  end;
end;

procedure TfrmInteligent.adSqlClick(Sender: TObject);
begin
  if not MostrarFormChild('frmSqlManager') then
  begin
    Application.CreateForm(TfrmSqlManager, frmSqlManager);
    frmSqlManager.show;
  end;
end;

procedure TfrmInteligent.opSQLAnexoClick(Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmimportaciondedatos') then
    begin
      Application.CreateForm(TfrmImportaciondeDatos, frmImportaciondeDatos);
      frmImportaciondeDatos.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Importación de Datos', 0);
    end;
  end;

end;

procedure TfrmInteligent.cFactorCostoClick(Sender: TObject);
begin
  if not MostrarFormChild('frmfactordecosto') then
  begin
    Application.CreateForm(TfrmFactordeCosto, frmFactordeCosto);
    frmFactordeCosto.show
  end;
end;

procedure TfrmInteligent.mnCopiarParametroClick(Sender: TObject);
begin
  {if not MostrarFormChild('frmCopiaParametros') then
  begin
    Application.CreateForm(TFrmcopiaparametros, frmcopiaparametros);
    frmcopiaparametros.show;
  end;}
end;

procedure TfrmInteligent.mniGeneradoresClick(Sender: TObject);
begin
  Application.CreateForm(TFrmGeneradores,FrmGeneradores);
  FrmGeneradores.Show;
end;

procedure TfrmInteligent.mniListadoPersonalClick(Sender: TObject);
begin
    if not MostrarFormChild('frmListaPersonalV2') then
  begin
    Application.CreateForm(TfrmListaPersonalV2,frmListaPersonalV2);
    frmListaPersonalV2.Show;
  end;
end;

procedure TfrmInteligent.mniRendimientoClick(Sender: TObject);
begin
  {Application.CreateForm(TfrmRepExcelNew,frmRepExcelNew);
  frmRepExcelNew.ShowModal;}
end;

procedure TfrmInteligent.geAvFiFinClick(Sender: TObject);
begin
  if not MostrarFormChild('frmAvancesFinancieros') then
  begin
    Application.CreateForm(TfrmAvancesFinancieros, frmAvancesFinancieros);
    frmAvancesFinancieros.show
  end;
end;

procedure TfrmInteligent.Generaciondeinformes2Click(Sender: TObject);
begin
  rComparativo.Click
end;

procedure TfrmInteligent.Generadores2Click(Sender: TObject);
begin
  opGeneradores.Click
end;

procedure TfrmInteligent.adImportarClick(Sender: TObject);
begin
  if not MostrarFormChild('frmSqlImportar') then
  begin
    Application.CreateForm(TfrmSqlImportar, frmSqlImportar);
    frmSqlImportar.show;
  end;
end;

procedure TfrmInteligent.adExportarClick(Sender: TObject);
begin
  if not MostrarFormChild('frmSqlExportar') then
  begin
    Application.CreateForm(TfrmSqlExportar, frmSqlExportar);
    frmSqlExportar.show;
  end;
end;

procedure TfrmInteligent.adResidenciasClick(Sender: TObject);
begin
  if not MostrarFormChild('frmresidencias') then
  begin
    Application.CreateForm(TfrmResidencias, frmResidencias);
    frmResidencias.show
  end;
end;

procedure TfrmInteligent.adProgramasClick(Sender: TObject);
begin
  if not MostrarFormChild('frmprogramas') then
  begin
    Application.CreateForm(TfrmProgramas, frmProgramas);
    frmProgramas.show;
  end;
end;

procedure TfrmInteligent.adGruposClick(Sender: TObject);
begin
  if not MostrarFormChild('frmgrupos') then
  begin
    Application.CreateForm(TfrmGrupos, frmGrupos);
    frmGrupos.show;
  end;

end;

procedure TfrmInteligent.adGrupoPClick(Sender: TObject);
begin
  if not MostrarFormChild('frmgruposxprograma') then
  begin
    Application.CreateForm(TfrmGruposxPrograma, frmGruposxPrograma);
    frmGruposxPrograma.show;
  end;
end;

procedure TfrmInteligent.opProgramacionClick(Sender: TObject);
begin
  if not MostrarFormChild('frmpersonalprogramado') then
  begin
    Application.CreateForm(TfrmPersonalProgramado, frmPersonalProgramado);
    frmPersonalProgramado.show;
  end;
end;

procedure TfrmInteligent.gePersonalProgClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmPersonalConsolidado') then
  begin
    Application.CreateForm(TfrmPersonalConsolidado, frmPersonalConsolidado);
    frmPersonalConsolidado.show;
  end;
end;

procedure TfrmInteligent.GerencialBarco1Click(Sender: TObject);
begin
  if not MostrarFormChild('frmOpcionesGerencial') then
  begin
      Application.CreateForm(TfrmOpcionesGerencial, frmOpcionesGerencial);
      frmOpcionesGerencial.showModal;
  end;
end;

procedure TfrmInteligent.GraficadorClick(Sender: TObject);
begin

  if not MostrarFormChild('frmGraficador') then
  begin
    Application.CreateForm(TfrmGraficador, frmGraficador);
    frmGraficador.show;
  end;
end;

procedure TfrmInteligent.adUsuariosCClick(Sender: TObject);
begin
  if not MostrarFormChild('frmcontratosxusuario') then
  begin
    Application.CreateForm(TfrmContratosxUsuario, frmContratosxUsuario);
    frmContratosxUsuario.show;
  end;
end;

procedure TfrmInteligent.grPenasRetClick(Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    Application.CreateForm(TfrmRetencionesyPenas, frmRetencionesyPenas);
    frmRetencionesyPenas.showModal;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Reportes Diarios', 0);
    end;
  end;

end;

procedure TfrmInteligent.HorariosGerenciales1Click(Sender: TObject);
begin
  Application.CreateForm( TfrmHorariosGerenciales, frmHorariosGerenciales );
  frmHorariosGerenciales.ShowModal;
end;

procedure TfrmInteligent.cConsulta4Click(Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmconsultaactividad4') then
    begin
      Application.CreateForm(TfrmConsultaActividad4, frmConsultaActividad4);
      frmConsultaActividad4.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Reportes Diarios', 0);
    end;
  end;
end;

procedure TfrmInteligent.adImportarOkClick(Sender: TObject);
begin
  if not MostrarFormChild('frmImportarDiarios') then
  begin
    Application.CreateForm(TfrmImportarDiarios, frmImportarDiarios);
    frmImportarDiarios.show;
  end;
end;

procedure TfrmInteligent.SpeedItem4Click0(Sender: TObject);
begin
  opValida.Click
end;

procedure TfrmInteligent.SpeedItem15Click(Sender: TObject);
begin
  opAbre.Click
end;

procedure TfrmInteligent.SpeedItem16Click(Sender: TObject);
begin
  opFirmas.Click
end;

procedure TfrmInteligent.SpeedItem17Click(Sender: TObject);
begin
  rComparativo.Click
end;

procedure TfrmInteligent.SpeedItem16Click0(Sender: TObject);
begin
  opFirmas.Click
end;

procedure TfrmInteligent.AccesoObraClick(Sender: TObject);
begin
//<ROJAS>
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
//
  if not MostrarFormChild('frmControlDirecto') then
  begin
    Application.CreateForm(TfrmControlDirecto, frmControlDirecto);
    frmControlDirecto.show;
  end;
end;

procedure TfrmInteligent.adReg04Click(Sender: TObject);
begin
  if not MostrarFormChild('frmcalculoavancespaquetes') then
  begin
    Application.CreateForm(TfrmCalculoAvancesPaquetes, frmCalculoAvancesPaquetes);
    frmCalculoAvancesPaquetes.show
  end;
end;

procedure TfrmInteligent.opValidaEstClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;

  if not MostrarFormChild('frmvalidaestimacion') then
  begin
    Application.CreateForm(TfrmValidaEstimacion, frmValidaEstimacion);
    frmValidaEstimacion.show
  end;
end;

procedure TfrmInteligent.Panel2CanResize(Sender: TObject; var NewWidth,
  NewHeight: Integer; var Resize: Boolean);
begin
  Resize := True;
end;

procedure TfrmInteligent.permisosUsuarios;
var
  iElemento: Integer;
  sPrograma: string;
  Component: tMenuItem;
begin
  //Aqui se realiza la seguridad  del Sistema, permisos, privilegios..
  for iElemento := 0 to frmInteligent.ComponentCount - 1 do
  begin
    sPrograma := frmInteligent.Components[iElemento].GetNamePath;
    if frmInteligent.Components[iElemento].ClassName = 'TMenuItem' then
    begin
      Component := frmInteligent.Components[iElemento] as tMenuItem;
      if not Component.IsLine then
        if frmInteligent.Components[iElemento].Tag = 123 then
          Component.Enabled := True
        else
          if Global_Grupo = '' then
          begin
            //soad -> cambio de la consulta para la seguridad personalizada..
            Connection.UsuariosxPrograma.Active := False;
            Connection.UsuariosxPrograma.SQL.Clear;
            Connection.UsuariosxPrograma.SQL.Add('Select sIdPrograma from usuariosxprograma ' +
              'Where sIdUsuario = :Usuario and sIdPrograma = :Programa ');
            Connection.UsuariosxPrograma.Params.ParamByName('Usuario').DataType := ftString;
            Connection.UsuariosxPrograma.Params.ParamByName('Usuario').Value := global_usuario;
            Connection.UsuariosxPrograma.Params.ParamByName('Programa').DataType := ftString;
            Connection.UsuariosxPrograma.Params.ParamByName('Programa').Value := sPrograma;
            Connection.UsuariosxPrograma.Open;

            if Connection.UsuariosxPrograma.RecordCount > 0 then
              Component.Enabled := True
            else
              Component.Enabled := False;
          end
          else
          begin
            Connection.gruposxPrograma.Active := False;
            Connection.gruposxPrograma.Params.ParamByName('Grupo').DataType := ftString;
            Connection.gruposxPrograma.Params.ParamByName('Grupo').Value := global_grupo;
            Connection.gruposxPrograma.Params.ParamByName('Programa').DataType := ftString;
            Connection.gruposxPrograma.Params.ParamByName('Programa').Value := sPrograma;
            Connection.gruposxPrograma.Open;
            if Connection.gruposxPrograma.RecordCount > 0 then
              Component.Enabled := True
            else
              Component.Enabled := False;
          end
    end
  end;
end;

procedure TfrmInteligent.opOrdenCamClick(Sender: TObject);
begin
  if not MostrarFormChild('frmordendecambio') then
  begin
    Application.CreateForm(TfrmOrdendeCambio, frmOrdendeCambio);
    frmOrdendeCambio.show
  end;
end;

procedure TfrmInteligent.opAvisodeEmbClick(Sender: TObject);
begin
  stMenu := '';
  global_orden := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmentradaanex') then
  begin
    Application.CreateForm(TfrmEntradaAnex, frmEntradaAnex);
    frmEntradaAnex.Show;
  end;
end;

procedure TfrmInteligent.cProveedoresClick(Sender: TObject);
begin
  if not MostrarFormChild('frmproveedores') then
  begin
    Application.CreateForm(TfrmProveedores, frmProveedores);
    frmProveedores.show
  end;
end;

procedure TfrmInteligent.SubAdministradorClick(Sender: TObject);
begin
  if uppercase(Global_Usuario) = 'ADMIN' then
  begin
    if not MostrarFormChild('frmadministrarbd') then
    begin
      application.CreateForm(TFrmAdministrarBd, FrmAdministrarBd);
      FrmAdministrarBd.Show;
    end
  end
  else
    MessageDlg('Usted no Puede Acceder a Este Modulo.', mtInformation, [mbOk], 0);
end;

procedure TfrmInteligent.SubidadePersonal1Click(Sender: TObject);
begin
//     Application.CreateForm(TfrmEmpleados_subida, frmEmpleados_subida);
//     frmEmpleados_subida.show
end;

procedure TfrmInteligent.SpeedItem18Click(Sender: TObject);
begin
  opAvisodeEmb.Click
end;

procedure TfrmInteligent.SpeedItem19Click(Sender: TObject);
begin
  opTiemposM.Click
end;

procedure TfrmInteligent.SpeedItem20Click(Sender: TObject);
begin
  cConsulta5.Click
end;

procedure TfrmInteligent.cConsulta5Click(Sender: TObject);
begin
  if not MostrarFormChild('frmconsultaactividad3') then
  begin
    Application.CreateForm(TfrmConsultaActividad3, frmConsultaActividad3);
    frmConsultaActividad3.show
  end;
end;

procedure TfrmInteligent.cConsulta6Click(
  Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmconsultaxdescripcion') then
    begin
      Application.CreateForm(TfrmConsultaxDescripcion, frmConsultaxDescripcion);
      frmConsultaxDescripcion.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Reportes Diarios', 0);
    end;
  end;


end;

procedure TfrmInteligent.SpeedItem13Click(Sender: TObject);
begin
  cConsulta5.Click
end;

procedure TfrmInteligent.SpeedItem10Click0(Sender: TObject);
begin
  opComparativo1.Click
end;

procedure TfrmInteligent.SpeedItem22Click(Sender: TObject);
begin
  cConsulta1.Click
end;

procedure TfrmInteligent.cConsulta3Click(Sender: TObject);
begin
  if not MostrarFormChild('frmconsultaactividad') then
  begin
    Application.CreateForm(TfrmConsultaActividad, frmConsultaActividad);
    frmConsultaActividad.show
  end;
end;

procedure TfrmInteligent.adExportar2Click(Sender: TObject);
begin
  if not MostrarFormChild('frmExportaGeneral') then
  begin
    Application.CreateForm(TfrmExportaGeneral, frmExportaGeneral);
    frmExportaGeneral.show;
  end;
end;

procedure TfrmInteligent.cConsulta7Click(Sender: TObject);
begin
  if not MostrarFormChild('frmactividadesanexo') then
  begin
    Application.CreateForm(TfrmActividadesAnexo, frmActividadesAnexo);
    frmActividadesAnexo.show
  end;
end;

procedure TfrmInteligent.FormResize(Sender: TObject);
begin
  ImgKardex.Top := frmInteligent.Height - 132;
  ImgKardex.Left := frmInteligent.Width - 63;

end;

procedure TfrmInteligent.imgKardexClick(Sender: TObject);
begin
  Application.CreateForm(TfrmKardex, frmKardex);
  frmKardex.show
end;

procedure TfrmInteligent.InformedeSincronizado1Click(Sender: TObject);
begin

  if not MostrarFormChild('frmInformeSincronizacion') then
  begin
    Application.CreateForm(TfrmInformeSincronizacion, frmInformeSincronizacion);
    frmInformeSincronizacion.show ;
  end;
end;

procedure TfrmInteligent.inteligentpopPopup(Sender: TObject);
begin
  cambiarimagendefondo1.Enabled := TRUE;
  irareportesdiarios1.ENABLED := TRUE;
  iraestimaciones1.ENABLED := TRUE;
  irageneradores1.ENABLED := TRUE;
  irageneradoresdeinformes1.ENABLED := TRUE;
  cambiarmododefondo1.Enabled := true;
  estirado1.Enabled := true;
  mosaico1.Enabled := true;
  centrado1.Enabled := true;
  Ventanasen1.Enabled := true;
  Cascada1.Enabled := true;
  MosaicoVertical1.Enabled := true;
  Mosaicohorizontal1.Enabled := true;
end;

procedure TfrmInteligent.Iraestimaciones1Click(Sender: TObject);
begin
  opEstimaciones.Click
end;

procedure TfrmInteligent.Irageneradores1Click(Sender: TObject);
begin
  opGeneradores.Click;
end;

procedure TfrmInteligent.Irageneradoresdeinformes1Click(Sender: TObject);
begin
  rComparativo.Click;
end;

procedure TfrmInteligent.Irareportesdiarios1Click(Sender: TObject);
begin
  rDiario.Click;
end;

procedure TfrmInteligent.sCambiaPClick(Sender: TObject);
begin
  Application.CreateForm(TfrmCambioPassword, frmCambioPassword);
  frmcambiopassword.ShowModal
end;

procedure TfrmInteligent.adCancelacionClick(
  Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmcancelacion') then
  begin
    Application.CreateForm(TfrmCancelacion, frmCancelacion);
    frmCancelacion.Show
  end;
end;

procedure TfrmInteligent.rGerencialClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmreportegerencial') then
  begin
    Application.CreateForm(TfrmReporteGerencial, frmReporteGerencial);
    frmReporteGerencial.Show
  end;
end;

procedure TfrmInteligent.opRequisicionesClick(Sender: TObject);
begin
   // if Connection.configuracion.FieldValues['sCampPerf'] = 'No' then
   //  begin
   //    Application.CreateForm(TfrmRequisicion, frmRequisicion);
   //    frmRequisicion.show
   //  end
   //  else
   //  begin
  stMenu := '';
  if Sender is TMenuItem then
    stMenu := (Sender as TMenuItem).Name;
  if not MostrarFormChild('frmrequisicionperf') then
  begin
    Application.CreateForm(TfrmRequisicionPerf, frmRequisicionPerf);
    frmRequisicionPerf.show
  end;
 // end;

end;

procedure TfrmInteligent.opPedidosClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmpedidos') then
  begin
    Application.CreateForm(TfrmPedidos, frmPedidos);
    frmPedidos.show
  end;
end;

procedure TfrmInteligent.opComparativo5Click(Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;
    if not MostrarFormChild('frmconsultaactividad5') then
    begin
      Application.CreateForm(TfrmConsultaActividad5, frmConsultaActividad5);
      frmConsultaActividad5.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Historico de Movimientos de Partidas Anexo', 0);
    end;
  end;
end;

procedure TfrmInteligent.SpeedItem23Click(Sender: TObject);
begin
  opComparativo5.Click
end;

procedure TfrmInteligent.rSintesisClick(Sender: TObject);
begin
  if not MostrarFormChild('frmsintesisgerencial') then
  begin
    Application.CreateForm(TfrmSintesisGerencial, frmSintesisGerencial);
    frmSintesisGerencial.show
  end;
end;

procedure TfrmInteligent.rt1Click(Sender: TObject);
begin
  showmessage('ok');
end;

procedure TfrmInteligent.sTipsClick(Sender: TObject);
begin
  if not MostrarFormChild('frmtipsdia') then
  begin
    Application.CreateForm(TfrmTipsDia, frmTipsDia);
    frmTipsDia.Show;
  end;
end;

procedure TfrmInteligent.sWarningClick(Sender: TObject);
begin
  if not MostrarFormChild('frmwarningdia') then
  begin
    Application.CreateForm(TfrmWarningDia, frmWarningDia);
    frmWarningDia.Show;
  end;
end;

procedure TfrmInteligent.TiempoTimer(Sender: TObject);
begin
  application.Terminate;
end;

procedure TfrmInteligent.ToolButton10Click(Sender: TObject);
begin
  qComentarios.Click
end;

procedure TfrmInteligent.ToolButton11Click(Sender: TObject);
begin
  opComparativo5.Click
end;

procedure TfrmInteligent.ToolButton13Click(Sender: TObject);
begin
  rDiario.Click
end;

procedure TfrmInteligent.ToolButton14Click(Sender: TObject);
begin
  Application.CreateForm(TfrmDiarioBarco, frmDiarioBarco);
  frmDiarioBarco.show
end;

procedure TfrmInteligent.ToolButton15Click(Sender: TObject);
begin
  opAvisodeEmb.Click
end;

procedure TfrmInteligent.ToolButton17Click(Sender: TObject);
begin
  opTiemposM.Click
end;

procedure TfrmInteligent.ToolButton18Click(Sender: TObject);
begin
  opOrdenCam.Click
end;

procedure TfrmInteligent.ToolButton1Click(Sender: TObject);
begin
  adConfiguracion.Click
end;

procedure TfrmInteligent.ToolButton21Click(Sender: TObject);
begin
  opFirmas.Click
end;

procedure TfrmInteligent.ToolButton22Click(Sender: TObject);
begin
  opValida.Click
end;

procedure TfrmInteligent.ToolButton23Click(Sender: TObject);
begin
  opAbre.Click
end;

procedure TfrmInteligent.ToolButton24Click(Sender: TObject);
begin
  opEstimaciones.Click
end;

procedure TfrmInteligent.ToolButton25Click(Sender: TObject);
begin
  opGeneradores.Click
end;

procedure TfrmInteligent.ToolButton27Click(Sender: TObject);
begin
  rComparativo.Click
end;

procedure TfrmInteligent.ToolButton2Click(Sender: TObject);
begin
  adentro := True;
  sSeleccion.Click;
end;

procedure TfrmInteligent.ToolButton3Click(Sender: TObject);
begin
  cPaquetesPer.Click
end;

procedure TfrmInteligent.ToolButton4Click(Sender: TObject);
begin
  cPaquetesEQ.Click
end;

procedure TfrmInteligent.ToolButton6Click(Sender: TObject);
begin
  cConsulta1.Click
end;

procedure TfrmInteligent.ToolButton7Click(Sender: TObject);
begin
  opComparativo1.Click
end;

procedure TfrmInteligent.ToolButton8Click(Sender: TObject);
begin
  cConsulta5.Click
end;

procedure TfrmInteligent.ToolButton9Click(Sender: TObject);
begin
  cConsulta4.Click
end;

procedure TfrmInteligent.adAvisosClick(Sender: TObject);
begin
  if not MostrarFormChild('frmavisosalertas') then
  begin
    Application.CreateForm(TfrmAvisosAlertas, frmAvisosAlertas);
    frmAvisosAlertas.show
  end;
end;

procedure TfrmInteligent.adActivosClick(Sender: TObject);
begin
  try
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;

    if not MostrarFormChild('frmactivos') then
    begin
      Application.CreateForm(TfrmActivos, frmActivos);
      frmActivos.show
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Comparativo de Avances', 0);
    end;
  end;

end;

procedure TfrmInteligent.opFiltroClick(Sender: TObject);
begin
  if not MostrarFormChild('frmfiltros') then
  begin
    Application.CreateForm(TfrmFiltros, frmFiltros);
    frmFiltros.show
  end;
end;

procedure TfrmInteligent.cTrinomiosClick(Sender: TObject);
begin
  Application.CreateForm(TfrmTrinomios, frmTrinomios);
  frmTrinomios.show
end;

procedure TfrmInteligent.opadmonCatalogoClick(Sender: TObject);
begin
  if not MostrarFormChild('frmadmoncatalogos') then
  begin
    Application.CreateForm(TfrmAdmonCatalogos, frmAdmonCatalogos);
    frmAdmonCatalogos.show
  end;
end;

procedure TfrmInteligent.opGeneradorSubClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmestimaproveedor') then
  begin
    Application.CreateForm(TfrmEstimaProveedor, frmEstimaProveedor);
    frmEstimaProveedor.show
  end;

end;

procedure TfrmInteligent.adSubContratosClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
    ///JLR
  if not MostrarFormChild('frmsubcontratos') then
  begin
    Application.CreateForm(TfrmSubContratos, frmSubContratos);
    frmSubContratos.show
  end;
end;

procedure TfrmInteligent.FormActivate(Sender: TObject);
var
  iElemento: Integer;
  sPrograma: string;
  Component: tMenuItem;

begin
end;

procedure TfrmInteligent.opAyudaClick(Sender: TObject);
begin
  Application.HelpFile := 'C:\inteligent\inteligenthelp.chm';
  Application.HelpCommand(HELP_CONTENTS, 0);
end;

procedure TfrmInteligent.mnuPersonal2Click(Sender: TObject);
begin
  Application.CreateForm(TFrmMoe,FrmMoe);
  with FrmMoe do
    try
      ShowModal;
    finally
      Destroy;
    end;
end;

procedure TfrmInteligent.mnuPozosdePerfoClick(Sender: TObject);
begin
  if not MostrarFormChild('frmeqpozos') then
  begin
    Application.CreateForm(TfrmEqPozos, frmEqPozos);
    frmEqPozos.show
  end;
end;

procedure TfrmInteligent.MnuSalAlmacenClick(Sender: TObject);
begin
  stMenu := '';
  if Sender is TMenuItem then
    stMenu := (Sender as TMenuItem).Name;
  if not MostrarFormChild('frmsalidaalmacen') then
  begin
    Application.CreateForm(TfrmSalidaAlmacen, frmSalidaAlmacen);
    frmSalidaAlmacen.show
  end;
end;

procedure TfrmInteligent.MOE1Click(Sender: TObject);
begin
  Application.CreateForm( TFrmMoeBordo, FrmMoeBordo );
  FrmMoeBordo.ShowModal;
end;

procedure TfrmInteligent.mosaico1Click(Sender: TObject);
begin
  detectar := ExtractFilePath(Application.Exename) + 'image.ini'; //extraepath
  if leeini(detectar) <> 'no' then
    modofondo(JvBackground1, 'bmTile', detectar)
  else
    escribeinidefault(detectar, 'bmTile');
end;

procedure TfrmInteligent.MovimientosdeBarcoPT1Click(Sender: TObject);
begin
  if not MostrarFormChild('frmrecursosmovimientos') then
  begin
    Application.CreateForm(TfrmRecursosMovimientos, frmRecursosMovimientos);
    FrmRecursosMovimientos.show;
  end;
end;

procedure TfrmInteligent.MovimientosdeBarcoPT2Click(Sender: TObject);
begin
  if not MostrarFormChild('frmrecursospersonal') then
  begin
    Application.CreateForm(TfrmRecursosPersonal, frmRecursosPersonal);
    FrmRecursosPersonal.show;
  end;
end;

procedure TfrmInteligent.MovPeroficClick(Sender: TObject);
begin
  if not MostrarFormChild('frmmovtospersonalxoficio') then
  begin
    Application.CreateForm(TfrmMovtosPersonalxoficio, frmMovtosPersonalxoficio);
    FrmMovtosPersonalxoficio.show;
  end;
end;

procedure TfrmInteligent.NombresdeFirmantes1Click(Sender: TObject);
begin
  if not MostrarFormChild('frmcatnomfirmates') then
  begin
    Application.CreateForm(Tfrmcatnomfirmates, frmcatnomfirmates);
    frmcatnomfirmates.show
  end;
end;

procedure TfrmInteligent.rInstaladoClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmdetalledeinstalacion') then
  begin
    Application.CreateForm(TfrmDetalledeInstalacion, frmDetalledeInstalacion);
    frmDetalledeInstalacion.show
  end;
end;

procedure TfrmInteligent.tripulacionClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  Application.CreateForm(TfrmTripulacion, frmTripulacion);
  frmTripulacion.Show;
end;

procedure TfrmInteligent.Unificadordeequipos1Click(Sender: TObject);
begin
  if not MostrarFormChild('FrmUnificadorEquipos') then
  begin
    Application.CreateForm(TFrmUnificadorEquipos, FrmUnificadorEquipos);
    FrmUnificadorEquipos.show
  end;
end;

procedure TfrmInteligent.optValidaClick(Sender: TObject);
begin
  if not MostrarFormChild('frmvalidaestimaciongral') then
  begin
    Application.CreateForm(TfrmValidaEstimacionGral, frmValidaEstimacionGral);
    frmValidaEstimacionGral.show
  end;
end;

procedure TfrmInteligent.PernoctasPT1Click(Sender: TObject);
begin
  if not MostrarFormChild('frmrecursospernocta') then
  begin
    Application.CreateForm(TfrmRecursosPernocta, frmRecursosPernocta);
    FrmRecursosPernocta.show;
  end;
end;

procedure TfrmInteligent.ProyecciondeActClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmProyeccion2') then
  begin
    Application.CreateForm(TfrmProyeccion2, frmProyeccion2);
    frmProyeccion2.show;
  end;
end;

procedure TfrmInteligent.mnuAgrupacionPClick(Sender: TObject);
begin
  if not MostrarFormChild('frmgrupospersonal') then
  begin
    Application.CreateForm(TfrmGruposPersonal, frmGruposPersonal);
    frmGruposPersonal.show
  end;
end;

procedure TfrmInteligent.MnuAlmacenClick(Sender: TObject);
begin
  stMenu := '';
  if Sender is TMenuItem then
    stMenu := (Sender as TMenuItem).Name;
  if not MostrarFormChild('frmAlmacenes') then
  begin
    Application.CreateForm(TfrmAlmacenes, frmAlmacenes);
    frmAlmacenes.show
  end;
end;

procedure TfrmInteligent.tbbPaqEqClick(Sender: TObject);
begin
    cPaquetesEQ.Click
end;

procedure TfrmInteligent.tbbRepDiarioClick(Sender: TObject);
begin
    rDiario.Click
end;

procedure TfrmInteligent.tbbRepBarcoClick(Sender: TObject);
begin
    reporteBarco.Click;
end;

procedure TfrmInteligent.tbbAvisoEmbClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmdetalledeinstalacion') then
  begin
    Application.CreateForm(TfrmDetalledeInstalacion, frmDetalledeInstalacion);
    frmDetalledeInstalacion.show
  end;
end;

procedure TfrmInteligent.tbbFotosClick(Sender: TObject);
begin
  opTiemposM.Click
end;

procedure TfrmInteligent.tbbOrdenCambioClick(Sender: TObject);
begin
  opOrdenCam.Click
end;

procedure TfrmInteligent.tbbFirmantesClick(Sender: TObject);
begin
  opFirmas.Click
end;

procedure TfrmInteligent.tbbAutorizaClick(Sender: TObject);
begin
  opValida.Click
end;

procedure TfrmInteligent.tbbSetupClick(Sender: TObject);
begin
  adConfiguracion.Click
end;

procedure TfrmInteligent.tbbDesautorizaClick(Sender: TObject);
begin
  opAbre.Click
end;

procedure TfrmInteligent.tbbEstimaClick(Sender: TObject);
begin
  opEstimaciones.Click
end;

procedure TfrmInteligent.tbbGeneraClick(Sender: TObject);
begin
  opGeneradores.Click;

end;

procedure TfrmInteligent.tbbInformesClick(Sender: TObject);
begin
  rComparativo.Click
end;

procedure TfrmInteligent.AdvToolBarButton24Click(Sender: TObject);
begin
  rComparativo.Click
end;

procedure TfrmInteligent.tbbCambiaContratoClick(Sender: TObject);
begin
  adentro := True;
  sSeleccion.Click;
end;

procedure TfrmInteligent.tbbConsult1Click(Sender: TObject);
begin
  cConsulta1.Click
end;

procedure TfrmInteligent.tbbPaquetePerClick(Sender: TObject);
begin
  cPaquetesPer.Click
end;

procedure TfrmInteligent.tbbProcGeneradorClick(Sender: TObject);
begin
  Application.CreateForm(TfrmProcesaGenerador, frmProcesaGenerador);
  frmProcesaGenerador.Show;
end;

procedure TfrmInteligent.tbbConsult2Click(Sender: TObject);
begin
 // cConsulta4.Click
 mniRendimiento.Click;
end;

procedure TfrmInteligent.tbbConsult3Click(Sender: TObject);
begin
  opComparativo1.Click
end;

procedure TfrmInteligent.tbbConsult4Click(Sender: TObject);
begin
  cConsulta5.Click
end;

procedure TfrmInteligent.tbbConsult5Click(Sender: TObject);
begin
  qComentarios.Click
end;

procedure TfrmInteligent.tbbConsult6Click(Sender: TObject);
begin
//  opComparativo5.Click
  ContenidoNotaCampo1.Click;
end;

procedure TfrmInteligent.AdvToolBarMenuButton1Click(Sender: TObject);
begin
  adConfiguracion.Click
end;

procedure TfrmInteligent.AgrupadordePersonal1Click(Sender: TObject);
begin
  Application.CreateForm(TfrmGruposDePersonal, frmGruposDePersonal);
  frmGruposDePersonal.Show;
end;

procedure TfrmInteligent.AsignaciondeOrdClick(
  Sender: TObject);
begin
  if not MostrarFormChild('frmordenesxusuario') then
  begin
    Application.CreateForm(TfrmOrdenesxUsuario, frmOrdenesxusuario);
    frmOrdenesxUsuario.show;
  end;
end;

procedure TfrmInteligent.Button1Click(Sender: TObject);
begin
  Application.CreateForm(TFrmImportaProject,FrmImportaProject);
  FrmImportaProject.show;
end;

procedure TfrmInteligent.RepBarcoClick(Sender: TObject);
begin
  if not MostrarFormChild('frmDiarioBarco') then
  begin
    Application.CreateForm(TfrmDiarioBarco, frmDiarioBarco);
    frmDiarioBarco.show;
  end;
end;

procedure TfrmInteligent.reporteBarcoClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if (not MostrarFormChild('frmDiarioBarco')) then
  begin
    //if global_contrato = global_contrato_barco then
    //begin
    Application.CreateForm(TfrmDiarioBarco, frmDiarioBarco);
    frmDiarioBarco.show;
    //end
    //else begin
    //  showmessage('Para acceder a este módulo, debe seleccionar el contrato de barco (' + global_contrato_barco + ')' );
    //end;
  end;
end;

procedure TfrmInteligent.ReportedeProduccionClick(Sender: TObject);
begin
  if not MostrarFormChild('frmOpcionesReporteProduccion') then
  begin
      Application.CreateForm(TfrmOpcionesReporteProduccion, frmOpcionesReporteProduccion);
      frmOpcionesReporteProduccion.showModal;
  end;
end;

procedure TfrmInteligent.Reportesdiarios1Click(Sender: TObject);
begin
  rDiario.Click;
end;

procedure TfrmInteligent.ReportesProduccionSabanas1Click(Sender: TObject);
begin
  rComparativo.Click
end;

procedure TfrmInteligent.ReprogramacionesXFolios1Click(Sender: TObject);
begin
  Application.CreateForm( TfrmReprogramacionFolio, frmReprogramacionFolio );
  frmReprogramacionFolio.ShowModal; 
end;

procedure TfrmInteligent.ResumendePersonal1Click(Sender: TObject);
begin
  if not MostrarFormChild('FrmResumenPersonal') then
  begin
    Application.CreateForm(TFrmResumenPersonal, FrmResumenPersonal);
    frmTripulacion.Show;
  end;
 
end;

procedure TfrmInteligent.sTripulacionClick(Sender: TObject);
begin
  if not MostrarFormChild('frmtripulacion') then
  begin
    Application.CreateForm(TfrmTripulacion, frmtripulacion);
    frmTripulacion.Show;
  end;
end;

procedure TfrmInteligent.opInventarioClick(Sender: TObject);
begin
  try
    //ROJAS
    stMenu := '';
    if sender is TMenuItem then
      stMenu := (sender as TMenuItem).Name;
   //
    if not MostrarFormChild('frmconsultaactividad5') then
    begin
      Application.CreateForm(TfrmConsultaActividad5, frmConsultaActividad5);
      frmConsultaActividad5.show
    end;
   //ROJAS
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ventana Principal', 'Al abrir ventana Historico de Movimientos de Partidas Anexo', 0);
    end;
  end;
   //
end;

procedure TfrmInteligent.opGeneradoresClick(Sender: TObject);
begin
  stMenu := '';
  if sender is TMenuItem then
    stMenu := (sender as TMenuItem).Name;
  if not MostrarFormChild('frmestimainstalado') then
  begin
    Application.CreateForm(TfrmEstimaInstalado, frmEstimaInstalado);
    frmEstimaInstalado.show
  end;
end;

procedure TfrmInteligent.rPartidasIsomClick(Sender: TObject);
begin
  stMenu := '';
  if Sender is TMenuItem then
    stMenu := (Sender as TMenuItem).Name;

  if not MostrarFormChild('frmpartidasxisometrico') then
  begin
    Application.CreateForm(TfrmPartidasxIsometrico, frmPartidasxIsometrico);
    frmPartidasxIsometrico.show
  end;
end;

procedure TfrmInteligent.mnuConversionesClick(Sender: TObject);
begin
  if not MostrarFormChild('frmConversiones') then
  begin
    Application.CreateForm(TfrmConversiones, frmConversiones);
    frmConversiones.show;
  end;
end;

procedure TfrmInteligent.MnuEntAlmaceClick(Sender: TObject);
begin
  stMenu := '';
  if Sender is TMenuItem then
    stMenu := (Sender as TMenuItem).Name;
  if not MostrarFormChild('frmentradaalmacen') then
  begin
    Application.CreateForm(TfrmEntradaAlmacen, frmEntradaAlmacen);
    frmEntradaAlmacen.show
  end;
end;

procedure TfrmInteligent.MnuFamiliadeProClick(Sender: TObject);
begin
  if not MostrarFormChild('frmgrupofamilias') then
  begin
    Application.CreateForm(TfrmGrupoFamilias, frmGrupoFamilias);
    frmGrupoFamilias.show
  end;
end;

procedure TfrmInteligent.MnuImpAvContClick(Sender: TObject);
begin
  if not MostrarFormChild('frmActualizacionRemota') then
  begin
    application.CreateForm(TfrmActualizacionRemota, frmActualizacionRemota);
    frmActualizacionRemota.Show;
  end;
end;

procedure TfrmInteligent.mnuKardexClick(Sender: TObject);
begin
  if not MostrarFormChild('frmKardex') then
  begin
    Application.CreateForm(TfrmKardex, frmKardex);
    frmKardex.Show;
  end;

end;

end.

