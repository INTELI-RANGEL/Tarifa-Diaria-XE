unit frm_bitacoradepartamental_2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms, UnitTBotonesPermisos,
  Dialogs, Grids, DBGrids, StdCtrls, ComCtrls, frm_Connection, DB, DateUtils, StrUtils,
  frm_barra, DBCtrls, Mask, Global, Menus, Buttons, Utilerias, ExtCtrls, Math, ComObj,
  frxClass, ImgList, ActnList, PanelDown, Newpanel, ZAbstractRODataset, ZDataset, ADODB,
  ZAbstractDataset, RxLookup, RXDBCtrl, rxCurrEdit, rxToolEdit, ClipBrd, OleServer,
  frm_EditorBitacoraDepartamental, RxMemDS, udbgrid, UnitTarifa,
  unitactivapop, UFunctionsGHH, UnitValidacion, RXCtrls,
  AdvGlowButton, masUtilerias, NxPageControl, JvExDBGrids, JvDBGrid, Jpeg,
  JvDBUltimGrid, JvExControls, JvDBLookup,
  commctrl, NxCollection, unt_Actividades,
  NxEdit, UnitPatrick, UnitExcel,
  cxLabel, cxImage, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit,
  ZSqlProcessor, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint,
  dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide,
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
  cxButtons, cxTextEdit, cxGroupBox, cxDBEdit, cxMaskEdit,
  cxDropDownEdit, cxStyles, dxSkinscxPCPainter, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxNavigator, cxDBData, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxClasses, cxGridCustomView, cxGrid,
  AdvGlassButton, cxSpinEdit, cxPC, dxDockControl, dxDockPanel;
type
Evalidaciones = class(Exception)
end;
type
  TfrmBitacoraDepartamental_2 = class(TForm)
    {$REGION 'Componentes'}
    ds_ordenesdetrabajo: TDataSource;
    ds_tiposdemovimiento: TDataSource;
    ds_bitacora: TDataSource;
    ds_actividadesiguales: TDataSource;
    ImageGrupos: TImageList;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    ds_ParidasEfectivas: TDataSource;
    ActividadesIguales: TZReadOnlyQuery;
    ActividadesIgualessWbs: TStringField;
    ActividadesIgualessNumeroActividad: TStringField;
    ActividadesIgualesmDescripcion: TMemoField;
    ActividadesIgualesdCantidad: TFloatField;
    ActividadesIgualesdInstalado: TFloatField;
    ActividadesIgualesdExcedente: TFloatField;
    ActividadesIgualesdPonderado: TFloatField;
    ActividadesIgualessMedida: TStringField;
    ActividadesIgualesdRestante: TFloatField;
    ordenesdetrabajo: TZReadOnlyQuery;
    TiposdeMovimiento: TZReadOnlyQuery;
    MaximoDiario: TZReadOnlyQuery;
    AvanceMaximo: TZReadOnlyQuery;
    ReporteDiario: TZReadOnlyQuery;
    QryPartidasEfectivas: TZReadOnlyQuery;
    QryBitacora: TZReadOnlyQuery;
    QryBitacorasContrato: TStringField;
    QryBitacorasNumeroOrden: TStringField;
    QryBitacoraiIdDiario: TIntegerField;
    QryBitacorasIdTurno: TStringField;
    QryBitacorasWbs: TStringField;
    QryBitacorasNumeroActividad: TStringField;
    QryBitacorasIdTipoMovimiento: TStringField;
    QryBitacoradCantidad: TFloatField;
    QryBitacoradAvance: TFloatField;
    QryBitacoramDescripcion: TMemoField;
    QryBitacorasDescripcion: TStringField;
    QryBitacorasMedida: TStringField;
    QryBitacoradVentaMN: TFloatField;
    QryBitacoradVentaDLL: TFloatField;
    QryBitacoradTotalMN: TCurrencyField;
    QryExistePartida: TZReadOnlyQuery;
    ActividadesIgualessWbsAnterior: TStringField;
    rDiario: TfrxReport;
    QryBitacorasTurno: TStringField;
    dsTiemposExtras: TDataSource;
    QryBitacorasWbsAnterior: TStringField;
    QryBitacorasHoraInicio: TStringField;
    QryBitacorasHoraFinal: TStringField;
    N7: TMenuItem;
    RevisarOrtografia2: TMenuItem;
    ActividadesIgualessTipoAnexo: TStringField;
    RxAvances: TRxMemoryData;
    RxAvancesdCantidad: TFloatField;
    RxAvancessMedida: TStringField;
    RxAvancesdCantidadActual: TFloatField;
    RxAvancesdCantidadAnterior: TFloatField;
    RxAvancesdCantidadAcumulada: TFloatField;
    QryBitacoralImprime: TStringField;
    QryBitacoralCancelada: TStringField;
    QryBitacorasDescripcionPartida: TStringField;
    ds_notasGerencial: TDataSource;
    popNotas: TPopupMenu;
    popAdd: TMenuItem;
    popEdit: TMenuItem;
    popPost: TMenuItem;
    popCancel: TMenuItem;
    popDelete: TMenuItem;
    N3: TMenuItem;
    N6: TMenuItem;
    CosolidarNotasGerenciales1: TMenuItem;
    QryBitacorahGerencial: TStringField;
    SaveDialog1: TSaveDialog;
    QryBitacorasNumeroActividad_ADM: TStringField;
    QryBitacorasWbs_ADM: TStringField;
    QryBitacorasTipoAnexo: TStringField;
    ImgBtns: TImageList;
    dsClasificacion: TDataSource;
    QrNotasDetalle: TZReadOnlyQuery;
    dsNotasDetalle: TDataSource;
    QrOt: TZReadOnlyQuery;
    dsOt: TDataSource;
    QrFrentes: TZReadOnlyQuery;
    dsFrentes: TDataSource;
    pnl7: TPanel;
    pnl8: TPanel;
    Label1: TLabel;
    tdIdFecha: TDateTimePicker;
    pnl9: TPanel;
    tNewGroupBox1: tNewGroupBox;
    tdAvanceGlobal: TCurrencyEdit;
    pnl10: TPanel;
    pnl12: TPanel;
    pnl16: TPanel;
    ds_Plataformas: TDataSource;
    Plataformas: TZReadOnlyQuery;
    LblReportados: TNxLinkLabel;
    QryBitacoradIdFecha: TDateField;
    btn1: TSpeedButton;
    Panel2: TPanel;
    lblFolio: TLabel;
    GrdOrden: TDBGrid;
    LblTodos: TNxLinkLabel;
    SpeedButton1: TSpeedButton;
    zqryDetalle: TZQuery;
    dsDetalle: TDataSource;
    QrFrentessNumeroOrden: TStringField;
    QrFrentessContrato: TStringField;
    QrFrentescIdStatus: TStringField;
    QrFrentessIdTipoMovimiento: TStringField;
    Panel4: TPanel;
    cxImage1: TcxImage;
    cxLabel4: TcxLabel;
    cxLabel1: TcxLabel;
    cxImage2: TcxImage;
    cxImage3: TcxImage;
    cxLabel2: TcxLabel;
    zfol: TZReadOnlyQuery;
    QryBitacoralaplicalibro: TStringField;
    tNewGroupBox2: tNewGroupBox;
    Label4: TLabel;
    Label5: TLabel;
    Label11: TLabel;
    tsIdTipoMovimiento: TDBLookupComboBox;
    tsHoraInicio: TMaskEdit;
    tsHoraFinal: TMaskEdit;
    tmDescripcion: TMemo;
    chkCancelada: TCheckBox;
    GridNotas: TRxDBGrid;
    btnUp: TAdvGlowButton;
    btnDown: TAdvGlowButton;
    tdAvance: TRxCalcEdit;
    ActividadesIgualesAvance: TFloatField;
    ImprimirSeccion1: TMenuItem;
    Actividades1: TMenuItem;
    tdCantidad: TRxCalcEdit;
    lbl19: TLabel;
    QrClasificacion: TZReadOnlyQuery;
    tssIdClasificacion: TDBLookupComboBox;
    tsNumeroActividad_PU: TRxDBLookupCombo;
    tsNumeroActividad_ADM: TRxDBLookupCombo;
    QryPartidasEfectivas_ADM: TZReadOnlyQuery;
    ds_PartidasEfectivas_ADM: TDataSource;
    ds_PartidasEfectivas_PU: TDataSource;
    QryPartidasEfectivas_PU: TZReadOnlyQuery;
    cxImageList1: TcxImageList;
    cxButton1: TcxButton;
    cxButton2: TcxButton;
    grpTipoActividad: TcxGroupBox;
    cxLabel3: TcxLabel;
    cxButton3: TcxButton;
    btnSalirTiposActividad: TcxButton;
    dbTipoActividad: TcxDBTextEdit;
    cbbTipoActividad: TComboBox;
    ActividadesIgualessWbsContrato: TStringField;
    grid_iguales: TRxDBGrid;
    cxComboGerencial: TcxComboBox;
    Grid_Bitacora: TcxGrid;
    BView_Actividades: TcxGridDBTableView;
    sNumeroActividad: TcxGridDBColumn;
    mDescripcion: TcxGridDBColumn;
    hGerencial: TcxGridDBColumn;
    dAvance: TcxGridDBColumn;
    Grid_BitacoraLevel1: TcxGridLevel;
    csStilos: TcxStyleRepository;
    cxDiurno: TcxStyle;
    cxNocturno: TcxStyle;
    cxNormal: TcxStyle;
    btnAddN: TAdvGlowButton;
    btnEditN: TAdvGlowButton;
    btnPostN: TAdvGlowButton;
    btnCancelN: TAdvGlowButton;
    btnDeleteN: TAdvGlowButton;
    cxAbierto: TcxButton;
    AdvNota: TAdvGlassButton;
    chkEstatusHora: TDBCheckBox;
    cxCerrado: TcxButton;
    popImprimeGerencial: TMenuItem;
    AgruparHorarios1: TMenuItem;
    DesagruparHorarios1: TMenuItem;
    MantenerHorario1: TMenuItem;
    N4: TMenuItem;
    Actividadesenproceso1: TMenuItem;
    AvancesActividades1: TMenuItem;
    QryBitacoraiNumeroGerencial: TIntegerField;
    nxNumGerencial: TcxSpinEdit;
    QryBitacorasIdGerencial: TStringField;
    sIdGerencial: TcxGridDBColumn;
    VisualizarGerencial1: TMenuItem;
    cxAlCorte: TcxStyle;
    QrFrentesConvenio: TStringField;
    QryBitacoradAvancePartida: TFloatField;
    gridMaterialesxPartida: TDBGrid;
    bitacorademateriales: TZQuery;
    bitacoradematerialesdIdFecha: TDateField;
    bitacoradematerialesiIdDiario: TIntegerField;
    bitacoradematerialessWbs: TStringField;
    bitacoradematerialessIdMaterial: TStringField;
    bitacoradematerialesdCantidad: TFloatField;
    bitacoradematerialessMedida: TStringField;
    bitacoradematerialessContrato: TStringField;
    bitacoradematerialessDescripcion: TStringField;
    bitacoradematerialesdSolicitado: TFloatField;
    bitacoradematerialessAnexo: TStringField;
    bitacoradematerialesdCostoMN: TFloatField;
    bitacoradematerialesdCostoDLL: TFloatField;
    bitacoradematerialesdCantidadComercial: TFloatField;
    bitacoradematerialessTrazabilidad: TStringField;
    bitacoradematerialessPertenece: TStringField;
    bitacoradematerialessTextoAux: TStringField;
    bitacoradematerialesidMat: TIntegerField;
    bitacoradematerialessColumnaAux: TStringField;
    bitacoradematerialessTrazabilidadAux: TStringField;
    ds_bitacorademateriales: TDataSource;
    BuscaObjeto: TZReadOnlyQuery;
    ds_buscaobjeto: TDataSource;
    bitacoradematerialessNumeroOrden: TStringField;
    tdEditaAvance: TCurrencyEdit;
    tdAvanceAcumulado: TCurrencyEdit;
    tdAvanceAnterior: TCurrencyEdit;
    cxEdita: TcxButton;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    panelPassword: TPanel;
    GroupBox1: TGroupBox;
    Label16: TLabel;
    tsPassword: TEdit;
    cxAceptar: TcxButton;
    cxCancelar: TcxButton;
    frmBarra1: TfrmBarra;
    Panel: tNewGroupBox;
    ListaObjeto: TRxDBGrid;
    QrFrentessIdFolio: TStringField;
    QryBitacoradPonderado: TFloatField;
    Label3: TLabel;
    tsNumeroActividad: TRxDBLookupCombo;
    chkImprime: TCheckBox;
    Label9: TLabel;
    Label10: TLabel;
    Label2: TLabel;
    tsTipoAct: TDBLookupComboBox;
    zqTiposAct: TZReadOnlyQuery;
    ds_tiposAct: TDataSource;
    Label8: TLabel;
    tsIdPernocta: TDBLookupComboBox;
    tsIdPlataforma: TDBLookupComboBox;
    Label6: TLabel;
    pernoctan: TZReadOnlyQuery;
    ds_Pernoctan: TDataSource;
    QrFrentessIdPernocta: TStringField;
    QrFrentessIdPlataforma: TStringField;
    qryNotasGerencial: TZQuery;
    qryNotasGerencialiIdDiario: TIntegerField;
    qryNotasGerencialmDescripcion: TMemoField;
    qryNotasGerencialsHoraInicio: TStringField;
    qryNotasGerencialsHoraFinal: TStringField;
    qryNotasGerencialsConceptoGerencial: TStringField;
    qryNotasGerencialsIdClasificacion: TStringField;
    qryNotasGerencialdCantidad: TFloatField;
    qryNotasGerencialdAvance: TStringField;
    qryNotasGerencialsWbs: TStringField;
    qryNotasGerencialNota: TStringField;
    qryNotasGerenciallImprime: TStringField;
    qryNotasGerencialiHermano: TIntegerField;
    qryNotasGerencialiIdTarea: TIntegerField;
    qryNotasGerencialiIdActividad: TIntegerField;
    qryNotasGerencialeTipoActividad: TStringField;
    qryNotasGerencialsTipoObra: TStringField;
    qryNotasGerencialsIdPernocta: TStringField;
    qryNotasGerencialsIdPlataforma: TStringField;
    qryNotasGerencialdCantidadAjuste: TFloatField;
    qryNotasGerencialdCantidadAjusteNext: TFloatField;
    qryNotasGerencialdCantidadAjusteNext2: TFloatField;
    bitacoradematerialessIdPernocta: TStringField;
    qryNotasGerencialdRestaEspacio: TFloatField;
    ActividadesDetalle1: TMenuItem;
    {$ENDREGION}
    {$REGION 'Procedimientos'}
    procedure FormShow(Sender: TObject);
    procedure tdIdFechaExit(Sender: TObject);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroActividadKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdTipoMovimientoKeyPress(Sender: TObject; var Key: Char);
    procedure tdAvanceKeyPress(Sender: TObject; var Key: Char);
    function lExisteActividadAnexo(sActividad: string): Boolean;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tdCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure grid_bitacoraEnter(Sender: TObject);
    procedure tdIdFechaEnter(Sender: TObject);
    procedure tsNumeroActividadEnter(Sender: TObject);
    procedure tsIdTipoMovimientoEnter(Sender: TObject);
    procedure tsIdTipoMovimientoExit(Sender: TObject);
    procedure tdCantidadEnter(Sender: TObject);
    procedure tdCantidadExit(Sender: TObject);
    procedure grid_igualesGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);

    procedure TiposdeMovimientoAfterScroll(DataSet: TDataSet);
    procedure tsNumeroActividadChange(Sender: TObject);
    procedure ActividadesIgualesAfterScroll(DataSet: TDataSet);
    procedure QryBitacoraCalcFields(DataSet: TDataSet);
    function fnValidaPartidaAnexo(sParamNumeroActividad: string): boolean;
    function fnValidaPartidaOrden(sParamWbs, sParamNumeroActividad: string): boolean;
    function fnValidaPartidaOrdenPorcentaje(sParamWbs, sParamNumeroActividad: string): boolean;


    function fnActualizaAcumuladosOrden(sParamOpcion, sParamWbs, sParamNumeroActividad: string;
      dParamCantidadInstalar, dParamInstalado, dParamExcedente, dParamCantidad: double): Boolean;
    function fnActualizaAcumuladosContrato(sParamOpcion, sParamNumeroActividad: string;
      dParamCantidadInstalar, dParamInstalado, dParamExcedente, dParamCantidad: double): Boolean;
    procedure QryBitacoraAfterScroll(DataSet: TDataSet);
    procedure rDiarioGetValue(const VarName: string; var Value: Variant);
    procedure btnMayusClick(Sender: TObject);
    procedure RevisarOrtografia2Click(Sender: TObject);
    procedure tmDescripcionDblClick(Sender: TObject);
    procedure CopiaMemo(Sender: TObject);
    procedure tsNumeroActividadMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_BitacoraGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure ActualizaIdDiario(dParamContrato: string; dParamFecha: tDate; dParamIdDiario, dParamIdDiarioOld: Integer);
    procedure Grid_BitacoraTitleClick(Column: TColumn);
    procedure grid_igualesTitleClick(Column: TColumn);
    procedure tdCantidadChange(Sender: TObject);
    procedure btnDeleteNClick(Sender: TObject);
    procedure btnUpClick(Sender: TObject);
    procedure OrdenarNotas(sParamOrden : string);
    procedure btnDownClick(Sender: TObject);
    procedure CosolidarNotasGerenciales1Click(Sender: TObject);
    procedure Grid_BitacoraCellClick(Column: TColumn);
    procedure btnDeleteDClick(Sender: TObject);
    function GetItemByName(Wnd : hWnd;  hItem : HTREEITEM;szItemName : LPCTSTR) : HTREEITEM ;
    procedure TvFrentesCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode;
      State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure TvFrentesChange(Sender: TObject; Node: TTreeNode);
    procedure TvFrentesCollapsing(Sender: TObject; Node: TTreeNode;
      var AllowCollapse: Boolean);
    procedure ds_bitacoradeequiposDataChange(Sender: TObject; Field: TField);
    procedure btn1Click(Sender: TObject);
    procedure GrdOrdenDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure GrdOrdenDblClick(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure LblReportadosClick(Sender: TObject);
    procedure FiltrarFolios(Sender : TObject);
    procedure LblTodosClick(Sender: TObject);
    procedure tdAvanceEnter(Sender: TObject);
    procedure tdAvanceExit(Sender: TObject);
    procedure tsHoraInicioKeyPress(Sender: TObject; var Key: Char);
    procedure tsHoraFinalKeyPress(Sender: TObject; var Key: Char);

    {$ENDREGION}
    procedure ActualizaImprime();
    procedure formatoEncabezado;
    procedure LeerDatasets(PsContrato,PsNumeroOrden:string);
    procedure Vigencias();
    function ValidaBarco(dParamPersonal: string): boolean;
    function SumaMaterial(sParamMaterial :string): double;
    function MaterialDisponible(sParamMaterial :string): double;
    procedure tmDescripcion_GerencialEnter(Sender: TObject);
    function CompararFechas(fecha1: TDate; fecha2: TDate) : Boolean;
    procedure QrFrentesAfterScroll(DataSet: TDataSet);
    procedure CargaActividades;
    function ValidaReporteDiario : boolean;
    function AvanceActual(sParamActividad, sParamWbs : string) : double;
    function MaximoItem : integer;
    procedure PartidaExistente(sParamActividad, sParamWbs, sParamTE : string);
    Function InstaladoOrden(sParamActividad, sParamWbs :string; sParamFecha : integer) : double;
    Function CantidadAnexoC  : double;
    Function InstaladoAnexoC(sParamActividad, sParamWbs :string) : double;
    procedure btnAddNClick(Sender: TObject);
    procedure btnEditNClick(Sender: TObject);
    procedure btnCancelNClick(Sender: TObject);
    procedure ActividadesIgualesCalcFields(DataSet: TDataSet);
    procedure btnPostNClick(Sender: TObject);
    procedure GridNotasCellClick(Column: TColumn);
   // Procedure GeneraReporteDiario_PDF(RTipo:FtTipo;RImpresion:FtActividades);
   // Procedure GeneraReporteAvances_PDF(RTipo:FtTipo; RImpresion:FtActividades);
    procedure Actividades1Click(Sender: TObject);
    function AvanceFolio : double;
    procedure Minimo_id;
    procedure tsHoraInicioEnter(Sender: TObject);
    procedure tsHoraInicioExit(Sender: TObject);
    procedure tsHoraFinalEnter(Sender: TObject);
    procedure tsHoraFinalExit(Sender: TObject);
    procedure tssIdClasificacionEnter(Sender: TObject);
    procedure tssIdClasificacionExit(Sender: TObject);
    procedure GrdOrdenCellClick(Column: TColumn);
    procedure tsNumeroActividad_PUEnter(Sender: TObject);
    procedure tsNumeroActividad_PUExit(Sender: TObject);
    procedure tsNumeroActividad_ADMEnter(Sender: TObject);
    procedure tsNumeroActividad_ADMExit(Sender: TObject);
    procedure tsNumeroActividadExit(Sender: TObject);
    procedure cxButton3Click(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
    procedure cxButton2Click(Sender: TObject);
    procedure tdAvanceChange(Sender: TObject);
    procedure tsHoraInicioChange(Sender: TObject);
    procedure tsHoraFinalChange(Sender: TObject);
    procedure BView_ActividadesDblClick(Sender: TObject);
    procedure cxComboBox1PropertiesChange(Sender: TObject);
    procedure sNumeroActividadStylesGetContentStyle(
      Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
      AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
    procedure mDescripcionStylesGetContentStyle(Sender: TcxCustomGridTableView;
      ARecord: TcxCustomGridRecord; AItem: TcxCustomGridTableItem;
      var AStyle: TcxStyle);
    procedure hGerencialStylesGetContentStyle(Sender: TcxCustomGridTableView;
      ARecord: TcxCustomGridRecord; AItem: TcxCustomGridTableItem;
      var AStyle: TcxStyle);
    procedure dAvanceStylesGetContentStyle(Sender: TcxCustomGridTableView;
      ARecord: TcxCustomGridRecord; AItem: TcxCustomGridTableItem;
      var AStyle: TcxStyle);
    procedure tsNumeroActividad_ADMKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroActividad_PUKeyPress(Sender: TObject; var Key: Char);
    procedure AdvNotaClick(Sender: TObject);
    procedure cxCerradoClick(Sender: TObject);
    procedure cxAbiertoClick(Sender: TObject);
    procedure popImprimeGerencialClick(Sender: TObject);
    procedure tssIdClasificacionKeyPress(Sender: TObject; var Key: Char);
    procedure AgruparHorarios1Click(Sender: TObject);
    procedure DesagruparHorarios1Click(Sender: TObject);
    procedure MantenerHorario1Click(Sender: TObject);
    procedure nxNumGerencialChange(Sender: TObject);
    procedure lblFolioDblClick(Sender: TObject);
    procedure bitacoradematerialessIdMaterialChange(Sender: TField);
    procedure bitacoradematerialesCalcFields(DataSet: TDataSet);
    procedure bitacoradematerialesAfterEdit(DataSet: TDataSet);
    procedure bitacoradematerialesAfterInsert(DataSet: TDataSet);
    procedure bitacoradematerialesBeforeDelete(DataSet: TDataSet);
    procedure bitacoradematerialesBeforePost(DataSet: TDataSet);
    procedure ListaObjetoExit(Sender: TObject);
    procedure ListaObjetoDblClick(Sender: TObject);
    procedure ListaObjetoKeyPress(Sender: TObject; var Key: Char);
    procedure bitacoradematerialesBeforeInsert(DataSet: TDataSet);
    procedure cxEditaClick(Sender: TObject);
    procedure tdAvanceGlobalChange(Sender: TObject);
    procedure cxAceptarClick(Sender: TObject);
    procedure cxCancelarClick(Sender: TObject);
    procedure tsPasswordEnter(Sender: TObject);
    procedure tsPasswordExit(Sender: TObject);
    procedure tdEditaAvanceEnter(Sender: TObject);
    procedure tdEditaAvanceExit(Sender: TObject);
    procedure tdEditaAvanceKeyPress(Sender: TObject; var Key: Char);
    procedure tsTipoActKeyPress(Sender: TObject; var Key: Char);
    procedure tsTipoActEnter(Sender: TObject);
    procedure tsTipoActExit(Sender: TObject);
    procedure GridNotasGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure popPostClick(Sender: TObject);
    procedure popAddClick(Sender: TObject);
    procedure popEditClick(Sender: TObject);
    procedure popCancelClick(Sender: TObject);
    procedure popDeleteClick(Sender: TObject);
    procedure qryNotasGerencialAfterScroll(DataSet: TDataSet);
    procedure frmBarra1btnPrinterClick(Sender: TObject);
    procedure Actividadesenproceso1Click(Sender: TObject);

    Procedure GeneraReporteDiario_PDF(RTipo:FtTipo;RImpresion:FtSeccion);Overload;
    procedure GridNotasKeyPress(Sender: TObject; var Key: Char);
    procedure bitacoradematerialesBeforeEdit(DataSet: TDataSet);
    procedure qryNotasGerencialBeforeDelete(DataSet: TDataSet);
    procedure ActividadesDetalle1Click(Sender: TObject);
    procedure qryNotasGerencialBeforeEdit(DataSet: TDataSet);

  private
    ExistelAplicaDiario:Boolean;
    sMenuP: string;
    ListaContratos:TStringList;
    lKardex: boolean;
    BotonPermisoV: TBotonesPermisos;
    sWbsKardex, opcKardex, fechaKardex: string;
    UtGrid: TicDbGrid;
    Utgrid2: TicDbGrid;
    sPernocta,stipoPersonal,Categoria: string;
    sPlataforma: string;
    sOpcion,d1,d2,d3,d4: string;
    total: Byte;
    zTipoPersonal : tzReadOnlyQuery;

    TiposDeActividad : TTiposActividad;

    { Private declarations }
    function ValidarConfJor:Boolean;
  public
    { Public declarations }
    Param_Frente:string;
  end;



const MAXTEXTLEN=50;

var
  frmBitacoraDepartamental_2: TfrmBitacoraDepartamental_2;
  sDescripcion: string;
  sWbsFormulario: string;
  sSegur: string;
  SavePlace: TBookmark;
  dExcedenteOrden: Double;
  dExcedenteAnexo: Double;
  dInstaladoOrden,
  dInstaladoOrdenAnt: Double;
  dInstaladoAnexo: Double;
  dCantidadAnexo: Double;
  dCantidadOrden,solicitadoE,solicitadop, dCantidadMaterial : Double;
  dError: Currency;
  txtMensaje: string;
  ListaPEQ: array[1..100] of integer;
  i: integer;
  dCantidadOld: Double;
  iIdDiarioOld: Integer;
  lRespuesta, lBorra,encontrado,Bandera: Boolean;
  sInformacionCliente: string;
  NuevoRegistro: Boolean;
  TipoFolio : string;
  IdAct, indice : integer;

  lAutentica : boolean;
  {Variables para Kardex del sistema..}


  myYear, myMonth, myDay: Word;

  {------------------------------------}
  lMostrarNotas,BandTE : boolean;

  //Exporta elementos a Excel..
  Excel, Libro, Hoja: Variant;
  nodoseleccionado: Integer;

  Indicar,Busqueda:Byte;
  dParamFecha,dFechaAnterior,dFechaActual:TDate;
  insertedit, modoEdit,isCharged : Boolean;
  valactividad: string;
  global_nota : string;
  sSQlCadena  : string;

  dAvanceCronologia,
  dAvanceEditar: Double;

  Zupd, zQryDatos :TZQuery;

  function GetTempDir: string;
  function NombreAleatorio(Longitud: Integer):String;
  procedure Split(Delimiter: Char; Str: string; ListOfStrings: TStrings);
implementation

uses frm_comentariosxanexo, UnitExcepciones, frm_OpcionesActividades,
  frm_ReprogramacionFolio, frm_DesgloceActividadesPEQ;

{$R *.dfm}

procedure Split(Delimiter: Char; Str: string; ListOfStrings: TStrings);
begin
   ListOfStrings.Clear;
   ListOfStrings.Delimiter     := Delimiter;
   ListOfStrings.DelimitedText := Str;
end;

function GetTempDir: string;
var
    Buffer : Array[0..Max_path] of char;
begin
    FillChar(Buffer,Max_Path + 1, 0);
    GetTempPath(Max_path, Buffer);
    Result := String(Buffer);
    if Result[Length(Result)] <> '\' then Result := Result + '\';
end;

function NombreAleatorio(Longitud: Integer):String;
const
  Chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
var
  S: string;
  i, N: integer;
begin
  Randomize;
  S := '';
  for i := 1 to Longitud do begin
    N := Random(Length(Chars)) + 1;
    S := S + Chars[N];
  end;
  Result := S;
end;

procedure TfrmBitacoraDepartamental_2.FiltrarFolios(Sender : TObject);
begin
  LblTodos.Font.Style :=[];
  LblReportados.Font.Style :=[];
  QryBitacora.close;
  QrFrentes.Active:=False;
  QrFrentes.SQL.Clear;
  QrFrentes.SQL.Add('SELECT b.sidtipomovimiento,ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.sIdPernocta, ot.sIdPlataforma, ot.cIdStatus, c.sIdConvenio as Convenio '+
                    'FROM ordenesdetrabajo AS ot  '+
                    'left join convenios c on (c.sContrato = ot.sContrato and c.sNumeroOrden = ot.sNumeroOrden) ');
  if  (TNxLinkLabel(Sender).Name = 'LblTodos') or (TNxLinkLabel(Sender).Name = 'LblCerrados')  then
  begin
    QrFrentes.SQL.Add('left join bitacoradeactividades b on (b.sContrato = ot.sContrato ');
  end;
  if (TNxLinkLabel(Sender).Name = 'LblReportados') or (TNxLinkLabel(Sender).Name = 'LblRProceso') then
  begin
    QrFrentes.SQL.Add('inner join bitacoradeactividades b on (b.sContrato = ot.sContrato ');
  end;
  QrFrentes.SQL.Add('and b.dIdFecha =:fecha and b.sNumeroOrden = ot.sNumeroOrden) '+
                    'WHERE ot.sContrato = :contrato and ot.cIdStatus ="P" ');
  if (TNxLinkLabel(Sender).Name = 'LblRProceso')  then
    QrFrentes.SQL.Add(' and ot.LEstado = "Proceso" ');
  QrFrentes.SQL.Add(' group by ot.sNumeroOrden, c.sIdConvenio ORDER BY ot.iOrden, b.sidtipomovimiento');
  QrFrentes.ParamByName('Contrato').AsString := param_global_contrato;
  QrFrentes.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
  QrFrentes.Open;
  try
    TipoFolio := QrFrentes.FieldByName('cIdStatus').AsString;
  except
    ;
  end;

  TNxLinkLabel(Sender).Font.Style := [fsBold,fsUnderline];
  if not QrFrentes.Locate('sidtipomovimiento','E',[]) then
  begin
    QrFrentes.First;
  end;
  if (QrFrentes.recordcount = 0) and (TNxLinkLabel(Sender).name = 'LblReportados') then
    FiltrarFolios(LblTodos);

  QryBitacora.Active := False;
  QryBitacora.Params.ParamByName('contrato').DataType := ftString;
  QryBitacora.Params.ParamByName('contrato').Value := param_global_contrato;
  QryBitacora.Params.ParamByName('convenio').DataType := ftString;
  QryBitacora.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
  QryBitacora.Params.ParamByName('orden').DataType := ftString;
  QryBitacora.Params.ParamByName('orden').Value := '%';
  QryBitacora.Params.ParamByName('fecha').DataType := ftDate;
  QryBitacora.Params.ParamByName('fecha').Value := tdIdFecha.Date;
  QryBitacora.Params.ParamByName('Alcance').DataType := ftString;
  QryBitacora.Params.ParamByName('Alcance').Value := Connection.configuracion.FieldValues['sTipoAlcance'];
  QryBitacora.Params.ParamByName('Ordenado').DataType := ftString;
  QryBitacora.Params.ParamByName('Ordenado').Value := 'iItemOrden';
  QryBitacora.Open;

end;

function TfrmBitacoraDepartamental_2.ValidarConfJor;
var
  QrConf:TZReadOnlyQuery;
  resp:Boolean;
begin
  
  resp:=True;
  QrConf:=TZReadOnlyQuery.Create(nil);
  try
    QrConf.Connection:=connection.zConnection;
    QrConf.SQL.Text:='select AvJornadas from configuracion where scontrato=:Contrato';
    QrConf.ParamByName('Contrato').AsString:=global_Contrato_Barco;
    QrConf.Open;

    if QrConf.RecordCount=1 then
    begin
      if QrConf.FieldByName('AvJornadas').AsString='No' then
        resp:=False;
    end;
  finally
    FreeAndNil(QrConf);
  end;
  Result:=resp;
end;

function TfrmBitacoraDepartamental_2.ValidaBarco(dParamPersonal: string): boolean;
var
  Q_ValidaBarco: TZReadOnlyQuery;
begin
  Q_ValidaBarco := TZReadOnlyQuery.Create(self);
  Q_ValidaBarco.Connection := connection.zConnection;

  Q_ValidaBarco.Active := False;
  Q_ValidaBarco.SQL.Clear;
  Q_ValidaBarco.SQL.Add('select lValidaBarco from configuracion where sContrato =:Contrato and lValidaBarco ="Si" ');
  Q_ValidaBarco.ParamByName('Contrato').AsString := param_global_contrato;
  Q_ValidaBarco.Open;

  if Q_ValidaBarco.RecordCount > 0 then
  begin
    Q_ValidaBarco.Active := False;
    Q_ValidaBarco.SQL.Clear;
    Q_ValidaBarco.SQL.Add('select sIdPersonal from personal where sContrato =:Contrato and sIdPersonal =:Personal and lProrrateo = "Si"');
    Q_ValidaBarco.ParamByName('Contrato').AsString := param_global_contrato;
    Q_ValidaBarco.ParamByName('Personal').AsString := dParamPersonal;
    Q_ValidaBarco.Open;

    if Q_ValidaBarco.RecordCount > 0 then
    begin
      Q_ValidaBarco.Active := False;
      Q_ValidaBarco.SQL.Clear;
      Q_ValidaBarco.SQL.Add('select lStatus from reportediario where sContrato =:Contrato and dIdFecha =:Fecha and lStatus <> "Pendiente" ');
      Q_ValidaBarco.ParamByName('Contrato').AsString := global_contrato_barco;
      Q_ValidaBarco.ParamByName('Fecha').AsDate := tdIdFecha.Date;
      Q_ValidaBarco.Open;

      if Q_ValidaBarco.RecordCount > 0 then
        result := True
      else
        result := False;
    end
    else
      result := False;
  end
  else
    result := False;

end;

procedure TfrmBitacoraDepartamental_2.Vigencias();
begin
  //Aqui leo las categorias de Personal Y Verifico que existan en el Oficio
  Connection.QryBusca2.Active;
  Connection.QryBusca2.SQL.Clear;
  Connection.QryBusca2.SQL.Add('Select dFechaVigencia from ordenesdetrabajogral  ' +
    'Where sContrato =:Contrato order by dFechaVigencia ');
  Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
  Connection.QryBusca2.Params.ParamByName('Contrato').Value := param_global_Contrato;
  Connection.QryBusca2.Open;
  if Bandera = True then
  begin
    dParamFecha := Global_Fecha;
    d3 := DateToStr(Global_Fecha);
  end
  else
  begin
    dParamFecha := tdIdFecha.DateTime;
    d3 := DateToStr(tdIdFecha.DateTime);
  end;

  if Connection.QryBusca2.RecordCount > 0 then
  begin
    while not Connection.QryBusca2.Eof do
    begin
      dFechaAnterior := Connection.QryBusca2.FieldValues['dFechaVigencia'];
      d1 := Connection.QryBusca2.FieldValues['dFechaVigencia'];
      Connection.QryBusca2.Next;
      dFechaActual := Connection.QryBusca2.FieldValues['dFechaVigencia'];
      d2 := Connection.QryBusca2.FieldValues['dFechaVigencia'];
      if ((dParamFecha > dFechaAnterior) and (dParamFecha > dFechaActual) or (dParamFecha > dFechaAnterior) and (dParamFecha < dFechaActual) or (dParamFecha >= dFechaAnterior) and (dParamFecha <= dFechaActual)) then
        d4 := d1;
    end;
  end;
end;

function TfrmBitacoraDepartamental_2.fnActualizaAcumuladosContrato(sParamOpcion, sParamNumeroActividad: string;
  dParamCantidadInstalar, dParamInstalado, dParamExcedente, dParamCantidad: double): Boolean;
begin
  try
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('UPDATE actividadesxanexo SET dInstalado = :Instalado, dExcedente = :Excedente ' +
      'where sContrato = :contrato And sNumeroActividad = :Actividad And sTipoActividad = "Actividad"');
    Connection.zCommand.Params.ParamByName('contrato').DataType  := ftString;
    Connection.zCommand.Params.ParamByName('contrato').value     := param_global_contrato;
    Connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
    Connection.zCommand.Params.ParamByName('Actividad').value    := sParamNumeroActividad;
    if sParamOpcion = 'Eliminar' then
      if dParamExcedente > 0 then
        if (dParamExcedente > dParamCantidad) then
        begin
          Connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
          Connection.zCommand.Params.ParamByName('Instalado').value := dParamCantidadInstalar;
          Connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
          Connection.zCommand.Params.ParamByName('Excedente').value := dParamExcedente - dParamCantidad
        end
        else
        begin
          Connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
          Connection.zCommand.Params.ParamByName('Instalado').value := dParamCantidadInstalar - (dParamCantidad - dParamExcedente);
          Connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
          Connection.zCommand.Params.ParamByName('Excedente').value := 0;
        end
      else
      begin
        Connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
        Connection.zCommand.Params.ParamByName('Instalado').value := dParamInstalado - dParamCantidad;
        Connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
        Connection.zCommand.Params.ParamByName('Excedente').value := 0;
      end
    else if (dExcedenteAnexo > 0) then
    begin
      Connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
      Connection.zCommand.Params.ParamByName('Instalado').value := dParamCantidadInstalar;
      Connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
      Connection.zCommand.Params.ParamByName('Excedente').value := dParamExcedente;
    end
    else
    begin
      Connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
      Connection.zCommand.Params.ParamByName('Instalado').value := dParamInstalado + dParamCantidad;
      Connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
      Connection.zCommand.Params.ParamByName('Excedente').value := 0;
    end;
    connection.zCommand.ExecSQL;
    fnActualizaAcumuladosContrato := True
  except
    fnActualizaAcumuladosContrato := False
  end
end;


function TfrmBitacoraDepartamental_2.GetItemByName(Wnd : hWnd;  hItem : HTREEITEM;szItemName : LPCTSTR) : HTREEITEM ;
var
  szBuffer: array [0..MAXTEXTLEN+1] of char;
  item : TTVItem;
  hItemFound, hItemChild : HTREEITEM;
begin
    if (hItem = nil) then
        hItem := HTREEITEM(SendMessage(Wnd, TVM_GETNEXTITEM, TVGN_ROOT, 0));
    while (hItem <> nil) do
    begin
        item.hItem := hItem;
        item.mask := TVIF_TEXT OR TVIF_CHILDREN;
        item.pszText := szBuffer;
        item.cchTextMax := MAXTEXTLEN;
        SendMessage(Wnd, TVM_GETITEM, 0, longint(@item));
        if (lstrcmp(szBuffer, szItemName) = 0) then
          begin
             Result := hItem;
             Exit;
          end;
        if (item.cChildren > 0) then
        begin
            hItemChild := HTREEITEM(SendMessage(Wnd, TVM_GETNEXTITEM,
                                                TVGN_CHILD, longint(hItem)));

            hItemFound := GetItemByName(Wnd, hItemChild, szItemName);
            if (hItemFound <> nil) then
             begin
                Result :=  hItemFound;
                Exit;
             end;
        end;
        hItem := HTREEITEM(SendMessage(Wnd, TVM_GETNEXTITEM,
                                       TVGN_NEXT, LPARAM(hItem)));
    end;
    Result := nil;
end;

function TfrmBitacoraDepartamental_2.fnActualizaAcumuladosOrden(sParamOpcion, sParamWbs, sParamNumeroActividad: string;
  dParamCantidadInstalar, dParamInstalado, dParamExcedente, dParamCantidad: double): Boolean;
begin
  try
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('UPDATE actividadesxorden SET dInstalado = :Instalado, dExcedente = :Excedente ' +
      'where sContrato = :contrato And sIdConvenio = :Convenio And sNumeroOrden = :Orden And sWbs = :wbs And sNumeroActividad = :Actividad And sTipoActividad = "Actividad"');
    connection.zCommand.Params.ParamByName('contrato').DataType  := ftString;
    connection.zCommand.Params.ParamByName('contrato').value     := param_global_contrato;
    connection.zCommand.Params.ParamByName('convenio').DataType  := ftString;
    connection.zCommand.Params.ParamByName('Convenio').Value     := QrFrentes.FieldByName('Convenio').AsString;
    connection.zCommand.Params.ParamByName('Orden').DataType     := ftString;
    connection.zCommand.Params.ParamByName('Orden').value        := QrFrentes.FieldValues['sNumeroOrden'];
    connection.zCommand.Params.ParamByName('wbs').DataType       := ftString;
    connection.zCommand.Params.ParamByName('wbs').value          := sParamWbs;
    Connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
    connection.zCommand.Params.ParamByName('Actividad').value    := sParamNumeroActividad;
    if sParamOpcion = 'Eliminar' then
    begin
        if dParamExcedente > 0 then
        begin
            if (dParamExcedente > dParamCantidad) then
            begin
              Connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
              Connection.zCommand.Params.ParamByName('Instalado').value    := dParamCantidadInstalar;
              Connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
              Connection.zCommand.Params.ParamByName('Excedente').value    := dParamExcedente - dParamCantidad
            end
            else
            begin
              Connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
              Connection.zCommand.Params.ParamByName('Instalado').value := dParamCantidadInstalar - (dParamCantidad - dParamExcedente);
              Connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
              Connection.zCommand.Params.ParamByName('Excedente').value := 0;
            end
        end
        else
        begin
            Connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
            Connection.zCommand.Params.ParamByName('Instalado').value := dParamInstalado - dParamCantidad;
            Connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
            Connection.zCommand.Params.ParamByName('Excedente').value := 0;
        end
    end
    else
    begin
        if (dParamExcedente > 0) then
        begin
            connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
            connection.zCommand.Params.ParamByName('Instalado').value    := dParamCantidadInstalar;
            connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
            connection.zCommand.Params.ParamByName('Excedente').value    := dParamExcedente;
        end
        else
        begin
            connection.zCommand.Params.ParamByName('Instalado').DataType := ftFloat;
            connection.zCommand.Params.ParamByName('Instalado').value    := dParamInstalado + dParamCantidad;
            connection.zCommand.Params.ParamByName('Excedente').DataType := ftFloat;
            connection.zCommand.Params.ParamByName('Excedente').value    := 0;
        end;
    end;
    connection.zCommand.ExecSQL;
    fnActualizaAcumuladosOrden := True
  except
    fnActualizaAcumuladosOrden := False
  end;

end;

function TfrmBitacoraDepartamental_2.fnValidaPartidaAnexo(sParamNumeroActividad: string): boolean;
begin
  dExcedenteAnexo := 0;
  dInstaladoAnexo := 0;
  dCantidadAnexo := 0;

  connection.qryBusca.Active := False;
  connection.qryBusca.SQL.Clear;
  connection.qryBusca.SQL.Add('Select (dInstalado + dExcedente) as dInstalado, dCantidadAnexo from actividadesxanexo where ' +
    'sContrato = :contrato  ' +
    'And sNumeroActividad = :Actividad And sTipoActividad = "Actividad" ');
  Connection.qryBusca.Params.ParamByName('contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('contrato').Value := param_global_contrato;
  Connection.qryBusca.Params.ParamByName('actividad').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('actividad').Value := sParamNumeroActividad;
  connection.qryBusca.Open;
  if (Connection.qryBusca.RecordCount > 0) then
  begin
    if Connection.qryBusca.FieldByName('dInstalado').IsNull then
      dInstaladoAnexo := 0
    else
      dInstaladoAnexo := Connection.qryBusca.FieldValues['dInstalado'];

    if Connection.qryBusca.FieldByName('dCantidadAnexo').IsNull then
      dCantidadAnexo := 0
    else
      dCantidadAnexo := Connection.qryBusca.FieldValues['dCantidadAnexo'];

    dError := (dInstaladoAnexo + tdCantidad.Value);
    dError := dError - dCantidadAnexo;
    if (dError > 0) then
    begin
      txtMensaje := 'No se puede asignar mas cantidad de la cantidad estipulada en el contrato vigente, ' +
        'Cantidad a instalar segun contrato = ' + floattostr(dCantidadAnexo) +
        ', Cantidad instalada a la fecha = ' + floattostr(dInstaladoAnexo) +
        ', si continua se creara un volumen de adicional a lo estipulado en el contrato vigente. Desea Continuar?';
      if MessageDlg(txtMensaje, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        dExcedenteAnexo := (dInstaladoAnexo + tdCantidad.Value) - dCantidadAnexo;
        fnValidaPartidaAnexo := True;
      end
      else
        fnValidaPartidaAnexo := False;
    end
    else
      fnValidaPartidaAnexo := True
  end
  else
    fnValidaPartidaAnexo := False;
end;

function TfrmBitacoraDepartamental_2.fnValidaPartidaOrden(sParamWbs, sParamNumeroActividad: string): boolean;
begin
  dExcedenteOrden := 0;
  dInstaladoOrden := 0;
  dCantidadOrden  := 0;

  connection.qryBusca.Active := False;
  connection.qryBusca.SQL.Clear;
  connection.qryBusca.SQL.Add('Select (dInstalado + dExcedente) as dInstalado , dCantidad, sTipoAnexo from actividadesxorden where ' +
    'sContrato = :contrato And sIdConvenio = :Convenio And sNumeroOrden = :orden And sWbs = :Wbs And ' +
    'sNumeroActividad = :Actividad And sTipoActividad = :Tipo');
  Connection.qryBusca.Params.ParamByName('contrato').DataType  := ftString;
  Connection.qryBusca.Params.ParamByName('contrato').Value     := param_global_contrato;
  Connection.qryBusca.Params.ParamByName('convenio').DataType  := ftString;
  Connection.qryBusca.Params.ParamByName('Convenio').Value     := QrFrentes.FieldByName('Convenio').AsString;
  Connection.qryBusca.Params.ParamByName('orden').DataType     := ftString;
  Connection.qryBusca.Params.ParamByName('orden').Value        := QrFrentes.FieldValues['sNumeroOrden'];
  Connection.qryBusca.Params.ParamByName('Wbs').DataType       := ftString;
  Connection.qryBusca.Params.ParamByName('Wbs').Value          := sParamWbs;
  Connection.qryBusca.Params.ParamByName('actividad').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('actividad').Value    := sParamNumeroActividad;
  Connection.qryBusca.Params.ParamByName('Tipo').DataType      := ftString;
  Connection.qryBusca.Params.ParamByName('Tipo').Value         := 'Actividad';
  connection.qryBusca.Open;

  if Connection.qryBusca.RecordCount > 0 then
  begin
    if connection.qryBusca.FieldByName('dInstalado').IsNull then
      dInstaladoOrden := 0
    else
      if connection.qryBusca.FieldValues['dInstalado'] < 0 then
        dInstaladoOrden := 0
      else
        dInstaladoOrden := connection.qryBusca.FieldValues['dInstalado'];

    if connection.qryBusca.FieldByName('dCantidad').IsNull then
      dCantidadOrden := 0
    else
      dCantidadOrden := connection.qryBusca.FieldValues['dCantidad'];

    if opcKardex = 'Edita' then
       dError := (dInstaladoOrden)
    else
       dError := (dInstaladoOrden + tdCantidad.Value);
    dError := dError - dCantidadOrden;
    if (dError > 0) then
    begin
      txtMensaje := 'No se puede reportar mas de lo propuesto en la Partida seleccionada. ' +
        'cantidad a reportar para la partida y Folio Seleccionado = ' + floattostr(dCantidadOrden) +
        ', Cantidad reportada a la fecha = ' + floattostr(dInstaladoOrden) +
        ', Sidesea continuar?';
      if MessageDlg(txtMensaje, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        dExcedenteOrden := (dInstaladoOrden + tdCantidad.Value) - dCantidadOrden;
        fnValidaPartidaOrden := True
      end
      else
      begin
        fnValidaPartidaOrden := False;
        lRespuesta := False;
      end;
    end
    else
      fnValidaPartidaOrden := True
  end
  else
    fnValidaPartidaOrden := False
end;

function TfrmBitacoraDepartamental_2.fnValidaPartidaOrdenPorcentaje(sParamWbs, sParamNumeroActividad: string): boolean;
begin
  dExcedenteOrden := 0;
  dInstaladoOrden := 0;
  dCantidadOrden := 0;

  connection.qryBusca.Active := False;
  connection.qryBusca.SQL.Clear;
  connection.qryBusca.SQL.Add('Select (dInstalado + dExcedente) as dInstalado , dCantidad, sTipoAnexo from actividadesxorden where ' +
    'sContrato = :contrato And sIdConvenio = :Convenio And sNumeroOrden = :orden And sWbs = :Wbs And ' +
    'sNumeroActividad = :Actividad And sTipoActividad = :Tipo');
  Connection.qryBusca.Params.ParamByName('contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('contrato').Value := param_global_contrato;
  Connection.qryBusca.Params.ParamByName('convenio').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
  Connection.qryBusca.Params.ParamByName('orden').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
  Connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Wbs').Value := sParamWbs;
  Connection.qryBusca.Params.ParamByName('actividad').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('actividad').Value := sParamNumeroActividad;
  Connection.qryBusca.Params.ParamByName('Tipo').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Tipo').Value := 'Actividad';
  connection.qryBusca.Open;

  if Connection.qryBusca.RecordCount > 0 then
  begin
    if connection.qryBusca.FieldByName('dInstalado').IsNull then
      dInstaladoOrden := 0
    else
      if connection.qryBusca.FieldValues['dInstalado'] < 0 then
        dInstaladoOrden := 0
      else
        dInstaladoOrden := connection.qryBusca.FieldValues['dInstalado'];

    if connection.qryBusca.FieldByName('dCantidad').IsNull then
      dCantidadOrden := 0
    else
      dCantidadOrden := connection.qryBusca.FieldValues['dCantidad'];

      Result := True
  end
  else
    Result := False
end;

procedure TfrmBitacoraDepartamental_2.FormShow(Sender: TObject);
var
  qryGrupos: TZReadOnlyQuery;
  qryPuntos: TZReadOnlyQuery;
begin

  sPernocta := '';
  d4 := '';
  sPlataforma := '';

  lBorra := False;
  sMenuP := stMenu;
  Self.BotonPermisoV := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'rDiario', PopupPrincipal);

  chkcancelada.Checked := false;
  OpcButton := '';
  sWbsFormulario := '';
  chkImprime.Checked := True;
  frmBarra1.btnCancel.Click;
  btnCancelN.Click;

  zQryDatos := tzQuery.Create(self);
  zQryDatos.Connection := connection.zConnection;

  QrFrentes.Active:=False;
  QrFrentes.SQL.Clear;
  QrFrentes.SQL.Add('SELECT  b.sidtipomovimiento,ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.sIdPernocta, ot.sIdPlataforma, ot.cIdStatus, c.sIdConvenio as Convenio '+
                    'FROM ordenesdetrabajo AS ot  '+
                    'left join convenios c on (c.sContrato = ot.sContrato and c.sNumeroOrden = ot.sNumeroOrden) '+
                    'left join bitacoradeactividades b on (b.sContrato = ot.sContrato '+
                    'and b.dIdFecha =:fecha and b.sNumeroOrden = ot.sNumeroOrden) '+
                    'WHERE ot.sContrato = :contrato and ot.cIdStatus = "P" '+
                    'group by ot.sNumeroOrden, c.sIdConvenio ORDER BY iOrden');
  QrFrentes.ParamByName('contrato').AsString := param_global_contrato;
  QrFrentes.ParamByName('fecha').AsDate      := tdIdFecha.Date;
  QrFrentes.Open;

  if QrFrentes.RecordCount > 0 then
     Param_Frente := QrFrentes.FieldValues['sNumeroOrden'];

  TipoFolio := QrFrentes.FieldByName('cIdStatus').AsString;

  zTipoPersonal := tzReadOnlyQuery.Create(self);
  zTipoPersonal.Connection := connection.zConnection;

  tsNumeroActividad.ReadOnly  := True;
  tsIdtipoMovimiento.ReadOnly := True;
  tdCantidad.ReadOnly         := True;
  tmDescripcion.ReadOnly      := True;

  tdIdFecha.Date := date;
  connection.configuracion.refresh;

  // Inicializo el Query Bitacora y actualizo los querys necesarios en este modulo
  TiposdeMovimiento.Active := False;
  TiposdeMovimiento.Params.ParamByName('Contrato').DataType := ftString;
  TiposdeMovimiento.Params.ParamByName('Contrato').Value := global_Contrato_Barco;
  TiposdeMovimiento.Params.ParamByName('Clasificacion').DataType := ftString;
  TiposdeMovimiento.Params.ParamByName('Clasificacion').Value := 'Tiempo Muerto';
  TiposdeMovimiento.Params.ParamByName('Clasificacion2').DataType := ftString;
  TiposdeMovimiento.Params.ParamByName('Clasificacion2').Value := 'Movimiento de Barco';
  TiposdeMovimiento.Params.ParamByName('Alcance').DataType := ftString;
  TiposdeMovimiento.Params.ParamByName('Alcance').Value := Connection.configuracion.FieldValues['sTipoAlcance'];
  TiposdeMovimiento.Open;

  OrdenesdeTrabajo.Active := False;
  OrdenesdeTrabajo.SQL.Clear;
  if global_grupo = 'INTEL-CODE' then
    Ordenesdetrabajo.SQL.Add('select ot.cIdStatus,ot.sNumeroOrden, ot.iJornada, ot.bTipoAdmon, ot.iDecimales, ot.sIdPernocta, ot.sIdPlataforma from ordenesdetrabajo ot where ot.sContrato =:Contrato ' +
      'And ot.cIdStatus =:Status order by ot.sNumeroOrden')
  else
    OrdenesdeTrabajo.SQL.Add('Select ot.cIdStatus,ot.sNumeroOrden, ot.iJornada, ot.bTipoAdmon, ot.iDecimales, ot.sIdPernocta, ot.sIdPlataforma from ordenesdetrabajo ot ' +
      'INNER JOIN ordenesxusuario ou On (ou.sContrato=ot.sContrato ' +
      'And ou.sNumeroOrden=ot.sNumeroOrden) ' +
      'where ot.sContrato =:Contrato And ou.sDerechos<>"BLOQUEADO" ' +
      'And ou.sIdUsuario =:Usuario And ot.cIdStatus =:Status order by ot.iOrden');
  OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString;
  OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := param_Global_Contrato;
  OrdenesdeTrabajo.Params.ParamByName('status').DataType := ftString;
  OrdenesdeTrabajo.Params.ParamByName('status').Value := connection.configuracion.FieldValues['cStatusProceso'];
  if global_grupo <> 'INTEL-CODE' then
  begin
    OrdenesdeTrabajo.Params.ParamByName('Usuario').DataType := ftString;
    OrdenesdeTrabajo.Params.ParamByName('Usuario').Value := Global_Usuario;
  end;
  OrdenesdeTrabajo.Open;

  tdIdFecha.Date := global_fecha;

  QryPartidasEfectivas.Active := False;
  QryPartidasEfectivas.Params.ParamByName('contrato').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('contrato').Value := param_global_contrato;
  QryPartidasEfectivas.Params.ParamByName('convenio').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
  QryPartidasEfectivas.Params.ParamByName('Orden').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('Orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
  QryPartidasEfectivas.Open;
  if QryPartidasEfectivas.RecordCount < 1 then
    tsNumeroActividad.Enabled := false;

  Plataformas.Active := False;
  Plataformas.Open;

  LblReportadosClick(LblReportados);

  QryPartidasEfectivas_ADM.Active := False;
  QryPartidasEfectivas_ADM.Params.ParamByName('contrato').DataType := ftString;
  QryPartidasEfectivas_ADM.Params.ParamByName('contrato').Value := param_global_contrato;
  QryPartidasEfectivas_ADM.Params.ParamByName('convenio').DataType := ftString;
  QryPartidasEfectivas_ADM.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
  QryPartidasEfectivas_ADM.Params.ParamByName('Orden').DataType := ftString;
  QryPartidasEfectivas_ADM.Params.ParamByName('Orden').Value    := QrFrentes.FieldValues['sNumeroOrden'];
  QryPartidasEfectivas_ADM.Params.ParamByName('fecha').DataType := ftDate;
  QryPartidasEfectivas_ADM.Params.ParamByName('fecha').Value    := global_fecha;
  QryPartidasEfectivas_ADM.Open;

  QryPartidasEfectivas_PU.Active := False;
  QryPartidasEfectivas_PU.Params.ParamByName('contrato').DataType := ftString;
  QryPartidasEfectivas_PU.Params.ParamByName('contrato').Value := param_global_contrato;
  QryPartidasEfectivas_PU.Params.ParamByName('convenio').DataType := ftString;
  QryPartidasEfectivas_PU.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
  QryPartidasEfectivas_PU.Params.ParamByName('Orden').DataType := ftString;
  QryPartidasEfectivas_PU.Params.ParamByName('Orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
  QryPartidasEfectivas_PU.Open;

  QryBitacora.Active := False;
  QryBitacora.Params.ParamByName('contrato').DataType := ftString;
  QryBitacora.Params.ParamByName('contrato').Value := param_global_contrato;
  QryBitacora.Params.ParamByName('convenio').DataType := ftString;
  QryBitacora.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
  QryBitacora.Params.ParamByName('orden').DataType := ftString;
  QryBitacora.Params.ParamByName('orden').Value := '%';
  QryBitacora.Params.ParamByName('fecha').DataType := ftDate;
  QryBitacora.Params.ParamByName('fecha').Value := tdIdFecha.Date;
  QryBitacora.Params.ParamByName('Alcance').DataType := ftString;
  QryBitacora.Params.ParamByName('Alcance').Value := Connection.configuracion.FieldValues['sTipoAlcance'];
  QryBitacora.Params.ParamByName('Ordenado').DataType := ftString;
  QryBitacora.Params.ParamByName('Ordenado').Value := 'iItemOrden';
  QryBitacora.Open;

  if QryBitacora.RecordCount > 0 then
  begin
      tdAvanceGlobal.Value := AvanceFolio;
      Grid_Bitacora.SetFocus;
  end;

  QrClasificacion.Active:=False;
  QrClasificacion.ParamByName('Contrato').AsString := global_Contrato_Barco;
  QrClasificacion.Open;

  if QrClasificacion.RecordCount > 0 then
     tssIdClasificacion.KeyValue := 'TE';

  zqTiposAct.Active:=False;
  zqTiposAct.Open;

  TiposDeActividad := TTiposActividad.Create( 'bitacoradeactividades', 'eTipoActividad' );
  TiposDeActividad.LinkWithLookup( cbbTipoActividad );
  TiposDeActividad.LinkWithTextEdit( dbTipoActividad );

  Plataformas.Active := False ;
  Plataformas.Open ;

  Pernoctan.Active := False ;
  Pernoctan.Open ;

  pernoctan.First;
  while not pernoctan.Eof do
  begin
      gridMaterialesxPartida.Columns[4].PickList.Add(pernoctan.FieldByName('sIdPernocta').AsString);
      pernoctan.Next;
  end;

  lBorra := True;
end;

procedure TfrmBitacoraDepartamental_2.tdIdFechaExit(Sender: TObject);
begin
  lBorra := False;

  ReporteDiario.Active := False;
  ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
  ReporteDiario.Params.ParamByName('contrato').Value := global_Contrato_Barco;
  ReporteDiario.Params.ParamByName('turno').DataType := ftString;
  ReporteDiario.Params.ParamByName('turno').Value := global_turno_reporte;
  ReporteDiario.Params.ParamByName('Fecha').DataType := ftDate;
  ReporteDiario.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
  ReporteDiario.Params.ParamByName('Orden').DataType := ftString;
  ReporteDiario.Params.ParamByName('Orden').Value := param_global_contrato;
  ReporteDiario.Open;

  // Limpia valores
  tdCantidad.Value := 0;
  tmDescripcion.Text := '';

  // Termina Limpia
  QryBitacora.Active := False;
  QryBitacora.Params.ParamByName('contrato').DataType := ftString;
  QryBitacora.Params.ParamByName('contrato').Value := param_global_contrato;
  QryBitacora.Params.ParamByName('convenio').DataType := ftString;
  QryBitacora.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
  QryBitacora.Params.ParamByName('orden').DataType := ftString;
  QryBitacora.Params.ParamByName('orden').Value := '%';
  QryBitacora.Params.ParamByName('fecha').DataType := ftDate;
  QryBitacora.Params.ParamByName('fecha').Value := tdIdFecha.Date;
  QryBitacora.Params.ParamByName('Alcance').DataType := ftString;
  QryBitacora.Params.ParamByName('Alcance').Value := Connection.configuracion.FieldValues['sTipoAlcance'];
  QryBitacora.Params.ParamByName('Ordenado').DataType := ftString;
  QryBitacora.Params.ParamByName('Ordenado').Value := 'iItemOrden';
  QryBitacora.Open;

   if QryBitacora.RecordCount > 0 then
      tdAvanceGlobal.Value := AvanceFolio;

  QrFrentes.Active:=False;
  QrFrentes.SQL.Clear;
  QrFrentes.SQL.Add('SELECT  b.sidtipomovimiento,ot.sContrato, ot.sNumeroOrden, ot.sIdFolio, ot.sIdPernocta, ot.sIdPlataforma, ot.cIdStatus, c.sIdConvenio as Convenio '+
                    'FROM ordenesdetrabajo AS ot  '+
                    'left join convenios c on (c.sContrato = ot.sContrato and c.sNumeroOrden = ot.sNumeroOrden) '+
                    'left join bitacoradeactividades b on (b.sContrato = ot.sContrato '+
                    'and b.dIdFecha =:fecha and b.sNumeroOrden = ot.sNumeroOrden) '+
                    'WHERE ot.sContrato = :contrato and ot.cIdStatus = "P" '+
                    'group by ot.sNumeroOrden, c.sIdConvenio ORDER BY iOrden');
  QrFrentes.ParamByName('Contrato').AsString := param_global_contrato;
  QrFrentes.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
  QrFrentes.Open;

  QryBitacoraAfterScroll(QryBitacora);
  tdIdFecha.Color := global_color_salida

end;

procedure TfrmBitacoraDepartamental_2.frmBarra1btnAddClick(Sender: TObject);
var
   sHora : string;
begin
    if ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
    begin
        MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
        exit;
    end;

  try
      frmBarra1.btnAddClick(Sender);
      btnCancelN.Click;
      CargaActividades;

      cxComboGerencial.ItemIndex := -1;
      tsHoraInicio.Text := '00:00';
      tsHoraFinal.Text  := '00:00';
      sHora := ObtenerHora + ':00';
      if sHora <= '05:00'  then
      begin
         cxComboGerencial.ItemIndex := 1;
         tsHoraInicio.Text := '17:00';
         tsHoraFinal.Text  := '05:00';
      end;

      cbbtipoActividad.ItemIndex := 0;
      cxCerrado.Enabled := False;
      cxAbierto.Enabled := False;
      chkEstatusHora.Enabled := False;

      tsNumeroActividad_ADM.KeyValue := Null;
      tsNumeroActividad_PU.KeyValue  := Null;
      tsNumeroActividad_PU.Color     := global_color_pantalla;
      tsNumeroActividad_ADM.Color    := global_color_pantalla;

      Grid_Bitacora.Enabled       := False;
      opcKardex                   := 'Crea ';
      chkImprime.Checked          := True;
      tsIdTipoMovimiento.KeyValue := connection.configuracion.FieldValues['sTipoOperacion'];
      Grid_Iguales.Enabled        := True;
      tsNumeroActividad.Enabled   := True;
      tsNumeroActividad.ReadOnly  := False;
      tsIdtipoMovimiento.ReadOnly := False;
      tsIdTipoMovimiento.Enabled  := true;
      tmDescripcion.ReadOnly      := False;
      tssIdClasificacion.Enabled  := False;
      nxNumGerencial.Enabled      := True;

      tsNumeroActividad.KeyValue     := '';
      tsNumeroActividad_PU.KeyValue  := '';
      tsNumeroActividad_ADM.KeyValue := '';
      tdCantidad.Enabled         := False;
      tdAvance.Enabled           := False;
      tsNumeroActividad.KeyValue := '';
      tmDescripcion.Text         := '';
      tdCantidad.Value  := 0;
      tdAvance.Value    := 0;
      tsNumeroActividad.SetFocus;

      if param_global_contrato = Global_contrato_Barco then
      begin
          tsIdTipoMovimiento.KeyValue := 'B';
          tsIdTipoMovimiento.SetFocus;
      end;

      if QryPartidasEfectivas.RecordCount > 0 then
      begin
          tsNumeroActividad.Enabled := True;
          tsNumeroActividad.ReadOnly := False;
          tsNumeroActividad_PU.ReadOnly := False;
          tsNumeroActividad_ADM.ReadOnly := False;
      end;
      OpcButton := 'New';

      btnAddN.Enabled    := False;
      btnEditN.Enabled   := False;
      btnPostN.Enabled   := False;
      btnCancelN.Enabled := False;
      btnDeleteN.Enabled := False;

      Self.BotonPermisoV.permisosBotones(frmBarra1);
  except
    on e:Exception do
      ShowMessage('No se puede continuar con un registro nuevo por el siguiente motivo: '+#10+e.Message);
  end;

end;

procedure TfrmBitacoraDepartamental_2.frmBarra1btnEditClick(Sender: TObject);
var
  sNumeroActividad : string;
begin
   try
      if (QryBitacora.RecordCount > 0) and (QryBitacora.FieldValues['sIdTurno'] = global_turno_reporte) then
      begin
          if ValidaReporteDiario = False then
          begin
              frmBarra1.btnCancel.Click;
              exit;
          end;
          //activapop(frmBitacoraDepartamental_2, popupprincipal);
      end;
      frmBarra1.btnEditClick(Sender);
      btnCancelN.Click;
      CargaActividades;

      Grid_Bitacora.Enabled := false;
      opcKardex             := 'Edita';
      OpcButton             := 'Edit';
      iIdDiarioOld          := QryBitacora.FieldValues['iIdDiario'];
      tsIdTipoMovimiento.Enabled := False;
      sNumeroActividad           := tsNumeroActividad.Text;
      tsNumeroActividad.ReadOnly := False;
      cxCerrado.Enabled := False;
      cxAbierto.Enabled := False;
      chkEstatusHora.Enabled := False;
      tsHoraInicio.ReadOnly  := False;
      tsHoraFinal.ReadOnly   := False;
      nxNumGerencial.Enabled := True;

      // Localizar el registro que se está editando
      if not ActividadesIguales.Locate('swbs', QryBitacora.FieldByName('swbs').AsString, []) then
         ActividadesIguales.First;
      dCantidadOld := tdCantidad.Value;

      //Grid_Iguales.Enabled        := False;
      tsIdtipoMovimiento.ReadOnly := False;
      tdCantidad.Enabled          := False;
      tdAvance.Enabled            := False;
      tmDescripcion.ReadOnly      := False;
      tssIdClasificacion.Enabled  := False;
      tmDescripcion.SetFocus;

      if (tiposdemovimiento.FieldValues['sIdTipoMovimiento'] = 'E') and not (tdCantidad.ReadOnly) then
         tdCantidad.SetFocus;

      if param_global_contrato = Global_contrato_Barco then
      begin
          tsIdTipoMovimiento.KeyValue := 'B';
          tmDescripcion.ReadOnly := False;
          tsIdTipoMovimiento.SetFocus;
      end;
      Self.BotonPermisoV.permisosBotones(frmBarra1);
  except
    on e:Exception do
      ShowMessage('No se puede editar el regsitro por el siguiente motivo: '+#10+e.Message);
  end;
end;

procedure TfrmBitacoraDepartamental_2.frmBarra1btnPostClick(Sender: TObject);
var
    lFiltro,
    lEfectivo,
    Consolidado,
    lIncorrecto : Boolean;

    sTiempoEfectivo,
    sNumeroActividad, sTipoAnexo, sParametro : string;
  iDiario, iRegistro: Integer;
  dAvance,
    dAvanceAnterior,
    MaxCantidad : Extended;

  Q_GuardaDatos,
  Q_BuscaAvance: TZReadOnlyQuery;
  QryUpdate: TZQuery;

  nombres, cadenas: TStringList;
  sWbsAux : string;
  iCheck: byte;
  iReporte      : Real ;
  Caracteres:Byte;
  bContinue:Boolean;
  AvanceTotal: Real;
  function xRound(Valor: Real; Dec: Integer): Real;
  var

    Desarrollo: Real;
    Decimales: string;
    Multiplo: Integer;
    sValor: string;

  begin
    { Redondear con limite mínimo superior de 5
         Delphi y mysql redondean los decimales de 0.1 a 0.5 hacia abajo, es decir al cero y de 0.6 a 0.9 hacia arriba, es decir a 1.
         Pues bien, EXCEL lo hace de la siguiente manera: de 0.1 a 0.4 hacia abajo y de 0.5 a 0.9 hacia arriba.
         Debido a que el personal de BMPI realiza sus cálculos en EXCEL es por esto que nunca llegan a los mismo avances
         de acuerdo al sistema. }
    Multiplo := 1;
    for I := 1 to Dec do
      Multiplo := Multiplo * 10;

    sValor := FloatToStr(Valor * Multiplo);
    if Pos('.', sValor) = 0 then sValor := sValor + '.00';

    Desarrollo := StrToFloat(Copy(sValor, 1, Pos('.', sValor) - 1));
    Decimales := Copy(sValor, Pos('.', sValor) + 1, Length(sValor));
    if StrToInt(Decimales[1]) > 4 then
      Desarrollo := Desarrollo + 1;
    Result := Desarrollo / Multiplo;
  end;

begin
  try
    //################jjivan 2014-08-12 --- Funcion Depurada##############
    Q_GuardaDatos := TZReadOnlyQuery.Create(self);
    Q_GuardaDatos.Connection := connection.zConnection;

    Q_BuscaAvance := TZReadOnlyQuery.Create(self);
    Q_BuscaAvance.Connection := connection.zConnection;

    dAvanceEditar      := 0;
    sNumeroActividad   := tsNumeroActividad.Text;
    sWbsFormulario     := '';
    sWbsAux            := '';
    lKardex            := False;
    lRespuesta         := True;
    bContinue          := True;
    swbsFormulario     := ActividadesIguales.FieldByName('sWbs').AsString;
    MaxCantidad        := 0;
    dInstaladoOrden    := 0;
    dInstaladoOrdenAnt := 0;
    dCantidadAnexo     := 0;
    dExcedenteOrden    := 0;
    dInstaladoAnexo    := 0;
    dExcedenteAnexo    := 0;

    if (tsNumeroActividad.KeyValue <> Null) then
       sTipoAnexo := QryPartidasEfectivas.FieldByName('sTipoAnexo').AsString
    else
       if (tsNumeroActividad_ADM.KeyValue <> Null) then
       begin
          sTipoAnexo       := 'PU';
          sNumeroActividad := tsNumeroActividad_ADM.Text;
          swbsFormulario   := QryPartidasEfectivas_PU.FieldByName('sWbs').AsString;
       end;

    if ActividadesIguales.Active then
       sWbsAux := 'x';

    if (tdCantidad.Value = 0) and (tdAvance.Value > 0)  then
        tsHoraInicio.SetFocus;

    {Validamos los horarios}
    if (tsHoraInicio.Text = tsHoraFinal.Text) and (tsHoraInicio.Text <> '00:00') and (tsHoraFinal.Text <> '00:00') then
    begin
        messageDLG('No se puede Guardar Horarios iguales', mtInformation, [mbOk], 0);
        exit;
    end;

    if bContinue then
    begin
         {$REGION 'Editar Registro'}
        if OpcButton = 'Edit' then
           if TiposdeMovimiento.FieldValues['sClasificacion'] <> 'Notas' then
           begin
              {Se hace una copia del registro}
              Q_GuardaDatos.Active := False;
              Q_GuardaDatos.SQL.Clear;
              Q_GuardaDatos.SQL.Add('select * from bitacoradeactividades where sContrato =:Contrato and sIdConvenio =:Convenio and dIdFecha =:Fecha and iIdDiario =:Diario and sWbs =:Wbs and sIdTipoMovimiento = "E"');
              Q_GuardaDatos.ParamByName('Contrato').AsString := param_global_contrato;
              Q_GuardaDatos.ParamByName('Convenio').AsString := QrFrentes.FieldByName('Convenio').AsString;
              Q_GuardaDatos.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
              Q_GuardaDatos.ParamByName('Diario').AsInteger  := QryBitacora.FieldValues['iIdDiario'];
              Q_GuardaDatos.ParamByName('Wbs').AsString      := sWbsFormulario;
              Q_GuardaDatos.Open;

              iDiario := QryBitacora.FieldValues['iIdDiario'];
              {Se elimina el registo con IdDiario Especificado..}
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('delete from bitacoradeactividades where sContrato = :contrato And sIdConvenio =:Convenio And dIdFecha = :fecha And iIdDiario = :diario and sWbs =:Wbs and sIdTipoMovimiento = "E" ');
              connection.zCommand.Params.ParamByName('contrato').AsString := param_Global_Contrato;
              connection.zCommand.Params.ParamByName('Convenio').AsString := QrFrentes.FieldByName('Convenio').AsString;
              connection.zCommand.Params.ParamByName('fecha').AsDate      := tdIdFecha.Date;
              connection.zCommand.Params.ParamByName('diario').AsInteger  := QryBitacora.FieldValues['iIdDiario'];
              connection.zCommand.Params.ParamByName('Wbs').AsString      := sWbsFormulario;
              connection.zCommand.ExecSQL;

              {Leer las cantiades reportadas a esta fecha.. Funcion}
              dInstaladoOrden := InstaladoOrden(sNumeroActividad, sWbsFormulario, 0);

              if opcKardex = 'Edita' then
                 dInstaladoOrden := dInstaladoOrden + tdCantidad.Value;

              //Verifica si lo reportado no excede lo permitido en el Frente/Folio..
              if dInstaladoOrden > ActividadesIguales.FieldValues['dCantidad'] then
              begin
                 dExcedenteOrden := dInstaladoOrden - ActividadesIguales.FieldValues['dCantidad'];
                 dInstaladoOrden := ActividadesIguales.FieldValues['dCantidad'];
              end;

              {Leer solamente las cantiades reportadas anteriormente a esta fecha.. Funcion}
              //dInstaladoOrdenAnt := InstaladoOrden(sNumeroActividad, sWbsFormulario, 1);

              {Obtener Cantidad Anexo ...Funcion..}
              dCantidadAnexo := CantidadAnexoC;

              {Cantidad Instalada en el Anexo C ... Funcion}
              dInstaladoAnexo := InstaladoAnexoC(sNumeroActividad, sWbsFormulario);

              if opcKardex = 'Edita' then
                 dInstaladoAnexo :=  dInstaladoAnexo + tdCantidad.Value;

              if dInstaladoAnexo > dCantidadAnexo then
              begin
                  dExcedenteAnexo := dInstaladoAnexo - dCantidadAnexo;
                  dInstaladoAnexo := dCantidadAnexo;
              end;

              {Actualiza la cantidad Instalada en el Folio/Frente.. Funcion}
              lEfectivo := fnActualizaAcumuladosOrden('', QryBitacora.FieldByName('sWbs').AsString, sNumeroActividad, ActividadesIguales.FieldValues['dCantidad'], dInstaladoOrden, dExcedenteOrden, 0);

              {Actualiza la cantidad Instalada en el Anexo C.. Funcion}
              lEfectivo := fnActualizaAcumuladosContrato('', sNumeroActividad, dCantidadAnexo, dInstaladoAnexo, dExcedenteAnexo, 0);

              OpcButton     := 'New';
              global_Editor := 'Edit';
              lRespuesta:=false;
           end;
        {$ENDREGION}
        lEfectivo := False;
        if OpcButton = 'New' then
        begin
            dAvance         := 0;
            lFiltro         := False;
            sTiempoEfectivo := tsIdTipoMovimiento.KeyValue;

            if TiposdeMovimiento.FieldValues['sClasificacion'] = 'Notas' then
               lFiltro := True
            else
               if ActividadesIguales.RecordCount > 0 then
                  if TiposdeMovimiento.FieldValues['sClasificacion'] = 'Tiempo en Operacion' then
                     lEfectivo := True;

            sWbsFormulario := ActividadesIguales.FieldByName('sWbs').AsString;
            SavePlace      := ActividadesIguales.GetBookmark;

            if lEfectivo then
            begin
                {Validamos que la cantidad maxima a reportar...}
                if Connection.configuracion.FieldValues['sAvanceBitacora'] = 'Volumen' then
                  lFiltro := fnValidaPartidaOrden(sWbsFormulario, sNumeroActividad)
                else
                  lFiltro := fnValidaPartidaOrdenPorcentaje(sWbsFormulario, sNumeroActividad);
            end;

            //Si la respuesta es No, Regresamos el registro eliminado de la bitacoradeactividades..
            if lRespuesta = False then
            begin
              if global_Editor =  'Edit' then
                if Q_GuardaDatos.RecordCount > 0 then
                begin
                    Connection.zCommand.Active := False;
                    Connection.zCommand.SQL.Clear;
                    Connection.zCommand.SQL.Add(funcsql(Q_GuardaDatos, 'bitacoradeactividades'));
                    Connection.zCommand.Active := False;
                    {Se inserta nuevamente el registor a bitacoradeactividades..}
                    for iRegistro := 1 to Q_GuardaDatos.fieldcount do
                    begin
                        sParametro := 'param' + trim(inttostr(iRegistro));
                        connection.zCommand.Params.parambyname(sParametro).datatype := Q_GuardaDatos.fields[iRegistro - 1].datatype;
                        if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'sHoraInicio') then
                            connection.zCommand.Params.parambyname(sparametro).value := tsHoraInicio.Text
                          else
                            if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'sHoraFinal') then
                              connection.zCommand.Params.parambyname(sparametro).value := tsHoraFinal.Text
                            else
                               if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'dCantidad') then
                                connection.zCommand.Params.parambyname(sparametro).value := tdCantidad.value
                               else
                                if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'dAvance') then
                                  connection.zCommand.Params.parambyname(sparametro).value := tdAvance.value
                                else
                                   if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'iNumeroGerencial') then
                                       connection.zCommand.Params.parambyname(sparametro).value := StrToInt(nxNumGerencial.Text)
                                   else
                                      if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'lImprime') then
                                      begin
                                          if chkImprime.Checked then
                                             connection.zCommand.Params.parambyname(sparametro).value := 'Si'
                                          else
                                             connection.zCommand.Params.parambyname(sparametro).value := 'No';
                                      end
                                      else
                                          connection.zCommand.Params.parambyname(sparametro).value    := Q_GuardaDatos.fields[iRegistro - 1].value;
                    end;
                    connection.zCommand.ExecSQL;
                    lFiltro:=false;
                end;
                lRespuesta := True;
            end;

            //Continua proceso normal del sistema..
            if lFiltro then
            begin
                Consolidado := False;
                {Busqueda de la partida... Funcion}
                PartidaExistente(sNumeroActividad, sWbsFormulario, sTiempoEfectivo);

                if QryExistePartida.RecordCount > 0 then
                   if MessageDlg('Se encontro una coincidencia del Wbs-Partida en los registros de la fecha y orden seleccionada, ¿Desea consolidar el movimiento?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                      Consolidado := True;

                   if Consolidado then
                   begin
                       {$REGION 'Consolida Partida'}
                       if lEfectivo then
                       begin
                           dAvance := 0;
                           if ActividadesIguales.FieldValues['dCantidad'] > 0 then
                              dAvance := AvanceActual(sNumeroActividad,sWbsFormulario );

                           {Se consolida el movimiento se suman los volumenes de la partida..}
                           connection.zCommand.Active := False;
                           connection.zCommand.SQL.Clear;
                           connection.zCommand.SQL.Add('UPDATE bitacoradeactividades SET dCantidad = :Cantidad, dAvance = :Avance ' +
                                              'where sContrato = :contrato and sIdConvenio =:Convenio And dIdFecha = :fecha And iIdDiario = :diario ');
                           connection.zCommand.Params.ParamByName('Contrato').AsString := param_Global_Contrato;
                           connection.zCommand.Params.ParamByName('Convenio').AsString := QrFrentes.FieldByName('Convenio').AsString;
                           connection.zCommand.Params.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
                           connection.zCommand.Params.ParamByName('Diario').AsInteger  := QryExistePartida.FieldValues['iIdDiario'];
                           connection.zCommand.Params.ParamByName('Avance').AsFloat    := dAvance;
                           connection.zCommand.Params.ParamByName('Cantidad').AsFloat  := tdCantidad.Value + QryExistePartida.FieldValues['dCantidad'];
                           connection.zCommand.ExecSQL;

                           try
                             if fnActualizaAcumuladosOrden('', sWbsFormulario, sNumeroActividad, dCantidadOrden, dInstaladoOrden, dExcedenteOrden, tdCantidad.Value) then
                                if fnActualizaAcumuladosContrato('', sNumeroActividad, dCantidadAnexo, dInstaladoAnexo, dExcedenteAnexo, tdCantidad.Value) then
                           except
                               MessageDlg('ERROR: Ocurrio un error al actualizar en concepto. ' + sNumeroActividad + ', notificar al administrador del sistema', mtWarning, [mbOk], 0)
                           end;
                       end;
                       {$ENDREGION}
                   end
                   else
                   begin
                       {$REGION 'Inserta Nuevo Registro'}
                       {Obtener IdDiarioMaximo... Funcion}
                       iDiario := MaximoItem;
                       IdAct   := iDiario;

                       if lEfectivo then
                       begin
                           dAvance := 0;
                           if ActividadesIguales.FieldValues['dCantidad'] > 0 then
                              dAvance := AvanceActual(sNumeroActividad,sWbsFormulario );
                       end;

                       {Guardamos los datos en la bitacora de actividades...}
                       connection.zCommand.Active := False;
                       connection.zCommand.SQL.Clear;
                       connection.zCommand.SQL.Add('INSERT INTO bitacoradeactividades (sContrato, sIdConvenio, dIdFecha, iIdDiario, sIdTurno, sNumeroOrden, sWbs, sNumeroActividad, ' +
                         ' sIdTipoMovimiento, dAvance, dCantidad, sHoraInicio, sHoraFinal, mDescripcion, lImprime, lCancelada, iIdTarea, sNumeroActividad_ADM, sWbs_ADM, iNumeroGerencial) ' +
                         ' VALUES (:contrato, :Convenio, :fecha, :diario, :turno, :orden, :Wbs, :actividad, ' +
                         ' :tipo,            :avance, :cantidad,  :inicio,     :final,     :descripcion, :imprime, :cancela, :diario, :ActividadADM, :WbsADM, :Numero) ');
                       Connection.zCommand.Params.ParamByName('contrato').AsString    := param_Global_Contrato;
                       connection.zCommand.Params.ParamByName('Convenio').AsString    := QrFrentes.FieldByName('Convenio').AsString;
                       Connection.zCommand.Params.ParamByName('fecha').AsDate         := tdIdFecha.Date;
                       Connection.zCommand.Params.ParamByName('diario').AsInteger     := iDiario;
                       Connection.zCommand.Params.ParamByName('turno').AsString       := global_turno_reporte;
                       Connection.zCommand.Params.ParamByName('orden').AsString       := QrFrentes.FieldValues['sNumeroOrden'];
                       if sWbsFormulario <> '' then
                          Connection.zCommand.Params.ParamByName('wbs').AsString      := sWbsFormulario;
                       Connection.zCommand.Params.ParamByName('actividad').AsString   := sNumeroActividad;
                       Connection.zCommand.Params.ParamByName('tipo').AsString        := sTiempoEfectivo;
                       Connection.zCommand.Params.ParamByName('avance').AsFloat       := tdAvance.value;
                       Connection.zCommand.Params.ParamByName('cantidad').AsFloat     := tdCantidad.Value;
                       Connection.zCommand.Params.ParamByName('inicio').AsString      := tsHoraInicio.Text;
                       Connection.zCommand.Params.ParamByName('final').AsString       := tsHoraFinal.Text;
                       Connection.zCommand.Params.ParamByName('descripcion').AsMemo   := ActividadesIguales.FieldByName('mDescripcion').AsString;
                       Connection.zCommand.Params.ParamByName('WbsADM').DataType := ftString;
                       Connection.zCommand.Params.ParamByName('ActividadADM').DataType := ftString;
                       if tsNumeroActividad_ADM.KeyValue = Null then
                       begin
                           Connection.zCommand.Params.ParamByName('WbsADM').value       := '';
                           Connection.zCommand.Params.ParamByName('ActividadADM').value := '';
                       end
                       else
                       begin
                           Connection.zCommand.Params.ParamByName('WbsADM').value       := QryPartidasEfectivas_ADM.FieldValues['sWbs'];
                           Connection.zCommand.Params.ParamByName('ActividadADM').value := QryPartidasEfectivas_ADM.FieldValues['sNumeroActividad'];
                       end;

                       if chkImprime.Checked then
                          Connection.zCommand.Params.ParamByName('Imprime').AsString  := 'Si'
                       else
                          Connection.zCommand.Params.ParamByName('Imprime').AsString  := 'No';
                       if chkCancelada.Checked then
                          Connection.zCommand.Params.ParamByName('Cancela').AsString  := 'Si'
                       else
                          Connection.zCommand.Params.ParamByName('Cancela').AsString  := 'No';
                       Connection.zCommand.Params.ParamByName('Numero').AsInteger     := StrToInt(nxNumGerencial.Text);
                       connection.zCommand.ExecSQL;

                       //Registra movimiento en kardex del sistema.
                       lKardex    := true;
                       sWbsKardex := sWbsFormulario;

                       if lEfectivo then
                          if fnActualizaAcumuladosOrden('', sWbsFormulario, sNumeroActividad, dCantidadOrden, dInstaladoOrden, dExcedenteOrden, tdCantidad.Value) then
                            if fnActualizaAcumuladosContrato('', sNumeroActividad, dCantidadAnexo, dInstaladoAnexo, dExcedenteAnexo, tdCantidad.Value) then

                            else
                                MessageDlg('ERROR: Ocurrio un error al Insertar la Partida. ' + sNumeroActividad + ', notificar al administrador del sistema', mtWarning, [mbOk], 0)
                       else
                          MessageDlg('ERROR: Ocurrio un error al actualizar en concepto. ' + sWbsFormulario + ' de la orden de trabajo seleccionada, notificar al administrador del sistema', mtWarning, [mbOk], 0);
                       {$ENDREGION}
                   end;
            end;//Fin lFiltro

            {Actualizacion de personal y equipo con el idDiario Anterior}
            if (Global_Editor = 'Edit')  then
               ActualizaIdDiario(param_global_contrato, tdIdFecha.Date, iDiario, iIdDiarioOld);
        end
        else
        begin
            {$REGION 'Actuliza Notas'}
            try
                QryUpdate := TZQuery.Create(nil);
                QryUpdate.Connection := Connection.zConnection;
                QryUpdate.Active := False;
                QryUpdate.SQL.Clear;
                QryUpdate.SQL.Add('UPDATE bitacoradeactividades SET mDescripcion = :descripcion, sHoraInicio = :HoraInicio, sHoraFinal = :HoraFinal, lImprime =:Imprime, iNumeroGerencial =:NumGerencial ' +
                                  'where sContrato = :contrato and sIdConvenio =:Convenio And dIdFecha = :fecha And iIdDiario = :diario  ');
                QryUpdate.Params.ParamByName('contrato').AsString    := param_Global_Contrato;
                QryUpdate.Params.ParamByName('Convenio').AsString    := QrFrentes.FieldByName('Convenio').AsString;
                QryUpdate.Params.ParamByName('fecha').AsDate         := tdIdFecha.Date;
                QryUpdate.Params.ParamByName('diario').AsInteger     := QryBitacora.FieldValues['iIdDiario'];
                QryUpdate.Params.ParamByName('descripcion').AsMemo   := tmDescripcion.Text;
                if chkImprime.Checked then
                   QryUpdate.Params.ParamByName('Imprime').AsString  := 'Si'
                else
                   QryUpdate.Params.ParamByName('Imprime').AsString  := 'No';
                QryUpdate.Params.ParamByName('HoraInicio').AsString  := tsHoraInicio.Text;
                QryUpdate.Params.ParamByName('HoraFinal').AsString   := tsHoraFinal.Text;
                QryUpdate.Params.ParamByName('NumGerencial').AsInteger := StrToInt(nxNumGerencial.Text);
                QryUpdate.ExecSQL;

                //Asignamos movieintos a kardex del sistema..}
                lKardex := true;
                sWbsKardex := QryBitacora.FieldByName('sWbs').AsString;
                DecodeDate(tdIdFecha.Date, myYear, myMonth, myDay);
                fechaKardex := inttostr(myDay) + '/' + inttostr(myMonth) + '/' + inttostr(myYear);
            except
              on e: exception do
              begin
                UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Registros de Volumenes de Obra y Notas', 'Al actualizar el registro', 0);
              end;
            end
            {$ENDREGION}
        end;

        //Registrar la operacion en el kardex
        if lKardex then
        begin
            if sWbsKardex = '' then
               opcKardex := opcKardex + ' Comentario'
            else
               opcKardex := opcKardex + ' Partida ' + sWbsKardex;
            try
               Kardex('Reporte Diario', opcKardex, fechaKardex, 'Fecha', Param_Frente, '', '', 'Tarifa Diaria','Volumenes de Obra Y Nota');
            except
              on e: exception do
              begin
                UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Registros de Volumenes de Obra y Notas', 'Al registrar actualizacion de registro en kardex', 0);
              end;
            end;
        end;

        QryPartidasEfectivas_ADM.Active := False;
        QryPartidasEfectivas_ADM.Params.ParamByName('contrato').DataType := ftString;
        QryPartidasEfectivas_ADM.Params.ParamByName('contrato').Value := param_global_contrato;
        QryPartidasEfectivas_ADM.Params.ParamByName('convenio').DataType := ftString;
        QryPartidasEfectivas_ADM.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
        QryPartidasEfectivas_ADM.Params.ParamByName('Orden').DataType := ftString;
        QryPartidasEfectivas_ADM.Params.ParamByName('Orden').Value    := QrFrentes.FieldValues['sNumeroOrden'];
        QryPartidasEfectivas_ADM.Params.ParamByName('fecha').DataType := ftDate;
        QryPartidasEfectivas_ADM.Params.ParamByName('fecha').Value    := global_fecha;
        QryPartidasEfectivas_ADM.Open;
    end;


    if opcKardex = 'Edita' then
    begin
        QryBitacora.Refresh;
        BView_Actividades.DataController.ClearSelection;
        BView_Actividades.DataController.DataSource.DataSet.Locate('iIdDiario', IdAct,[]);
        indice := BView_Actividades.DataController.GetRootDataController.GetSelectionAnchorRowIndex;
        BView_Actividades.DataController.SelectRows(indice, indice);
    end;

    SavePlace := BView_Actividades.DataController.DataSource.DataSet.GetBookmark;
    if opcKardex <> 'Edita' then
    begin
        QryBitacora.Active := False;
        QryBitacora.Open;
        BView_Actividades.DataController.ClearSelection;
        BView_Actividades.DataController.DataSource.DataSet.Locate('iIdDiario', IdAct,[]);
        indice := BView_Actividades.DataController.GetRootDataController.GetSelectionAnchorRowIndex;
        BView_Actividades.DataController.SelectRows(indice, indice);
    end;
    try
      BView_Actividades.DataController.DataSource.DataSet.GotoBookmark(SavePlace);
    except
    else
      BView_Actividades.DataController.DataSet.FreeBookmark(SavePlace);
    end;

    if QryBitacora.RecordCount > 0 then
       tdAvanceGlobal.Value := AvanceFolio;

    Q_BuscaAvance.Destroy;
    Q_GuardaDatos.Destroy;

    ActividadesIguales.Active := False;
    Grid_Bitacora.Enabled     := True;

    if Global_Editor = 'Edit' then
      frmBarra1.btnCancelClick(Sender)
    else
      frmBarra1.btnPostClick(Sender);

    btnAddN.Enabled    := True;
    btnEditN.Enabled   := True;
    btnPostN.Enabled   := False;
    btnCancelN.Enabled := False;
    btnDeleteN.Enabled := True;

    Self.BotonPermisoV.permisosBotones(frmBarra1);

    frmbarra1.btnEdit.Enabled := False;
    tsIdTipoMovimiento.Enabled := true;
  except
    on e:Evalidaciones do
    begin
        MessageDlg(''+#13+e.message, mtinformation, [mbOK], 0);
    end;
  end;
end;

procedure TfrmBitacoraDepartamental_2.frmBarra1btnPrinterClick(Sender: TObject);
begin
    GeneraReporteDiario_PDF(FtAbordo,FtsAll);
end;

procedure TfrmBitacoraDepartamental_2.frmBarra1btnCancelClick(Sender: TObject);
begin
  tsNumeroActividad.ReadOnly  := True;
  tsIdtipoMovimiento.ReadOnly := True;
  tdCantidad.ReadOnly         := True;
  tdAvance.ReadOnly           := True;
  tmDescripcion.ReadOnly      := True;
  tsHoraInicio.ReadOnly       := True;
  tsHoraFinal.ReadOnly        := True;
  tdCantidad.Enabled          := True;
  tdAvance.Enabled            := True;

  global_Editor               := '';
  tsNumeroActividad.KeyValue  := '';
  tmDescripcion.Text          := '';
  tdCantidad.Value            := 0;

  tsNumeroActividad.KeyValue     := '';
  tsNumeroActividad_PU.KeyValue  := '';
  tsNumeroActividad_ADM.KeyValue := '';

  cxCerrado.Enabled := False;
  cxAbierto.Enabled := False;
  chkEstatusHora.Enabled := False;

  Insertar1.Enabled  := True;
  Editar1.Enabled    := True;
  Registrar1.Enabled := False;
  Can1.Enabled       := False;
  Eliminar1.Enabled  := True;
  Refresh1.Enabled   := True;
  ActividadesIguales.Active := False;
  nxNumGerencial.Enabled    := False;

  btnAddN.Enabled    := True;
  btnEditN.Enabled   := True;
  btnPostN.Enabled   := False;
  btnCancelN.Enabled := False;
  btnDeleteN.Enabled := True;

  try
     tsNumeroActividad.KeyValue := QryPartidasEfectivas.FieldByName('sNumeroActividad').AsString;
  Except
  end;

  tsIdTipoMovimiento.Enabled := true;
  Grid_Bitacora.Enabled:=true;
  frmBarra1.btnCancelClick(Sender);
  Self.BotonPermisoV.permisosBotones(frmBarra1);
  frmbarra1.btnEdit.Enabled := False;
end;

procedure TfrmBitacoraDepartamental_2.frmBarra1btnRefreshClick(Sender: TObject);
begin
  QryBitacora.Active := False;
  QryBitacora.Open;

  connection.configuracion.refresh;

  ordenesdetrabajo.Active := False;
  ordenesdetrabajo.Open;

  TiposdeMovimiento.Active := False;
  TiposdeMovimiento.Open;
end;

procedure TfrmBitacoraDepartamental_2.GridNotasCellClick(Column: TColumn);
begin
    frmbarra1.btnCancel.Click;
    btnCancelN.Click;
    lMostrarNotas := True;
    cxCerrado.Enabled := True;
    cxAbierto.Enabled := True;
    chkEstatusHora.Enabled := True;
    if qryNotasGerencial.RecordCount > 0 then
    begin
         if (sOpcion = '') and (lMostrarNotas) then
         begin
             tsHoraInicio.Text  := QryNOtasGerencial.FieldValues['sHoraInicio'];
             tsHoraFinal.Text   := QryNOtasGerencial.FieldValues['sHoraFinal'];
             tmDescripcion.Text := QryNOtasGerencial.FieldValues['mDescripcion'];
             tdCantidad.Value   := QryNOtasGerencial.FieldValues['dCantidad'];
             tdAvance.Value     := QryNOtasGerencial.FieldValues['dAvance'];
         end;       
     end;
end;

procedure TfrmBitacoraDepartamental_2.GridNotasGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
  try
    if (Sender as TrxDBGrid).DataSource.DataSet.State = dsBrowse then
      if QryNotasGerencial.RecordCount > 0 then
      begin
          Background   := $00FEC6BA;
        if (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sTipoObra').AsString = 'PU' then
        begin
            Background  := $00CBFFB9;
        end;

        if (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sTipoObra').AsString = 'FO' then
        begin
            Background  := $0091C8FF;
        end;

      end
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Conceptos/Partidas x Frente de Trabajo', 'Al cambiar de registro de actividades', 0);
    end;
  end;
end;

procedure TfrmBitacoraDepartamental_2.GridNotasKeyPress(Sender: TObject;
  var Key: Char);
begin
 if key=#13 then
    if qrynotasGerencial.State in [dsEdit] then
        qryNotasGerencial.Post;
end;

procedure TfrmBitacoraDepartamental_2.Grid_BitacoraCellClick(Column: TColumn);
begin
    lMostrarNotas := False;
    //Colores Descripcion

    if (QryBitacora.FieldValues['sIdTipoMovimiento'] = 'N') then
        tmDescripcion.Color := $00ADE86C ;
        QryBitacora.Refresh;
end;

procedure TfrmBitacoraDepartamental_2.frmBarra1btnDeleteClick(Sender: TObject);
var
  lBorra: Boolean;
begin
  try
    if QryBitacora.RecordCount > 0 then
       if ValidaReporteDiario = False then
       begin
           frmBarra1.btnCancel.Click;
           exit;
       end;

    lBorra := True;
    //Primero validamos que no existan notas gerenciales asignadas a dicha partida o comentario,,
    if QryNotasGerencial.RecordCount > 0 then
    begin
        lBorra := False;
        messageDLG('Existen Cortes de Actividades asignadas a la Partida, Favor de Eliminarlos!', mtInformation, [mbOk], 0);
    end;

    if lBorra then
      if (QryBitacora.RecordCount > 0) and (QryBitacora.FieldValues['sIdTurno'] = global_turno_reporte) then
        if (QryBitacora.FieldValues['sIdTipoMovimiento'] = 'E') then
        begin
          if MessageDlg('Desea eliminar la actividad y todo el personal y equipo asignado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
              if QryBitacora.FieldValues['dCantidad'] > 0 then
              begin
                  connection.qryBusca.Active := False;
                  connection.qryBusca.SQL.Clear;
                  connection.qryBusca.SQL.Add('select dCantidad, dInstalado, dExcedente from actividadesxorden where sContrato = :Contrato and sIdConvenio = :Convenio and ' +
                            'sNumeroOrden = :Orden and sWbs = :Wbs And sNumeroActividad = :Actividad');
                  Connection.qryBusca.Params.ParamByName('Contrato').AsString   := param_global_contrato;
                  Connection.qryBusca.Params.ParamByName('Convenio').AsString   := QrFrentes.FieldByName('Convenio').AsString;
                  Connection.qryBusca.Params.ParamByName('Orden').AsString      := QryBitacora.FieldValues['sNumeroOrden'];
                  Connection.qryBusca.Params.ParamByName('Wbs').AsString        := QryBitacora.FieldValues['sWbs'];
                  Connection.qryBusca.Params.ParamByName('Actividad').AsString  := QryBitacora.FieldValues['sNumeroActividad'];
                  connection.qryBusca.Open;
                  if connection.qryBusca.RecordCount > 0 then
                  begin
                     if not fnActualizaAcumuladosOrden('Eliminar', QryBitacora.FieldValues['sWbs'], QryBitacora.FieldValues['sNumeroActividad'],
                        Connection.qryBusca.FieldValues['dCantidad'], Connection.qryBusca.FieldValues['dInstalado'],
                        Connection.qryBusca.FieldValues['dExcedente'], QryBitacora.FieldValues['dCantidad']) then
                      MessageDlg('ERROR: Ocurrio un error al actualizar en concepto. ' + QryBitacora.FieldValues['sWbs'] + ' de la orden de trabajo seleccionada, notificar al administrador del sistema', mtWarning, [mbOk], 0);
                  end
                  else
                    MessageDlg('ERROR: Ocurrio un error al actualizar en concepto. ' + QryBitacora.FieldValues['sWbs'] + ' de la orden de trabajo seleccionada, notificar al administrador del sistema', mtWarning, [mbOk], 0);

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zCommand.SQL.Add('Delete from bitacoradepersonal where sContrato = :contrato and dIdFecha = :fecha and iIdDiario = :diario');
                  connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('Contrato').Value    := param_Global_Contrato;
                  connection.zCommand.Params.ParamByName('Fecha').DataType    := ftDate;
                  connection.zCommand.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
                  connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
                  connection.zCommand.Params.ParamByName('diario').Value      := QryBitacora.FieldValues['iIdDiario'];
                  connection.zCommand.ExecSQL();

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zCommand.SQL.Add('Delete from bitacoradeequipos where sContrato = :contrato and dIdFecha = :fecha and iIdDiario = :diario');
                  connection.zCommand.Params.ParamByName('Contrato').Value    := param_Global_Contrato;
                  connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('Fecha').DataType    := ftDate;
                  connection.zCommand.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
                  connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
                  connection.zCommand.Params.ParamByName('diario').Value      := QryBitacora.FieldValues['iIdDiario'];
                  connection.zCommand.ExecSQL;

                  connection.zCommand.Active := False;
                  connection.zCommand.SQL.Clear;
                  connection.zCommand.SQL.Add('Delete from bitacorademateriales where sContrato = :contrato and dIdFecha = :fecha and iIdDiario = :diario');
                  connection.zCommand.Params.ParamByName('Contrato').Value    := param_Global_Contrato;
                  connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
                  connection.zCommand.Params.ParamByName('Fecha').DataType    := ftDate;
                  connection.zCommand.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
                  connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
                  connection.zCommand.Params.ParamByName('diario').Value      := QryBitacora.FieldValues['iIdDiario'];
                  connection.zCommand.ExecSQL;
              end;
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('Delete from bitacoradeactividades where sContrato = :contrato and dIdFecha = :fecha and iIdDiario = :diario');
              connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
              connection.zCommand.Params.ParamByName('Contrato').Value    := param_Global_Contrato;
              connection.zCommand.Params.ParamByName('Fecha').DataType    := ftDate;
              connection.zCommand.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
              connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
              connection.zCommand.Params.ParamByName('diario').Value      := QryBitacora.FieldValues['iIdDiario'];
              connection.zCommand.ExecSQL();

              {Registramos movimiento en Kardex del sistema..}
              sWbsKardex := QryBitacora.FieldByName('sWbs').AsString;
              DecodeDate(tdIdFecha.Date, myYear, myMonth, myDay);
              fechaKardex := inttostr(myDay) + '/' + inttostr(myMonth) + '/' + inttostr(myYear);
              if sWbsKardex = '' then
                opcKardex := 'Borra Comentario'
              else
                opcKardex := 'Borra Partida ' + sWbsKardex;
              Kardex('Reporte Diario', opcKardex, fechaKardex, 'Fecha', Param_Frente, '', '','Tarifa Diaria','Volumenes de Obra Y Nota');

              SavePlace := BView_Actividades.DataController.DataSource.DataSet.GetBookmark;
              QryBitacora.Active := False;
              QryBitacora.Open;
              try
                BView_Actividades.DataController.DataSource.DataSet.GotoBookmark(SavePlace);
              except

                BView_Actividades.DataController.DataSet.FreeBookmark(SavePlace);
              end;

              Grid_Bitacora.SetFocus
          end
        end
        else
          MessageDlg('La partida no puede eliminarse, elimine los alcances registrados a la partida en el dia para poder realizar la eliminación.', mtInformation, [mbOk], 0)
      else
        MessageDlg('No existe registro a eliminar o talves el registro pertenece a otro turno, verifique su información.', mtInformation, [mbOk], 0);
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Registros de Volumenes de Obra y Notas', 'Al eliminar registro', 0);
    end;
  end;
end;

procedure TfrmBitacoraDepartamental_2.frmBarra1btnExitClick(Sender: TObject);
begin
  global_Editor := '';
  frmBarra1.btnExitClick(Sender);
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  close
end;

procedure TfrmBitacoraDepartamental_2.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    Grid_Bitacora.SetFocus
end;

procedure TfrmBitacoraDepartamental_2.tsPasswordEnter(Sender: TObject);
begin
    tsPassword.Color := global_color_entrada;
end;

procedure TfrmBitacoraDepartamental_2.tsPasswordExit(Sender: TObject);
begin
    tsPassword.Color := global_color_salida;
end;

procedure TfrmBitacoraDepartamental_2.tssIdClasificacionEnter(Sender: TObject);
begin
    tssIdClasificacion.Color := global_color_entrada;
end;

procedure TfrmBitacoraDepartamental_2.tssIdClasificacionExit(Sender: TObject);
begin
    tssIdClasificacion.Color := global_color_salida;
end;

procedure TfrmBitacoraDepartamental_2.tssIdClasificacionKeyPress(
  Sender: TObject; var Key: Char);
begin
    if key =#13 then
       tsHoraInicio.SetFocus;
end;

procedure TfrmBitacoraDepartamental_2.tsTipoActEnter(Sender: TObject);
begin
    tsTipoAct.Color := global_color_entrada;
end;

procedure TfrmBitacoraDepartamental_2.tsTipoActExit(Sender: TObject);
begin
    tsTipoAct.Color := global_color_salida;
    tsIdPlataforma.Visible := False;
    tsIdPernocta.Visible   := False;
    Label6.Visible         := False;
    Label8.Visible         := False;
    if tsTipoAct.KeyValue = 'FO' then
    begin
       tsIdPlataforma.Visible := True;
       tsIdPernocta.Visible   := True;
       Label6.Visible         := True;
       Label8.Visible         := True;
    end;
end;

procedure TfrmBitacoraDepartamental_2.tsTipoActKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       tmDescripcion.SetFocus;
end;

procedure TfrmBitacoraDepartamental_2.TvFrentesChange(Sender: TObject;
  Node: TTreeNode);
begin

    nodoseleccionado:=Integer(node.ItemId);
    if not isCharged then
      frmBarra1.btnCancel.Click;

    param_global_contrato:=TContrato(ListaContratos.Objects[ListaContratos.IndexOf(IntToStr(Integer(Node.ItemId)))]).sContrato;
    Param_Frente:=TContrato(ListaContratos.Objects[ListaContratos.IndexOf(IntToStr(Integer(Node.ItemId)))]).sNumeroOrden;

    OrdenesdeTrabajo.Active:=False;
    OrdenesdeTrabajo.Params.ParamByName('Contrato').AsString := param_Global_Contrato;
    OrdenesdeTrabajo.Params.ParamByName('status').DataType := ftString;
    OrdenesdeTrabajo.Params.ParamByName('status').Value := connection.configuracion.FieldValues['cStatusProceso'];
    if global_grupo <> 'INTEL-CODE' then
    begin
      OrdenesdeTrabajo.Params.ParamByName('Usuario').DataType := ftString;
      OrdenesdeTrabajo.Params.ParamByName('Usuario').Value := Global_Usuario;
    end;
    OrdenesdeTrabajo.Open;


    LeerDatasets(param_global_contrato,Param_Frente);
    tdIdFechaExit(Sender);


end;

procedure TfrmBitacoraDepartamental_2.TvFrentesCollapsing(Sender: TObject;
  Node: TTreeNode; var AllowCollapse: Boolean);
begin
  AllowCollapse:=False;
end;

procedure TfrmBitacoraDepartamental_2.TvFrentesCustomDrawItem(
  Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState;
  var DefaultDraw: Boolean);
begin
  Sender.Canvas.Font.Color := clblack;
  Sender.Canvas.Font.Style := [];
  //Sender.Canvas.Font.Size:=tam;
  if (cdsFocused in State) then
  begin
    Sender.Canvas.Font.Color := clwhite;
    sender.Canvas.Font.Style:=[fsbold];
    //sender.Canvas.Brush.Color:=clwindow;
  end
  else
    if Integer(node.ItemId)=nodoseleccionado then
    begin
      Sender.Canvas.Font.Color := clwhite;
      sender.Canvas.Font.Style:=[fsbold];
    end;
end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
  begin
      if QryBitacora.FieldByName('sTipoAnexo').AsString  = 'ADM' then
         if tdAvance.Visible = True then
            tdAvance.SetFocus;
  end;
end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividadMouseMove(
  Sender: TObject; Shift: TShiftState; X, Y: Integer);
begin
  if QryPartidasEfectivas.RecordCount > 0 then
    tsNumeroActividad.Hint := ' Paquete  [' + QryPartidasEfectivas.FieldValues['sWbsAnterior'] + ']';
end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividad_ADMEnter(
  Sender: TObject);
begin
    tsNumeroActividad_ADM.Color := global_color_entrada;
    tsNumeroActividad.Color     := global_color_pantalla;
end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividad_ADMExit(
  Sender: TObject);
begin
    tsNumeroActividad_ADM.Color := global_color_salida;
    if (tsNumeroActividad_PU.KeyValue <> Null) and (tsNumeroActividad_ADM.KeyValue = Null) then
       tsNumeroActividad_ADM.SetFocus;

   if (frmBarra1.btnCancel.Enabled = True) and (not tsNumeroActividad_PU.ReadOnly) then
    if tsNumeroActividad_PU.Text <> '' then
    begin

      {Se buscan todas las actiivdades que tengan el mismo nombre...}
      ActividadesIguales.Active := False;
      ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
      ActividadesIguales.Params.ParamByName('contrato').Value := param_global_contrato;
      ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
      ActividadesIguales.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
      ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
      ActividadesIguales.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
      ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
      ActividadesIguales.Params.ParamByName('actividad').Value := tsNumeroActividad_ADM.Text;
      ActividadesIguales.ParamByName('turno').AsString := global_turno_reporte;
      ActividadesIguales.ParamByName('fecha').AsDate := tdIdFecha.Date;
      ActividadesIguales.Open;

    end
    else
    begin
      ActividadesIguales.Active := False;
      ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
      ActividadesIguales.Params.ParamByName('contrato').Value := param_global_contrato;
      ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
      ActividadesIguales.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
      ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
      ActividadesIguales.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
      ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
      ActividadesIguales.Params.ParamByName('actividad').Value := '';
      ActividadesIguales.ParamByName('turno').AsString := global_turno_reporte;
      ActividadesIguales.ParamByName('fecha').AsDate := tdIdFecha.Date;
      ActividadesIguales.Open;

    end;
end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividad_ADMKeyPress(
  Sender: TObject; var Key: Char);
begin
    if tsNumeroActividad_ADM.Enabled then
    begin
        if key=#13 then
          tdCantidad.SetFocus
    end;
end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividad_PUEnter(
  Sender: TObject);
begin
    tsNumeroActividad_PU.Color  := global_color_entrada;
    tsNumeroActividad_ADM.Color := global_color_salida;
    tsNumeroActividad.Color     := global_color_pantalla;
    tsNumeroActividad.KeyValue  := Null;
end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividad_PUExit(Sender: TObject);
begin
    tdCantidad.Value   := 0;
    tmDescripcion.Text := '';
    tsNumeroActividad_PU.Color := global_color_salida;

    if (frmBarra1.btnCancel.Enabled = True) and (not tsNumeroActividad_PU.ReadOnly) then
    if tsNumeroActividad_PU.Text <> '' then
    begin

      {Se buscan todas las actiivdades que tengan el mismo nombre...}
      ActividadesIguales.Active := False;
      ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
      ActividadesIguales.Params.ParamByName('contrato').Value := param_global_contrato;
      ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
      ActividadesIguales.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
      ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
      ActividadesIguales.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
      ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
      ActividadesIguales.Params.ParamByName('actividad').Value := tsNumeroActividad_PU.Text;
      ActividadesIguales.ParamByName('turno').AsString := global_turno_reporte;
      ActividadesIguales.ParamByName('fecha').AsDate := tdIdFecha.Date;
      ActividadesIguales.Open;
    end
    else
    begin
      ActividadesIguales.Active := False;
      ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
      ActividadesIguales.Params.ParamByName('contrato').Value := param_global_contrato;
      ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
      ActividadesIguales.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
      ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
      ActividadesIguales.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
      ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
      ActividadesIguales.Params.ParamByName('actividad').Value := '';
      ActividadesIguales.ParamByName('turno').AsString := global_turno_reporte;
      ActividadesIguales.ParamByName('fecha').AsDate := tdIdFecha.Date;
      ActividadesIguales.Open;
    end;

end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividad_PUKeyPress(
  Sender: TObject; var Key: Char);
begin
    if key=#13 then
       tsNumeroActividad_ADM.SetFocus;
end;

procedure TfrmBitacoraDepartamental_2.tsIdTipoMovimientoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tdCantidad.SetFocus
end;

procedure TfrmBitacoraDepartamental_2.tdAvanceChange(Sender: TObject);
begin
     //activapop(frmBitacoraDepartamental_2, popupprincipal);
end;

procedure TfrmBitacoraDepartamental_2.tdAvanceEnter(Sender: TObject);
begin
    tdAvance.Color := global_color_entrada;
end;

procedure TfrmBitacoraDepartamental_2.tdAvanceExit(Sender: TObject);
begin
    tdAvance.color := global_color_salida;
    try
      if actividadesiguales.RecordCount > 0 then
      begin
          if (tdAvance.Value >= 0) and (tdAvance.Value <= 100) then
          begin
             if actividadesiguales.FieldValues['dCantidad'] > 0 then
                tdCantidad.Value := (tdAvance.Value * actividadesiguales.FieldValues['dCantidad']) / 100;
          end
          else
             tdCantidad.Value := actividadesiguales.FieldValues['dCantidad'];
      end;
    except

    end;
end;

procedure TfrmBitacoraDepartamental_2.tdAvanceGlobalChange(Sender: TObject);
begin
    if tdEditaAvance.Value > 0 then
       tdAvanceglobal.Font.Color := clRed
    else
       tdAvanceglobal.Font.Color := clBlack;
end;

procedure TfrmBitacoraDepartamental_2.tdAvanceKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsTipoAct.SetFocus
end;

function TfrmBitacoraDepartamental_2.lExisteActividadAnexo(sActividad: string): Boolean;
begin
  if TiposdeMovimiento.FieldValues['sClasificacion'] = 'Notas' then
  begin
    sDescripcion := '';
    lExisteActividadAnexo := True
  end
  else
  begin
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('select mDescripcion from actividadesxanexo a where a.sContrato = :Contrato ' +
      'And a.sNumeroActividad = :Actividad');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := param_global_contrato;
    Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Actividad').Value := sActividad;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
    begin
      sDescripcion := Connection.qryBusca.FieldValues['mDescripcion'];
      lExisteActividadAnexo := True
    end
    else
    begin
      sDescripcion := '';
      lExisteActividadAnexo := False
    end
  end
end;

procedure TfrmBitacoraDepartamental_2.ListaObjetoDblClick(Sender: TObject);
begin
    GridMaterialesxPartida.SetFocus;
end;

procedure TfrmBitacoraDepartamental_2.ListaObjetoExit(Sender: TObject);
begin
  if Panel.Visible = True then
  begin
    bitacorademateriales.FieldValues['sIdMaterial'] := BuscaObjeto.FieldValues['sNumeroActividad'];
    bitacorademateriales.FieldValues['sAnexo']      := BuscaObjeto.FieldValues['sAnexo'];
    GridMaterialesxPartida.SetFocus ;
    Panel.Visible := False;
  end
end;

procedure TfrmBitacoraDepartamental_2.ListaObjetoKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key = #13 then
       GridMaterialesxPartida.SetFocus;
end;

procedure TfrmBitacoraDepartamental_2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  zqryDatos.Destroy;
  Action:=caFree;
end;

procedure TfrmBitacoraDepartamental_2.Insertar1Click(Sender: TObject);
begin
  frmBarra1.btnAdd.Click
end;

procedure TfrmBitacoraDepartamental_2.Editar1Click(Sender: TObject);
begin
  frmBarra1.btnEdit.Click
end;

procedure TfrmBitacoraDepartamental_2.popAddClick(Sender: TObject);
begin
    btnAddN.OnClick(sender);
end;

procedure TfrmBitacoraDepartamental_2.popCancelClick(Sender: TObject);
begin
    btnCancelN.OnClick(sender);
end;

procedure TfrmBitacoraDepartamental_2.popDeleteClick(Sender: TObject);
begin
    btnDeleteN.OnClick(sender);
end;

procedure TfrmBitacoraDepartamental_2.popEditClick(Sender: TObject);
begin
    btnEditN.OnClick(sender);
end;

procedure TfrmBitacoraDepartamental_2.Registrar1Click(Sender: TObject);
begin
  frmBarra1.btnPost.Click
end;

procedure TfrmBitacoraDepartamental_2.RevisarOrtografia2Click(Sender: TObject);
var
  WindowName: string;
  WindowHandle: Cardinal;
  WordApp, Document, Selection: OleVariant;
  exito: boolean;
  actualizar: boolean;
  registro: tbookmark;
begin
  registro := qrybitacora.GetBookmark;
  actualizar := false;
  if (OpcButton <> 'New') and (OpcButton <> 'Edit') then
  begin
    if MessageDlgpos('Para Corregir el comentario se necesita editar el registro' + #13 + #13 + '¿Desea Actualizar el comentario?',
      mtConfirmation, [mbYes, mbNo], 0, self.Left + round(self.Width / 4) + 10, self.Top + round(self.Height / 2)) = mrYes then
      actualizar := true
    else exit;
  end;


  if (length(trim(self.tmDescripcion.Text)) > 0) then
  begin
    exito := true;
    try
      WordApp := CreateOleObject('Word.Application');
    except
      exito := false;
    end;
    if exito then
    begin
      Document := WordApp.Documents.Add;
      Selection := WordApp.Selection;
      Selection.TypeText(tmDescripcion.Text);
      WindowName := WordApp.ActiveDocument.FullName + ' - ' + WordApp.Application.Caption;
      WindowHandle := 0;
      WindowHandle := FindWindow(nil, pChar(WindowName));
      SetWindowRgn(WindowHandle, CreateRectRgn(0, 0, 0, 0), true);
      if wordapp.Options.IgnoreUppercase = true then
        wordapp.Options.IgnoreUppercase := false;
      WordApp.ActiveDocument.CheckGrammar;

      Selection.WholeStory;
      Selection.Copy;
      if actualizar then
        frmBarra1btnEditClick(sender);
      tmDescripcion.Text := Clipboard.AsText;
      wordapp.quit(false);
      if actualizar then
        frmBarra1btnPostClick(sender);
    end else
      MessageDlg('Para Verificar la ortografia necesita tener instalado Microsoft Word xp o versiones posteriores de office word.', mtWarning, [mbOk], 0);

  end;
  if actualizar then
  begin
    try
      qrybitacora.GotoBookmark(registro);
    except
      qrybitacora.FreeBookmark(registro);
    end;
    self.Grid_Bitacora.SetFocus;
  end;
  exito := False;
  actualizar := False;


  if (length(trim(self.tmDescripcion.Text)) > 0) then
  begin
    exito := true;
    try
      WordApp := CreateOleObject('Word.Application');
    except
      exito := false;
    end;
    if exito then
    begin
      Document := WordApp.Documents.Add;
      Selection := WordApp.Selection;
      Selection.TypeText(tmDescripcion.Text);
      WindowName := WordApp.ActiveDocument.FullName + ' - ' + WordApp.Application.Caption;
      WindowHandle := 0;
      WindowHandle := FindWindow(nil, pChar(WindowName));
      SetWindowRgn(WindowHandle, CreateRectRgn(0, 0, 0, 0), true);
      if wordapp.Options.IgnoreUppercase = true then
        wordapp.Options.IgnoreUppercase := false;
      WordApp.ActiveDocument.CheckGrammar;

      Selection.WholeStory;
      Selection.Copy;
      if actualizar then
        frmBarra1btnEditClick(sender);
      tmDescripcion.Text := Clipboard.AsText;
      wordapp.quit(false);
      if actualizar then
        frmBarra1btnPostClick(sender);
    end else
      MessageDlg('Para Verificar la ortografia necesita tener instalado Microsoft Word xp o versiones posteriores de office word.', mtWarning, [mbOk], 0);

  end;
  if actualizar then
  begin
    try
      qrybitacora.GotoBookmark(registro);
    except
      qrybitacora.FreeBookmark(registro);
    end;
    self.Grid_Bitacora.SetFocus;
  end;




end;

procedure TfrmBitacoraDepartamental_2.bitacoradematerialesAfterEdit(
  DataSet: TDataSet);
begin
  if lBorra = True then
  begin
    if bitacorademateriales.RecordCount = 0 then
      bitacorademateriales.Cancel
    else
      if (QryBitacora.FieldValues['sIdTurno'] <> global_turno_reporte) then
        bitacorademateriales.Cancel
  end
  else
  begin
    bitacorademateriales.Cancel;
    MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
  end;
end;

procedure TfrmBitacoraDepartamental_2.bitacoradematerialesAfterInsert(
  DataSet: TDataSet);
begin
   if lBorra = True then
    if QryBitacora.RecordCount > 0 then
      if (QryBitacora.FieldValues['sIdTurno'] = global_turno_reporte) then
      begin
        bitacorademateriales.FieldValues['dIdFecha'] := tdIdFecha.Date;
        bitacorademateriales.FieldValues['sContrato'] := param_Global_Contrato;
        bitacorademateriales.FieldValues['iIdDiario'] := QryNOtasGerencial.FieldValues['iIdDiario'];
        bitacorademateriales.FieldValues['sWbs'] := QryBitacora.FieldValues['sWbs'];
        bitacorademateriales.FieldValues['dCantidad'] := 0;
        bitacorademateriales.FieldValues['sTrazabilidad'] := 'S/T';
        bitacorademateriales.FieldValues['sPertenece'] := 'PEP';
        bitacorademateriales.FieldValues['sNumeroOrden'] := QrFrentes.FieldValues['sNumeroOrden'];
        bitacorademateriales.FieldValues['sIdPernocta']  := QrFrentes.FieldValues['sIdPernocta'];

      end
      else
        bitacorademateriales.Cancel
    else
      bitacorademateriales.Cancel
  else
  begin
    MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
    bitacorademateriales.Cancel
  end;
end;

procedure TfrmBitacoraDepartamental_2.bitacoradematerialesBeforeDelete(
  DataSet: TDataSet);
begin
  if lBorra = False then
  begin
    MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
    Abort;
  end
end;

procedure TfrmBitacoraDepartamental_2.bitacoradematerialesBeforeEdit(
  DataSet: TDataSet);
begin
     dCantidadMaterial := bitacoradematerialesdCantidad.Value;
end;

procedure TfrmBitacoraDepartamental_2.bitacoradematerialesBeforeInsert(
  DataSet: TDataSet);
begin
    If ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
    begin
        MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
        abort;
    end;

    if qryNotasGerencial.RecordCount = 0 then
    begin
        messageDLG('Debe registrar "Horarios de actividades"! ', mtInformation, [mbOk], 0);
        bitacorademateriales.Cancel;
    end;
    
end;

procedure TfrmBitacoraDepartamental_2.bitacoradematerialesBeforePost(
  DataSet: TDataSet);
var
   dCantidadSuma, dCantidadSalidas : double;
begin
 try
    bitacorademateriales.FieldByName('swbs').AsString := QryBitacora.FieldByName('swbs').AsString;

    if (bitacorademateriales.FieldValues['sIdMaterial'] = null) then
      bitacorademateriales.Cancel
    else
      if BitacoradeMateriales.fieldbyName('dcantidad').asfloat = 0 then
      begin
          messageDLG('No se Aceptan Cantidades en 0', mtInformation, [mbOk], 0);
          BitacoradeMateriales.Cancel;
      end
      else
      begin
          dCantidadSalidas := MaterialDisponible(bitacorademateriales.FieldByName('sIdMaterial').AsString);
          dCantidadSuma    := SumaMaterial(bitacorademateriales.FieldByName('sIdMaterial').AsString);

         if dCantidadMaterial > 0 then
            dCantidadSuma := dCantidadSuma - dCantidadMaterial;

          if xRound((dCantidadSuma + bitacorademateriales.FieldByName('dCantidad').AsFloat),2) > xRound(dCantidadSalidas,2) then
          begin
              messageDLG('No existen salidas de Almacen suficientes para reportar dicho material!' +#13 + 'Cantidad disponinle = '+ FloatTostr(dCantidadSalidas - dCantidadSuma), mtInformation, [mbOk], 0);
              abort;
          end;
          dCantidadMaterial := 0;
      end;

  except
    abort;
    MessageDlg('Ocurrio un error al Actualizar el registro.', mtInformation, [mbOk], 0);
  end;
end;

procedure TfrmBitacoraDepartamental_2.bitacoradematerialesCalcFields(
  DataSet: TDataSet);
begin
  if not bitacorademateriales.FieldByName('sIdMaterial').IsNull then
  begin
    connection.qryBusca2.Active := False;
    connection.qryBusca2.SQL.Clear;
    connection.qryBusca2.SQL.Add('select sMedida from actividadesxanexo ' +
      'where sContrato = :contrato And sNumeroActividad = :actividad and sTipoActividad = "Actividad" ');
    connection.qryBusca2.Params.ParamByName('contrato').DataType := ftString;
    connection.qryBusca2.Params.ParamByName('contrato').Value := global_Contrato_Barco;
    connection.qryBusca2.Params.ParamByName('actividad').DataType := ftString;
    connection.qryBusca2.Params.ParamByName('actividad').Value := bitacorademateriales.FieldValues['sIdMaterial'];
    connection.qryBusca2.Open;
    if connection.qryBusca2.RecordCount > 0 then
      bitacoradematerialessMedida.Text := connection.qryBusca2.FieldValues['sMedida']
    else
      bitacoradematerialessMedida.Text := '';

  end
end;

procedure TfrmBitacoraDepartamental_2.bitacoradematerialessIdMaterialChange(
  Sender: TField);
begin

 //aqui va para cuando cambia por partida
  if not bitacorademateriales.FieldByName('sIdMaterial').IsNull then
  begin
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('select SubStr(aa.mDescripcion, 1, 255) as sDescripcion, aa.sMedida, aa.sAnexo from actividadesxanexo aa '+
              'inner join anexos a on (aa.sAnexo = a.sAnexo and (a.sTipo = "MATERIAL" or a.sTipo = "ANEXO")) '+
              'where aa.sContrato =:Contrato and aa.sNumeroActividad =:Actividad and aa.sTipoActividad = "Actividad" order by sNumeroActividad');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value :=global_Contrato_Barco;
    Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('Actividad').Value := bitacoradematerialessIdMaterial.Text;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
    begin
      bitacorademateriales.FieldValues['sDescripcion'] := Connection.QryBusca.FieldValues['sDescripcion'];
      bitacorademateriales.FieldValues['sMedida']  := Connection.QryBusca.FieldValues['sMedida'];
      bitacorademateriales.FieldValues['sAnexo']   := Connection.QryBusca.FieldValues['sAnexo'];
    end
    else
      if not bitacorademateriales.FieldByName('sIdMaterial').IsNull then
        if Trim(bitacorademateriales.FieldValues['sIdMaterial']) <> '' then
        begin

          sDescripcion := '%' + Trim(bitacorademateriales.FieldValues['sIdMaterial']) + '%';
          BuscaObjeto.Active := False;
          ListaObjeto.Columns.Clear;
          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[0].FieldName := 'sNumeroActividad';
          ListaObjeto.Columns[0].Width := 100;
          ListaObjeto.Columns[0].Title.Caption := 'Anexo';
          ListaObjeto.Columns.Add;
          ListaObjeto.Columns[1].FieldName := 'sDescripcion';
          ListaObjeto.Columns[1].Width := 550;
          ListaObjeto.Columns[1].Title.Caption := 'Descripcion';
          BuscaObjeto.SQL.Clear;
          BuscaObjeto.SQL.Add('select sNumeroActividad, SubStr(aa.mDescripcion, 1, 255) as sDescripcion, aa.sMedida, aa.sAnexo from actividadesxanexo aa '+
                      'inner join anexos a on (aa.sAnexo = a.sAnexo and (a.sTipo = "MATERIAL" or a.sTipo = "ANEXO")) '+
                      'where aa.sContrato =:Contrato and aa.sTipoActividad = "Actividad" '+
                      'and mDescripcion like :Descripcion order by sNumeroActividad');
          BuscaObjeto.Params.ParamByName('Contrato').DataType    := ftString;
          BuscaObjeto.Params.ParamByName('Contrato').Value       := global_contrato_barco;
          BuscaObjeto.Params.ParamByName('Descripcion').DataType := ftString;
          BuscaObjeto.Params.ParamByName('Descripcion').Value    := sDescripcion;
          BuscaObjeto.Open;
               // yavienedeRegreso := 'No' ;
          Panel.Visible := True;
          Panel.Height  := 358;
          Panel.Width   := 676;
          ListaObjeto.SetFocus
        end
  end;
end;

procedure TfrmBitacoraDepartamental_2.btn1Click(Sender: TObject);
begin
  FiltrarFolios(LblTodos);
end;

procedure TfrmBitacoraDepartamental_2.btnAddNClick(Sender: TObject);
begin
    if ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
    begin
        MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
        exit;
    end;

    if QryBitacora.RecordCount > 0 then
    begin
        frmBarra1.btnCancelClick(Sender);
        frmbarra1.btnEdit.Enabled := False;
        btnAddN.Enabled    := False;
        btnEditN.Enabled   := False;
        btnPostN.Enabled   := True;
        btnCancelN.Enabled := True;
        btnDeleteN.Enabled := False;

        popAdd.Enabled    := False;
        popEdit.Enabled   := False;
        popPost.Enabled   := True;
        popCancel.Enabled := True;
        popDelete.Enabled := False;

        btnUp.Enabled     := False;
        btnDown.Enabled   := False;

        cxCerrado.Enabled := True;
        cxAbierto.Enabled := True;
        chkEstatusHora.Enabled  := True;
        chkEstatusHora.ReadOnly := False;
        chkEstatusHora.Checked  := True;

        if global_nota = '' then
           advNota.Click;     

        cbbtipoActividad.ItemIndex := 0;

        //Deahabilitamos los demas botones
        tsNumeroActividad.Enabled := False;
        tsIdTipoMovimiento.Enabled := False;
        tdCantidad.Enabled := True;
        tdAvance.Enabled   := True;
        tdCantidad.ReadOnly := False;
        tdAvance.ReadOnly   := False;
        chkImprime.Enabled := False;
        chkCancelada.Enabled := False;
        tsHoraInicio.ReadOnly := False;
        tsHoraFinal.ReadOnly := False;
        tmDescripcion.ReadOnly := False;
        tssIdClasificacion.Enabled  := True;
        tssIdClasificacion.ReadOnly := False;
        tsTipoAct.Enabled  := True;
        tsTipoAct.ReadOnly := False;
        cxComboGerencial.Enabled := False;
        tsNumeroActividad_Pu.Enabled  := False;
        tsNumeroActividad_ADM.Enabled := False;
        nxNumGerencial.Enabled        := True;

        tdCantidad.Value := 0;
        tdAvance.Value   := 0;

        tmDescripcion.Text := '';
        global_nota := 'Nota';
        tmDescripcion.Color := $00E6FEFF;
        advNota.Enabled      := True;

        if tsTipoAct.KeyValue = 'FO' then
        begin
           tsIdPlataforma.Visible := True;
           tsIdPernocta.Visible   := True;
           Label6.Visible         := True;
           Label8.Visible         := True;
        end;

        QryNotasGerencial.Last;
        if QryNotasGerencial.RecordCount > 0 then
        begin
            tsHoraInicio.Text := QryNotasGerencial.FieldValues['sHorafinal'];
            tsHoraFinal.Text  := QryNotasGerencial.FieldValues['sHoraFinal'];
        end
        else
        begin
            tsHoraInicio.Text := '00:00';
            tsHoraFinal.Text  := '00:00';
        end;

        sOpcion := 'Nuevo';
        OpcButton := 'New';

        tsIdPernocta.KeyValue   := qrFrentes.FieldByName('sIdPernocta').AsString;
        tsIdPlataforma.KeyValue := qrFrentes.FieldByName('sIdPlataforma').AsString;

        tssIdClasificacion.KeyValue := 'TE';
        tssIdClasificacion.SetFocus;

        advNota.Caption := 'Descripcion';
        advNota.Click;
    end
    else
       messageDLG('No existen Partidas / Notas Generales Registradas', mtInformation, [mbOk], 0);
end;

procedure TfrmBitacoraDepartamental_2.btnCancelNClick(Sender: TObject);
begin
    btnAddN.Enabled    := True;
    btnEditN.Enabled   := True;
    btnPostN.Enabled   := False;
    btnCancelN.Enabled := False;
    btnDeleteN.Enabled := True;

    popAdd.Enabled    := True;
    popEdit.Enabled   := True;
    popPost.Enabled   := False;
    popCancel.Enabled := False;
    popDelete.Enabled := True;

    chkEstatusHora.ReadOnly := True;

    btnUp.Enabled     := True;
    btnDown.Enabled   := True;

    tsNumeroActividad.Enabled := True;
    tsIdTipoMovimiento.Enabled := True;
    tdCantidad.Enabled := True;
    tdAvance.Enabled   := True;
    chkImprime.Enabled := True;
    chkCancelada.Enabled := True;
    tsHoraInicio.ReadOnly := True;
    tsHoraFinal.ReadOnly := True;
    tmDescripcion.ReadOnly := True;
    tssIdClasificacion.Enabled  := False;
    tssIdClasificacion.ReadOnly := True;
    tsTipoAct.Enabled  := True;
    tsTipoAct.ReadOnly := False;
    cxComboGerencial.Enabled := True;
    tsNumeroActividad_Pu.Enabled  := True;
    tsNumeroActividad_ADM.Enabled := True;
    nxNumGerencial.Enabled        := False;

    GridNotas.Enabled := True;
    advNota.Enabled   := True;
    sOpcion := '';
    OpcButton := '';
end;

procedure TfrmBitacoraDepartamental_2.btnDeleteNClick(Sender: TObject);
var
  SavePlace   : TBookmark;
begin
   if ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
   begin
        MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
        exit;
   end;

   if MessageDlg('Desea eliminar El Horario y todo el personal y equipo asignado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
   begin
    connection.zConnection.StartTransaction;
    try
      try
        if qryNotasGerencial.FieldByName('iHermano').AsInteger<>-1 then
        begin
          Connection.QryBusca.Active:=False;
          Connection.QryBusca.SQL.Text:='select * FROM bitacoradeactividades WHERE sContrato = :contrato and ' +
          'dIdFecha = :fecha and sIdTipoMovimiento="ED" and iHermano = :Hermano and sHoraFinal=:HoraFinal';
          Connection.QryBusca.ParamByName('Contrato').AsString:=param_Global_Contrato;
          Connection.QryBusca.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
          Connection.QryBusca.ParamByName('Hermano').AsInteger:=qryNotasGerencial.FieldByName('iHermano').AsInteger;
          Connection.QryBusca.ParamByName('HoraFinal').AsString:=qryNotasGerencial.FieldByName('sHoraInicio').AsString;
          Connection.QryBusca.Open;
          //sIdTipoMovimiento, sHoraInicio, sHoraFinal
          if Connection.QryBusca.RecordCount=1 then
          begin
            connection.zCommand.Active := False;
            connection.zCommand.SQL.text:='update bitacoradeactividades set sHoraFinal=:HoraFinal where '+
            'sContrato=:Contrato and didfecha=:Fecha and iIdDiario=:Diario and iIdtarea=:tarea and iIdActividad=:Actividad';
            Connection.zCommand.ParamByName('Contrato').AsString:=Connection.QryBusca.FieldByName('sContrato').AsString;
            Connection.zCommand.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
            Connection.zCommand.ParamByName('Diario').AsInteger:=Connection.QryBusca.FieldByName('iIdDiario').AsInteger;
            Connection.zCommand.ParamByName('tarea').AsInteger:=Connection.QryBusca.FieldByName('iIdtarea').AsInteger;
            Connection.zCommand.ParamByName('Actividad').AsInteger:=Connection.QryBusca.FieldByName('iIdActividad').AsInteger;
            Connection.zCommand.ParamByName('HoraFinal').AsString:=qryNotasGerencial.FieldByName('sHoraFinal').AsString;
            Connection.zCommand.ExecSQL;
          end
          else
          begin
            Connection.QryBusca.Active:=False;
            Connection.QryBusca.SQL.Text:='select * FROM bitacoradeactividades WHERE sContrato = :contrato and ' +
            'dIdFecha = :fecha and sIdTipoMovimiento="ED" and iHermano = :Hermano and sHoraInicio=:HoraInicio';
            Connection.QryBusca.ParamByName('Contrato').AsString:=param_Global_Contrato;
            Connection.QryBusca.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
            Connection.QryBusca.ParamByName('Hermano').AsInteger:=qryNotasGerencial.FieldByName('iHermano').AsInteger;
            Connection.QryBusca.ParamByName('HoraInicio').AsString:=qryNotasGerencial.FieldByName('sHoraFinal').AsString;
            Connection.QryBusca.Open;
            //sIdTipoMovimiento, sHoraInicio, sHoraFinal
            if Connection.QryBusca.RecordCount=1 then
            begin
              connection.zCommand.Active := False;
              connection.zCommand.SQL.text:='update bitacoradeactividades set sHoraInicio=:HoraInicio where '+
              'sContrato=:Contrato and didfecha=:Fecha and iIdDiario=:Diario and iIdtarea=:tarea and iIdActividad=:Actividad';
              Connection.zCommand.ParamByName('Contrato').AsString:=Connection.QryBusca.FieldByName('sContrato').AsString;
              Connection.zCommand.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
              Connection.zCommand.ParamByName('Diario').AsInteger:=Connection.QryBusca.FieldByName('iIdDiario').AsInteger;
              Connection.zCommand.ParamByName('tarea').AsInteger:=Connection.QryBusca.FieldByName('iIdtarea').AsInteger;
              Connection.zCommand.ParamByName('Actividad').AsInteger:=Connection.QryBusca.FieldByName('iIdActividad').AsInteger;
              Connection.zCommand.ParamByName('HoraInicio').AsString:=qryNotasGerencial.FieldByName('sHoraInico').AsString;
              Connection.zCommand.ExecSQL;
            end
            else
              if Messagedlg('El Registro que sera eliminado cuenta con un Agrupador, pero no se encontro su Horario Sucesor ni Predecesor.'+#13 + #10+
                        'Desea Continuar?', mtConfirmation,[mbyes,Mbno],0) =MrNo then
              begin
                Connection.zConnection.Rollback;
                Exit;
              end;
          end;
        end;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('DELETE FROM bitacoradepersonal WHERE sContrato = :contrato and ' +
          'dIdFecha = :fecha and iIdDiario = :diario and iIdTarea=:Tarea and iIdActividad=:ACtividad');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := qryNotasGerencial.FieldValues['iIdDiario'];
        Connection.zCommand.Params.ParamByName('Tarea').AsInteger      := qryNotasGerencial.FieldValues['iIdTarea'];
        Connection.zCommand.Params.ParamByName('ACtividad').AsInteger      := qryNotasGerencial.FieldValues['iIdActividad'];
        connection.zCommand.ExecSQL;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('DELETE FROM bitacoradeequipos WHERE sContrato = :contrato and ' +
          'dIdFecha = :fecha and iIdDiario = :diario and iIdTarea=:Tarea and iIdActividad=:ACtividad');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := qryNotasGerencial.FieldValues['iIdDiario'];
        Connection.zCommand.Params.ParamByName('Tarea').AsInteger      := qryNotasGerencial.FieldValues['iIdTarea'];
        Connection.zCommand.Params.ParamByName('ACtividad').AsInteger      := qryNotasGerencial.FieldValues['iIdActividad'];
        connection.zCommand.ExecSQL;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('DELETE FROM bitacorademateriales WHERE sContrato = :contrato and ' +
          'dIdFecha = :fecha and iIdDiario = :diario ');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := qryNotasGerencial.FieldValues['iIdDiario'];
        connection.zCommand.ExecSQL;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('DELETE FROM bitacoradeactividades WHERE sContrato = :contrato and ' +
          'dIdFecha = :fecha and iIdDiario = :diario and iIdTarea=:Tarea and iIdActividad=:ACtividad');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := qryNotasGerencial.FieldValues['iIdDiario'];
        Connection.zCommand.Params.ParamByName('Tarea').AsInteger      := qryNotasGerencial.FieldValues['iIdTarea'];
        Connection.zCommand.Params.ParamByName('ACtividad').AsInteger      := qryNotasGerencial.FieldValues['iIdActividad'];
       
        connection.zCommand.ExecSQL;

        Connection.QryBusca.Active:=False;
        Connection.QryBusca.SQL.Text:='select * FROM bitacoradeactividades WHERE sContrato = :contrato and ' +
        'dIdFecha = :fecha and sIdTipoMovimiento="ED" and iHermano = :Hermano';
        Connection.QryBusca.ParamByName('Contrato').AsString:=param_Global_Contrato;
        Connection.QryBusca.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
        Connection.QryBusca.ParamByName('Hermano').AsInteger:=qryNotasGerencial.FieldByName('iHermano').AsInteger;
        Connection.QryBusca.Open;

        if Connection.QryBusca.RecordCount=1 then
        begin
          connection.zCommand.Active := False;
          connection.zCommand.SQL.text:='update bitacoradeactividades set iHermano=-1 where '+
          'sContrato=:Contrato and didfecha=:Fecha and iIdDiario=:Diario and iIdtarea=:tarea and iIdActividad=:Actividad';
          Connection.zCommand.ParamByName('Contrato').AsString:=Connection.QryBusca.FieldByName('sContrato').AsString;
          Connection.zCommand.ParamByName('Fecha').AsDate:=tdIdFecha.Date;
          Connection.zCommand.ParamByName('Diario').AsInteger:=Connection.QryBusca.FieldByName('iIdDiario').AsInteger;
          Connection.zCommand.ParamByName('tarea').AsInteger:=Connection.QryBusca.FieldByName('iIdtarea').AsInteger;
          Connection.zCommand.ParamByName('Actividad').AsInteger:=Connection.QryBusca.FieldByName('iIdActividad').AsInteger;
          Connection.zCommand.ExecSQL;
        end;
        connection.zConnection.Commit;
        try
           SavePlace := QryNotasGerencial.GetBookmark ;
           QryNotasGerencial.Refresh;
           QryNotasGerencial.GotoBookmark(SavePlace);
        Except
          QryNotasGerencial.FreeBookmark(SavePlace);
        end;
      except
        if connection.zConnection.InTransaction then
          connection.zConnection.Rollback;
      end;
    finally
      if connection.zConnection.InTransaction then
        connection.zConnection.Rollback;
    end;
   end;
end;

procedure TfrmBitacoraDepartamental_2.btnDeleteDClick(Sender: TObject);
var
  SavePlace   : TBookmark;
begin

    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('DELETE FROM bitacoradeactividades WHERE sContrato = :contrato and ' +
      'dIdFecha = :fecha and iIdDiario = :diario');
    Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
    Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
    Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
    Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
    Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
    Connection.zCommand.Params.ParamByName('diario').value      := QrNotasDetalle.FieldValues['iIdDiario'];
    connection.zCommand.ExecSQL;

    try
       SavePlace := QrNotasDetalle.GetBookmark ;
       QrNotasDetalle.Refresh;
       QrNotasDetalle.GotoBookmark(SavePlace);
       QrNotasDetalle.FreeBookmark(SavePlace);
    Except
    end;
end;

procedure TfrmBitacoraDepartamental_2.btnDownClick(Sender: TObject);
begin
    OrdenarNotas('Abajo');
end;

procedure TfrmBitacoraDepartamental_2.btnEditNClick(Sender: TObject);
begin
    if ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
    begin
        MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
        exit;
    end;

   if QryBitacora.RecordCount > 0 then
    begin
        frmBarra1.btnCancelClick(Sender);
        frmbarra1.btnEdit.Enabled := False;
        btnAddN.Enabled    := False;
        btnEditN.Enabled   := False;
        btnPostN.Enabled   := True;
        btnCancelN.Enabled := True;
        btnDeleteN.Enabled := False;

        popAdd.Enabled    := False;
        popEdit.Enabled   := False;
        popPost.Enabled   := True;
        popCancel.Enabled := True;
        popDelete.Enabled := False;

        cxCerrado.Enabled := True;
        cxAbierto.Enabled := True;
        chkEstatusHora.Enabled  := True;
        chkEstatusHora.ReadOnly := False;

        if global_nota = '' then
           advNota.Click;

        btnUp.Enabled     := False;
        btnDown.Enabled   := False;

        tdCantidad.Value   := QryNOtasGerencial.FieldValues['dCantidad'];
        tdAvance.Value     := QryNOtasGerencial.FieldValues['dAvance'];

        tsNumeroActividad.Enabled := False;
        tsIdTipoMovimiento.Enabled := False;
        tdCantidad.Enabled := False;
        tdAvance.Enabled   := True;
        tdCantidad.ReadOnly := False;
        tdAvance.ReadOnly   := False;
        chkImprime.Enabled := True;
        chkCancelada.Enabled := False;
        tsHoraInicio.ReadOnly := False;
        tsHoraFinal.ReadOnly := False;
        tmDescripcion.ReadOnly := False;
        tssIdClasificacion.Enabled  := True;
        tssIdClasificacion.ReadOnly := False;
        tsTipoAct.Enabled  := True;
        tsTipoAct.ReadOnly := False;
        cxComboGerencial.Enabled := False;
        tsNumeroActividad_Pu.Enabled  := False;
        tsNumeroActividad_ADM.Enabled := False;
        nxNumGerencial.Enabled        := True;
        sOpcion := 'Edit';
        OpcButton := 'Edit';
        opcKardex := 'Edit';
        advNota.Caption := 'Descripcion';
        advNota.Click;
        dCantidadOld := tdCantidad.Value;
        tsHoraInicio.Text :=  qryNotasGerencial.FieldByName('sHoraInicio').AsString;
        tsHoraFinal.Text  :=  qryNotasGerencial.FieldByName('sHoraFinal').AsString;
    end
    else
       messageDLG('No existen Partidas / Notas Generales Registradas', mtInformation, [mbOk], 0);
end;

procedure TfrmBitacoraDepartamental_2.btnMayusClick(Sender: TObject);
begin
  tmDescripcion.Text := UpperCase(tmDescripcion.Text);
end;

procedure TfrmBitacoraDepartamental_2.btnPostNClick(Sender: TObject);
var
   iDiario, iRegistro     : integer;
   lIncorrecto : boolean;
   SavePlace, SavePlacePartida   : TBookmark;

   lFiltro,
   lEfectivo,
   Consolidado,
   bContinue: Boolean;

   sTiempoEfectivo,
   sNumeroActividad, sTipoAnexo, sParametro : string;

   dAvance,
   dAvanceAnterior,
   MaxCantidad : Extended;

   Q_GuardaDatos,
   Q_BuscaAvance: TZReadOnlyQuery;
   QryUpdate: TZQuery;

   nombres, cadenas: TStringList;
   sWbsAux : string;

begin
    tdAvance.OnExit(sender);
    if tmDescripcion.Text = '' then
    begin
       messageDLG('Debe escribir una descripcion de la nota!', mtInformation, [mbOk], 0);
       exit;
    end;   

    {Validacion de jornada hora > 24:00}
    lIncorrecto := False;
    if (StrToInt(copy(tsHoraInicio.Text, 1, 2)) = 24) and (StrToInt(copy(tsHoraInicio.Text, 4, 5)) > 0) then
      lIncorrecto := True
    else
      if (StrToInt(copy(tsHoraInicio.Text, 1, 2)) > 24) then
        lIncorrecto := True;

    if lIncorrecto then
    begin
      messageDLG('La Hora de Inicio es Mayor a 24:00 Hrs', mtInformation, [mbOk], 0);
      tsHoraInicio.SetFocus;
      exit;
    end;

    if (StrToInt(copy(tsHoraFinal.Text, 1, 2)) = 24) and (StrToInt(copy(tsHoraFinal.Text, 4, 5)) > 0) then
      lIncorrecto := True
    else
      if (StrToInt(copy(tsHoraFinal.Text, 1, 2)) > 24) then
        lIncorrecto := True;

    if lIncorrecto then
    begin
      messageDLG('La Hora de Final es Mayor a 24:00 Hrs', mtInformation, [mbOk], 0);
      tsHoraFinal.SetFocus;
      exit;
    end;

    //Validaciones de notas dentro del corte del gerencial
    if (tshoraInicio.Text <> '00:00') and (tsHoraFinal.Text <> '00:00') then
    begin
       if tsHoraInicio.Text < tsHoraFinal.Text then       
          if tsHorafinal.Text > QryBitacora.FieldValues['sHoraFinal'] then
             //messageDLG('La hora de Termino de la Actividad es mayor a '+ tsHoraFinal.Text , mtWarning, [mbOk], 0);
    end;

   {Continua insercion de datos}
    if (tsHoraInicio.Text = '  :  ') or (tsHoraFinal.Text = '  :  ') then
    begin
        ShowMessage('Los horarios no deben estar vacios!!');
        tsHoraFinal.SetFocus;
    end;

    //################jjivan 2014-08-12 --- Funcion Depurada##############
    Q_GuardaDatos := TZReadOnlyQuery.Create(self);
    Q_GuardaDatos.Connection := connection.zConnection;

    Q_BuscaAvance := TZReadOnlyQuery.Create(self);
    Q_BuscaAvance.Connection := connection.zConnection;

    sWbsAux            := '';
    lKardex            := False;
    lRespuesta         := True;
    bContinue          := True;

    sNumeroActividad   := QryBitacora.FieldByName('sNumeroActividad').AsString;
    swbsFormulario     := QryBitacora.FieldByName('sWbs').AsString;
    MaxCantidad        := 0;
    dInstaladoOrden    := 0;
    dInstaladoOrdenAnt := 0;
    dCantidadAnexo     := 0;
    dExcedenteOrden    := 0;
    dInstaladoAnexo    := 0;
    dExcedenteAnexo    := 0;

    opcKardex := '';

    if qryNotasGerencial.RecordCount > 0 then
        if xround(dCantidadOld,4) = xround(tdCantidad.Value,4) then
           bContinue := False;

    if OpcButton = 'New' then
       bContinue := True;

    if bContinue then
    begin
         {$REGION 'Editar Registro'}
        if OpcButton = 'Edit' then
        begin
              {Se hace una copia del registro}
              Q_GuardaDatos.Active := False;
              Q_GuardaDatos.SQL.Clear;
              Q_GuardaDatos.SQL.Add('select * from bitacoradeactividades where sContrato =:Contrato and sIdConvenio =:Convenio and dIdFecha =:Fecha and iIdDiario =:Diario and sWbs =:Wbs and sIdTipoMovimiento = "ED"');
              Q_GuardaDatos.ParamByName('Contrato').AsString := param_global_contrato;
              Q_GuardaDatos.ParamByName('Convenio').AsString := QrFrentes.FieldByName('Convenio').AsString;
              Q_GuardaDatos.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
              Q_GuardaDatos.ParamByName('Diario').AsInteger  := QryNotasGerencial.FieldValues['iIdDiario'];
              Q_GuardaDatos.ParamByName('Wbs').AsString      := sWbsFormulario;
              Q_GuardaDatos.Open;

              iDiario := QryBitacora.FieldValues['iIdDiario'];
              {Se elimina el registo con IdDiario Especificado..}
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('delete from bitacoradeactividades where sContrato = :contrato And sIdConvenio =:Convenio And dIdFecha = :fecha And iIdDiario = :diario and sWbs =:Wbs and sIdTipoMovimiento = "ED" ');
              connection.zCommand.Params.ParamByName('contrato').AsString := param_Global_Contrato;
              connection.zCommand.Params.ParamByName('Convenio').AsString := QrFrentes.FieldByName('Convenio').AsString;
              connection.zCommand.Params.ParamByName('fecha').AsDate      := tdIdFecha.Date;
              connection.zCommand.Params.ParamByName('diario').AsInteger  := QryNotasGerencial.FieldValues['iIdDiario'];
              connection.zCommand.Params.ParamByName('Wbs').AsString      := sWbsFormulario;
              connection.zCommand.ExecSQL;

              {Leer las cantiades reportadas a esta fecha.. Funcion}
              dInstaladoOrden := InstaladoOrden(sNumeroActividad, sWbsFormulario, 0);

              if opcKardex = 'Edita' then
                 dInstaladoOrden := dInstaladoOrden + tdCantidad.Value;

              //Verifica si lo reportado no excede lo permitido en el Frente/Folio..
              if dInstaladoOrden > ActividadesIguales.FieldValues['dCantidad'] then
              begin
                 dExcedenteOrden := dInstaladoOrden - ActividadesIguales.FieldValues['dCantidad'];
                 dInstaladoOrden := ActividadesIguales.FieldValues['dCantidad'];
              end;

              {Leer solamente las cantiades reportadas anteriormente a esta fecha.. Funcion}
              //dInstaladoOrdenAnt := InstaladoOrden(sNumeroActividad, sWbsFormulario, 1);

              {Obtener Cantidad Anexo ...Funcion..}
              dCantidadAnexo := CantidadAnexoC;

              {Cantidad Instalada en el Anexo C ... Funcion}
              dInstaladoAnexo := InstaladoAnexoC(sNumeroActividad, sWbsFormulario);

              if opcKardex = 'Edita' then
                 dInstaladoAnexo :=  dInstaladoAnexo + tdCantidad.Value;

              if dInstaladoAnexo > dCantidadAnexo then
              begin
                  dExcedenteAnexo := dInstaladoAnexo - dCantidadAnexo;
                  dInstaladoAnexo := dCantidadAnexo;
              end;

              {Actualiza la cantidad Instalada en el Folio/Frente.. Funcion}
              lEfectivo := fnActualizaAcumuladosOrden('', QryBitacora.FieldByName('sWbs').AsString, sNumeroActividad, ActividadesIguales.FieldValues['dCantidad'], dInstaladoOrden, dExcedenteOrden, 0);

              {Actualiza la cantidad Instalada en el Anexo C.. Funcion}
              lEfectivo := fnActualizaAcumuladosContrato('', sNumeroActividad, dCantidadAnexo, dInstaladoAnexo, dExcedenteAnexo, 0);

              OpcButton     := 'New';
              global_Editor := 'Edit';
              lRespuesta:=false;
        end;
        {$ENDREGION}
        lEfectivo := False;
        if OpcButton = 'New' then
        begin
            dAvance         := 0;
            lFiltro         := False;
            sTiempoEfectivo := tsIdTipoMovimiento.KeyValue;

            if TiposdeMovimiento.FieldValues['sClasificacion'] = 'Notas' then
               lFiltro := True
            else
               if ActividadesIguales.RecordCount > 0 then
                  if TiposdeMovimiento.FieldValues['sClasificacion'] = 'Tiempo en Operacion' then
                     lEfectivo := True;

            sWbsFormulario := ActividadesIguales.FieldByName('sWbs').AsString;
            SavePlace      := ActividadesIguales.GetBookmark;

            if lEfectivo then
            begin
                {Validamos que la cantidad maxima a reportar...}
                if Connection.configuracion.FieldValues['sAvanceBitacora'] = 'Volumen' then
                  lFiltro := fnValidaPartidaOrden(sWbsFormulario, sNumeroActividad)
                else
                  lFiltro := fnValidaPartidaOrdenPorcentaje(sWbsFormulario, sNumeroActividad);
            end;

            //Si la respuesta es No, Regresamos el registro eliminado de la bitacoradeactividades..
            if lRespuesta = False then
            begin
              if global_Editor =  'Edit' then
                if Q_GuardaDatos.RecordCount > 0 then
                begin
                    Connection.zCommand.Active := False;
                    Connection.zCommand.SQL.Clear;
                    Connection.zCommand.SQL.Add(funcsql(Q_GuardaDatos, 'bitacoradeactividades'));
                    Connection.zCommand.Active := False;
                    {Se inserta nuevamente el registor a bitacoradeactividades..}
                    for iRegistro := 1 to Q_GuardaDatos.fieldcount do
                    begin
                        sParametro := 'param' + trim(inttostr(iRegistro));
                        connection.zCommand.Params.parambyname(sParametro).datatype := Q_GuardaDatos.fields[iRegistro - 1].datatype;
                        if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'sHoraInicio') then
                            connection.zCommand.Params.parambyname(sparametro).value := tsHoraInicio.Text
                          else
                            if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'sHoraFinal') then
                              connection.zCommand.Params.parambyname(sparametro).value := tsHoraFinal.Text
                            else
                               if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'dCantidad') then
                                connection.zCommand.Params.parambyname(sparametro).value := tdCantidad.value
                               else
                                if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'dAvance') then
                                  connection.zCommand.Params.parambyname(sparametro).value := tdAvance.value
                                else
                                   if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'iNumeroGerencial') then
                                       connection.zCommand.Params.parambyname(sparametro).value := StrToInt(nxNumGerencial.Text)
                                   else
                                       if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'descripcion') then
                                           connection.zCommand.Params.parambyname(sparametro).value := tmDescripcion.Text
                                        else
                                            if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'tipoActividad') then
                                            begin
                                                if cbbtipoActividad.Text = '' then
                                                   connection.zCommand.Params.parambyname(sparametro).value := 'Desconocido'
                                                else
                                                   connection.zCommand.Params.parambyname(sparametro).value := cbbTipoActividad.Text;
                                            end
                                            else
                                                if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'clasificacion') then
                                                   connection.zCommand.Params.parambyname(sparametro).value := tssIdClasificacion.KeyValue
                                                else
                                                   if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'sTipoObra') then
                                                       connection.zCommand.Params.parambyname(sparametro).value := tsTipoAct.KeyValue
                                                   else
                                                      if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'sIdPernocta') then
                                                            connection.zCommand.Params.parambyname(sparametro).value := tsIdPernocta.KeyValue
                                                      else
                                                         if (Q_GuardaDatos.fields[iRegistro - 1].DisplayName = 'sIdPlataforma') then
                                                             connection.zCommand.Params.parambyname(sparametro).value := tsIdPlataforma.KeyValue
                                                           else
                                                               connection.zCommand.Params.parambyname(sparametro).value    := Q_GuardaDatos.fields[iRegistro - 1].value;
                    end;

                    connection.zCommand.ExecSQL;
                    lFiltro:=false;
                end;
                lRespuesta := True;
            end;

            //Continua proceso normal del sistema..
            if lFiltro then
            begin
                {Busqueda de la partida... Funcion}
                PartidaExistente(sNumeroActividad, sWbsFormulario, sTiempoEfectivo);
                   begin
                       {$REGION 'Inserta Nuevo Registro'}
                       {Obtener IdDiarioMaximo... Funcion}
                       iDiario := MaximoItem;
                       IdAct   := iDiario;

                       if lEfectivo then
                       begin
                           dAvance := 0;
                           if ActividadesIguales.FieldValues['dCantidad'] > 0 then
                              dAvance := AvanceActual(sNumeroActividad,sWbsFormulario );
                       end;

                      {Obtenemos el Maximo Item ... Funcion.}
                      iDiario := MaximoItem;
                      idAct   := iDiario;

                      connection.zCommand.Active := False;
                      connection.zCommand.SQL.Clear;
                      connection.zCommand.SQL.Add('INSERT INTO bitacoradeactividades ( sContrato, sIdConvenio, dIdFecha , iIdDiario, sIdTurno, sNumeroActividad, ' +
                        ' sNumeroOrden, sWbs, sIdTipoMovimiento, sHoraInicio, sHoraFinal, mDescripcion, lImprime, sConceptoGerencial, iIdDiarioNota, sIdClasificacion, iIdTarea, eTipoActividad, dCantidad, dAvance, sTipoObra, sIdPernocta, sIdPlataforma) ' +
                        ' VALUES (:contrato, :Convenio, :fecha, :diario, :turno, :actividad, :orden, :Wbs, :tipo, :inicio, :final, :descripcion, :imprime, :Concepto, :Id, :clasificacion, :diario, :tipoActividad, :cantidad, :avance, :tipoAct, :pernocta, :plataforma)');
                      Connection.zCommand.Params.ParamByName('contrato').AsString      := param_Global_Contrato;
                      connection.zCommand.Params.ParamByName('Convenio').AsString      := QrFrentes.FieldByName('Convenio').AsString;
                      Connection.zCommand.Params.ParamByName('fecha').AsDate           := tdIdFecha.Date;
                      Connection.zCommand.Params.ParamByName('diario').AsInteger       := iDiario;
                      Connection.zCommand.Params.ParamByName('turno').AsString         := global_turno_reporte;
                      Connection.zCommand.Params.ParamByName('orden').AsString         := QrFrentes.FieldByName('sNumeroOrden').AsString;
                      if QryBitacora.FieldValues['sIdTipoMovimiento'] = 'N' then
                      begin
                          Connection.zCommand.Params.ParamByName('wbs').AsString       := Null;
                          Connection.zCommand.Params.ParamByName('actividad').AsString := '';
                      end
                      else
                      begin
                          Connection.zCommand.Params.ParamByName('wbs').AsString       := QryBitacora.FieldValues['sWbs'];
                          Connection.zCommand.Params.ParamByName('actividad').AsString := QryBitacora.FieldValues['sNumeroActividad'];
                      end;
                      Connection.zCommand.Params.ParamByName('avance').AsFloat         := tdAvance.value;
                      Connection.zCommand.Params.ParamByName('cantidad').AsFloat       := tdCantidad.Value;
                      Connection.zCommand.Params.ParamByName('tipo').AsString          := 'ED';
                      Connection.zCommand.Params.ParamByName('inicio').AsString        := tsHoraInicio.Text;
                      Connection.zCommand.Params.ParamByName('final').AsString         := tsHoraFinal.Text;
                      Connection.zCommand.Params.ParamByName('descripcion').AsMemo     := tmDescripcion.Text;
                      Connection.zCommand.Params.ParamByName('concepto').AsMemo        := '';
                      Connection.zCommand.Params.ParamByName('Imprime').AsString       := 'Si';
                      Connection.zCommand.Params.ParamByName('Id').AsInteger           := QryBitacora.FieldValues['iIdDiario'];
                      Connection.zCommand.Params.ParamByName('clasificacion').AsString := tssIdClasificacion.KeyValue;
                      Connection.zCommand.Params.ParamByName('tipoAct').AsString          := tsTipoAct.KeyValue;
                      if cbbtipoActividad.Text = '' then
                         Connection.zCommand.ParamByName('tipoActividad').AsString    := 'Desconocido'
                      else
                         Connection.zCommand.ParamByName('tipoActividad').AsString    := cbbTipoActividad.Text;
                      Connection.zCommand.Params.ParamByName('pernocta').AsString     := tsIdPernocta.KeyValue;
                      Connection.zCommand.Params.ParamByName('plataforma').AsString   := tsIdPlataforma.KeyValue;
                      connection.zCommand.ExecSQL;

                      tsHoraInicio.Text := tsHoraFinal.Text;
                      tmDescripcion.Text := '';

                       if lEfectivo then
                          if fnActualizaAcumuladosOrden('', sWbsFormulario, sNumeroActividad, dCantidadOrden, dInstaladoOrden, dExcedenteOrden, tdCantidad.Value) then
                            if fnActualizaAcumuladosContrato('', sNumeroActividad, dCantidadAnexo, dInstaladoAnexo, dExcedenteAnexo, tdCantidad.Value) then

                   {$ENDREGION}
                   end;
            end;
        end ;
    end
    else
    begin
        {$REGION 'Actualiza Notas'}
          try
              QryUpdate := TZQuery.Create(nil);
              QryUpdate.Connection := Connection.zConnection;
              QryUpdate.Active := False;
              QryUpdate.SQL.Clear;
              QryUpdate.SQL.Add('UPDATE bitacoradeactividades SET mDescripcion = :descripcion, sHoraInicio = :HoraInicio, sHoraFinal = :HoraFinal, sTipoObra =:Tipo, sIdPernocta =:Pernocta, sIdPlataforma =:Plataforma, sIdClasificacion =:Clasificacion ' +
                                'where sContrato = :contrato and sIdConvenio =:Convenio And dIdFecha = :fecha And iIdDiario = :diario  ');
              QryUpdate.Params.ParamByName('contrato').AsString    := param_Global_Contrato;
              QryUpdate.Params.ParamByName('Convenio').AsString    := QrFrentes.FieldByName('Convenio').AsString;
              QryUpdate.Params.ParamByName('fecha').AsDate         := tdIdFecha.Date;
              QryUpdate.Params.ParamByName('diario').AsInteger     := qryNotasGerencial.FieldValues['iIdDiario'];
              QryUpdate.Params.ParamByName('descripcion').AsMemo   := tmDescripcion.Text;
              QryUpdate.Params.ParamByName('HoraInicio').AsString  := tsHoraInicio.Text;
              QryUpdate.Params.ParamByName('HoraFinal').AsString   := tsHoraFinal.Text;
              QryUpdate.Params.ParamByName('Tipo').AsString        := tsTipoAct.KeyValue;
              QryUpdate.Params.ParamByName('Pernocta').AsString    := tsIdPernocta.KeyValue;
              QryUpdate.Params.ParamByName('Plataforma').AsString  := tsIdPlataforma.KeyValue;
              QryUpdate.Params.ParamByName('Clasificacion').AsString  := tssIdClasificacion.KeyValue;
              QryUpdate.ExecSQL;

              //Asignamos movieintos a kardex del sistema..}
              lKardex := true;
              sWbsKardex := QryBitacora.FieldByName('sWbs').AsString;
              DecodeDate(tdIdFecha.Date, myYear, myMonth, myDay);
              fechaKardex := inttostr(myDay) + '/' + inttostr(myMonth) + '/' + inttostr(myYear);
          except
            on e: exception do
            begin
              UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Registros de Volumenes de Obra y Notas', 'Al actualizar el registro', 0);
            end;
          end
       {$ENDREGION}
    end;

    sOpcion := '';
    opcKardex := '';

    //if OpcButton = 'Edit' then
       btnCancelN.Click;

    cbbtipoActividad.ItemIndex := 0;

    SavePlace := QryNotasGerencial.GetBookmark ;
    QryNotasGerencial.Refresh;
    QryNotasGerencial.GotoBookmark(SavePlace);
    QryNotasGerencial.FreeBookmark(SavePlace);
//
//    SavePlacePartida  := BView_Actividades.DataController.DataSource.DataSet.GetBookmark;
//    QryBitacora.Active := False;
//    QryBitacora.Open;
//    try
//      BView_Actividades.DataController.DataSource.DataSet.GotoBookmark(SavePlacePartida);
//    except
//    else
//      BView_Actividades.DataController.DataSet.FreeBookmark(SavePlacePartida);
//    end;

    
    if OpcButton = 'New' then
    begin
        QryNotasGerencial.Last;
        tshoraInicio.Text := QryNotasGerencial.FieldByName('sHoraFinal').AsString;
        tsHoraFinal.Text  := QryNotasGerencial.FieldByName('sHoraFinal').AsString;
    end;
    tsHoraInicio.SetFocus;
end;

procedure TfrmBitacoraDepartamental_2.btnUpClick(Sender: TObject);
begin
    OrdenarNotas('Arriba');
end;

procedure TfrmBitacoraDepartamental_2.BView_ActividadesDblClick(
  Sender: TObject);
begin
    if BView_Actividades.OptionsView.CellAutoHeight then
       BView_Actividades.OptionsView.CellAutoHeight := False
    else
       BView_Actividades.OptionsView.CellAutoHeight := True;
end;

procedure TfrmBitacoraDepartamental_2.Can1Click(Sender: TObject);
begin
  frmBarra1.btnCancel.Click
end;

procedure TfrmBitacoraDepartamental_2.Eliminar1Click(Sender: TObject);
begin
  frmBarra1.btnDelete.Click
end;

procedure TfrmBitacoraDepartamental_2.Refresh1Click(Sender: TObject);
begin
  frmBarra1.btnRefresh.Click
end;

procedure TfrmBitacoraDepartamental_2.Salir1Click(Sender: TObject);
begin
  frmBarra1.btnExit.Click
end;

procedure TfrmBitacoraDepartamental_2.sNumeroActividadStylesGetContentStyle(
  Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
  AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
begin
//    AStyle:=cxNormal;
//    AItem:= (Sender as TcxGridDBTableView).GetColumnByFieldName('sIdGerencial');
//     if VarToStr(ARecord.Values[AItem.Index])='2' then
//       AStyle:=cxDiurno;
//
//     if VarToStr(ARecord.Values[AItem.Index])='1' then
//         AStyle:=cxNocturno;
//
//     if VarToStr(ARecord.Values[AItem.Index])='3' then
//        AStyle:=cxAlCorte;
end;

procedure TfrmBitacoraDepartamental_2.SpeedButton1Click(Sender: TObject);
begin
  FiltrarFolios(LblReportados);
end;

procedure TfrmBitacoraDepartamental_2.LblReportadosClick(Sender: TObject);
begin
  FiltrarFolios(Sender);
end;

procedure TfrmBitacoraDepartamental_2.LblTodosClick(Sender: TObject);
begin
 FiltrarFolios(Sender);
end;

procedure TfrmBitacoraDepartamental_2.LeerDatasets(PsContrato: string; PsNumeroOrden: string);
var
  ListItem: TListItem;
  qryGrupos: tZReadOnlyQuery;
begin
  ReporteDiario.Active := False;
  ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
  ReporteDiario.Params.ParamByName('contrato').Value := global_Contrato_Barco;
  ReporteDiario.Params.ParamByName('turno').DataType := ftString;
  ReporteDiario.Params.ParamByName('turno').Value := global_turno_reporte;
  ReporteDiario.Params.ParamByName('Fecha').DataType := ftDate;
  ReporteDiario.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
  ReporteDiario.Params.ParamByName('Orden').DataType := ftString;
  ReporteDiario.Params.ParamByName('Orden').Value := param_global_contrato;
  ReporteDiario.Open;

  // Termina Limpia
  QryPartidasEfectivas.Active := False;
  QryPartidasEfectivas.Params.ParamByName('contrato').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('contrato').Value := param_global_contrato;
  QryPartidasEfectivas.Params.ParamByName('convenio').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
  QryPartidasEfectivas.Params.ParamByName('Orden').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('Orden').Value := PsNumeroOrden;
  QryPartidasEfectivas.Open;

  QryBitacora.Active := False;
  QryBitacora.Params.ParamByName('contrato').DataType := ftString;
  QryBitacora.Params.ParamByName('contrato').Value := param_global_contrato;
  QryBitacora.Params.ParamByName('convenio').DataType := ftString;
  QryBitacora.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
  QryBitacora.Params.ParamByName('orden').DataType := ftString;
  QryBitacora.Params.ParamByName('orden').Value := '%';
  QryBitacora.Params.ParamByName('fecha').DataType := ftDate;
  QryBitacora.Params.ParamByName('fecha').Value := tdIdFecha.Date;
  QryBitacora.Params.ParamByName('Alcance').DataType := ftString;
  QryBitacora.Params.ParamByName('Alcance').Value := Connection.configuracion.FieldValues['sTipoAlcance'];
  QryBitacora.Params.ParamByName('Ordenado').DataType := ftString;
  QryBitacora.Params.ParamByName('Ordenado').Value := 'iItemOrden';
  QryBitacora.Open;
  QryBitacoraAfterScroll(QryBitacora);
end;

procedure TfrmBitacoraDepartamental_2.tdCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
 if not keyFiltroTRxCalcEdit(tdCantidad, key) then
      key := #0;
      if Key = #13 then
         tdAvance.SetFocus ;
end;

procedure TfrmBitacoraDepartamental_2.tdEditaAvanceEnter(Sender: TObject);
begin
    tdEditaAvance.Color := global_color_entrada;
end;

procedure TfrmBitacoraDepartamental_2.tdEditaAvanceExit(Sender: TObject);
begin
    tdEditaAvance.Color := clActiveCaption;
end;

procedure TfrmBitacoraDepartamental_2.tdEditaAvanceKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key=#13 then
       cxedita.SetFocus;
end;

procedure TfrmBitacoraDepartamental_2.grid_bitacoraEnter(Sender: TObject);
var
  iCheck: Byte;
  sDescompone, sFase, sNumeroActividad: string;
begin
  if frmbarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;

  if tsNumeroActividad.KeyValue <> Null then
     sNumeroActividad := tsNumeroActividad.Text ;

  lMostrarNotas := False;

  if (QryPartidasEfectivas.Active) and (QryBitacora.Active) then
  begin
      ActividadesIguales.Active := False;
      ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
      ActividadesIguales.Params.ParamByName('contrato').Value := param_global_contrato;
      ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
      ActividadesIguales.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
      ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
      ActividadesIguales.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
      ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
      ActividadesIguales.Params.ParamByName('actividad').Value := sNumeroActividad;
      ActividadesIguales.ParamByName('turno').AsString := global_turno_reporte;
      ActividadesIguales.ParamByName('fecha').AsDate := tdIdFecha.Date;
      ActividadesIguales.Open;

      tsHoraInicio.Text  := QryBitacora.FieldValues['sHoraInicio'];
      tsHoraFinal.Text   := QryBitacora.FieldValues['sHoraFinal'];
  end;
end;

procedure TfrmBitacoraDepartamental_2.tdIdFechaEnter(Sender: TObject);
begin
  frmBarra1.btnCancel.Click;
  tdIdFecha.Color := global_color_entrada;
end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividadEnter(
  Sender: TObject);
begin
  tsNumeroActividad.Color := global_color_entrada;
  tsNumeroActividad_ADM.KeyValue := Null;
  tsNumeroActividad_PU.KeyValue  := Null;
  tsNumeroActividad_PU.Color  := global_color_pantalla;
  tsNumeroActividad_ADM.Color := global_color_pantalla;

end;


procedure TfrmBitacoraDepartamental_2.tsNumeroActividadExit(Sender: TObject);
begin

   tdCantidad.Value   := 0;
   tmDescripcion.Text := '';

    if tsNumeroActividad.Text <> '' then
    begin
      {Se buscan todas las actiivdades que tengan el mismo nombre...}
      ActividadesIguales.Active := False;
      ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
      ActividadesIguales.Params.ParamByName('contrato').Value := param_global_contrato;
      ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
      ActividadesIguales.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
      ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
      ActividadesIguales.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
      ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
      ActividadesIguales.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
      ActividadesIguales.ParamByName('turno').AsString := global_turno_reporte;
      ActividadesIguales.ParamByName('fecha').AsDate := tdIdFecha.Date;
      ActividadesIguales.Open;
    end
    else
    begin
      ActividadesIguales.Active := False;
      ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
      ActividadesIguales.Params.ParamByName('contrato').Value := param_global_contrato;
      ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
      ActividadesIguales.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
      ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
      ActividadesIguales.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
      ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
      ActividadesIguales.Params.ParamByName('actividad').Value := '';
      ActividadesIguales.ParamByName('turno').AsString := global_turno_reporte;
      ActividadesIguales.ParamByName('fecha').AsDate := tdIdFecha.Date;
      ActividadesIguales.Open;
    end;
end;

procedure TfrmBitacoraDepartamental_2.tsHoraFinalChange(Sender: TObject);
begin
   //activapop(frmBitacoraDepartamental_2, popupprincipal);
end;

procedure TfrmBitacoraDepartamental_2.tsHoraFinalEnter(Sender: TObject);
begin
    tsHoraFinal.Color := global_color_entrada;   
end;

procedure TfrmBitacoraDepartamental_2.tsHoraFinalExit(Sender: TObject);
begin
    tsHoraFinal.Color := global_color_salida;
end;

procedure TfrmBitacoraDepartamental_2.tsHoraFinalKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key = #13 then
       tdAvance.SetFocus;

end;

procedure TfrmBitacoraDepartamental_2.tsHoraInicioChange(Sender: TObject);
begin
    //activapop(frmBitacoraDepartamental_2, popupprincipal);
end;

procedure TfrmBitacoraDepartamental_2.tsHoraInicioEnter(Sender: TObject);
begin
    tsHoraInicio.Color := global_color_entrada;
    if tssIdClasificacion.Text = '' then
       tssIdClasificacion.KeyValue := QrClasificacion.FieldValues['sIdTipoMovimiento'];
end;

procedure TfrmBitacoraDepartamental_2.tsHoraInicioExit(Sender: TObject);
begin
    tsHoraInicio.Color := global_color_salida;
end;

procedure TfrmBitacoraDepartamental_2.tsHoraInicioKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key = #13 then
       tsHoraFinal.SetFocus;
end;

procedure TfrmBitacoraDepartamental_2.tsIdTipoMovimientoEnter(
  Sender: TObject);
begin
  tsIdTipoMovimiento.Color := global_color_entrada
end;

procedure TfrmBitacoraDepartamental_2.tsIdTipoMovimientoExit(
  Sender: TObject);
begin
  tsIdTipoMovimiento.Color := global_color_salida;
end;

procedure TfrmBitacoraDepartamental_2.tdCantidadChange(Sender: TObject);
begin
  TRxCalcEditChangef(tdCantidad, 'Cantidad');
end;

procedure TfrmBitacoraDepartamental_2.tdCantidadEnter(Sender: TObject);
begin
    tdCantidad.Color := global_color_entrada;
end;

procedure TfrmBitacoraDepartamental_2.tdCantidadExit(Sender: TObject);
begin
    tdCantidad.Color := global_color_salida ;  
    try
      if actividadesiguales.RecordCount > 0 then
      begin
          if (tdCantidad.Value > 0) and (tdCantidad.Value <= actividadesiguales.FieldValues['dCantidad']) then
          begin
             if actividadesiguales.FieldValues['dCantidad'] > 0 then
                tdAvance.Value := (tdCantidad.Value / actividadesiguales.FieldValues['dCantidad']) * 100;
          end
          else
          begin
              if tdCantidad.Value >= actividadesiguales.FieldValues['dCantidad'] then
                 tdAvance.Value := 100
              else
                 tdAvance.Value := 0;
          end;
      end;
    Except

    end;
end;

procedure TfrmBitacoraDepartamental_2.tmDescripcionDblClick(Sender: TObject);
begin
  if global_Editor <> 'Nuevo' then
  begin
    sTituloVentana := ' DESCRIPCION PARTIDA / NOTAS GENERALES';
    Application.CreateForm(TfrmEditorBitacoraDepartamental, frmEditorBitacoraDepartamental);
    frmEditorBitacoraDepartamental.ShowModal;
  end;
end;

procedure TfrmBitacoraDepartamental_2.tmDescripcion_GerencialEnter(
  Sender: TObject);
begin
  if (modoEdit) OR (insertedit) then begin
    if tsIdTipoMovimiento.Text <> 'VOLUMEN DE OBRA' then begin
      TMemo(Sender).ReadOnly := False;
    end;
  end;
end;

procedure TfrmBitacoraDepartamental_2.popImprimeGerencialClick(Sender: TObject);
begin
    if qryNotasGerencial.RecordCount > 0 then
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE bitacoradeactividades SET lImprime_Gerencial = :imprime ' +
                                    'where sContrato = :contrato And dIdFecha = :fecha And iIdActividad = :id ');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
        Connection.zCommand.Params.ParamByName('Id').DataType       := ftInteger;
        Connection.zCommand.Params.ParamByName('Id').value          := qryNotasGerencial.FieldValues['iIdActividad'];
        Connection.zCommand.Params.ParamByName('imprime').DataType  := ftString;
        if popImprimeGerencial.ImageIndex = 64 then
           Connection.zCommand.Params.ParamByName('imprime').value  := 'No'
        else
           Connection.zCommand.Params.ParamByName('imprime').value  := 'Si';
        connection.zCommand.ExecSQL;
        qryNotasGerencial.Refresh;
    end;
end;

procedure TfrmBitacoraDepartamental_2.popPostClick(Sender: TObject);
begin
    btnPostN.OnClick(sender);
end;

procedure TfrmBitacoraDepartamental_2.grid_igualesGetCellParams(
  Sender: TObject; Field: TField; AFont: TFont; var Background: TColor;
  Highlight: Boolean);
begin
    if (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('dCantidad').AsFloat = (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('dInstalado').AsFloat then
    begin
       Afont.Color  := clBlue;
       Afont.Style  := [fsBold];
    end;

    if (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('dExcedente').AsFloat > 0 then
    begin
       Afont.Color  := clRed;
       Afont.Style  := [fsBold];
    end;
end;


procedure TfrmBitacoraDepartamental_2.grid_igualesTitleClick(Column: TColumn);
begin
  UtGrid2.DbGridTitleClick(Column);
end;

procedure TfrmBitacoraDepartamental_2.hGerencialStylesGetContentStyle(
  Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
  AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
begin
//    AStyle := cxNormal;
//
//   AItem:= (Sender as TcxGridDBTableView).GetColumnByFieldName('sIdGerencial');
//   if VarToStr(ARecord.Values[AItem.Index])='2' then
//       AStyle:=cxDiurno;
//
//   if VarToStr(ARecord.Values[AItem.Index])='1' then
//       AStyle:=cxNocturno;
//
//   if VarToStr(ARecord.Values[AItem.Index])='3' then
//      AStyle:=cxAlCorte;
end;

procedure TfrmBitacoraDepartamental_2.Grid_BitacoraGetCellParams(
  Sender: TObject; Field: TField; AFont: TFont; var Background: TColor;
  Highlight: Boolean);
var x: integer;
begin
  if (Sender as TrxDBGrid).DataSource.DataSet.State = dsBrowse then
    if QryBitacora.RecordCount > 0 then
    begin
      AFont.Color := clBlack;
      if i > 1 then
      begin
        for x := 1 to i - 1 do
        begin
          if ListaPEQ[x] = QryBitacora.FieldValues['iIdDiario'] then
          begin
            Afont.Style := [fsBold];
            AFont.Color := clBlue;
          end;
        end;
      end;

//      if (QryBitacora.FieldValues['sIdTipoMovimiento']= 'E') then
//      begin
//         Background := $0083EDF5 ;
//         AFont.Color := clBlue;
//         Afont.Style := [fsBold];
//      end;
    end;
end;

procedure TfrmBitacoraDepartamental_2.Grid_BitacoraTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmBitacoraDepartamental_2.TiposdeMovimientoAfterScroll(
  DataSet: TDataSet);
begin
  if TiposdeMovimiento.FieldValues['sClasificacion'] <> 'Tiempo en Operacion' then
  begin
    tsNumeroActividad.Color := global_color_pantalla;
    tdCantidad.Color := global_color_pantalla;
    tsNumeroActividad.KeyValue := '';
    tsNumeroActividad.Enabled := False;
    tmDescripcion.Text := '';

    ActividadesIguales.Active := False;
    tdCantidad.Enabled := False;
  end
  else
  begin
    tsNumeroActividad.Color := global_color_text;
    tdCantidad.Color := global_color_text;
    tsNumeroActividad.Enabled := True;
    tdCantidad.Enabled := True;
  end;

end;

procedure TfrmBitacoraDepartamental_2.mDescripcionStylesGetContentStyle(
  Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
  AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
begin
//    AStyle := cxNormal;
//    if BView_Actividades.OptionsView.CellAutoHeight = False then
//    begin
//        AItem:= (Sender as TcxGridDBTableView).GetColumnByFieldName('sIdGerencial');
//        if VarToStr(ARecord.Values[AItem.Index])='2' then
//           AStyle:=cxDiurno;
//
//        if VarToStr(ARecord.Values[AItem.Index])='1' then
//           AStyle:=cxNocturno;
//
//        if VarToStr(ARecord.Values[AItem.Index])='3' then
//           AStyle:=cxAlCorte;
//    end;
end;

procedure TfrmBitacoraDepartamental_2.nxNumGerencialChange(Sender: TObject);
begin
    if nxNumGerencial.Text = '1' then
       cxComboGerencial.ItemIndex := 0;

    if nxNumGerencial.Text = '2' then
       cxComboGerencial.ItemIndex := 1;

    if nxNumGerencial.Text = '2' then
       cxComboGerencial.ItemIndex := 2;

end;

procedure TfrmBitacoraDepartamental_2.tsNumeroActividadChange(
  Sender: TObject);
begin
  global_partida := tsNumeroActividad.Text;

  tdCantidad.Value := 0;
  tmDescripcion.Text := '';
  tsHoraInicio.ReadOnly := False;
  tsHoraFinal.ReadOnly := False;
  //tsNumeroActividad.Hint := QryPartidasEfectivas.FieldByName('mDescripcion').AsString;

  if (frmBarra1.btnCancel.Enabled = True) and (not tsNumeroActividad.ReadOnly) then
    if tsNumeroActividad.Text <> '' then
    begin

      {Se buscan todas las actiivdades que tengan el mismo nombre...}
      ActividadesIguales.Active := False;
      ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
      ActividadesIguales.Params.ParamByName('contrato').Value := param_global_contrato;
      ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
      ActividadesIguales.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
      ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
      ActividadesIguales.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
      ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
      ActividadesIguales.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
      ActividadesIguales.ParamByName('turno').AsString := global_turno_reporte;
      ActividadesIguales.ParamByName('fecha').AsDate := tdIdFecha.Date;
      ActividadesIguales.Open;
    end
    else
    begin
      ActividadesIguales.Active := False;
      ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
      ActividadesIguales.Params.ParamByName('contrato').Value := param_global_contrato;
      ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
      ActividadesIguales.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
      ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
      ActividadesIguales.Params.ParamByName('orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
      ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
      ActividadesIguales.Params.ParamByName('actividad').Value := '';
      ActividadesIguales.ParamByName('turno').AsString := global_turno_reporte;
      ActividadesIguales.ParamByName('fecha').AsDate := tdIdFecha.Date;
      ActividadesIguales.Open;

    end;
  tsNumeroActividad.Color := global_color_salida  
end;

procedure TfrmBitacoraDepartamental_2.Actividades1Click(Sender: TObject);
begin
   // GeneraReporteDiario_PDF(FtAbordo,FtsActividad);
end;

procedure TfrmBitacoraDepartamental_2.ActividadesDetalle1Click(Sender: TObject);
begin
    global_FrenteTrabajo := QrFrentes.FieldByName('sNumeroOrden').AsString;
    Application.CreateForm(TfrmDesgloceActividadesPEQ, frmDesgloceActividadesPEQ);
    frmDesgloceActividadesPEQ.showModal
end;

procedure TfrmBitacoraDepartamental_2.Actividadesenproceso1Click(
  Sender: TObject);
begin
    Application.CreateForm(TfrmOpcionesActividades, frmOpcionesActividades);
    frmOpcionesActividades.showModal
end;

procedure TfrmBitacoraDepartamental_2.ActividadesIgualesAfterScroll(
  DataSet: TDataSet);
begin
  if ActividadesIguales.State <> dsInactive then
  begin
    if ActividadesIguales.Active and (ActividadesIguales.RecordCount > 0) then
    begin
      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select sWbsAnterior from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden and sWbs =:Wbs and sTipoActividad ="Paquete" ');
      connection.QryBusca.ParamByName('Contrato').AsString := param_global_contrato;
      connection.QryBusca.ParamByName('Convenio').AsString := QrFrentes.FieldByName('Convenio').AsString;
      connection.QryBusca.ParamByName('Orden').AsString := Param_Frente;
      connection.QryBusca.ParamByName('Wbs').AsString := ActividadesIguales.FieldValues['sWbsAnterior'];
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then
      begin
        connection.QryBusca2.Active := False;
        connection.QryBusca2.SQL.Clear;
        connection.QryBusca2.SQL.Add('select sNumeroActividad, mDescripcion from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden and sWbs =:Wbs and sTipoActividad ="Paquete" ');
        connection.QryBusca2.ParamByName('Contrato').AsString := param_global_contrato;
        connection.QryBusca2.ParamByName('Convenio').AsString := QrFrentes.FieldByName('Convenio').AsString;
        connection.QryBusca2.ParamByName('Orden').AsString := Param_Frente;
        connection.QryBusca2.ParamByName('Wbs').AsString := connection.QryBusca.FieldValues['sWbsAnterior'];
        connection.QryBusca2.Open;
      end;
    end;
  end;
end;

procedure TfrmBitacoraDepartamental_2.QryBitacoraCalcFields(
  DataSet: TDataSet);
begin
  try
    QryBitacoradTotalMN.Value := QryBitacoradCantidad.Value * QryBitacoradVentaMN.Value;
    QryBitacorasIdGerencial.Value := QryBitacora.FieldByName('iNumeroGerencial').AsString;

    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('select SUM(b.dAvance) as Avance, o.dPonderado from bitacoradeactividades b '+
              'inner join actividadesxorden o on (o.sContrato = b.sContrato and o.sNumeroOrden = b.sNumeroOrden and b.sIdConvenio = o.sIdConvenio and b.sWbs = o.sWbs and o.sTipoActividad = "Actividad") '+
              'where b.sContrato =:Contrato ' +
              'and b.dIdFecha =:Fecha and b.sNumeroOrden =:Orden and b.sNumeroActividad =:Actividad and b.sWbs =:Wbs group by b.sContrato ');
    connection.QryBusca.ParamByName('contrato').AsString  := param_global_contrato;
    connection.QryBusca.ParamByName('Fecha').AsDate       := tdIdFecha.DateTime;
    connection.QryBusca.ParamByName('Orden').AsString     := QryBitacora.FieldByName('sNumeroOrden').AsString;
    connection.QryBusca.ParamByName('Wbs').AsString       := QryBitacora.FieldByName('sWbs').AsString;
    connection.QryBusca.ParamByName('Actividad').AsString := QryBitacora.FieldByName('sNumeroActividad').AsString;
    connection.QryBusca.Open;

    if connection.QryBusca.RecordCount > 0 then
    begin
       QryBitacoradAvancePartida.Value := connection.QryBusca.FieldByName('Avance').AsFloat;
       QryBitacoradPonderado.Value     := connection.QryBusca.FieldByName('dPonderado').AsFloat;
    end
    else
       QryBitacoradAvancePartida.Value := 0;



  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Registros de Volumenes de Obra y Notas', 'Al cambiar de registro', 0);
    end;
  end;
end;

procedure TfrmBitacoraDepartamental_2.qryNotasGerencialAfterScroll(
  DataSet: TDataSet);
begin
    if qryNotasGerencial.RecordCount > 0 then
    begin
        if  (sOpcion = '') and (lMostrarNotas) then
        begin
            tsHoraInicio.Text  := QryNOtasGerencial.FieldValues['sHoraInicio'];
            tsHoraFinal.Text   := QryNOtasGerencial.FieldValues['sHoraFinal'];
        end;

        bitacorademateriales.Active := False;
        bitacorademateriales.Params.ParamByName('contrato').DataType := ftString;
        bitacorademateriales.Params.ParamByName('contrato').Value := param_global_contrato;
        bitacorademateriales.Params.ParamByName('fecha').DataType := ftDate;
        bitacorademateriales.Params.ParamByName('fecha').Value    := tdIdFecha.Date;
        bitacorademateriales.Params.ParamByName('Diario').DataType := ftInteger;
        bitacorademateriales.Params.ParamByName('Diario').Value := QryNOtasGerencial.FieldByName('iIdDiario').AsInteger;
        bitacorademateriales.Params.ParamByName('Wbs').DataType := ftString;
        bitacorademateriales.Params.ParamByName('Wbs').Value := QryBitacora.FieldByName('sWbs').AsString;
        bitacorademateriales.Open;

        tmDescripcion.Text          := QryNOtasGerencial.FieldValues['mDescripcion'];
        tssIdClasificacion.KeyValue := QryNOtasGerencial.FieldValues['sIdClasificacion'];
        tsTipoAct.KeyValue          := QryNOtasGerencial.FieldValues['sTipoObra'];
        tsTipoAct.KeyValue          := QryNOtasGerencial.FieldValues['sTipoObra'];

        tsIdPernocta.KeyValue       := qryNotasGerencial.FieldByName( 'sIdPernocta' ).AsString;
        tsIdPlataforma.KeyValue     := qryNotasGerencial.FieldByName( 'sIdPlataforma' ).AsString;

        if tsTipoAct.KeyValue = 'FO' then
        begin
            tsIdPlataforma.Visible := True;
            tsIdPernocta.Visible   := True;
            Label6.Visible         := True;
            Label8.Visible         := True;
        end;

    end;
end;

procedure TfrmBitacoraDepartamental_2.qryNotasGerencialBeforeDelete(
  DataSet: TDataSet);
begin
   abort;
end;

procedure TfrmBitacoraDepartamental_2.qryNotasGerencialBeforeEdit(
  DataSet: TDataSet);
begin
    if ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
    begin
        MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
        abort;
    end;
end;

procedure TfrmBitacoraDepartamental_2.QrFrentesAfterScroll(DataSet: TDataSet);
begin
    QryBitacora.Filtered := False;
    QryBitacora.Filter   := 'sNumeroOrden = '+ QuotedStr(QrFrentes.FieldByName('sNumeroOrden').AsString);
    QryBitacora.Filtered := True;

    //Titulo de la ventana informacion del folio..
    if QrFrentes.FieldByName('Convenio').AsString = '-1' then
      lblFolio.Caption := '[FOLIO: '+ QrFrentes.FieldByName('sIdFolio').AsString + ']'
    else
       lblFolio.Caption := '[FOLIO: '+ QrFrentes.FieldByName('sIdFolio').AsString + ']';
    LeerDatasets(QrFrentes.FieldByName('sContrato').AsString,QrFrentes.FieldByName('sNumeroOrden').AsString);


    QryPartidasEfectivas_ADM.Active := False;
    QryPartidasEfectivas_ADM.Params.ParamByName('contrato').DataType := ftString;
    QryPartidasEfectivas_ADM.Params.ParamByName('contrato').Value := param_global_contrato;
    QryPartidasEfectivas_ADM.Params.ParamByName('convenio').DataType := ftString;
    QryPartidasEfectivas_ADM.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
    QryPartidasEfectivas_ADM.Params.ParamByName('Orden').DataType := ftString;
    QryPartidasEfectivas_ADM.Params.ParamByName('Orden').Value    := QrFrentes.FieldValues['sNumeroOrden'];
    QryPartidasEfectivas_ADM.Params.ParamByName('fecha').DataType := ftDate;
    QryPartidasEfectivas_ADM.Params.ParamByName('fecha').Value    := global_fecha;
    QryPartidasEfectivas_ADM.Open;

    QryPartidasEfectivas_PU.Active := False;
    QryPartidasEfectivas_PU.Params.ParamByName('contrato').DataType := ftString;
    QryPartidasEfectivas_PU.Params.ParamByName('contrato').Value := param_global_contrato;
    QryPartidasEfectivas_PU.Params.ParamByName('convenio').DataType := ftString;
    QryPartidasEfectivas_PU.Params.ParamByName('Convenio').Value := QrFrentes.FieldByName('Convenio').AsString;
    QryPartidasEfectivas_PU.Params.ParamByName('Orden').DataType := ftString;
    QryPartidasEfectivas_PU.Params.ParamByName('Orden').Value := QrFrentes.FieldValues['sNumeroOrden'];
    QryPartidasEfectivas_PU.Open;

    Minimo_Id;
    tdEditaAvance.Value := zQryDatos.FieldByName('dAvanceActual').AsFloat;

end;

procedure TfrmBitacoraDepartamental_2.QryBitacoraAfterScroll(
  DataSet: TDataSet);
var
   sCondicion, sNumeroActividad, SQLAnt : string;
begin
    if QryBitacora.RecordCount > 0 then
    begin
        {Activar o no check box..}
        if QryBitacora.FieldValues['lImprime'] = 'Si' then
        begin
            chkImprime.Checked := True;
            chkImprime.Font.Color := clBlack;
            chkImprime.Font.Style := [];
        end
        else
        begin
            chkImprime.Checked := False;
            chkImprime.Font.Color := clRed;
            chkImprime.Font.Style := [fsBold];
        end;

        cxComboGerencial.ItemIndex := 2;

        if (QryBitacora.FieldByName('iNumeroGerencial').AsInteger = 1) or (QryBitacora.FieldByName('iNumeroGerencial').AsInteger = 3) then
           cxComboGerencial.ItemIndex := 0;

        if QryBitacora.FieldByName('iNumeroGerencial').AsInteger = 1 then
           cxComboGerencial.ItemIndex := 1;

        if QryBitacora.FieldByName('iNumeroGerencial').AsInteger = 0 then
           cxComboGerencial.ItemIndex := 2;

        if QryBitacora.FieldValues['lCancelada'] = 'Si' then
          chkCancelada.Checked := True
        else
          chkCancelada.Checked := False;

        nxNumGerencial.Text := QryBitacora.FieldByName('iNumeroGerencial').AsString;   

        tdAvance.Value              := 0;
        tsHoraInicio.Text           := QryBitacora.FieldValues['sHoraInicio'];
        tsHoraFinal.Text            := QryBitacora.FieldValues['sHoraFinal'];
        tsIdTipoMovimiento.KeyValue := QryBitacora.FieldValues['sIdTipoMovimiento'];
        tsNumeroActividad.KeyValue  := QryBitacora.FieldValues['sNumeroActividad'];

        tsNumeroActividad_ADM.KeyValue := QryBitacora.FieldValues['sNumeroActividad'];
        tsNumeroActividad_PU.KeyValue  := QryBitacora.FieldValues['sNumeroActividad_ADM'];

        if tsNumeroActividad.KeyValue <> Null then
        begin
           sNumeroActividad        := tsNumeroActividad.Text;
           tsNumeroActividad_PU.Color  := global_color_pantalla;
           tsNumeroActividad_ADM.Color := global_color_pantalla;
           tsNumeroActividad.Color     := global_color_salida;
        end
        else
        begin
            sNumeroActividad := tsNumeroActividad_PU.Text;
            tsNumeroActividad_PU.Color  := global_color_salida;
            tsNumeroActividad_ADM.Color := global_color_salida;
            tsNumeroActividad.Color     := global_color_pantalla;
        end;

        rxAvances.Active := True;
        rxAvances.EmptyTable;

        tdCantidad.Text := FormatFloat('0.00000', 0);

        ActividadesIguales.Active := False;
        ActividadesIguales.Params.ParamByName('contrato').DataType  := ftString;
        ActividadesIguales.Params.ParamByName('contrato').Value     := param_global_contrato;
        ActividadesIguales.Params.ParamByName('convenio').DataType  := ftString;
        ActividadesIguales.Params.ParamByName('Convenio').Value     := QrFrentes.FieldByName('Convenio').AsString;
        ActividadesIguales.Params.ParamByName('orden').DataType     := ftString;
        ActividadesIguales.Params.ParamByName('orden').Value        := QrFrentes.FieldValues['sNumeroOrden'];;
        ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
        ActividadesIguales.Params.ParamByName('actividad').Value    := QryBitacora.FieldValues['sNumeroActividad'];
        ActividadesIguales.ParamByName('turno').AsString            := global_turno_reporte;
        ActividadesIguales.ParamByName('fecha').AsDate              := tdIdFecha.Date;
        ActividadesIguales.Open;

        ActividadesIguales.Locate('sWbs', QryBitacora.FieldValues['sWbs'], [loPartialKey]);

        if QryBitacora.FieldValues['sIdTipoMovimiento'] <> 'N' then
           sCondicion := ' and sWbs =:Wbs '
        else
           sCondicion := '';

        QryNotasGerencial.Active := false;
        QryNotasGerencial.SQL.Clear;
        QryNotasGerencial.SQL.Add('select iIdDiario, mDescripcion, sHoraInicio, sHoraFinal, sConceptoGerencial, sIdClasificacion, dCantidad, format(dAvance,2) as dAvance, sWbs, '+
                                  'concat(sHoraInicio," - ",sHoraFinal) as Nota, lImprime,iHermano,iIdTarea,iIdActividad, eTipoActividad, sTipoObra, sIdPernocta, sIdPlataforma, dCantidadAjuste, dCantidadAjusteNext, dCantidadAjusteNext2, dRestaEspacio '+
                                  'from bitacoradeactividades where sContrato =:Contrato and sIdConvenio =:Convenio '+
                                  'and dIdFecha =:Fecha && snumeroorden = :folio && sNumeroActividad =:actividad && sIdTipoMovimiento = "ED" order by sHoraInicio ');
        QryNotasGerencial.ParamByName('Contrato').AsString  := param_global_Contrato;
        QryNotasGerencial.ParamByName('Convenio').AsString  := QrFrentes.FieldByName('Convenio').AsString;
        QryNotasGerencial.ParamByName('Fecha').AsDate       := tdIdFecha.Date;
        QryNotasGerencial.ParamByName('folio').asstring     := QrFrentes.FieldByName('snumeroorden').asstring;
        QryNotasGerencial.ParamByName('actividad').asString := QryBitacora.fieldbyname('sNumeroActividad').AsString;
        QryNotasGerencial.Open;


        bitacorademateriales.Active := False;
        bitacorademateriales.Params.ParamByName('contrato').DataType := ftString;
        bitacorademateriales.Params.ParamByName('contrato').Value := param_global_contrato;
        bitacorademateriales.Params.ParamByName('fecha').DataType := ftDate;
        bitacorademateriales.Params.ParamByName('fecha').Value    := tdIdFecha.Date;
        bitacorademateriales.Params.ParamByName('Diario').DataType := ftInteger;
        bitacorademateriales.Params.ParamByName('Diario').Value := QryNOtasGerencial.FieldByName('iIdDiario').AsInteger;
        bitacorademateriales.Params.ParamByName('Wbs').DataType := ftString;
        bitacorademateriales.Params.ParamByName('Wbs').Value := QryBitacora.FieldByName('sWbs').AsString;
        bitacorademateriales.Open;
    end;
end;

procedure TfrmBitacoraDepartamental_2.rDiarioGetValue(
  const VarName: string; var Value: Variant);
begin
  if CompareText(VarName, 'ORDEN') = 0 then
    Value := 'DE LA ORDEN DE TRABAJO ' + Param_Frente;

  if CompareText(VarName, 'FECHA_INICIO') = 0 then
    Value := tdIdFecha.Date;

  if CompareText(VarName, 'FECHA_FINAL') = 0 then
    Value := tdIdFecha.Date;

  if CompareText(VarName, 'SEMANA') = 0 then
    Value := WeekOfTheMonth(tdIdFecha.Date);

  if CompareText(VarName, 'DIAS_SEMANA') = 0 then
    Value := '1';

  if CompareText(VarName, 'MONEDA') = 0 then
    Value := 'M.N.';


  If CompareText(VarName, 'SUPERINTENDENTE') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
          Value := sSuperIntendente
      Else
          Value := sSuperIntendentePatio ;

  If CompareText(VarName, 'SUPERVISOR') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
          Value := sSupervisor
      Else
          Value := sSupervisorPatio ;

  If CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
          Value := sSupervisorTierra
      Else
          Value := sResidente ;

  If CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
      begin
          if pos('#', sPuestoSuperIntendente) > 0 then
             Value := copy(sPuestoSuperIntendente,0, pos('#', sPuestoSuperIntendente)-1) +#13+ copy(sPuestoSuperIntendente,pos('#', sPuestoSuperIntendente)+1, length(sPuestoSuperIntendente))
          else
             Value := sPuestoSuperIntendente
      end
      Else
          Value := sPuestoSuperIntendentePatio ;

  If CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
      begin
          if pos('#', sPuestoSupervisor) > 0 then
             Value := copy(sPuestoSupervisor,0, pos('#', sPuestoSupervisor)-1) +#13+ copy(sPuestoSupervisor,pos('#', sPuestoSupervisor)+1, length(sPuestoSupervisor))
          else
             Value := sPuestoSupervisor
      end
      Else
          Value := sPuestoSupervisorPatio ;

  If CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
      begin
          if pos('#', sPuestoSupervisorTierra) > 0 then
             Value := copy(sPuestoSupervisorTierra,0, pos('#', sPuestoSupervisorTierra)-1) +#13+ copy(sPuestoSupervisorTierra,pos('#', sPuestoSupervisorTierra)+1, length(sPuestoSupervisorTierra))
          else
             Value := sPuestoSupervisorTierra
      end
      Else
          Value := sPuestoResidente ;


  If CompareText(VarName, 'DESCRIPCION_ORDEN') = 0 then
      Value := mDescripcionOrden  ;

  If CompareText(VarName, 'Oficio_Orden') = 0 then
      Value := sFolio  ;

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
    If CompareText(VarName, 'SUMPERSONAL') = 0 then
      Value := SumaPersonal ;
  If CompareText(VarName, 'SUMEQUIPOS') = 0 then
      Value := SumaEquipos ;

end;

procedure Tfrmbitacoradepartamental_2.CopiaMemo(Sender: TObject);
begin
  tmDescripcion.Text := (Sender as tMemo).Text;
end;

procedure TfrmBitacoraDepartamental_2.CosolidarNotasGerenciales1Click(
  Sender: TObject);
var
   sConcepto   : string;
   Diario      : integer;
   Wbs         : string;
   sComentario, sHoraInicio, sHoraFinal : string;
begin
    if QryBitacora.RecordCount > 0 then
    begin
        sHoraInicio := '';
        sHoraFinal  := '';

        //Se Toma la primera nota o comentario del Reporte Diario
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select min(iIdDiario) as iIdDiario, mDescripcion '+
                                    'from bitacoradeactividades where sContrato =:Contrato '+
                                    'and dIdFecha =:Fecha and sNumeroOrden =:Orden and sIdTipoMovimiento = "N" group by sContrato ');
        connection.zCommand.ParamByName('Contrato').AsString := param_global_contrato;
        connection.zCommand.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
        connection.zCommand.ParamByName('Orden').AsString    := Param_Frente;
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
        begin
            Diario      := connection.zCommand.FieldValues['iIdDiario'];
            sComentario := connection.zCommand.FieldValues['mDescripcion'];
        end
        else
        begin
            messageDLG('No se puede Continuar, No se encontro una Nota General del Reporte Diario.', mtWarning, [mbOk], 0);
            exit;
        end;

        //Consultamos la fecha de inicio del dia anterior del Gerencial que continua al siguiente dia...
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select sHoraInicio, sHoraFinal from bitacoradeactividades where sContrato =:Contrato '+
                                    'and dIdFecha =:Fecha and sNumeroOrden =:Orden and sIdTipoMovimiento <> "G" and sHoraInicio > sHoraFinal group by sHoraInicio order by sHoraInicio DESC ');
        connection.zCommand.ParamByName('Contrato').AsString := param_global_contrato;
        connection.zCommand.ParamByName('Fecha').AsDate      := tdIdFecha.Date - 1;
        connection.zCommand.ParamByName('Orden').AsString    := Param_Frente;
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
        begin
            sHoraInicio := connection.zCommand.FieldValues['sHoraInicio'];
            sHoraFinal  := connection.zCommand.FieldValues['sHoraFinal'];
        end;

        //Tomamos los horarios maximos y minimos del reporte
        if sHoraInicio = '' then
        begin
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select MIN(sHoraInicio) as sHoraInicio from bitacoradeactividades where sContrato =:Contrato ' +
                                        'and dIdFecha =:fecha and sNumeroOrden =:orden and sIdTipoMovimiento <> "G" group by sContrato');
            connection.zCommand.ParamByName('Contrato').AsString := param_global_contrato;
            connection.zCommand.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
            connection.zCommand.ParamByName('Orden').AsString    := Param_Frente;
            connection.zCommand.Open;

            if connection.zCommand.RecordCount > 0 then
               sHoraFinal  := connection.zCommand.FieldValues['sHoraInicio'];
        end;

        //Consolidar las notas del gerencial a las notas del reporte diario
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select b.sWbs, b.sNumeroActividad, b.sHoraInicio, b.sHoraFinal, b.sConceptoGerencial, b.mDescripcion '+
                                    'from bitacoradeactividades b '+
                                    'left join actividadesxorden o on (o.sContrato = b.sContrato and o.sIdConvenio =:Convenio '+
                                    'and o.sNumeroOrden = b.sNumeroOrden and o.sWbs = b.sWbs and o.sTipoActividad = "Actividad") '+
                                    'where b.sContrato =:Contrato and b.dIdFecha >=:fechaI and b.dIdFecha <=:fechaF and b.sNumeroOrden =:Orden and b.sIdTipoMovimiento <> "G" group by b.sWbs order by o.iItemOrden, b.iIdDiario');
        connection.zCommand.ParamByName('Contrato').AsString := param_global_contrato;
        connection.zCommand.ParamByName('FechaI').AsDate     := tdIdFecha.Date - 1;
        connection.zCommand.ParamByName('FechaF').AsDate     := tdIdFecha.Date;
        connection.zCommand.ParamByName('Orden').AsString    := Param_Frente;
        connection.zCommand.ParamByName('Convenio').AsString := QrFrentes.FieldByName('Convenio').AsString;
        connection.zCommand.Open;

        sConcepto := '.l.';
        Wbs       := '.l.';
        tmDescripcion.Text := sComentario;
        tmDescripcion.Lines.Add(#13);
        while not connection.zCommand.Eof do
        begin
            if sConcepto <> connection.zCommand.FieldValues['sConceptoGerencial'] then
            begin
               sConcepto := connection.zCommand.FieldValues['sConceptoGerencial'];
            end;

            if Wbs <> connection.zCommand.FieldValues['sWbs'] then
            begin
                if connection.zCommand.FieldValues['sWbs'] <> Null then
                begin
                    tmDescripcion.Lines.Add('PARTIDA '+ connection.zCommand.FieldValues['sNumeroActividad']);
                    Wbs := connection.zCommand.FieldValues['sWbs'];

                    //Consolidamos movimientos anteriores del dia anterior..
                    if sHoraInicio <> '' then
                    begin
                        connection.QryBusca2.Active := False;
                        connection.QryBusca2.SQL.Clear;
                        connection.QryBusca2.SQL.Add('select b.sWbs, b.sNumeroActividad, b.sHoraInicio, b.sHoraFinal, b.sConceptoGerencial, b.mDescripcion '+
                                                    'from bitacoradeactividades b '+
                                                    'where b.sContrato =:Contrato and b.dIdFecha =:fecha and b.sNumeroOrden =:Orden and b.sWbs =:Wbs ' +
                                                    'and b.sHoraFinal > "00:00" and b.sHoraFinal <=:Final and b.sIdTipoMovimiento = "G" order by b.sHoraInicio');
                        connection.QryBusca2.ParamByName('Contrato').AsString := param_global_contrato;
                        connection.QryBusca2.ParamByName('Fecha').AsDate      := tdIdFecha.Date - 1;
                        connection.QryBusca2.ParamByName('Orden').AsString    := Param_Frente;
                        connection.QryBusca2.ParamByName('Wbs').AsString      := Wbs;
                        connection.QryBusca2.ParamByName('Final').AsString    := sHoraFinal;
                        connection.QryBusca2.Open;

                        while not connection.QryBusca2.Eof do
                        begin
                            if connection.QryBusca2.FieldValues['sHoraInicio'] > sHoraFinal then
                                tmDescripcion.Lines.Add( '00:00'+ ' - ' +
                                      connection.QryBusca2.FieldValues['sHoraFinal'] + '  '+connection.QryBusca2.FieldValues['mDescripcion'])
                            else
                                tmDescripcion.Lines.Add(connection.QryBusca2.FieldValues['sHoraInicio']+ ' - ' +
                                      connection.QryBusca2.FieldValues['sHoraFinal'] + '  '+connection.QryBusca2.FieldValues['mDescripcion']);
                            tmDescripcion.Lines.Add('');
                            connection.QryBusca2.Next;
                        end;
                    end;

                    //Consolidamos movimientos actuales
                    connection.QryBusca2.Active := False;
                    connection.QryBusca2.SQL.Clear;
                    connection.QryBusca2.SQL.Add('select b.sWbs, b.sNumeroActividad, b.sHoraInicio, b.sHoraFinal, b.sConceptoGerencial, b.mDescripcion '+
                                                'from bitacoradeactividades b '+
                                                'where b.sContrato =:Contrato and b.dIdFecha =:fecha and b.sNumeroOrden =:Orden and b.sWbs =:Wbs ' +
                                                'and b.sHoraInicio >=:Inicio and b.sHoraInicio < "24:00" and b.sIdTipoMovimiento = "G" order by b.sHoraInicio');
                    connection.QryBusca2.ParamByName('Contrato').AsString := param_global_contrato;
                    connection.QryBusca2.ParamByName('Fecha').AsDate      := tdIdFecha.Date ;
                    connection.QryBusca2.ParamByName('Orden').AsString    := Param_Frente;
                    connection.QryBusca2.ParamByName('Wbs').AsString      := Wbs;
                    connection.QryBusca2.ParamByName('Inicio').AsString   := sHoraFinal;
                    connection.QryBusca2.Open;

                    while not connection.QryBusca2.Eof do
                    begin
                        if (connection.QryBusca2.FieldValues['sHoraFinal'] > '00:00') and (connection.QryBusca2.FieldValues['sHoraFinal'] <= sHoraFinal) then
                            tmDescripcion.Lines.Add(connection.QryBusca2.FieldValues['sHoraInicio']+ ' - ' +
                                   '24:00' + '  '+connection.QryBusca2.FieldValues['mDescripcion'])
                        else
                            tmDescripcion.Lines.Add(connection.QryBusca2.FieldValues['sHoraInicio']+ ' - ' +
                                   connection.QryBusca2.FieldValues['sHoraFinal'] + '  '+connection.QryBusca2.FieldValues['mDescripcion']);
                        tmDescripcion.Lines.Add('');
                        connection.QryBusca2.Next;
                    end;
                end;
            end;
            connection.zCommand.Next;
        end;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Update bitacoradeactividades set mDescripcion =:Descripcion '+
                                    'where sContrato =:Contrato and dIdFecha =:Fecha and iIdDiario =:Id ');
        connection.zCommand.ParamByName('Contrato').AsString  := param_global_contrato;
        connection.zCommand.ParamByName('Fecha').AsDate       := tdIdFecha.Date;
        connection.zCommand.ParamByName('Descripcion').AsMemo := tmDescripcion.Text;
        connection.zCommand.ParamByName('Id').AsInteger       := Diario;
        connection.zCommand.ExecSQL;

        QryBitacora.Refresh;
        messageDLG('Proceso Terminado con Exito!', mtInformation, [mbOk], 0);
    end;
end;

procedure TfrmBitacoraDepartamental_2.cxAbiertoClick(Sender: TObject);
begin
        chkEstatusHora.OnEnter(sender);
end;

procedure TfrmBitacoraDepartamental_2.cxAceptarClick(Sender: TObject);
begin
    if trim(tsPassword.Text) = '' then
    begin
        messageDLG('Favor de ingresar Password', mtInformation, [mbOk],0);
        tsPassword.SetFocus;
    end
    else
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select sPasswordEspecial from usuarios where sIdUsuario =:usuario ');
        connection.zCommand.ParamByName('usuario').AsString := global_usuario;
        connection.zCommand.Open;

        if connection.zCommand.RecordCount = 0 then
        begin
            messageDLG('Este usuario no cuenta con los provilegios para Modificar Avances!', mtInformation, [mbOk],0);
            panelPassword.Visible := False;
            exit;
        end;

        if tsPassword.Text = connection.zCommand.FieldByName('sPasswordEspecial').AsString then
        begin
            lAutentica := True;
            panelPassword.Visible := False;
            cxEdita.Caption := 'Actualizar';
            tdEditaAvance.SetFocus;
            tdEditaAvance.ReadOnly := False;
        end
        else
        begin
           messageDLG('Password incorrecto!', mtInformation, [mbOk],0);
           tsPassword.SetFocus;
        end;
    end;

end;

procedure TfrmBitacoraDepartamental_2.cxButton1Click(Sender: TObject);
var
  Form : TForm;
begin
  Form := TForm.Create( nil );
  Form.BorderStyle := bsDialog;
  Form.Width := 220;
  Form.Height := 140;
  Form.Position := poScreenCenter;

  grpTipoActividad.Parent := Form;
  grpTipoActividad.Align := alClient;
  grpTipoActividad.Visible := True;

  TiposDeActividad.cdValores.Append;

  Form.ShowModal;

  grpTipoActividad.Align := alNone;
  grpTipoActividad.Visible := False;
  grpTipoActividad.Width := 0;
  grpTipoActividad.Height := 0;
  grpTipoActividad.Left := 0;
  grpTipoActividad.Top := 0;

  TiposDeActividad.LinkWithLookup( cbbTipoActividad );

end;

procedure TfrmBitacoraDepartamental_2.cxButton2Click(Sender: TObject);
var
  zQuery : TZReadOnlyQuery;
begin

  zQuery := TZReadOnlyQuery.Create( nil );
  zQuery.Connection := connection.zConnection;
  zQuery.Active := False;
  zQuery.SQL.Text := 'select count( scontrato ) as iActividades '+
                     'from bitacoradeactividades '+
                     'where eTipoActividad = :tipo' ;
  zQuery.ParamByName( 'tipo' ).AsString := cbbTipoActividad.Text;
  zQuery.Open;

  if zQuery.FieldByName( 'iActividades' ).asinteger > 0 then
  begin
    MessageDlg( 'Ya hay actividades registradas con este tipo de actividad, no se puede eliminar', mtInformation, [ mbOK ], 0 );
    Exit;
  end;
  

  if MessageDlg( '¿Desea eliminar el tipo de actividad?', mtConfirmation, [ mbYes, mbCancel ], 0 ) = mrYes then
  begin
    TiposDeActividad.cdValores.Locate( 'TipoActividad', cbbTipoActividad.Text, [] );
    TiposDeActividad.cdValores.Delete;
    TiposDeActividad.AlterTable;
    TiposDeActividad.LinkWithLookup( cbbTipoActividad );
    cbbTipoActividad.Text := '';
  end;
end;

procedure TfrmBitacoraDepartamental_2.cxButton3Click(Sender: TObject);
var
  c : Integer;
begin

  try

    if Length( Trim( dbTipoActividad.Text ) ) = 0 then
      raise Exception.Create( 'No se ha especificado un valor' );


    if LowerCase( ( dbTipoActividad.Text ) ) = 'enum' then
      raise Exception.Create( 'Valor no valido' );

    for c := 0 to Length( dbTipoActividad.Text ) do
    begin

      if ( dbTipoActividad.Text[ c ] = ',' ) or ( dbTipoActividad.Text[ c ] = Char( 39 ) ) then
        raise Exception.Create( 'No se permiten caracteres especiales' );

    end;

    if MessageDlg( '¿Confirma crear el nuevo tipo de Actividad?', mtConfirmation, [ mbYes, mbCancel ], 0 ) =  mrYes then
    begin
      TiposDeActividad.cdValores.Post;
      TiposDeActividad.AlterTable;

      btnSalirTiposActividad.Click;
    end;

  except
    on e:Exception do
    begin
      MessageDlg( e.Message, mtInformation, [ mbOK ], 0 );
      Exit;

    end;

  end;

end;

procedure TfrmBitacoraDepartamental_2.cxCancelarClick(Sender: TObject);
begin
    PanelPassword.Visible := False;
end;

procedure TfrmBitacoraDepartamental_2.cxCerradoClick(Sender: TObject);
begin
    chkEstatusHora.OnEnter(sender);
end;

procedure TfrmBitacoraDepartamental_2.cxComboBox1PropertiesChange(
  Sender: TObject);
var
   sHora : string;
begin
   if cxComboGerencial.Text = 'GERENCIAL (05:00 - 17:00 HRS.)' then
   begin
      cxComboGerencial.Style.Color := $0072C5FC;
      tsHoraInicio.Text := '05:00';
      tsHoraFinal.Text  := '17:00';
      nxNumGerencial.Value := 2;
   end;

   if cxComboGerencial.Text = 'GERENCIAL (17:00 - 05:00 HRS.)' then
   begin
      cxComboGerencial.Style.Color := clGradientActiveCaption;
      tsHoraInicio.Text := '00:00';
      tsHoraFinal.Text  := '05:00';
      sHora := ObtenerHora + ':00';
      if sHora <= '05:00'  then
         nxNumGerencial.Value := 1
      else
      begin
         nxNumGerencial.Value := 3;
         tsHoraFinal.Text  := '24:00';
      end;
   end;

   if cxComboGerencial.Text = 'NO APLICA GERENCIAL' then
   begin
      cxComboGerencial.Style.Color := clWhite;
      tsHoraInicio.Text := '00:00';
      tsHoraFinal.Text  := '00:00';
      nxNumGerencial.Value := 0;
   end;

end;

procedure TfrmBitacoraDepartamental_2.cxEditaClick(Sender: TObject);
begin
    if ObtenerEstatusReporte(global_contrato, tdIdFecha.Date) <> 'Pendiente' then
    begin
        MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
        exit;
    end;      

    if cxEdita.Caption = 'Nuevo' then
    begin
       lAutentica := False;
       panelPassword.Visible := True;
       tsPassword.Text := '';
       tsPassword.SetFocus;
    end;

    if lAutentica then
    begin             
        if tdEditaAvance.Value <> 0 then
        begin
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('UPDATE bitacoradeactividades SET lConsideraAvance = "No" ' +
                               'where sContrato = :contrato and sNumeroOrden =:Orden And dIdFecha = :fecha ');
            connection.zCommand.Params.ParamByName('contrato').AsString  := param_Global_Contrato;
            connection.zCommand.Params.ParamByName('Orden').AsString     := QrFrentes.FieldByName('sNumeroOrden').AsString;
            connection.zCommand.Params.ParamByName('Fecha').AsDate       := tdIdFecha.Date;
            connection.zCommand.ExecSQL;

            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('UPDATE bitacoradeactividades SET dAvanceActual := :Avance ' +
                               'where sContrato = :contrato and sNumeroOrden =:Orden And dIdFecha = :fecha And iIdDiario = :diario  ');
            connection.zCommand.Params.ParamByName('contrato').AsString  := param_Global_Contrato;
            connection.zCommand.Params.ParamByName('Orden').AsString     := QrFrentes.FieldByName('sNumeroOrden').AsString;
            connection.zCommand.Params.ParamByName('Fecha').AsDate       := tdIdFecha.Date;
            connection.zCommand.Params.ParamByName('diario').AsInteger   := zQryDatos.FieldByName('iIdDiario').AsInteger;
            connection.zCommand.Params.ParamByName('Avance').AsFloat     := tdEditaAvance.Value;
            connection.zCommand.ExecSQL;
        end;

        if tdEditaAvance.Value = 0 then
        begin
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('UPDATE bitacoradeactividades SET lConsideraAvance = "Si" ' +
                               'where sContrato = :contrato and sNumeroOrden =:Orden And dIdFecha = :fecha ');
            connection.zCommand.Params.ParamByName('contrato').AsString  := param_Global_Contrato;
            connection.zCommand.Params.ParamByName('Orden').AsString     := QrFrentes.FieldByName('sNumeroOrden').AsString;
            connection.zCommand.Params.ParamByName('Fecha').AsDate       := tdIdFecha.Date;
            connection.zCommand.ExecSQL;

            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('UPDATE bitacoradeactividades SET dAvanceActual := :Avance ' +
                               'where sContrato = :contrato and sNumeroOrden =:Orden And dIdFecha = :fecha And iIdDiario = :diario  ');
            connection.zCommand.Params.ParamByName('contrato').AsString  := param_Global_Contrato;
            connection.zCommand.Params.ParamByName('Orden').AsString     := QrFrentes.FieldByName('sNumeroOrden').AsString;
            connection.zCommand.Params.ParamByName('Fecha').AsDate       := tdIdFecha.Date;
            connection.zCommand.Params.ParamByName('diario').AsInteger   := zQryDatos.FieldByName('iIdDiario').AsInteger;
            connection.zCommand.Params.ParamByName('Avance').AsFloat     := tdEditaAvance.Value;
            connection.zCommand.ExecSQL;
        end;

        cxEdita.Caption := 'Nuevo';
        tdEditaAvance.ReadOnly := True;
        tdAvanceGlobal.Value := AvanceFolio;

    end;
end;

procedure TfrmBitacoraDepartamental_2.dAvanceStylesGetContentStyle(
  Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
  AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
begin
//    AStyle := cxNormal;
//
//    AItem:= (Sender as TcxGridDBTableView).GetColumnByFieldName('sIdGerencial');
//    if VarToStr(ARecord.Values[AItem.Index])='2' then
//       AStyle:=cxDiurno;
//
//    if VarToStr(ARecord.Values[AItem.Index])='1' then
//       AStyle:=cxNocturno;
//
//    if VarToStr(ARecord.Values[AItem.Index])='3' then
//       AStyle:=cxAlCorte;
end;

procedure TfrmBitacoraDepartamental_2.GrdOrdenCellClick(Column: TColumn);
begin
     if QryBitacora.RecordCount > 0 then
         tdAvanceGlobal.Value := AvanceFolio;
end;

procedure TfrmBitacoraDepartamental_2.GrdOrdenDblClick(Sender: TObject);
begin

  //frmBarra1btnAddClick(sender);
end;

procedure TfrmBitacoraDepartamental_2.GrdOrdenDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if GrdOrden.DataSource.DataSet.FieldByName('sidtipomovimiento').AsString = 'E' then
  begin
    GrdOrden.Canvas.Font.Color:=esColor(12);
    GrdOrden.Canvas.Font.Style := [fsBold];
    GrdOrden.Canvas.Brush.Color :=  $00E9CFB6;
  end;
  if TipoFolio = 'T' then
  begin
    GrdOrden.Canvas.Font.Color:= $00011421;
    GrdOrden.Canvas.Font.Style := [fsBold];
    GrdOrden.Canvas.Brush.Color :=  $001FA3FA;
    
  end;
  if TipoFolio = 'S' then
  begin
    GrdOrden.Canvas.Font.Color:= $00004646;
    GrdOrden.Canvas.Font.Style := [fsBold];
    GrdOrden.Canvas.Brush.Color :=  $0000D9D9;
  end;

  //si esta terminado se colorea en rojo los terminados


  if GrdOrden.DataSource.DataSet.FieldByName('cIdStatus').AsString = 'T' then
  begin
    GrdOrden.Canvas.Font.Color:= $00011421;
    GrdOrden.Canvas.Font.Style := [fsBold];
    GrdOrden.Canvas.Brush.Color :=  $001FA3FA;
  end;

  GrdOrden.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmBitacoraDepartamental_2.DesagruparHorarios1Click(Sender: TObject);
begin
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('UPDATE bitacoradepersonal SET sHoraInicioG = sHoraInicio, sHoraFinalG = sHoraFinal ' +
                                'where sContrato = :contrato And dIdFecha = :fecha And sWbs =:Wbs ');
    Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
    Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
    Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
    Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
    Connection.zCommand.Params.ParamByName('Wbs').AsString      := qryNotasGerencial.FieldValues['sWbs'];
    connection.zCommand.ExecSQL;

    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('UPDATE bitacoradeequipos SET sHoraInicioG = sHoraInicio, sHoraFinalG = sHoraFinal ' +
                                'where sContrato = :contrato And dIdFecha = :fecha And sWbs =:Wbs ');
    Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
    Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
    Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
    Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
    Connection.zCommand.Params.ParamByName('Wbs').AsString      := qryNotasGerencial.FieldValues['sWbs'];
    connection.zCommand.ExecSQL;

    messageDLG('Terminado!', mtInformation, [mbOK], 0);
end;

procedure TfrmBitacoraDepartamental_2.ds_bitacoradeequiposDataChange(
  Sender: TObject; Field: TField);
begin

end;

{_______________________________________________________________________________
 FUNCION PARA ACTUALIZAR EL ID DIARIO DE PERSONAL, EQUIPO Y PERNOCTA CARGADO A LA PARTIDA
--------------------------------------------------------------------------------}

procedure TfrmBitacoraDepartamental_2.ActividadesIgualesCalcFields(
  DataSet: TDataSet);
begin
     if ActividadesIguales.FieldByName('dInstalado').AsFloat = 0 then
        ActividadesIguales.FieldValues['Avance'] := 0
     else
     begin
          if (ActividadesIguales.FieldByName('dInstalado').AsFloat > 0) and (ActividadesIguales.FieldByName('dInstalado').AsFloat  <= actividadesiguales.FieldByName('dCantidad').AsFloat) then
          begin
              if actividadesiguales.FieldValues['dCantidad'] > 0 then
                 ActividadesIguales.FieldValues['Avance'] := (ActividadesIguales.FieldByName('dInstalado').AsFloat / actividadesiguales.FieldByName('dCantidad').AsFloat) * 100;
          end
          else
             ActividadesIguales.FieldValues['Avance'] := 100;
     end;
end;

procedure Tfrmbitacoradepartamental_2.ActualizaIdDiario(dParamContrato: string; dParamFecha: TDate; dParamIdDiario, dParamIdDiarioOld: Integer);
var
  Q_BuscaId: TZReadOnlyQuery;
begin
  Q_BuscaId := TZReadOnlyQuery.Create(self);
  Q_BuscaId.Connection := connection.zConnection;

    {Actualiza IdDiario Bitacora de Personal}
  Q_BuscaId.Active := False;
  Q_BuscaId.SQL.Clear;
  Q_BuscaId.SQL.Add('Update bitacoradepersonal set iIdDiario =:Id where sContrato =:Contrato and dIdFecha =:Fecha and iIdDiario =:IdOld');
  Q_BuscaId.ParamByName('Contrato').AsString := dParamContrato;
  Q_BuscaId.ParamByName('Fecha').AsDate := dParamFecha;
  Q_BuscaId.ParamByName('Id').AsInteger := dParamIdDiario;
  Q_BuscaId.ParamByName('IdOld').AsInteger := dParamIdDiarioOld;
  Q_BuscaId.ExecSQL;

    {Actualiza IdDiario Bitacora de Personal}
  Q_BuscaId.Active := False;
  Q_BuscaId.SQL.Clear;
  Q_BuscaId.SQL.Add('Update bitacoradeequipos set iIdDiario =:Id where sContrato =:Contrato and dIdFecha =:Fecha and iIdDiario =:IdOld');
  Q_BuscaId.ParamByName('Contrato').AsString := dParamContrato;
  Q_BuscaId.ParamByName('Fecha').AsDate := dParamFecha;
  Q_BuscaId.ParamByName('Id').AsInteger := dParamIdDiario;
  Q_BuscaId.ParamByName('IdOld').AsInteger := dParamIdDiarioOld;
  Q_BuscaId.ExecSQL;

    {Actualiza IdDiario Bitacora de Materiales}
  Q_BuscaId.Active := False;
  Q_BuscaId.SQL.Clear;
  Q_BuscaId.SQL.Add('Update bitacorademateriales set iIdDiario =:Id where sContrato =:Contrato and dIdFecha =:Fecha and iIdDiario =:IdOld');
  Q_BuscaId.ParamByName('Contrato').AsString := dParamContrato;
  Q_BuscaId.ParamByName('Fecha').AsDate := dParamFecha;
  Q_BuscaId.ParamByName('Id').AsInteger := dParamIdDiario;
  Q_BuscaId.ParamByName('IdOld').AsInteger := dParamIdDiarioOld;
  Q_BuscaId.ExecSQL;

     {Actualiza IdDiario Bitacora de Pernocta auxiliar..}
  Q_BuscaId.Active := False;
  Q_BuscaId.SQL.Clear;
  Q_BuscaId.SQL.Add('Update bitacoradepernocta_aux set iIdDiario =:Id where sContrato =:Contrato and dIdFecha =:Fecha and iIdDiario =:IdOld');
  Q_BuscaId.ParamByName('Contrato').AsString := dParamContrato;
  Q_BuscaId.ParamByName('Fecha').AsDate := dParamFecha;
  Q_BuscaId.ParamByName('Id').AsInteger := dParamIdDiario;
  Q_BuscaId.ParamByName('IdOld').AsInteger := dParamIdDiarioOld;
  Q_BuscaId.ExecSQL;       

   {Actualiza IdDiario Bitacora de Movimeinto de Gerencial..}
  Q_BuscaId.Active := False;
  Q_BuscaId.SQL.Clear;
  Q_BuscaId.SQL.Add('Update bitacoradeactividades set iIdDiarioNota =:Id where sContrato =:Contrato and dIdFecha =:Fecha and iIdDiarioNota =:IdOld');
  Q_BuscaId.ParamByName('Contrato').AsString := dParamContrato;
  Q_BuscaId.ParamByName('Fecha').AsDate := dParamFecha;
  Q_BuscaId.ParamByName('Id').AsInteger := dParamIdDiario;
  Q_BuscaId.ParamByName('IdOld').AsInteger := dParamIdDiarioOld;
  Q_BuscaId.ExecSQL;
end;

procedure Tfrmbitacoradepartamental_2.OrdenarNotas(sParamOrden: string);
var
   idAuxiliar, idAuxiliar2 : integer;
   SavePlace   : TBookmark;
begin
    if qryNotasGerencial.RecordCount > 0 then
    begin
        if sParamOrden = 'Arriba' then
        begin
            idAuxiliar2 := QryNotasGerencial.FieldValues['iIdDiario'];
            QryNotasGerencial.Prior;

            idAuxiliar  := QryNotasGerencial.FieldValues['iIdDiario'];
            QryNotasGerencial.Next;
        end;

        if sParamOrden = 'Abajo' then
        begin
            idAuxiliar2 := QryNotasGerencial.FieldValues['iIdDiario'];
            QryNotasGerencial.Next;

            idAuxiliar  := QryNotasGerencial.FieldValues['iIdDiario'];
            QryNotasGerencial.Prior;
        end;
        //Colocamos un id mayor para evitar duplicidad..
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE bitacoradeactividades SET iIdDiario = :DiarioNuevo ' +
                                    'where sContrato = :contrato And dIdFecha = :fecha And iIdDiario = :diario ');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := idAuxiliar2;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').DataType := ftInteger;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').value    := idAuxiliar + 500;
        connection.zCommand.ExecSQL;

        //Ahora actualizamos el item mayor
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE bitacoradeactividades SET iIdDiario = :DiarioNuevo ' +
                                    'where sContrato = :contrato And dIdFecha = :fecha And iIdDiario = :diario ');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := idAuxiliar ;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').DataType := ftInteger;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').value    := idAuxiliar2;
        connection.zCommand.ExecSQL;

         //Ahora actualizamos el item alterado
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE bitacoradeactividades SET iIdDiario = :DiarioNuevo ' +
                                    'where sContrato = :contrato And dIdFecha = :fecha And iIdDiario = :diario ');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := idAuxiliar + 500;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').DataType := ftInteger;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').value    := idAuxiliar;
        connection.zCommand.ExecSQL;

        if sParamOrden = 'Arriba' then
           QryNotasGerencial.Prior
        else
           QryNotasGerencial.Next;

        SavePlace := QryNotasGerencial.GetBookmark;
        QryNotasGerencial.Refresh;
        QryNotasGerencial.GotoBookmark(SavePlace);
        QryNotasGerencial.FreeBookmark(SavePlace);
    end;
end;


procedure TfrmBitacoradepartamental_2.ActualizaImprime;
var
    zQryActualiza : tzReadOnlyQuery;
begin
    zQryActualiza := tzReadOnlyQuery.Create(self);
    zQryActualiza.Connection := connection.zConnection;

    zQryActualiza.Active := False;
    zQryActualiza.SQL.Clear;
    zQryActualiza.SQL.Add('UPDATE bitacoradeactividades SET lImprime =:Imprime ' +
                      ' where sContrato =:contrato and sIdConvenio =:Convenio And dIdFecha = :fecha and sIdTurno =:Turno And sWbs =:Wbs and sIdTipoMovimiento = "G" ');
    zQryActualiza.Params.ParamByName('contrato').DataType := ftString;
    zQryActualiza.Params.ParamByName('contrato').value    := param_Global_Contrato;
    zQryActualiza.Params.ParamByName('Convenio').AsString := QrFrentes.FieldByName('Convenio').AsString;
    zQryActualiza.Params.ParamByName('fecha').DataType    := ftDate;
    zQryActualiza.Params.ParamByName('fecha').value       := tdIdFecha.Date;
    zQryActualiza.Params.ParamByName('Turno').DataType    := ftString;
    zQryActualiza.Params.ParamByName('Turno').value       := QryBitacora.FieldValues['sIdTurno'];
    zQryActualiza.Params.ParamByName('Wbs').DataType      := ftString;
    zQryActualiza.Params.ParamByName('Wbs').value         := QryBitacora.FieldValues['sWbs'];
    zQryActualiza.Params.ParamByName('Imprime').DataType  := ftString;
    if chkImprime.Checked then
       zQryActualiza.Params.ParamByName('Imprime').value  := 'Si'
    else
       zQryActualiza.Params.ParamByName('Imprime').value  := 'No';
    zQryActualiza.ExecSQL;

    zQryActualiza.Destroy;
end;

procedure TfrmBitacoraDepartamental_2.AdvNotaClick(Sender: TObject);
begin
   if advNota.Caption = 'Descripcion' then
   begin
      tmDescripcion.Color := $00E6FEFF;
      tmDescripcion.Text  := qryNotasGerencial.FieldByName('mDescripcion').AsString;
      global_nota := 'Nota';
      sSQlCadena  := 'mDescripcion ';
      tmDescripcion.Visible := True;
      gridmaterialesxpartida.Visible := False;
      advNota.Hint := 'Partidas de Anexo';
      advNota.Caption := 'Anexo';
   end
   else
   begin
      tmDescripcion.Visible := False;
      gridmaterialesxpartida.Visible := True;
      advNota.Hint := 'Descripcion Actividad';
      advNota.Caption := 'Descripcion';
   end;
end;

procedure TfrmBitacoraDepartamental_2.AgruparHorarios1Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
  sHoraInicio, sHoraFinal : string;

begin

  SavePlace := GridNotas.DataSource.DataSet.GetBookmark;
  try
    with GridNotas.DataSource.DataSet do
    begin
      sHoraInicio := qryNotasGerencial.FieldByName('sHoraInicio').AsString;
      sHoraFinal  := '00:00';
      for iGrid := 0 to GridNotas.SelectedRows.Count - 1 do
      begin
          GotoBookmark(pointer(GridNotas.SelectedRows.Items[iGrid]));

         if sHoraInicio > qryNotasGerencial.FieldByName('sHoraInicio').AsString then
            sHoraInicio := qryNotasGerencial.FieldByName('sHoraInicio').AsString;

         if sHoraFinal < qryNotasGerencial.FieldByName('sHoraFinal').AsString then
            sHoraFinal := qryNotasGerencial.FieldByName('sHoraFinal').AsString;
      end;
    end;

    with GridNotas.DataSource.DataSet do
    begin
      for iGrid := 0 to GridNotas.SelectedRows.Count - 1 do
      begin
          GotoBookmark(pointer(GridNotas.SelectedRows.Items[iGrid]));
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('UPDATE bitacoradepersonal SET sHoraInicioG = :inicio, sHoraFinalG = :final ' +
                                      'where sContrato = :contrato And dIdFecha = :fecha And sWbs =:Wbs and sHoraInicio =:hInicio and sHoraFinal =:hFinal ');
          Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
          Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
          Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
          Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
          Connection.zCommand.Params.ParamByName('inicio').DataType   := ftString;
          Connection.zCommand.Params.ParamByName('inicio').value      := sHoraInicio;
          Connection.zCommand.Params.ParamByName('final').DataType    := ftString;
          Connection.zCommand.Params.ParamByName('final').value       := sHoraFinal;
          Connection.zCommand.Params.ParamByName('Wbs').AsString      := qryNotasGerencial.FieldValues['sWbs'];
          Connection.zCommand.Params.ParamByName('hInicio').AsString  := qryNotasGerencial.FieldValues['sHoraInicio'];
          Connection.zCommand.Params.ParamByName('hFinal').AsString   := qryNotasGerencial.FieldValues['sHoraFinal'];
          connection.zCommand.ExecSQL;

          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('UPDATE bitacoradeequipos SET sHoraInicioG = :inicio, sHoraFinalG = :final ' +
                                      'where sContrato = :contrato And dIdFecha = :fecha And sWbs =:Wbs and sHoraInicio =:hInicio and sHoraFinal =:hFinal ');
          Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
          Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
          Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
          Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
          Connection.zCommand.Params.ParamByName('inicio').DataType   := ftString;
          Connection.zCommand.Params.ParamByName('inicio').value      := sHoraInicio;
          Connection.zCommand.Params.ParamByName('final').DataType    := ftString;
          Connection.zCommand.Params.ParamByName('final').value       := sHoraFinal;
          Connection.zCommand.Params.ParamByName('Wbs').AsString      := qryNotasGerencial.FieldValues['sWbs'];
          Connection.zCommand.Params.ParamByName('hInicio').AsString  := qryNotasGerencial.FieldValues['sHoraInicio'];
          Connection.zCommand.Params.ParamByName('hFinal').AsString   := qryNotasGerencial.FieldValues['sHoraFinal'];
          connection.zCommand.ExecSQL;
      end;
    end;
    messageDLG('Terminado!', mtInformation, [mbOK], 0);
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmBitacoradepartamental_2.formatoEncabezado;
begin
    Excel.Selection.MergeCells := False;
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.VerticalAlignment   := xlCenter;
    Excel.Selection.Font.Size := 10;
    Excel.Selection.Font.Bold := True;
    Excel.Selection.Font.Color:= clNavy;
    Excel.Selection.Font.Name := 'Arial';
end;

function TfrmBitacoraDepartamental_2.CompararFechas(fecha1: TDate; fecha2: TDate) : Boolean;
var
  sFechaInicio, sFechaTermino : String;
  DiaF, MesF, AnoF, DiaI, MesI, AnoI: Word;
  Dias, Meses, Anos: integer;
  dias1, dias2 : integer;
  factor : integer;
begin
  DecodeDate(fecha1, AnoI, MesI, DiaI);
  DecodeDate(fecha2, AnoF, MesF, DiaF);

  if AnoI <= AnoF then
  begin
    if MesI <= MesF then
    begin
      if DiaI <= DiaF then
      begin
        Result := True;
      end
      else
      begin
        Result := False;
      end;
    end;    
  end;
end;

procedure TfrmBitacoraDepartamental_2.CargaActividades;
begin
    //Muestra las partidas pertenecientes al folio...
    if qrFrentes.FieldByName('sNumeroOrden').AsString <> QryPartidasEfectivas.FieldByName('sNumeroOrden').AsString then
    begin
        QryPartidasEfectivas.Active := False;
        QryPartidasEfectivas.Params.ParamByName('contrato').DataType := ftString;
        QryPartidasEfectivas.Params.ParamByName('contrato').Value    := param_global_contrato;
        QryPartidasEfectivas.Params.ParamByName('convenio').DataType := ftString;
        QryPartidasEfectivas.Params.ParamByName('Convenio').Value    := QrFrentes.FieldByName('Convenio').AsString;
        QryPartidasEfectivas.Params.ParamByName('Orden').DataType    := ftString;
        QryPartidasEfectivas.Params.ParamByName('Orden').Value       := QrFrentes.FieldValues['sNumeroOrden'];
        QryPartidasEfectivas.Open;
    end;
end;

function TfrmBitacoraDepartamental_2.ValidaReporteDiario;
begin
    //Procedimiento para validar si está Autorizado el Reporte Diario.
    ReporteDiario.Active := False;
    ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
    ReporteDiario.Params.ParamByName('contrato').Value    := global_Contrato_Barco;
    ReporteDiario.Params.ParamByName('turno').DataType    := ftString;
    ReporteDiario.Params.ParamByName('turno').Value       := global_turno_reporte;
    ReporteDiario.Params.ParamByName('Fecha').DataType    := ftDate;
    ReporteDiario.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
    ReporteDiario.Params.ParamByName('Orden').DataType    := ftString;
    ReporteDiario.Params.ParamByName('Orden').Value       := param_global_contrato;
    ReporteDiario.Open;

    Result := True;
    if ReporteDiario.RecordCount > 0 then
    begin
        if ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' then
        begin
            MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
            Result := False;
        end
    end;
end;

function TfrmBitacoraDepartamental_2.AvanceActual(sParamActividad: string; sParamWbs: string): double;
var
   dAvanceAnterior,
   dAvance, dAvancePartida   :Extended;
begin
    {Procedimeinto para obetene el avance de la Partida.}
    AvanceMaximo.Active := False;
    AvanceMaximo.SQL.Clear;
    AvanceMaximo.SQL.Add('SELECT Sum(dAvance) as Avance FROM bitacoradeactividades where ' +
      'sContrato = :contrato and sIdConvenio =:Convenio and dIdFecha < :fecha And sNumeroOrden = :orden and ' +
      'sWbs = :wbs and sNumeroActividad = :Actividad Group By sContrato');
    AvanceMaximo.Params.ParamByName('Contrato').DataType  := ftString;
    AvanceMaximo.Params.ParamByName('Contrato').Value     := param_Global_Contrato;
    AvanceMaximo.Params.ParamByName('Convenio').AsString  := QrFrentes.FieldByName('Convenio').AsString;
    AvanceMaximo.Params.ParamByName('Fecha').DataType     := ftDate;
    AvanceMaximo.Params.ParamByName('Fecha').Value        := tdIdFecha.Date;
    AvanceMaximo.Params.ParamByName('orden').DataType     := ftString;
    AvanceMaximo.Params.ParamByName('orden').Value        := QrFrentes.FieldValues['sNumeroOrden'];
    AvanceMaximo.Params.ParamByName('wbs').DataType       := ftString;
    AvanceMaximo.Params.ParamByName('wbs').Value          := sParamWbs;
    AvanceMaximo.Params.ParamByName('Actividad').DataType := ftString;
    AvanceMaximo.Params.ParamByName('Actividad').Value    := sParamActividad;
    AvanceMaximo.Open;

    if AvanceMaximo.RecordCount > 0 then
       dAvanceAnterior := AvanceMaximo.FieldValues['Avance'];

    AvanceMaximo.Active := False;
    AvanceMaximo.SQL.Clear;
    AvanceMaximo.SQL.Add('SELECT Sum(dAvance) as Avance FROM bitacoradeactividades where ' +
      'sContrato = :contrato and sIdConvenio =:Convenio and dIdFecha = :fecha and sIdTurno < :Turno And ' +
      'sNumeroOrden = :orden and sWbs = :wbs and sNumeroActividad = :Actividad ' +
      'Group By sContrato');
    AvanceMaximo.Params.ParamByName('Contrato').DataType  := ftString;
    AvanceMaximo.Params.ParamByName('Contrato').Value     := param_Global_Contrato;
    AvanceMaximo.Params.ParamByName('Convenio').AsString  := QrFrentes.FieldByName('Convenio').AsString;
    AvanceMaximo.Params.ParamByName('Fecha').DataType     := ftDate;
    AvanceMaximo.Params.ParamByName('Fecha').Value        := tdIdFecha.Date;
    AvanceMaximo.Params.ParamByName('turno').DataType     := ftString;
    AvanceMaximo.Params.ParamByName('turno').Value        := global_turno_reporte;
    AvanceMaximo.Params.ParamByName('orden').DataType     := ftString;
    AvanceMaximo.Params.ParamByName('orden').Value        := QrFrentes.FieldValues['sNumeroOrden'];
    AvanceMaximo.Params.ParamByName('wbs').DataType       := ftString;
    AvanceMaximo.Params.ParamByName('wbs').Value          := sParamWbs;
    AvanceMaximo.Params.ParamByName('Actividad').DataType := ftString;
    AvanceMaximo.Params.ParamByName('Actividad').Value    := sParamActividad;
    AvanceMaximo.Open;
    if AvanceMaximo.RecordCount > 0 then
       dAvanceAnterior := dAvanceAnterior + AvanceMaximo.FieldValues['Avance'];

    if QryExistePartida.RecordCount > 0 then
       dAvancePartida := QryExistePartida.FieldValues['dAvance'];

    if Connection.configuracion.FieldValues['sAvanceBitacora'] = 'Volumen' then
    begin
        dAvance := (tdCantidad.Value / dCantidadOrden) * 100;
        dAvance := xRound(dAvance, 4);
        dError  := (dInstaladoOrden + tdCantidad.Value) - dCantidadOrden;
        if (dError >= 0) then
          dAvance := 100 - dAvanceAnterior
        else
          dAvance := dAvance +  dAvancePartida;
    end;

    if (dAvanceAnterior + dAvance) > 100 then
    begin
        tdCantidad.Value := dCantidadOrden - dInstaladoOrden;
        dAvance := 100 - dAvanceAnterior
    end;
    result := dAvance;
end;

procedure TfrmBitacoraDepartamental_2.MantenerHorario1Click(Sender: TObject);
begin
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('UPDATE bitacoradepersonal SET sHoraInicioG = sHoraInicio, sHoraFinalG = sHoraFinal ' +
                                'where sContrato = :contrato And dIdFecha = :fecha And iIdActividad =:Id ');
    Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
    Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
    Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
    Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
    Connection.zCommand.Params.ParamByName('Id').AsInteger      := qryNotasGerencial.FieldValues['iIdActividad'];
    connection.zCommand.ExecSQL;

    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('UPDATE bitacoradeequipos SET sHoraInicioG = sHoraInicio, sHoraFinalG = sHoraFinal ' +
                                'where sContrato = :contrato And dIdFecha = :fecha And iIdActividad =:Id ');
    Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
    Connection.zCommand.Params.ParamByName('contrato').value    := param_Global_Contrato;
    Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
    Connection.zCommand.Params.ParamByName('fecha').value       := tdIdFecha.Date;
    Connection.zCommand.Params.ParamByName('Id').AsInteger      := qryNotasGerencial.FieldValues['iIdActividad'];
    connection.zCommand.ExecSQL;

    messageDLG('Actualizado!', mtInformation, [mbOK], 0);
end;

function TfrmBitacoraDepartamental_2.MaximoItem;
begin
    if Pos('TIERRA', Param_Frente) > 0 then
      global_inicio := global_inicio + 8000;

    MaximoDiario.Active := False;
    MaximoDiario.Params.ParamByName('Contrato').DataType := ftString;
    MaximoDiario.Params.ParamByName('Contrato').Value    := param_Global_Contrato;
    MaximoDiario.Params.ParamByName('Fecha').DataType    := ftDate;
    MaximoDiario.Params.ParamByName('Fecha').Value       := tdIdFecha.Date;
    MaximoDiario.Params.ParamByName('Inicio').DataType   := ftInteger;
    MaximoDiario.Params.ParamByName('Inicio').Value      := global_inicio;
    MaximoDiario.Params.ParamByName('Final').DataType    := ftInteger;
    MaximoDiario.Params.ParamByName('Final').Value       := global_final;
    MaximoDiario.Open;

    if MaximoDiario.FieldByName('TotalDiario').IsNull then
       result := global_inicio + 1
    else
       result := MaximoDiario.FieldValues['TotalDiario'] + 1;
end;

procedure TfrmBitacoraDepartamental_2.PartidaExistente(sParamActividad: string; sParamWbs: string; sParamTE: string);
begin
    QryExistePartida.Active := False;
    QryExistePartida.Params.ParamByName('Contrato').DataType  := ftString;
    QryExistePartida.Params.ParamByName('Contrato').Value     := param_global_contrato;
    QryExistePartida.Params.ParamByName('Fecha').DataType     := ftDate;
    QryExistePartida.Params.ParamByName('Fecha').Value        := tdIdFecha.Date;
    QryExistePartida.Params.ParamByName('Orden').DataType     := ftString;
    QryExistePartida.Params.ParamByName('Orden').Value        := QrFrentes.FieldValues['sNumeroOrden'];
    QryExistePartida.Params.ParamByName('wbs').DataType       := ftString;
    QryExistePartida.Params.ParamByName('wbs').Value          := sParamWbs;
    QryExistePartida.Params.ParamByName('Actividad').DataType := ftString;
    QryExistePartida.Params.ParamByName('Actividad').Value    := sParamActividad;
    QryExistePartida.Params.ParamByName('Turno').DataType     := ftString;
    QryExistePartida.Params.ParamByName('Turno').Value        := global_turno_reporte;
    QryExistePartida.Params.ParamByName('Tipo').DataType      := ftString;
    QryExistePartida.Params.ParamByName('Tipo').Value         := sParamTE;
    QryExistePartida.Open;
end;

function TfrmBitacoraDepartamental_2.InstaladoOrden(sParamActividad: string; sParamWbs: string; sParamFecha: Integer) :double;
begin
    {Catidad Instalada de la Partida a la Fecha Actual ...}
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('Select Sum(dCantidad) as dCantidad from bitacoradeactividades where ' +
                                'sContrato = :contrato and sIdConvenio =:Convenio And sNumeroOrden = :Orden And sWbs = :wbs '+
                                'And sNumeroActividad = :Actividad Group By sWbs, sNumeroActividad');
    connection.QryBusca.Params.ParamByName('contrato').AsString   := param_global_contrato;
    connection.Qrybusca.Params.ParamByName('Convenio').AsString   := QrFrentes.FieldByName('Convenio').AsString;
    connection.QryBusca.Params.ParamByName('orden').AsString      := QrFrentes.FieldValues['sNumeroOrden'];
    connection.QryBusca.Params.ParamByName('wbs').AsString        := sParamWbs;
    connection.QryBusca.Params.ParamByName('actividad').AsString  := sParamActividad;
    connection.QryBusca.open;

    if connection.QryBusca.RecordCount > 0 then
       result := connection.QryBusca.FieldByName('dCantidad').AsFloat
    else
       result := 0;
end;

procedure TfrmBitacoraDepartamental_2.lblFolioDblClick(Sender: TObject);
begin
    Application.CreateForm( TfrmReprogramacionFolio, frmReprogramacionFolio );
    frmReprogramacionFolio.ShowModal;
end;

function TfrmBitacoraDepartamental_2.CantidadAnexoC;
begin
    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select dCantidadAnexo from actividadesxanexo where ' +
                       'sContrato = :contrato And sWbs = :wbs And sTipoActividad = "Actividad" ');
    Connection.qryBusca.Params.ParamByName('contrato').AsString := param_global_contrato;
    Connection.QryBusca.Params.ParamByName('wbs').AsString      := ActividadesIguales.FieldByName('sWbsContrato').AsString;
    Connection.qryBusca.Open;

    if connection.QryBusca.RecordCount > 0 then
       result := connection.QryBusca.FieldByName('dCantidadAnexo').AsFloat
    else
       result := 0;
end;

function TfrmBitacoraDepartamental_2.InstaladoAnexoC(sParamActividad: string; sParamWbs: string): double;
begin
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('Select Sum(dCantidad) as dCantidad from bitacoradeactividades where ' +
      'sContrato = :contrato and sWbs =:wbs And sNumeroActividad = :Actividad Group By sNumeroActividad');
    connection.QryBusca.Params.ParamByName('contrato').DataType  := ftString;
    connection.QryBusca.Params.ParamByName('contrato').Value     := param_global_contrato;
    connection.QryBusca.Params.ParamByName('actividad').DataType := ftString;
    connection.QryBusca.Params.ParamByName('actividad').Value    := sParamActividad;
    connection.QryBusca.Params.ParamByName('wbs').DataType       := ftString;
    connection.QryBusca.Params.ParamByName('wbs').Value          := sParamWbs;
    connection.QryBusca.open;

    if connection.QryBusca.RecordCount > 0 then
       result := connection.QryBusca.FieldByName('dCantidad').AsFloat
    else
       result := 0;
end;

Function TfrmBitacoraDepartamental_2.AvanceFolio;
var
   dDiaSiguiente  : TDateTime;
begin
    {Avances anteriores}
    dDiaSiguiente := QryBitacora.FieldByName('dIdFecha').AsDateTime;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 4) AS dAvanceAnterior;';
    connection.QryBusca.ParamByName('Orden').AsString := QryBitacora.FieldByName('sContrato').AsString;
    connection.QryBusca.ParamByName('Folio').AsString := QryBitacora.FieldByName('sNumeroOrden').AsString;
    connection.QryBusca.Open;

    tdAvanceAnterior.Value := connection.QryBusca.FieldByName('dAvanceAnterior').AsFloat;

    //Avances Acumulados
    dDiaSiguiente := IncDay(dDiaSiguiente);
    connection.QryBusca2.SQL.Clear;
    connection.QryBusca2.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 4) AS dAvanceAcumulado;';
    connection.QryBusca2.ParamByName('Orden').AsString := QryBitacora.FieldByName('sContrato').AsString;
    connection.QryBusca2.ParamByName('Folio').AsString := QryBitacora.FieldByName('sNumeroOrden').AsString;
    connection.QryBusca2.Open;

    tdAvanceAcumulado.Value := connection.QryBusca2.FieldByName('dAvanceAcumulado').AsFloat;

    result := connection.QryBusca2.FieldByName('dAvanceAcumulado').AsFloat -  connection.QryBusca.FieldByName('dAvanceAnterior').AsFloat;
end;

procedure TfrmBitacoraDepartamental_2.Minimo_id;
begin
    zQryDatos.Active := False;
    zQryDatos.SQL.Clear;
    zQryDatos.SQL.Add('select min(iIdDiario) iIddiario, dAvanceActual  from bitacoradeactividades where sContrato =:Contrato ' +
                      'and dIdFecha =:fecha and sNumeroOrden =:Orden ' +
                      'and sIdTipoMovimiento = "E" group by sContrato ');
    zQryDatos.ParamByName('contrato').AsString := QryBitacora.FieldByName('sContrato').AsString;
    zQryDatos.ParamByName('Orden').AsString    := QryBitacora.FieldByName('sNumeroOrden').AsString;
    zQryDatos.ParamByName('Fecha').AsDateTime  := tdIdFecha.DateTime;
    zQryDatos.Open;
end;

Function TfrmBitacoraDepartamental_2.SumaMaterial(sParamMaterial: string) : Double;
begin
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('select sum(dCantidad) as Cantidad from bitacorademateriales '+
                        'where sContrato =:Contrato and dIdFecha <=:fecha and sIdMaterial =:Insumo '+
                        'group by sContrato');
    connection.QryBusca.ParamByName('Contrato').AsString  := QryBitacora.FieldByName('sContrato').AsString;
    connection.QryBusca.ParamByName('fecha').AsDate    := tdIdFecha.DateTime;
    connection.QryBusca.ParamByName('Insumo').AsString := bitacorademateriales.FieldByName('sIdMaterial').AsString;
    connection.QryBusca.Open;

    result := connection.QryBusca.FieldByName('Cantidad').AsFloat;
end;

Function TfrmBitacoraDepartamental_2.MaterialDisponible(sParamMaterial: string) : Double;
begin
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('select sum(dCantidad) as Cantidad from bitacoradesalida '+
                        'where sContrato =:Contrato and dFechaSalida <=:fecha and sIdInsumo = :Insumo group by sContrato');
    connection.QryBusca.ParamByName('Contrato').AsString  := QryBitacora.FieldByName('sContrato').AsString;
    connection.QryBusca.ParamByName('fecha').AsDate    := tdIdFecha.DateTime;
    connection.QryBusca.ParamByName('Insumo').AsString := bitacorademateriales.FieldByName('sIdMaterial').AsString;
    connection.QryBusca.Open;

    result := connection.QryBusca.FieldByName('Cantidad').AsFloat;
end;

Procedure TfrmBitacoraDepartamental_2.GeneraReporteDiario_PDF(RTipo:FtTipo;RImpresion:FtSeccion);
var
   sSeccion: string;
begin
    EncabezadoPDF_Horizontal(ReporteDiario,rDiario,FtAbordo);
    FirmasPDF_Generales(ReporteDiario,     rDiario,FtAbordo);
    sSeccion := connection.configuracion.FieldByName('sSeccionImprime').AsString;
    {Clasificacion de secciones a Imprimir..}

    if pos('Avance Global', sSeccion) > 0 then
       ReportePDF_ActividadesPorFolio(QrFrentes.FieldByName('sNumeroOrden').AsString , ReporteDiario, rDiario,RTipo,RImpresion)
    else
       ReportePDF_ActividadesPorFolio(QrFrentes.FieldByName('sNumeroOrden').AsString , ReporteDiario, rDiario,RTipo,ftsNone);


    rDiario.LoadFromFile(global_files + global_Mireporte + '_TDReporteDiarioActividades.fr3') ;
    rDiario.ShowReport();
    ReportePDF_ClearDataset(rDiario);
end;


end.
