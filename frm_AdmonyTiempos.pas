unit frm_AdmonyTiempos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, ZAbstractRODataset, ZAbstractDataset, ZDataset, Menus,
  frxClass, frxDBSet, DB, StrUtils, Global, frm_connection, StdCtrls, ExtCtrls, Mask, frm_barra,
  DBCtrls, Grids, DBGrids, ADODB, Buttons, DateUtils, unitTarifa,
  RXToolEdit, RXCurrEdit, RXDBCtrl, Utilerias, udbgrid, unitexcepciones, ClipBrd,
  unitTbotonesPermisos, UnitValidaTexto, unitactivapop, DBDateTimePicker, UnitValidacion,
  AdvGlowButton, JvExMask, JvToolEdit, JvCombobox, Newpanel,ComObj,
  AdvSmoothPanel, AdvDateTimePicker, AdvCombo,jpeg,ShellAPI,
  JvExStdCtrls, JvCheckBox, JvDBCheckBox;

type
  Tembarcacion = class
    private
    Identificador:string;
    Descrp: string;

  end;
  TfrmAdmonyTiempos = class(TForm)
    pgAdmon: TPageControl;
    movembarcacion: TTabSheet;
    ConClimatologicas: TTabSheet;
    arriboembarcaciones: TTabSheet;
    ds_Clasificaciones: TDataSource;
    ds_movimientosdeembarcacion: TDataSource;
    dbMovimientos: TfrxDBDataset;
    dbOtrosMovimientos: TfrxDBDataset;
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
    N3: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    N4: TMenuItem;
    Salir1: TMenuItem;
    frmBarra2: TfrmBarra;
    Clasificaciones: TZReadOnlyQuery;
    Fases: TZReadOnlyQuery;
    ds_Fases: TDataSource;
    MovimientosdeBarco: TZQuery;
    frmBarra1: TfrmBarra;
    ds_Embarcaciones: TDataSource;
    movimientosdeembarcacion: TZQuery;
    ds_MovimientosdeBarcos: TDataSource;
    dsCondicionesClimatologicas: TDataSource;
    qryCondicionesClimatologicas: TZQuery;
    frmBarra3: TfrmBarra;
    dsDirecciones: TDataSource;
    qryDirecciones: TZQuery;
    qryCondicionesClimatologicasCalcCondiciones: TStringField;
    qryCondicionesClimatologicascalcDireccion: TStringField;
    qryCondiciones: TZQuery;
    dsCondiciones: TDataSource;
    qryCondicionesClimatologicassContrato: TStringField;
    qryCondicionesClimatologicasdIdFecha: TDateField;
    qryCondicionesClimatologicassIdTurno: TStringField;
    qryCondicionesClimatologicasmPronostico: TMemoField;
    qryCondicionesClimatologicassHorario: TStringField;
    qryCondicionesClimatologicasdCantidad: TFloatField;
    dbExistencias: TDBGrid;
    frmBarra4: TfrmBarra;
    qryRecursos: TZQuery;
    dsRecursos: TDataSource;
    qryRecursosCalcMezclas: TStringField;
    qryRecursossContrato: TStringField;
    qryRecursosdIdFecha: TDateField;
    qryRecursossIdTurno: TStringField;
    qryRecursosdProduccion: TFloatField;
    qryRecursosdRecibido: TFloatField;
    qryRecursosdConsumo: TFloatField;
    qryRecursosdConsumoEquipos: TFloatField;
    qryRecursosdPrestamos: TFloatField;
    qryRecursosdExistenciaActual: TFloatField;
    qryRecursosCalcMezclasMedidas: TStringField;
    dsMezclas: TDataSource;
    qryMezclas: TZQuery;
    qryRecursosdExistenciaAnterior: TFloatField;
    qryRecursosdAcumulado: TFloatField;
    Embarcaciones: TZReadOnlyQuery;
    fReporte: TfrxReport;
    qryCondicionesClimatologicasCalcCMedida: TStringField;
    qryRecursosiIdRecursoExistencia: TLargeintField;
    MovimientosdeBarcosContrato: TStringField;
    MovimientosdeBarcodIdFecha: TDateField;
    MovimientosdeBarcosIdDepartamento: TStringField;
    MovimientosdeBarcosClasificacion: TStringField;
    MovimientosdeBarcosIdFase: TStringField;
    MovimientosdeBarcosHoraInicio: TStringField;
    MovimientosdeBarcosHoraFinal: TStringField;
    MovimientosdeBarcosFactor: TStringField;
    MovimientosdeBarcomDescripcion: TMemoField;
    MovimientosdeBarcosIdEmbarcacion: TStringField;
    movimientosdeembarcacionsContrato: TStringField;
    movimientosdeembarcaciondIdFecha: TDateField;
    movimientosdeembarcacionsIdDepartamento: TStringField;
    movimientosdeembarcacionsIdEmbarcacion: TStringField;
    movimientosdeembarcacionsClasificacion: TStringField;
    movimientosdeembarcacionsIdFase: TStringField;
    movimientosdeembarcacionsHoraInicio: TStringField;
    movimientosdeembarcacionsHoraFinal: TStringField;
    movimientosdeembarcacionsFactor: TStringField;
    movimientosdeembarcacionmDescripcion: TMemoField;
    EmbarcacionessIdEmbarcacion: TStringField;
    EmbarcacionessDescripcion: TStringField;
    MovimientosdeBarcosDescripcion: TStringField;
    movimientosdeembarcacionsDescripcion: TStringField;
    Edit1: TEdit;
    Label1: TLabel;
    movimientosdeembarcacioniIdDiario: TIntegerField;
    TabExistencias: TTabSheet;
    movimientosdeembarcacionsTipo: TStringField;
    MovimientosdeBarcosTipo: TStringField;
    qryRecursossIdEmbarcacion: TStringField;
    qryRecursossEmbarcacion: TStringField;
    Embarcaciones2: TZReadOnlyQuery;
    StringField1: TStringField;
    StringField2: TStringField;
    ds_embarcaciones2: TDataSource;
    Embarcaciones2sTipo: TStringField;
    qryRecursosdTrasiego: TFloatField;
    qryRecursosdAjuste: TFloatField;
    qryCondicionesClimatologicasiIdCondicion: TLargeintField;
    qryCondicionesClimatologicasiIdDireccion: TLargeintField;
    tdFecha: TDBDateTimePicker;
    Plataformas: TZReadOnlyQuery;
    ds_Plataformas: TDataSource;
    MovimientosdeBarcosIdPlataforma: TStringField;
    dsrTiemposCia: TfrxDBDataset;
    dsTiemposCia: TDataSource;
    qryTiempoCia: TZQuery;
    Panel1: TPanel;
    Edit2: TEdit;
    Edit3: TEdit;
    ImprimirTiempoCia1: TMenuItem;
    Button1: TButton;
    Fechai: TDateTimePicker;
    Fechaf: TDateTimePicker;
    DBLookupComboBox1: TDBLookupComboBox;
    Label30: TLabel;
    Label31: TLabel;
    qryCondicionesClimatologicassCantidad: TStringField;
    PanelExistencias: TPanel;
    Label19: TLabel;
    tdEmbarcacionExist: TDBLookupComboBox;
    Label20: TLabel;
    iIdRecursoExistencia: TDBLookupComboBox;
    Label21: TLabel;
    dProduccion: TRxCalcEdit;
    Label22: TLabel;
    dRecibido: TRxCalcEdit;
    Label23: TLabel;
    dConsumo: TRxCalcEdit;
    dConsumoEquipos: TRxCalcEdit;
    Label25: TLabel;
    Label17: TLabel;
    dPrestamos: TRxCalcEdit;
    Label28: TLabel;
    dExistenciaAnterior: TRxCalcEdit;
    Label27: TLabel;
    dExistenciaActual: TRxCalcEdit;
    Label26: TLabel;
    dTrasiego: TRxCalcEdit;
    Label29: TLabel;
    dAjuste: TRxCalcEdit;
    chkDescuento: TCheckBox;
    PanelMovimientos: TPanel;
    Label11: TLabel;
    tsIdBarco: TDBLookupComboBox;
    Label7: TLabel;
    tsClasificaciones: TDBLookupComboBox;
    Label16: TLabel;
    tsIdFase: TDBLookupComboBox;
    Label8: TLabel;
    mkHora1: TMaskEdit;
    Label9: TLabel;
    mkHora2: TMaskEdit;
    sSuma: TEdit;
    Label15: TLabel;
    tmDescripcion2: TMemo;
    Label10: TLabel;
    PanelCondiciones: TPanel;
    PanelArribo: TPanel;
    Label2: TLabel;
    dbEmbarcaciones: TDBLookupComboBox;
    Label3: TLabel;
    tsHoraInicio: TMaskEdit;
    tsHoraFinal: TMaskEdit;
    Label4: TLabel;
    Label6: TLabel;
    rbArribo: TRadioButton;
    rbDisposicion: TRadioButton;
    rbDos: TRadioButton;
    tmDescripcion: TDBMemo;
    Label5: TLabel;
    chkContinuaMov: TCheckBox;
    chkContinuaArribo: TCheckBox;
    movimientosdeembarcacionlContinuo: TStringField;
    MovimientosdeBarcolContinuo: TStringField;
    dbMovtosEmbarcacion: TRxDBGrid;
    lblBusca: TEdit;
    dbMovBarco: TRxDBGrid;
    tdJornada: TEdit;
    btnAjusta: TBitBtn;
    Label32: TLabel;
    qryRecursossUbicacion: TStringField;
    Label34: TLabel;
    Label35: TLabel;
    dGalones: TRxCalcEdit;
    qryRecursosdGalones: TFloatField;
    CmbFolios: TJvCheckedComboBox;
    Label33: TLabel;
    MovimientosdeBarcoiIdDiario: TIntegerField;
    DBMemo1: TDBMemo;
    MovimientosdeBarcosOrden: TStringField;
    DBPlataformas: TfrxDBDataset;
    tsOrdenes: TDBLookupComboBox;
    Label18: TLabel;
    zqOrdenes: TZQuery;
    ds_ordenes: TDataSource;
    Label24: TLabel;
    tsNumeroActividad: TDBLookupComboBox;
    zqPartida: TZQuery;
    ds_partidas: TDataSource;
    MovimientosdeBarcosNumeroActividad: TStringField;
    Edit4: TEdit;
    tsOrdenesSeleccion: TComboBox;
    movimientosdeembarcacionsNumeroActividad: TStringField;
    PanelOrdena: tNewGroupBox;
    btnUp: TBitBtn;
    btnDown: TBitBtn;
    btnOk: TBitBtn;
    ds_OrdenaOrden: TDataSource;
    dbCondiciones: TDBGrid;
    Label13: TLabel;
    iIdCondiciones: TDBLookupComboBox;
    Label12: TLabel;
    iIdDireccion: TDBLookupComboBox;
    Label14: TLabel;
    sHorario: TMaskEdit;
    Label36: TLabel;
    dbedtsCantidad: TDBEdit;
    Label37: TLabel;
    sPronostico: TDBMemo;
    dCantidad: TRxCalcEdit;
    grid_bitacorapersonal: TRxDBGrid;
    zqOrdenaOrden: TZQuery;
    zqOrdenaOrdensOrden: TStringField;
    zqOrdenaOrdeniIdOrden: TIntegerField;
    movimientosdeembarcacionsOrden: TStringField;
    zqOrdenaOrdenlAplicaRecibidoDiesel: TStringField;
    zqOrdenaOrdenlAplicaRecibidoAgua: TStringField;
    Label38: TLabel;
    tsLocalizacion: TDBEdit;
    qryCondicionesClimatologicassLocalizacion: TStringField;
    zqOrdenaOrdenlAplicaProducidoAgua: TStringField;
    zqOrdenaOrdenlAplicaTrasiegoAgua: TStringField;
    MovimientosdeBarcosActividades: TStringField;
    CbPartidas: TJvCheckedComboBox;
    Label39: TLabel;
    chkAplicaFactor: TCheckBox;
    BtnImpPlantilla: TAdvGlowButton;
    BorrarEx: TCheckBox;
    BtnGenPlantilla: TAdvGlowButton;
    PanelMovimientosxfolio: tNewGroupBox;
    btnAplicar: TBitBtn;
    grid_movimientosxfolio: TDBGrid;
    ds_movimientosxfolio: TDataSource;
    zqMovimientosxfolio: TZQuery;
    zqMovimientosxfoliosNumeroOrden: TStringField;
    zqMovimientosxfoliosFolio: TStringField;
    zqMovimientosxfoliosFactor: TFloatField;
    zqMovimientosxfoliosHoraInicio: TStringField;
    zqMovimientosxfoliosHoraFinal: TStringField;
    zqMovimientosxfolioiIdDiario: TIntegerField;
    ChkCalcula: TCheckBox;
    BitBtn1: TBitBtn;
    Existenciasycosumos1: TMenuItem;
    PnlExistenciasC: TAdvSmoothPanel;
    BtnImprimir: TAdvGlowButton;
    btnExit: TAdvGlowButton;
    CmbMeses: TAdvComboBox;
    CmbAnno: TAdvComboBox;
    GuardaExcel: TSaveDialog;
    CmbEmb: TAdvComboBox;
    cmdEmbarcaciones: TBitBtn;
    dbchkNavegando: TJvDBCheckBox;
    MovimientosdeBarcoeNavegando: TStringField;
    tsUbicacionBarco: TDBEdit;
    rbMovimiento: TRadioButton;
    EmbarcacionessTipo: TStringField;
    Label40: TLabel;
    tsMovimiento: TDBLookupComboBox;
    procedure frmBarra2btnAddClick(Sender: TObject);
    procedure frmBarra2btnExitClick(Sender: TObject);
    procedure frmBarra2btnPostClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure _tsClasificacionesEnter(Sender: TObject);
    procedure _tsClasificacionesExit(Sender: TObject);
    procedure _tsClasificacionesKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdFaseEnter(Sender: TObject);
    procedure tsIdFaseExit(Sender: TObject);
    procedure tsIdFaseKeyPress(Sender: TObject; var Key: Char);
    procedure mkHora1KeyPress(Sender: TObject; var Key: Char);
    procedure mkHora1Enter(Sender: TObject);
    procedure mkHora1Exit(Sender: TObject);
    procedure mkHora2Enter(Sender: TObject);
    procedure mkHora2Exit(Sender: TObject);
    procedure mkHora2KeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);

    procedure frmBarra2btnDeleteClick(Sender: TObject);
    procedure tsHoraInicioEnter(Sender: TObject);
    procedure tsHoraFinalEnter(Sender: TObject);
    procedure tsHoraInicioExit(Sender: TObject);
    procedure tsHoraInicioKeyPress(Sender: TObject; var Key: Char);
    procedure tsHoraFinalExit(Sender: TObject);
    procedure tsHoraFinalKeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra2btnEditClick(Sender: TObject);
    procedure frmBarra3btnAddClick(Sender: TObject);
    procedure frmBarra3btnCancelClick(Sender: TObject);
    procedure frmBarra3btnRefreshClick(Sender: TObject);
    procedure frmBarra3btnDeleteClick(Sender: TObject);
    procedure frmBarra3btnPostClick(Sender: TObject);
    procedure frmBarra3btnEditClick(Sender: TObject);
    procedure frmBarra3btnExitClick(Sender: TObject);
    procedure dIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure qryCondicionesClimatologicasCalcFields(DataSet: TDataSet);
    procedure frmBarra2btnCancelClick(Sender: TObject);
    procedure frmBarra2btnRefreshClick(Sender: TObject);
    procedure tmDescripcionEnter(Sender: TObject);
    procedure tmDescripcionExit(Sender: TObject);
    procedure qryRecursosCalcFields(DataSet: TDataSet);
    procedure frmBarra4btnAddClick(Sender: TObject);
    procedure frmBarra4btnEditClick(Sender: TObject);
    procedure frmBarra4btnPostClick(Sender: TObject);
    procedure frmBarra4btnCancelClick(Sender: TObject);
    procedure frmBarra4btnDeleteClick(Sender: TObject);
    procedure frmBarra4btnRefreshClick(Sender: TObject);
    procedure frmBarra4btnExitClick(Sender: TObject);
    procedure qryRecursosBeforePost(DataSet: TDataSet);
    procedure qryRecursosAfterPost(DataSet: TDataSet);
    procedure qryRecursosAfterDelete(DataSet: TDataSet);
    procedure dProduccionEnter(Sender: TObject);
    procedure dProduccionExit(Sender: TObject);
    procedure dRecibidoEnter(Sender: TObject);
    procedure dRecibidoExit(Sender: TObject);
    procedure dConsumoEnter(Sender: TObject);
    procedure dConsumoExit(Sender: TObject);
    procedure dConsumoEquiposEnter(Sender: TObject);
    procedure dConsumoEquiposExit(Sender: TObject);
    procedure dTrasiegoEnter(Sender: TObject);
    procedure dTrasiegoExit(Sender: TObject);
    procedure dExistenciaActualEnter(Sender: TObject);
    procedure dExistenciaActualExit(Sender: TObject);
    procedure dExistenciaAnteriorExit(Sender: TObject);
    procedure dExistenciaAnteriorEnter(Sender: TObject);
    procedure iIdRecursoExistenciaEnter(Sender: TObject);
    procedure iIdRecursoExistenciaExit(Sender: TObject);
    procedure dIdFechaExistenciaExit(Sender: TObject);
    procedure dIdFechaExistenciaKeyPress(Sender: TObject; var Key: Char);
    procedure iIdRecursoExistenciaKeyPress(Sender: TObject; var Key: Char);
    procedure dProduccionKeyPress(Sender: TObject; var Key: Char);
    procedure dRecibidoKeyPress(Sender: TObject; var Key: Char);
    procedure dConsumoKeyPress(Sender: TObject; var Key: Char);
    procedure dConsumoEquiposKeyPress(Sender: TObject; var Key: Char);
    procedure dTrasiegoKeyPress(Sender: TObject; var Key: Char);
    procedure dExistenciaActualKeyPress(Sender: TObject; var Key: Char);
    procedure dExistenciaAnteriorKeyPress(Sender: TObject; var Key: Char);
    procedure tsClasificacionesEnter(Sender: TObject);
    procedure tsClasificacionesExit(Sender: TObject);
    procedure tsClasificacionesKeyPress(Sender: TObject; var Key: Char);
    procedure MovimientosdeBarcoCalcFields(DataSet: TDataSet);
    procedure movembarcacionEnter(Sender: TObject);
    procedure arriboembarcacionesEnter(Sender: TObject);
    procedure movimientosdeembarcacionAfterScroll(DataSet: TDataSet);
    procedure tsIdBarcoExit(Sender: TObject);
    procedure tsIdBarcoEnter(Sender: TObject);
    procedure tsIdBarcoKeyPress(Sender: TObject; var Key: Char);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure dbEmbarcacionesExit(Sender: TObject);
    procedure qryRecursosAfterScroll(DataSet: TDataSet);
    procedure movimientosdeembarcacionCalcFields(DataSet: TDataSet);
    procedure MovimientosdeBarcoAfterScroll(DataSet: TDataSet);
    procedure btnAjustaClick(Sender: TObject);
    procedure tdFechaExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tdFechaEnter(Sender: TObject);
    procedure tdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure dbMovBarcoCellClick(Column: TColumn);
    procedure dbMovBarcoMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure dbMovtosEmbarcacionMouseWheel(Sender: TObject; Shift: TShiftState;
      WheelDelta: Integer; MousePos: TPoint; var Handled: Boolean);
    procedure dbMovBarcoKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure dbMovtosEmbarcacionKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure dbExistenciasEnter(Sender: TObject);
    procedure TabExistenciasEnter(Sender: TObject);
    procedure tdEmbarcacionExistKeyPress(Sender: TObject; var Key: Char);
    procedure tdEmbarcacionExistEnter(Sender: TObject);
    function ReporteLock(): boolean;
    procedure frmBarra1btnPrinterClick(Sender: TObject);
    procedure frmBarra2btnPrinterClick(Sender: TObject);
    procedure frmBarra4btnPrinterClick(Sender: TObject);
    procedure Embarcaciones2AfterScroll(DataSet: TDataSet);
    procedure dAjusteKeyPress(Sender: TObject; var Key: Char);
    procedure dAjusteEnter(Sender: TObject);
    procedure dAjusteExit(Sender: TObject);
    procedure dPrestamosKeyPress(Sender: TObject; var Key: Char);
    procedure chkDescuentoKeyPress(Sender: TObject; var Key: Char);
    procedure dbMovBarcoMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure dbCondicionesMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure dbMovtosEmbarcacionMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure dbExistenciasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure dbMovBarcoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dbCondicionesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dbMovtosEmbarcacionMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dbExistenciasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure dbMovBarcoTitleClick(Column: TColumn);
    procedure dbCondicionesTitleClick(Column: TColumn);
    procedure dbMovtosEmbarcacionTitleClick(Column: TColumn);
    procedure dbExistenciasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure dbEmbarcacionesEnter(Sender: TObject);
    procedure dbEmbarcacionesKeyPress(Sender: TObject; var Key: Char);
    function existeReporte: boolean;
    procedure dExistenciaAnteriorChange(Sender: TObject);
    procedure dExistenciaActualChange(Sender: TObject);
    procedure dAjusteChange(Sender: TObject);
    procedure ImprimirTiempoCia1Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Split
      (const Delimiter: Char;
      Input: string;
      const Strings: TStrings);
    procedure tmDescripcion2Enter(Sender: TObject);
    procedure tmDescripcion2Exit(Sender: TObject);
    procedure dbMovtosEmbarcacionGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure dbEmbarcacionesKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure filtra;
    procedure filtraCondicion;
    procedure pgAdmonChange(Sender: TObject);
    procedure iIdDireccionKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure dbMovBarcoGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure tsUbicacionBarcoEnter(Sender: TObject);
    procedure tsUbicacionBarcoExit(Sender: TObject);
    procedure tdEmbarcacionExistExit(Sender: TObject);
    procedure dGalonesExit(Sender: TObject);
    procedure dGalonesEnter(Sender: TObject);
    procedure dGalonesKeyPress(Sender: TObject; var Key: Char);
    procedure CargarCheckCombo(ComboFolios:TJvCustomCheckedComboBox);
    procedure InicializarCheckCombo(ComboFolios:TJvCustomCheckedComboBox);
    procedure InicializarCheckComboNew(ComboFolios:TJvCustomCheckedComboBox; indice : Integer);
    procedure GrabarCheckCombo(ComboFolios:TJvCustomCheckedComboBox);
    procedure FormCreate(Sender: TObject);
    procedure CmbFoliosExit(Sender: TObject);
    procedure tsOrdenesExit(Sender: TObject);
    procedure tsOrdenesKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroActividadEnter(Sender: TObject);
    procedure zqOrdenesAfterScroll(DataSet: TDataSet);
    procedure tsOrdenesSeleccionExit(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure dbExistenciasDblClick(Sender: TObject);
    procedure btnUpClick(Sender: TObject);
    procedure btnDownClick(Sender: TObject);
    procedure OrdenarOrdenes(sParamOrden : string);
    procedure iIdCondicionesKeyPress(Sender: TObject; var Key: Char);
    procedure iIdCondicionesEnter(Sender: TObject);
    procedure iIdCondicionesExit(Sender: TObject);
    procedure iIdDireccionKeyPress(Sender: TObject; var Key: Char);
    procedure iIdDireccionEnter(Sender: TObject);
    procedure iIdDireccionExit(Sender: TObject);
    procedure sHorarioKeyPress(Sender: TObject; var Key: Char);
    procedure sHorarioEnter(Sender: TObject);
    procedure sHorarioExit(Sender: TObject);
    procedure dbedtsCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure dbedtsCantidadEnter(Sender: TObject);
    procedure dbedtsCantidadExit(Sender: TObject);
    procedure sPronosticoEnter(Sender: TObject);
    procedure sPronosticoExit(Sender: TObject);
    procedure qryCondicionesClimatologicasAfterScroll(DataSet: TDataSet);
    procedure qryCondicionesClimatologicasBeforePost(DataSet: TDataSet);
    procedure zqOrdenaOrdenBeforePost(DataSet: TDataSet);
    procedure zqOrdenaOrdenlAplicaRecibidoDieselChange(Sender: TField);
    procedure zqOrdenaOrdenlAplicaRecibidoAguaChange(Sender: TField);
    procedure tsLocalizacionEnter(Sender: TObject);
    procedure tsLocalizacionExit(Sender: TObject);
    procedure tsLocalizacionKeyPress(Sender: TObject; var Key: Char);
    procedure CbPartidasExit(Sender: TObject);
    procedure CbPartidasKeyPress(Sender: TObject; var Key: Char);
    procedure BtnImpPlantillaClick(Sender: TObject);
    procedure BtnGenPlantillaClick(Sender: TObject);
    procedure btnAplicarClick(Sender: TObject);
    procedure zqMovimientosxfolioCalcFields(DataSet: TDataSet);
    procedure BitBtn1Click(Sender: TObject);
    procedure BtnImprimirClick(Sender: TObject);
    procedure Existenciasycosumos1Click(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure CmbAnnoChange(Sender: TObject);
    procedure CmbMesesChange(Sender: TObject);
    procedure cmdEmbarcacionesClick(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    function ValidaHorario(sParamHorario : string) : boolean;
    procedure rbArriboEnter(Sender: TObject);
    procedure rbMovimientoEnter(Sender: TObject);
    procedure DBMemo1Enter(Sender: TObject);
    procedure DBMemo1Exit(Sender: TObject);
  private
    { Private declarations }
    ListaContratos:TStringList;
    ZMovtos,ZFolios:TZReadOnlyQuery;
    inciandoAgua:Boolean;
    QryBarcoVigencia  : TZReadOnlyQuery;
    procedure EstableceCheckPArtidas(Cadena: string; Combo: tjvcheckedcombobox);
    function OrdenarCadena(CadOriginal: String): string;
    procedure ImportarMovtosBarco(Contrat, Emb: String);
    function ColumnaNombre(Numero: Integer): String;
    procedure GeneraPlantillaImp;
    procedure ImprimirExistenciasConsumo(Finicial, Ffinal: TDateTime;
      ContBarco: string);
    procedure CargaEmb(saño,smes:string);
    procedure AjusteProrrateo;
    procedure ExistenciaAnterior;
  public
    { Public declarations }
  end;

  TContrato=class
    private
      FsContrato,FsNumeroOrden:string;
      FIdentificador:string;
    published
      property Id:string Read FIdentificador write FIdentificador;
      property sContrato:string read FsContrato write FsContrato;
      property sNumeroOrden:string read FsNumeroOrden write FsNumeroOrden;
  end;
  const
    CFecha = 8;
    CHClima = 12;
    CHInicio = 1;
    CHFin = 3;
    CDescripcion = 5;
    Cpartida = 8;
    CFolio = 9;
    COt = 10;
    CClasificacion = 12;
    MaxVacias = 10;
    CEmbarcaciones = 1;
    CHelicopteros = 1;
    CVar1 = 7;
    CVar2 = 10;

    //paleta de Colores para celdas
    ClFVacia = 19;//amarilla no se toma en cuenta
    ClHFinMenor = 40;//el horario fin es menor
    ClCeldaVacia = 46;//el valor no puede ser nulo
    ClNoBd = 45;//no existe en la base de datos

    ClFNoOt = 36;//El folio no corresponde a la ot
    ClOk = 35;//Todo ok .Interior.ColorIndex = 44
    ClLineaImp= 50;
    ClGris = 16;

    xlContinuous = $00000001;

var
  Entre: Boolean;
  frmAdmonyTiempos: TfrmAdmonyTiempos;
  qryGetAcumulado: TZReadOnlyQuery;
  sOpcionEmb: string;
  sEmbarcacion: string;
  Pagina: string;
  utgrid: ticdbgrid;
  utgrid2: ticdbgrid;
  utgrid3: ticdbgrid;
  utgrid4: ticdbgrid;
  botonpermiso: TBotonesPermisos;
  botonpermiso2: TBotonesPermisos;
  botonpermiso3: TBotonesPermisos;
  botonpermiso4: TBotonesPermisos;
  SavePlace     : TBookmark;
  lRecalculo : Boolean;

  //Aqui para activar el copiar y pegar en memos
  lCopiaObjeto : boolean;
  dConsumos, dEquipos : extended;
  sColumaOrden : string;
  FocusClasificacion,
  FocusOrdenes,
  FocusPartida  : string;
  FocusiFolio   : Integer;
  sLocalizacion : string;
  sHoraI, sHoraF : string;


implementation

uses frm_OpcionesBarco, frm_embarcaciones;


{$R *.dfm}


procedure TfrmAdmonyTiempos.GrabarCheckCombo(ComboFolios: TJvCustomCheckedComboBox);
var
  QrDatos:TZReadOnlyQuery;
  i:Integer;
begin
  QrDatos:=TZReadOnlyQuery.Create(nil);
  QrDatos.Connection:=connection.zConnection;
  try
    QrDatos.Active:=False;
    QrDatos.SQL.Text:='delete from movimientosxfolios where sContrato=:Contrato and dIdFecha=:fecha and iIdDiario=:Diario' ;
    QrDatos.ParamByName('Contrato').AsString := MovimientosdeBarco.FieldByName('sContrato').AsString;
    QrDatos.ParamByName('Fecha').AsDate      := MovimientosdeBarco.FieldByName('dIdFecha').AsDateTime;
    QrDatos.ParamByName('Diario').AsInteger  := MovimientosdeBarco.FieldByName('iidDiario').AsInteger;
    QrDatos.ExecSQL;

    if chkAplicaFactor.Checked then
    begin
        FocusiFolio   := Movimientosdebarco.FieldValues['iIdDiario'];

        QrDatos.Active:=False;
        QrDatos.SQL.Text:='INSERT INTO movimientosxfolios (sContrato, dIdFecha, iIdDiario, sNumeroOrden, sFolio, lFactor) VALUES (:Contrato, :Fecha, :Diario, :Ot, :Folio, :Aplica)';

        for I := 0 to ComboFolios.Items.Count-1 do
        begin
          if ComboFolios.Checked[I]=True then
          begin
            QrDatos.Active:=False;
            QrDatos.ParamByName('Contrato').AsString  := MovimientosdeBarco.FieldByName('sContrato').AsString;
            QrDatos.ParamByName('Fecha').AsDate       := MovimientosdeBarco.FieldByName('dIdFecha').AsDateTime;
            QrDatos.ParamByName('Diario').AsInteger   := MovimientosdeBarco.FieldByName('iidDiario').AsInteger;
            QrDatos.ParamByName('ot').AsString        := zqOrdenes.FieldValues['sContrato'];
            QrDatos.ParamByName('Folio').AsString     := ComboFolios.Items[i];
            QrDatos.ParamByName('Aplica').AsString    := 'Si';
            QrDatos.ExecSQL;
          end;
        end;
    end;
  finally
    FreeAndNil(QrDatos);
  end;
end;

procedure TfrmAdmonyTiempos.InicializarCheckCombo(ComboFolios: TJvCustomCheckedComboBox);
var
  QrDatos:TZReadOnlyQuery;
  i:Integer;
begin
  ComboFolios.SetUnCheckedAll();
  QrDatos:=TZReadOnlyQuery.Create(nil);
  QrDatos.Connection:=connection.zConnection;
  try
    QrDatos.SQL.Text:='select * from movimientosxfolios where sContrato=:Contrato and dIdFecha=:fecha and iIdDiario=:Diario';
    QrDatos.ParamByName('Contrato').AsString := MovimientosdeBarco.FieldByName('sContrato').AsString;
    QrDatos.ParamByName('Fecha').AsDate      := MovimientosdeBarco.FieldByName('dIdFecha').AsDateTime;
    QrDatos.ParamByName('Diario').AsInteger  := MovimientosdeBarco.FieldByName('iidDiario').AsInteger;
    QrDatos.Open;
    while not QrDatos.Eof do
    begin
        for I := 0 to ComboFolios.Items.Count-1 do
           if ComboFolios.Items.Strings[i]=QrDatos.FieldByName('sFolio').AsString then
             ComboFolios.Checked[I]:=True;
        if QrDatos.FieldValues['lFactor'] = 'Si' then
           chkAplicaFactor.Checked := True
        else
           chkAplicaFactor.Checked := False;
      QrDatos.Next;
    end;

    if clasificaciones.FieldValues['lGenera'] = 'No' then
    begin
        zqPartida.Active := False;
        zqPartida.ParamByName('Contrato').AsString  := zqOrdenes.FieldValues['sContrato'];
        zqPartida.ParamByName('Convenio').AsString  := global_convenio;
        zqPartida.ParamByName('Orden').AsString     := cmbFolios.Text;
        zqPartida.Open;
    end;

  finally
    FreeAndNil(QrDatos);
  end;
end;


procedure TfrmAdmonyTiempos.InicializarCheckComboNew(ComboFolios: TJvCustomCheckedComboBox; indice: Integer);
var
  QrDatos:TZReadOnlyQuery;
  i:Integer;
begin
  ComboFolios.SetUnCheckedAll();
  QrDatos:=TZReadOnlyQuery.Create(nil);
  QrDatos.Connection:=connection.zConnection;
  try
    QrDatos.SQL.Text:='select * from movimientosxfolios where sContrato=:Contrato and dIdFecha=:fecha and iIdDiario=:Diario';
    QrDatos.ParamByName('Contrato').AsString := global_Contrato_Barco;
    QrDatos.ParamByName('Fecha').AsDate      := tdFecha.Date;
    QrDatos.ParamByName('Diario').AsInteger  := indice;
    QrDatos.Open;
    while not QrDatos.Eof do
    begin
      for I := 0 to ComboFolios.Items.Count-1 do
        if ComboFolios.Items.Strings[i]=QrDatos.FieldByName('sFolio').AsString then
          ComboFolios.Checked[I]:=True;

      QrDatos.Next;
    end;
  finally
    FreeAndNil(QrDatos);
  end;
end;


procedure TfrmAdmonyTiempos.CargarCheckCombo(ComboFolios: TJvCustomCheckedComboBox);
var
  QrFolios:TZReadOnlyQuery;
  ObjContrato:TContrato;
begin
    CmbFolios.Clear;
    QrFolios:=TZReadOnlyQuery.Create(nil);
    QrFolios.Connection:=connection.zConnection;
    ListaContratos.Clear;
    try
      with QrFolios,CmbFolios do
      begin
        Connection:=frm_connection.connection.zConnection;
        SQL.Text:='select ot.* from ordenesdetrabajo ot ' +
                  'inner join contratos c on(ot.sContrato=c.sContrato) ' +
                  'inner join bitacoradeactividades ba on(ba.sContrato=c.sContrato and ba.sNumeroOrden=ot.sNumeroOrden) ' +
                  'inner join tiposdemovimiento tm on(tm.sContrato=:Contrato and tm.sIdTipoMovimiento=ba.sIdTipoMovimiento and tm.sClasificacion="Tarifa Diaria") ' +
                  'where c.sContrato=:ContratoNormal and ba.dIdFecha=:Fecha group by ot.sContrato,ot.sNumeroorden';
        ParamByName('ContratoNormal').AsString := zqOrdenes.FieldValues['sContrato'];
        ParamByName('Contrato').AsString       := global_Contrato_Barco;
        ParamByName('Fecha').AsDate            := tdFecha.Date;
        Open;

        while not Eof do
        begin
          Items.Add(FieldByName('sNumeroOrden').AsString);
          ObjContrato:=TContrato.Create;
          ObjContrato.sContrato:=FieldByName('sContrato').AsString;
          ObjContrato.sNumeroOrden:=FieldByName('sNumeroOrden').AsString;
          ObjContrato.Id:=FieldByName('sNumeroOrden').AsString;
          ListaContratos.AddObject(ObjContrato.Id,ObjContrato);
          Next;
        end;
      end;
    finally
      FreeAndNil(QrFolios);
    end;

end;

procedure TfrmAdmonyTiempos.CbPartidasExit(Sender: TObject);
var cat:string;
begin
  cat :=  OrdenarCadena(CbPartidas.Text);
  CbPartidas.Text:=cat;
end;

procedure TfrmAdmonyTiempos.CbPartidasKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tmDescripcion2.SetFocus;
end;

procedure TfrmAdmonyTiempos.frmBarra2btnAddClick(Sender: TObject);
begin

  if ReporteLock then
  begin
    messageDLg('El Reporte Dairio se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;
  dbMovtosEmbarcacion.Enabled := False;

  tabexistencias.PageControl.Pages[0].TabVisible := false;
  tabexistencias.PageControl.Pages[1].TabVisible := false;
  tabexistencias.PageControl.Pages[3].TabVisible := false;

  frmBarra2.btnAddClick(Sender);
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;

  dBEmbarcaciones.Enabled := True;
  tsHoraInicio.ReadOnly := False;
  tsHoraFinal.ReadOnly := False;
  tmDescripcion.ReadOnly := False;
  tsHoraInicio.Text := '00:00';
  tsHoraFinal.Text := '00:00';

  chkContinuaMov.Checked := False;
  PanelArribo.Enabled := True;
  sOpcionEmb := 'Insert';
  MovimientosdeEmbarcacion.Append;
  movimientosdeEmbarcacion.FieldValues['sContrato'] := global_contrato_barco;
  movimientosdeEmbarcacion.FieldValues['dIdFecha']  := tdFecha.Date;
  rbArribo.Checked := true;
  movimientosdeembarcacion.FieldValues['sTipo'] := 'ARRIBO';
  movimientosdeEmbarcacion.FieldValues['sIdEmbarcacion'] := sEmbarcacion;
  movimientosdeEmbarcacion.FieldValues['sHoraInicio'] := '00:00';
  movimientosdeEmbarcacion.FieldValues['sHoraFinal'] := '00:00';
  tsHoraInicio.SetFocus;

  embarcaciones.Locate('sIdEmbarcacion', 'N/A', [loCaseInsensitive]);
  dbEmbarcaciones.KeyValue := embarcaciones.FieldByName('sIdEmbarcacion').AsString;
  clasificaciones.First;
  tsMovimiento.KeyValue := clasificaciones.FieldByName('sIdTipoMovimiento').AsString;
  BotonPermiso.permisosBotones(frmBarra2);

end;

procedure TfrmAdmonyTiempos.frmBarra2btnExitClick(Sender: TObject);
begin
  frmBarra2.btnExitClick(Sender);
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  close;
end;

procedure TfrmAdmonyTiempos.frmBarra2btnPostClick(Sender: TObject);
var
  sDescripcion: string;
  lContinua, lIncorrecto : boolean;
  nombres, cadenas: TStringList;
  sHoraInicio : string;
begin

  {Validacion de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('Embarcacion'); nombres.Add('Hora Inicio'); nombres.Add('Hora Final');
  cadenas.Add(dbEmbarcaciones.Text); cadenas.Add(tsHoraInicio.Text); cadenas.Add(tsHoraFinal.Text);
  if not validaTexto(nombres, cadenas, '', '') then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;

   {Validacion de barco hora > 24:00}
  lIncorrecto := False;
  if (StrToInt(copy(tsHoraInicio.Text, 1, 2)) = 24) and (StrToInt(copy(tsHoraInicio.Text, 4, 5)) > 0) then
    lIncorrecto := True
  else
    if (StrToInt(copy(tsHoraInicio.Text, 1, 2)) = 25) then
      lIncorrecto := True;

  if lIncorrecto then
  begin
    messageDLG('La Hora de Inicio es Mayor a 24:00 Hrs', mtInformation, [mbOk], 0);
    exit;
  end;

  if (StrToInt(copy(tshoraFinal.Text, 1, 2)) = 24) and (StrToInt(copy(tshoraFinal.Text, 4, 5)) > 0) then
    lIncorrecto := True
  else
    if (StrToInt(copy(tshoraFinal.Text, 1, 2)) = 25) then
      lIncorrecto := True;

  if lIncorrecto then
  begin
    messageDLG('La Hora de Final es Mayor a 24:00 Hrs', mtInformation, [mbOk], 0);
    exit;
  end;

  if (tsHoraInicio.Text = '  :  ') or (tshoraFinal.Text = '  :  ') then
  begin
      ShowMessage('Los horarios no deben estar vacios!!');
      tshoraFinal.SetFocus;
      exit;
  end;

  if tsHoraInicio.Text > tshoraFinal.Text then
  begin
      ShowMessage('La hora de inicio es mayor que la Hora Final!');
      tsHoraInicio.SetFocus;
      exit;
  end;


  {Continua insercion de datos}
  //manejo de tab carmen
  tabexistencias.PageControl.Pages[0].TabVisible := true;
  tabexistencias.PageControl.Pages[1].TabVisible := true;
  tabexistencias.PageControl.Pages[3].TabVisible := true;
  try

      sEmbarcacion := movimientosdeembarcacion.FieldValues['sIdEmbarcacion'];
      if tmDescripcion.Text = '' then
        sDescripcion := '*'
      else
        sDescripcion := tmDescripcion.Text;

      movimientosdeembarcacion.FieldValues['sIdEmbarcacion'] := dbEmbarcaciones.KeyValue;
      movimientosdeembarcacion.FieldValues['sClasificacion'] := '';
      movimientosdeembarcacion.FieldValues['sIdFase'] := '';
      movimientosdeembarcacion.FieldValues['sHoraInicio'] := tsHoraInicio.Text;
      movimientosdeembarcacion.FieldValues['sHoraFinal'] := tsHoraFinal.Text;
      movimientosdeembarcacion.FieldValues['sFactor'] := '0';
      movimientosdeembarcacion.FieldValues['mDescripcion'] := sDescripcion;
      movimientosdeembarcacion.FieldValues['sNumeroActividad'] := tsMovimiento.KeyValue;

      if rbArribo.Checked = true then
        movimientosdeembarcacion.FieldValues['sTipo'] := 'ARRIBO';
      if rbDisposicion.Checked = true then
        movimientosdeembarcacion.FieldValues['sTipo'] := 'DISPOSICION';
      if (rbDos.Checked = true) and (sOpcionEmb = 'Insert') then
        movimientosdeembarcacion.FieldValues['sTipo'] := 'ARRIBO';
      if chkContinuaArribo.Checked then
          movimientosdeembarcacion.FieldValues['lContinuo'] := 'Si'
      else
          movimientosdeembarcacion.FieldValues['lContinuo'] := 'No';
      if rbMovimiento.Checked = true then
        movimientosdeembarcacion.FieldValues['sTipo'] := 'MOVIMIENTO';
      movimientosdeembarcacion.Post;
      //desactivapop(popupprincipal);
      sHoraInicio := movimientosdeembarcacion.FieldValues['sHoraInicio'];

      //hacemos el recorrido del qry

      movimientosdebarco.First;
      while not movimientosdebarco.Eof do
      begin
          movimientosdeembarcacion.First;
          while not movimientosdeembarcacion.Eof do
          begin
              if (movimientosdeembarcacion.FieldValues['sHoraInicio'] >= movimientosdebarco.FieldValues['sHoraInicio'])
                   and (movimientosdeembarcacion.FieldValues['sHoraInicio'] <= movimientosdebarco.FieldValues['sHoraFinal'])then

               //  and (movimientosdeembarcacion.FieldValues['sHoraFinal'] <= movimientosdebarco.FieldValues['sHoraFinal'])then
                 begin
                     movimientosdeembarcacion.Edit;
                     movimientosdeembarcacion.FieldValues['sOrden']  := movimientosdebarco.FieldValues['sOrden'];
                     movimientosdeembarcacion.Post;
                 end;
              movimientosdeembarcacion.Next;
          end;
          movimientosdebarco.Next;
      end;

      movimientosdeembarcacion.Locate('sHoraInicio', sHoraInicio, [loCaseInsensitive]);
      movimientosdebarco.First;
      MovimientosdeBarco.Locate('sHoraInicio', sHoraInicio, [loCaseInsensitive]);

      if (rbDos.Checked = true) and (sOpcionEmb = 'Insert') then
      begin
        movimientosdeEmbarcacion.insert;
        movimientosdeEmbarcacion.FieldValues['sContrato'] := global_contrato_barco;
        movimientosdeEmbarcacion.FieldValues['dIdFecha'] := tdFecha.Date;
        movimientosdeEmbarcacion.FieldValues['sIdEmbarcacion'] := sEmbarcacion;
        movimientosdeembarcacion.FieldValues['sClasificacion'] := '';
        movimientosdeembarcacion.FieldValues['sIdFase'] := '';
        movimientosdeembarcacion.FieldValues['sHoraInicio'] := tsHoraInicio.Text;
        movimientosdeembarcacion.FieldValues['sHoraFinal'] := tsHoraFinal.Text;
        movimientosdeembarcacion.FieldValues['sFactor'] := '0';
        movimientosdeembarcacion.FieldValues['mDescripcion'] := sDescripcion;
        movimientosdeembarcacion.FieldValues['sTipo'] := 'DISPOSICION';
        movimientosdeembarcacion.Post;
      end;

      if sOpcionEmb = 'Edit' then
        sOpcionEmb := '';

      chkContinuaArribo.Checked := False;
      PanelArribo.Enabled := False;
      Insertar1.Enabled := True;
      Editar1.Enabled := True;
      Registrar1.Enabled := False;
      Can1.Enabled := False;
      Eliminar1.Enabled := True;
      Refresh1.Enabled := True;
      Salir1.Enabled := True;
      dbMovtosEmbarcacion.Enabled := True;
      frmBarra2.btnPostClick(Sender);

      lblBusca.Text := '';
      embarcaciones.Filtered := False;

  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al salvar registro en arribo de embarcaciones', 0);
      frmbarra2.btnCancel.Click;
    end;
  end;
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmAdmonyTiempos.frmBarra2btnPrinterClick(Sender: TObject);
begin
  Application.CreateForm(TfrmOpcionesBarco, frmOpcionesBarco);
  frmOpcionesBarco.showModal;
end;

procedure TfrmAdmonyTiempos.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  QryBarcoVigencia.Destroy;
  botonpermiso.Free;
  utgrid.Destroy;
  utgrid2.Destroy;
  utgrid3.Destroy;
  utgrid4.Destroy;
  hide
end;

procedure TfrmAdmonyTiempos.FormCreate(Sender: TObject);
begin
  ListaContratos:=TStringList.Create;
  QryBarcoVigencia := TZReadOnlyQuery.Create(self);
  QryBarcoVigencia.Connection := connection.zConnection;

  ZMovtos := TZReadOnlyQuery.Create(self);
  ZMovtos.Connection := connection.zConnection;
  ZMovtos.Active := False;
  ZMovtos.SQL.Clear;
  ZMovtos.SQL.Text := 'select scontrato,sIdTipoMovimiento,stipo from tiposdemovimiento where sclasificacion = "Movimiento de Barco"';


  ZFolios := TZReadOnlyQuery.Create(self);
  ZFolios.Connection := connection.zConnection;
  ZFolios.Active := False;
  ZFolios.SQL.Clear;
  ZFolios.SQL.Text := 'SELECT scontrato,snumeroorden FROM ordenesdetrabajo';

  Clasificaciones.Active := False;
  Clasificaciones.Sql.Clear;
  Clasificaciones.Sql.Add(' select sIdTipoMovimiento, sDescripcion, sTipo, lGenera from tiposdemovimiento ' +
    ' Where sClasificacion = "Movimiento de Barco" ' +
    'And sContrato = :Contrato Order by iOrden ');
  Clasificaciones.Params.ParamByName('Contrato').DataType := ftString;
  Clasificaciones.Params.ParamByName('Contrato').Value := Global_Contrato_Barco;
  Clasificaciones.Open;  
end;

procedure TfrmAdmonyTiempos.FormShow(Sender: TObject);
var
  sFactor: string;
  dProrrateo: Double;
  dAjuste: Double;
  iMultiplo: Integer;
  iDecimales: Integer;

  Zemb:TZReadOnlyQuery;

begin
  Zemb:=TZReadOnlyQuery.Create(nil);
  try
    Zemb.Connection := connection.zConnection;
    Zemb.Active := False;
    Zemb.SQL.Add('select * from embarcaciones limit 1');
    Zemb.Open;
    if Zemb.FieldDefs.IndexOf('lIniciaAgua') < 0 then
    begin
      ShowMessage('Falta el campo liniciaagua en el modulo de embarcaciones.'#10+'Puede generarlo ingresando al catalogo de embarcaciones/Vuelos del menú catálogo para evitar problemas relacionados a este campo.');
    end;
  finally
    Zemb.Free;
  end;

  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'reporteBarco', PopupPrincipal);
  UtGrid := TicdbGrid.create(dbmovbarco);
  UtGrid2 := TicdbGrid.create(dbcondiciones);
  UtGrid3 := TicdbGrid.create(dbmovtosembarcacion);
  UtGrid4 := TicdbGrid.create(dbexistencias);
  global_barco := Connection.configuracion.FieldValues['sIdEmbarcacion'];
  tdFecha.Date := Date;

  if Connection.Contrato.FieldByName('sIdResidencia').AsString = '04' then
  begin
      label24.Visible := False;
      tsNumeroActividad.Visible := False;
      Label39.Visible    := False;
      cbPartidas.Visible := False;
  end;

  zqOrdenes.Active:= False;
  zqOrdenes.Open;

  tsOrdenesSeleccion.Items.Clear;
  tsOrdenesSeleccion.Items.Add('<<Todas>>');
  while not zqOrdenes.Eof do
  begin
      if zqOrdenes.FieldValues['sContrato'] <> global_contrato_barco then
         tsOrdenesSeleccion.Items.Add(zqOrdenes.FieldValues['sContrato']);
      zqOrdenes.Next;
  end;

  Movimientosdebarco.Active := False;
  MovimientosdeBarco.Sql.Clear;
  MovimientosdeBarco.Sql.Add('select movimientosdeembarcacion.* from movimientosdeembarcacion ' +
    'inner join tiposdemovimiento  on ' +
    '(tiposdemovimiento.sContrato = :Contrato ' +
    ' And movimientosdeembarcacion.sClasificacion = tiposdemovimiento.sIdTipoMovimiento) ' +
    'where movimientosdeembarcacion.dIdFecha = :Fecha and movimientosdeembarcacion.sOrden like :ContratoNormal and sIdFase = "OPER" order by sActividades, sIdEmbarcacion, sHoraInicio ');
  movimientosdebarco.Params.ParamByName('Contrato').DataType := ftString;
  movimientosdebarco.Params.ParamByName('Contrato').Value    := Global_Contrato_Barco;
  movimientosdebarco.Params.ParamByName('ContratoNormal').DataType := ftString;
  if tsOrdenesSeleccion.Text = '<<Todas>>' then
     movimientosdebarco.Params.ParamByName('ContratoNormal').Value := '%'
  else
     movimientosdebarco.Params.ParamByName('ContratoNormal').Value := tsOrdenesSeleccion.Text;
  movimientosdebarco.Params.ParamByName('Fecha').DataType    := ftDate;
  movimientosdebarco.Params.ParamByName('Fecha').Value       := tdFecha.Date;
  movimientosdebarco.Open;
  try
    if movimientosdebarco.RecordCount > 0 then
      tdJornada.Text := sFnSumaBarco(movimientosdebarco.FieldValues['dIdFecha'], global_barco, frmAdmonyTiempos, -1);
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al mostrar ventana, al calcular total diario', 0);
    end;
  end;

  if Clasificaciones.RecordCount > 0 then
     FocusClasificacion := Clasificaciones.FieldValues['sIdTipoMovimiento'];

  sOpcionEmb := '';

  tdJornada.Text := '0.000000';

  iIdRecursoExistencia.KeyValue := 1;

  frmBarra1.btnCancel.Click;

  tsHoraInicio.ReadOnly := True;
  tsHoraFinal.ReadOnly := True;
  tmDescripcion.ReadOnly := True;

  Fases.Active := False;
  Fases.Open;

  Embarcaciones.Active := False;
  Embarcaciones.SQL.clear;
  Embarcaciones.sql.Add('select sIdEmbarcacion, sDescripcion, sTipo from embarcaciones ' +
    'Where sTipo="Principal" order by sDescripcion');
  Embarcaciones.Open;

  Embarcaciones2.Active := False;
  Embarcaciones2.Open;

  qryCondiciones.Active := True;
  qryDirecciones.Active := True;

  qryCondicionesClimatologicas.Active := False;
  qryCondicionesClimatologicas.Params.ParamByName('Fecha').DataType := ftDate;
  qryCondicionesClimatologicas.Params.ParamByName('Fecha').Value    := tdFecha.Date;
  qryCondicionesClimatologicas.Open;

  movimientosdeEmbarcacion.Active := False;
  movimientosdeEmbarcacion.Params.ParamByName('Contrato').DataType := ftString;
  movimientosdeEmbarcacion.Params.ParamByName('Contrato').Value    := Global_Contrato_barco;
  movimientosdeEmbarcacion.Params.ParamByName('Fecha').DataType    := ftDate;
  movimientosdeEmbarcacion.Params.ParamByName('Fecha').Value       := tdFecha.date;
  movimientosdeEmbarcacion.Open;

  qryRecursos.Active := False;
  qryRecursos.Params.ParamByName('Fecha').DataType := ftDate;
  qryRecursos.Params.ParamByName('Fecha').Value    := tdFecha.Date;
  qryRecursos.Open;

  qryMezclas.Active := False;
  qryMezclas.Open;

  Plataformas.Active := false;
  Plataformas.Open;

  iIdRecursoExistencia.Enabled := False;
  dProduccion.Enabled := False; self.dAjuste.Enabled := false;
  self.dAjuste.Enabled := False; //*****************BRITO 17/12/10********************
  dRecibido.Enabled := False;
  dConsumo.Enabled := false;
  dConsumoEquipos.Enabled := False;
  dPrestamos.Enabled := False;
  dExistenciaActual.Enabled := False;
  dExistenciaAnterior.Enabled := False;
  dTrasiego.Enabled := False;
  dbMovBarco.Columns[4].Width := 64;

  if FocusOrdenes = '' then
     FocusOrdenes :=  global_contrato;

  if global_fecha_rd = 0 then
     tdFecha.Date := now
  else
     tdFecha.Date := global_fecha_rd;
  tdFecha.OnExit(sender);
  dbMovbarco.SetFocus;

  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
  BotonPermiso.permisosBotones(frmBarra3);
  BotonPermiso.permisosBotones(frmBarra4);
end;

procedure TfrmAdmonyTiempos.frmBarra1btnAddClick(Sender: TObject);
var
  sCadena,
  sDato : string;
  i : Integer;
begin
  lRecalculo :=False;
  if ReporteLock then
  begin
      messageDLg('El Reporte Diario se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
      exit;
  end;

  chkContinuaMov.Checked := False;
  dbMovBarco.Enabled := False;
  tabexistencias.PageControl.Pages[1].TabVisible := false;
  tabexistencias.PageControl.Pages[2].TabVisible := false;
  tabexistencias.PageControl.Pages[3].TabVisible := false;
  sOpcionEmb := 'Inserta';

  frmBarra1.btnAddClick(Sender);
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;
  tsIdBarco.SetFocus;

  //soad -> Trae el horario anterior para mayor facilidad de captura..
  if MovimientosdeBarco.RecordCount > 0 then
  begin
    Movimientosdebarco.Last;
    if movimientosdebarco.FieldValues['sHoraFinal'] <> '24:00' then
    begin
      mkHora1.Text := movimientosdebarco.FieldValues['sHoraFinal'];
      mkHora2.Text := movimientosdebarco.FieldValues['sHoraFinal'];
    end
    else
    begin
      mkHora1.Text := '00:00';
      mkHora2.Text := '00:00';
    end
  end;
  sHoraI := '';
  sHoraF := '';
  CmbFolios.Enabled:=true;
  CbPartidas.Enabled := True;

  movimientosdebarco.Append;
  //activapop2(movembarcacion, popupprincipal);
  if tsOrdenesSeleccion.Text = '<<Todas>>' then
     tsOrdenes.KeyValue := global_contrato
  else
     tsOrdenes.KeyValue := tsOrdenesSeleccion.Text;
  tsClasificaciones.KeyValue := FocusClasificacion;
  tsOrdenes.KeyValue := FocusOrdenes;
  InicializarCheckComboNew(CmbFolios, FocusiFolio);
  tsNumeroActividad.KeyValue  :=  FocusPartida;

  movimientosdebarco.FieldValues['sContrato'] := Global_Contrato_Barco;
  movimientosdebarco.FieldValues['sOrden']    := zqOrdenes.FieldValues['sContrato'];
  movimientosdebarco.FieldValues['dIdFecha']  := tdFecha.Date;
  movimientosdebarco.FieldValues['eNavegando']  :='No';    
  tsIdBarco.KeyValue := embarcaciones.FieldByName('sIdEmbarcacion').AsString;
  movimientosdebarco.FieldValues['sIdEmbarcacion'] :=  tsIdBarco.KeyValue;

  BotonPermiso.permisosBotones(frmBarra1);
end;


procedure TfrmAdmonyTiempos.frmBarra1btnPostClick(Sender: TObject);
var
  sFactor: string;
  dProrrateo: Double;
  iJornada: Integer;
  dAjuste: Double;
  iMultiplo: Integer;
  iDecimales: Integer;
  lContinua: Boolean;
  lIncorrecto: Boolean;
  nombres, cadenas: TStringList;
begin
  frmBarra1.btnPost.SetFocus;
  {Validacion de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('Embarcacion'); nombres.Add('Clasificacion');
  nombres.Add('Hora Inicio'); nombres.Add('Hora Final');
  cadenas.Add(tsIdBarco.Text); cadenas.Add(tsClasificaciones.Text);
  cadenas.Add(mkHora1.Text); cadenas.Add(mkHora2.Text);
  if not validaTexto(nombres, cadenas, '', '') then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;

  {Validacion de barco hora > 24:00}
  lIncorrecto := False;
  if (StrToInt(copy(mkHora1.Text, 1, 2)) = 24) and (StrToInt(copy(mkHora1.Text, 4, 5)) > 0) then
    lIncorrecto := True
  else
    if (StrToInt(copy(mkHora1.Text, 1, 2)) = 25) then
      lIncorrecto := True;

  if lIncorrecto then
  begin
    messageDLG('La Hora de Inicio es Mayor a 24:00 Hrs', mtInformation, [mbOk], 0);
    exit;
  end;

  if (StrToInt(copy(mkHora2.Text, 1, 2)) = 24) and (StrToInt(copy(mkHora2.Text, 4, 5)) > 0) then
    lIncorrecto := True
  else
    if (StrToInt(copy(mkHora2.Text, 1, 2)) = 25) then
      lIncorrecto := True;

  if lIncorrecto then
  begin
    messageDLG('La Hora de Final es Mayor a 24:00 Hrs', mtInformation, [mbOk], 0);
    exit;
  end;

  {Continua insercion de datos}
  tabexistencias.PageControl.Pages[1].TabVisible := true;
  tabexistencias.PageControl.Pages[2].TabVisible := true;
  tabexistencias.PageControl.Pages[3].TabVisible := true;
  try
     if (mkHora1.Text = '  :  ') or (mkHora2.Text = '  :  ') then
     begin
         ShowMessage('Los horarios no deben estar vacios!!');
         mkHora2.SetFocus;
     end
     else
     begin
        if mkHora1.Text > mkHora2.Text then
           ShowMessage('La hora de inicio es menor que la hora final!!');
        if not lRecalculo then begin

        connection.configuracion.refresh;
        iDecimales := Connection.configuracion.FieldValues['iRedondeoEmbarcacion'];
        if movimientosdebarco.State <> dsInsert then
          lContinua := True
        else
          lContinua := False;
        movimientosdebarco.FieldValues['sIdDepartamento'] := global_depto;
        movimientosdebarco.FieldValues['mDescripcion']    := tmDescripcion2.Text;
        movimientosdebarco.FieldValues['sHoraInicio']     := mkHora1.Text;
        movimientosdebarco.FieldValues['sHoraFinal']      := mkHora2.Text;

        MovimientosDeBarco.FieldByName('sIdFase').AsString        := 'OPER';
        MovimientosDeBarco.FieldByName('sClasificacion').AsString := tsClasificaciones.KeyValue;

        if mkHora1.Text > mkHora2.Text then
           movimientosdebarco.FieldValues['sFactor'] := '0.000000'
        else
        begin
//            if (sHoraI <> mkHora1.Text) or (sHoraF <> mkHora2.Text) then
//            begin
                if (mkHora1.Text = '00:00') and (mkHora2.Text = '24:00') then
                  movimientosdebarco.FieldValues['sFactor'] := '1'
                else
                    movimientosdebarco.FieldValues['sFactor'] := sfnFactor(mkHora1.Text, mkHora2.Text, 24, iDecimales);
//            end;
        end;

        if chkAplicafactor.Checked = False then
        begin
           movimientosdebarco.FieldValues['sFactor'] := '0.000000';
           movimientosdebarco.FieldValues['sActividades'] := 'MOV';
        end
        else
           movimientosdebarco.FieldValues['sActividades'] := '';

        if tmDescripcion2.Text = '' then
          movimientosdebarco.FieldValues['mDescripcion'] := '*'
        else
          movimientosdebarco.FieldValues['mDescripcion'] := tmDescripcion2.Text;

        MovimientosdeBarco.FieldValues['sTipo'] := clasificaciones.FieldValues['sTipo'];

        if chkContinuaMov.Checked then
           MovimientosdeBarco.FieldValues['lContinuo'] := 'Si'
        else
           MovimientosdeBarco.FieldValues['lContinuo'] := 'No';

        if clasificaciones.FieldValues['lGenera'] = 'Si' then
           MovimientosdeBarco.FieldValues['sNumeroActividad'] := Clasificaciones.FieldValues['sIdTipoMovimiento']
        else
           MovimientosdeBarco.FieldValues['sNumeroActividad'] := zqPartida.FieldValues['sNumeroActividad'];
        //chkContinuaMov.Visible := False;
        try
         // MovimientosdeBarco.FieldByName('sactividades').AsString := OrdenarCadena(CbPartidas.text);
          movimientosdebarco.Post;

          GrabarCheckCombo(CmbFolios);

          AjusteProrrateo;

          //TdAjustaFolio(MovimientosdeBarco.FieldByName('sContrato').AsString,tdFecha.Date);

          desactivapop(popupprincipal);
        except
          on e: exception do begin
            UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al salvar registro en movimientos de embarcacion, al calcular total diario', 0);
          end;
        end;
        end else begin
          movimientosdebarco.Cancel;
        end;

        if sOpcionEmb = 'Inserta' then
        begin
            //soad -> Trae el horario anterior para mayor facilidad de captura..
            if lContinua = False then
            begin
              if MovimientosdeBarco.RecordCount > 0 then
              begin
                Movimientosdebarco.Last;
                if movimientosdebarco.FieldValues['sHoraFinal'] <> '24:00' then
                begin
                  mkHora1.Text := movimientosdebarco.FieldValues['sHoraFinal'];
                  mkHora2.Text := movimientosdebarco.FieldValues['sHoraFinal'];
                end
                else
                begin
                  mkHora1.Text := '00:00';
                  mkHora2.Text := '00:00';
                end
              end;
            end
            else
               sOpcionEmb := '';

           FocusClasificacion := tsClasificaciones.KeyValue;
           FocusOrdenes  := tsOrdenes.KeyValue;
           FocusPartida  := tsNumeroActividad.Text;
        end
        else
        begin
             SavePlace := dbMovBarco.DataSource.DataSet.GetBookmark ;

             Movimientosdebarco.Refresh;
             Try
                 dbMovBarco.DataSource.DataSet.GotoBookmark(SavePlace);
             Except
             Else
                 dbMovBarco.DataSource.DataSet.FreeBookmark(SavePlace);
             End ;
             sOpcionEmb := '';
        end;

        CmbFolios.Enabled:=False;
        CbPartidas.Enabled := False;
        chkContinuaMov.Checked := False;
        tmDescripcion2.Text := '';
        sHoraI := '';
        sHoraF := '';
        Insertar1.Enabled := True;
        Editar1.Enabled := True;
        Registrar1.Enabled := False;
        Can1.Enabled := False;
        Eliminar1.Enabled := True;
        Refresh1.Enabled := True;
        Salir1.Enabled := True;
        mkHora1.ReadOnly := true;
        dbMovBarco.Enabled := True;
        frmBarra1.btnCancelClick(Sender);

        if sOpcionEmb = 'Inserta' then
           frmBarra1.btnAdd.Click;
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al salvar registro en movimientos de embarcacion', 0);
      frmbarra1.btnCancel.Click;
    end;

  end;
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmAdmonyTiempos.frmBarra1btnPrinterClick(Sender: TObject);
begin
  Application.CreateForm(TfrmOpcionesBarco, frmOpcionesBarco);
  frmOpcionesBarco.showModal;
end;

procedure TfrmAdmonyTiempos._tsClasificacionesEnter(Sender: TObject);
begin
  tsHoraInicio.Color := Global_Color_Entrada;
end;

procedure TfrmAdmonyTiempos._tsClasificacionesExit(Sender: TObject);
begin
  tsClasificaciones.Color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos._tsClasificacionesKeyPress(Sender: TObject;
  var Key: Char);
begin
  {if key = #13 then
    tsIdFase.SetFocus; }
end;

procedure TfrmAdmonyTiempos.tsIdBarcoEnter(Sender: TObject);
begin
  tsIdBarco.Color := global_color_entrada
end;

procedure TfrmAdmonyTiempos.tsIdBarcoExit(Sender: TObject);
begin
    tsIdBarco.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.tsIdBarcoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsClasificaciones.SetFocus
end;

procedure TfrmAdmonyTiempos.tsIdFaseEnter(Sender: TObject);
begin
  tsIdfase.Color := Global_Color_Entrada;
end;

procedure TfrmAdmonyTiempos.tsIdFaseExit(Sender: TObject);
begin
  tsIdFase.Color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos.tsIdFaseKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    mkHora2.SetFocus;
end;

procedure TfrmAdmonyTiempos.tsLocalizacionEnter(Sender: TObject);
begin
    tsLocalizacion.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.tsLocalizacionExit(Sender: TObject);
begin
     tsLocalizacion.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.tsLocalizacionKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key =#13 then
       sPronostico.SetFocus;
end;

procedure TfrmAdmonyTiempos.tsNumeroActividadEnter(Sender: TObject);
begin
    if tsNumeroActividad.Text = '' then
    begin
        if clasificaciones.FieldValues['lGenera'] = 'No' then
           cmbFolios.OnExit(sender);
    end;

end;

procedure TfrmAdmonyTiempos.tsOrdenesExit(Sender: TObject);
begin
    CargarCheckCombo(CmbFolios);
end;

procedure TfrmAdmonyTiempos.tsOrdenesKeyPress(Sender: TObject; var Key: Char);
begin
    if key = #13 then
       CmbFolios.SetFocus;
end;

procedure TfrmAdmonyTiempos.tsOrdenesSeleccionExit(Sender: TObject);
var
   buscar, filtro : string;
begin
//    Movimientosdebarco.Active := False;
//    MovimientosdeBarco.Sql.Clear;
//    MovimientosdeBarco.Sql.Add('select movimientosdeembarcacion.* from movimientosdeembarcacion ' +
//      'inner join tiposdemovimiento  on ' +
//      '(tiposdemovimiento.sContrato = :Contrato ' +
//      ' And movimientosdeembarcacion.sClasificacion = tiposdemovimiento.sIdTipoMovimiento) ' +
//      'where movimientosdeembarcacion.dIdFecha = :Fecha and movimientosdeembarcacion.sOrden like :ContratoNormal and sIdFase = "OPER" order by sIdEmbarcacion, sHoraInicio ');
//    movimientosdebarco.Params.ParamByName('Contrato').DataType := ftString;
//    movimientosdebarco.Params.ParamByName('Contrato').Value    := Global_Contrato_Barco;
//    movimientosdebarco.Params.ParamByName('ContratoNormal').DataType := ftString;
//    if tsOrdenesSeleccion.Text = '<<Todas>>' then
//       movimientosdebarco.Params.ParamByName('ContratoNormal').Value := '%'
//    else
//       movimientosdebarco.Params.ParamByName('ContratoNormal').Value := tsOrdenesSeleccion.Text;
//    movimientosdebarco.Params.ParamByName('Fecha').DataType    := ftDate;
//    movimientosdebarco.Params.ParamByName('Fecha').Value       := tdFecha.Date;
//    movimientosdebarco.Open;

      if tsOrdenesSeleccion.Text = '<<Todas>>' then
         buscar := buscar + '*'
      else
      begin
          buscar := tsOrdenesSeleccion.Text;
          buscar := buscar + '*'
      end;
      filtro := ' sOrden like ' + QuotedStr(buscar);

      movimientosdebarco.Filtered := False;
      movimientosdebarco.Filter   := filtro;
      movimientosdebarco.Filtered := True;
end;

procedure TfrmAdmonyTiempos.tsUbicacionBarcoEnter(Sender: TObject);
begin
    tsUbicacionBarco.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.tsUbicacionBarcoExit(Sender: TObject);
begin
    tsUbicacionBarco.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.zqMovimientosxfolioCalcFields(DataSet: TDataSet);
var
   zHorarios: TZQuery;
begin
    zHorarios := TZQuery.Create(Self);
    zHorarios.Connection := Connection.zConnection;
    if zqMovimientosxFolio.RecordCount > 0 then
    begin
        zHorarios.Active := False;
        zHorarios.SQL.Clear;
        zHorarios.SQL.Add('select sHoraInicio, sHoraFinal from movimientosdeembarcacion where sContrato =:Contrato '+
                          'and dIdFecha =:Fecha and iIdDiario =:Diario');
        zHorarios.ParamByName('contrato').AsString := global_contrato_barco;
        zHorarios.ParamByName('fecha').AsDate      := tdFecha.Date;
        zHorarios.ParamByName('Diario').AsInteger  := zqMovimientosxfolio.FieldValues['iIdDiario'];
        zHorarios.Open;

        if zHorarios.RecordCount > 0 then
        begin
            zqMovimientosxfoliosHoraInicio.Value := zHorarios.FieldValues['sHoraInicio'];
            zqMovimientosxfoliosHoraFinal.Value  := zHorarios.FieldValues['sHoraFinal'];
        end;
    end;
   zHorarios.Destroy;
end;

procedure TfrmAdmonyTiempos.zqOrdenaOrdenBeforePost(DataSet: TDataSet);
begin
    if sColumaOrden = 'Diesel' then
    begin
        //Verificamos si existe una orden que ya tiene el Recibido de Diesel asignado.
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('select * from recursosordenados_orden where dIdFecha =:fecha and sOrden <> :orden and lAplicaRecibidoDiesel = "Si"');
        connection.QryBusca.ParamByName('Orden').AsString := zqOrdenaOrden.FieldValues['sOrden'];
        connection.QryBusca.ParamByName('fecha').AsDate   := tdFecha.Date;
        connection.QryBusca.Open;

        if connection.QryBusca.RecordCount > 0 then
        begin
            MessageDlg('No se puede Asignar a mas de una OT la Recepcion de Diesel!', mtWarning, [mbOk], 0);
            zqOrdenaOrden.Cancel;
        end;
    end;

    if sColumaOrden = 'Agua' then
    begin
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('select * from recursosordenados_orden where dIdFecha =:fecha and sOrden <> :orden and lAplicaRecibidoAgua = "Si"');
        connection.QryBusca.ParamByName('Orden').AsString := zqOrdenaOrden.FieldValues['sOrden'];
        connection.QryBusca.ParamByName('fecha').AsDate   := tdFecha.Date;
        connection.QryBusca.Open;

        if connection.QryBusca.RecordCount > 0 then
        begin
            MessageDlg('No se puede Asignar a mas de una OT la Recepcion de Agua!', mtWarning, [mbOk], 0);
            zqOrdenaOrden.Cancel;
        end;
    end;
    sColumaOrden := '';
end;

procedure TfrmAdmonyTiempos.zqOrdenaOrdenlAplicaRecibidoAguaChange(
  Sender: TField);
begin
    sColumaOrden := 'Agua';
end;

procedure TfrmAdmonyTiempos.zqOrdenaOrdenlAplicaRecibidoDieselChange(
  Sender: TField);
begin
    sColumaOrden := 'Diesel';
end;

procedure TfrmAdmonyTiempos.zqOrdenesAfterScroll(DataSet: TDataSet);
begin
    if zqOrdenes.RecordCount > 0 then
    begin
        if (Movimientosdebarco.State <> dsEdit) or (Movimientosdebarco.State <> dsInsert) then
            CargarCheckCombo(CmbFolios);

    end;
end;

procedure TfrmAdmonyTiempos.mkHora1KeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    mkHora2.SetFocus
end;

procedure TfrmAdmonyTiempos.mkHora1Enter(Sender: TObject);
begin
  mkHora1.Color := Global_Color_Entrada;
end;

procedure TfrmAdmonyTiempos.mkHora1Exit(Sender: TObject);
begin
    if ValidaHorario(mkHora1.Text) = false then
       mkHora1.SetFocus
    else
       mkHora1.Color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos.mkHora2Enter(Sender: TObject);
begin
  mkHora2.Color := Global_Color_Entrada;
end;

procedure TfrmAdmonyTiempos.mkHora2Exit(Sender: TObject);
begin
    if ValidaHorario(mkhora2.Text) = false then
       mkhora2.SetFocus
    else
       mkhora2.Color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos.mkHora2KeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsOrdenes.SetFocus;
end;

procedure TfrmAdmonyTiempos.movembarcacionEnter(Sender: TObject);
var
  x, i, hora_inicio, hora_final, minuto, total, aux: integer;
  I_hora, I_minuto, F_hora, F_minuto, cadena: string;
  numero: Double;
begin
  if (movimientosdebarco.state = dsInsert) or (movimientosdebarco.state = dsEdit) then
    exit;

  Embarcaciones.Active := False;
  Embarcaciones.SQL.clear;
  Embarcaciones.sql.Add('select sIdEmbarcacion, sDescripcion, sTipo from embarcaciones ' +
    'Where sTipo="Principal" order by sDescripcion');
  Embarcaciones.Open;

  //Consultamos la informacion de barco..
  connection.QryBusca.Active := False;
  connection.QryBusca.Sql.Clear;
  connection.QryBusca.Sql.Add('select movimientosdeembarcacion.* from movimientosdeembarcacion ' +
              'inner join tiposdemovimiento  on ' +
              '(tiposdemovimiento.sContrato = :Contrato ' +
              ' And movimientosdeembarcacion.sClasificacion = tiposdemovimiento.sIdTipoMovimiento) ' +
              'where movimientosdeembarcacion.dIdFecha = :Fecha  order by sActividades, sIdEmbarcacion, sHoraInicio ');
  connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
  connection.QryBusca.Params.ParamByName('Contrato').Value := Global_Contrato_Barco;
  connection.QryBusca.Params.ParamByName('Fecha').DataType := ftDate;
  connection.QryBusca.Params.ParamByName('Fecha').Value := tdFecha.Date;
  connection.QryBusca.Open;

  //soad -> Verificacion si estan completas las 24 horas..
  total := 0;
  if connection.QryBusca.RecordCount > 0 then
  begin
      //state := connection.QryBusca.state;
      while not connection.QryBusca.Eof do
      begin
             // Horarios...
        I_hora := LeftStr(connection.QryBusca.FieldValues['sHoraInicio'], 2);
        I_minuto := MidStr(connection.QryBusca.FieldValues['sHoraInicio'], 4, 2);

        F_hora := LeftStr(connection.QryBusca.FieldValues['sHoraFinal'], 2);
        F_minuto := MidStr(connection.QryBusca.FieldValues['sHoraFinal'], 4, 2);

             //Horarios a integer..
        hora_inicio := strToInt(I_hora);
        hora_final := strToInt(F_hora);
        aux := hora_final - hora_inicio;
        total := total + (aux * 60);

        hora_inicio := strToInt(I_minuto);
        hora_final := strToInt(F_minuto);
        aux := hora_final - hora_inicio;
        total := total + aux;
        connection.QryBusca.Next
      end;

      //Conversion de Horario..
      if (total < 1440) and (connection.QryBusca.RecordCount > 0) then
      begin
        numero := ((1440 - total) / 60);
        cadena := FloatToStr(int(numero));
        aux := strToint(cadena);
        total := ((1440 - total) - (aux * 60));
        if int(numero) < 10 then
          cadena := '0' + intTostr(aux)
        else
          cadena := intTostr(aux);

        if total < 10 then
          cadena := cadena + ':0' + intTostr(total)
        else
          cadena := cadena + ':' + intTostr(total);
             //Muestra mensaje de Horas faltantes...
//        messageDLG(' Movimientos de Embarcacion Incompletos, Faltan : ' + cadena + ' minutos', mtInformation, [mbOK], 0);
      end;
  end;
end;

procedure TfrmAdmonyTiempos.frmBarra1btnExitClick(Sender: TObject);
begin
  frmBarra1.btnExitClick(Sender);
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  close

end;

procedure TfrmAdmonyTiempos.frmBarra1btnDeleteClick(Sender: TObject);
var
   QrDatos  : TZReadOnlyQuery;
   IdDiario : integer;
begin
  //  if connection.contrato.FieldValues['sTipoObra'] <> 'BARCO' then
  //  begin
  //      messageDLG('SELECCIONE CONTRATO DE BARCO!', mtInformation, [mbOk], 0);
  //      exit;
  //  end;

  if ReporteLock then
  begin
      messageDLg('El Reporte Dairio se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
      exit;
  end;

  if movimientosdebarco.RecordCount > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
       mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
          // Aqui se verifica si el reporte no se encuentra autorizado ....
        IdDiario := MovimientosdeBarco.FieldByName('iidDiario').AsInteger;
        movimientosdebarco.Delete;

        QrDatos:=TZReadOnlyQuery.Create(nil);
        QrDatos.Connection:=connection.zConnection;

        QrDatos.Active:=False;
        QrDatos.SQL.Text:='delete from movimientosxfolios where sContrato=:Contrato and dIdFecha=:fecha and iIdDiario=:Diario' ;
        QrDatos.ParamByName('Contrato').AsString := MovimientosdeBarco.FieldByName('sContrato').AsString;
        QrDatos.ParamByName('Fecha').AsDate      := MovimientosdeBarco.FieldByName('dIdFecha').AsDateTime;
        QrDatos.ParamByName('Diario').AsInteger  := IdDiario;
        QrDatos.ExecSQL;

      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al eliminar registro en movimientos de embarcación', 0);
        end;
      end
    end

end;

procedure TfrmAdmonyTiempos.frmBarra1btnRefreshClick(Sender: TObject);
begin
  movimientosdebarco.Refresh;
  //PanelMovimientos.Enabled := False;
end;

procedure TfrmAdmonyTiempos.frmBarra1btnEditClick(Sender: TObject);
begin

  if movimientosdebarco.RecordCount < 1 then
    exit;

  if ReporteLock then
  begin
    messageDLg('El Reporte Diario se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;
  dbMovBarco.Enabled := False;
  tabexistencias.PageControl.Pages[1].TabVisible := false;
  tabexistencias.PageControl.Pages[2].TabVisible := false;
  tabexistencias.PageControl.Pages[3].TabVisible := false;

  try
    movimientosdebarco.Edit;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al editar registro en movimientos de embarcación', 0);
      exit;
    end;

  end;
  sHoraI := mkHora1.Text;
  sHoraF := mkHora2.Text;
  tmDescripcion2.Text := movimientosdebarco.FieldValues['mDescripcion'];
  
  CmbFolios.Enabled:=true;
  CbPartidas.Enabled := True;
  sOpcionEmb := 'Edita';
  frmBarra1.btnEditClick(Sender);
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;
  movimientosdebarco.Edit;
  tsIdBarco.SetFocus;
  BotonPermiso.permisosBotones(frmBarra1);

end;

procedure TfrmAdmonyTiempos.frmBarra1btnCancelClick(Sender: TObject);
begin
  tabexistencias.PageControl.Pages[1].TabVisible := true;
  tabexistencias.PageControl.Pages[2].TabVisible := true;
  tabexistencias.PageControl.Pages[3].TabVisible := true;
  sOpcionEmb := '';
  movimientosdebarco.Cancel;
  frmBarra1.btnCancelClick(Sender);
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  mkHora1.ReadOnly := False;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  dbMovBarco.Enabled := True;
  //PanelMovimientos.Enabled := False;
  CmbFolios.Enabled:=False;
  CbPartidas.Enabled := fALSE;
end;

Function TfrmAdmonyTiempos.ColumnaNombre(Numero: Integer): String;
Var
  Valor, NumLetras: Integer;
  Cad: String;
Begin
  NumLetras := 26;  // O1
  Cad := '';
  Valor := Numero Mod NumLetras;
  if Valor = 0 then Valor := 26;
  if Numero - Valor > 0 then Cad := Char(64 + Trunc((Numero - Valor) / NumLetras));
  Cad := Cad + Char(64 + Valor);

  Result := Cad;
End;

procedure TfrmAdmonyTiempos.BtnGenPlantillaClick(Sender: TObject);
var curgp:TCursor;
begin
  curgp := Screen.Cursor;
  try
    Screen.Cursor := crAppStart;
    GeneraPlantillaImp();
  finally
    Screen.Cursor := curgp;
  end;
end;

procedure TfrmAdmonyTiempos.GeneraPlantillaImp;
var ExcelEx,
    LibroEx,
    HojaEx: Variant;

    Cont:Integer;
begin
  //iniciando ExcelEx
  try
    ExcelEx:=CreateOleObject('Excel.Application');
  except
    raise Exception.Create('No tiene instalado microsoft excel o bién ocurre un problema con el mismo.');
  end;

  LibroEx := ExcelEx.Workbooks.Add;

  HojaEx := LibroEx.Sheets.Add;

  HojaEx.Name:='CARATULA';

  ExcelEx.Visible := True;

  {$REGION 'Ancho de todas columnas'}
  HojaEx.Columns[1].ColumnWidth:= 5.57;
  HojaEx.Columns[2].ColumnWidth:= 5.57;
  HojaEx.Columns[3].ColumnWidth:= 5.57;
  HojaEx.Columns[4].ColumnWidth:= 5.57;
  HojaEx.Columns[5].ColumnWidth:= 14.57;
  HojaEx.Columns[6].ColumnWidth:= 48;
  HojaEx.Columns[7].ColumnWidth:= 16.29;
  HojaEx.Columns[8].ColumnWidth:= 7.57;
  HojaEx.Columns[9].ColumnWidth:= 17.71;
  HojaEx.Columns[10].ColumnWidth:= 14.43;
  HojaEx.Columns[11].ColumnWidth:= 13.29;
  HojaEx.Columns[12].ColumnWidth:= 13;
  {$ENDREGION}

  {$REGION 'Cuadro superior izquierdo'}
  HojaEx.Range['C2:E2'].MergeCells := True;
  HojaEx.Range['C3:E3'].MergeCells := True;
  HojaEx.Range['C4:E4'].MergeCells := True;
  HojaEx.Range['C5:E5'].MergeCells := True;
  HojaEx.range['A2:E5'].Borders.LineStyle := xlContinuous;
  HojaEx.range['A2:E5'].HorizontalAlignment:=-4108;
  HojaEx.range['A2:E5'].Font.Bold:=True;
  HojaEx.range['A2:E5'].Font.size := 10;
  HojaEx.range['A2:E5'].interior.colorindex := clgris;
  HojaEx.range['A2:B5'].NumberFormat := '@';
  HojaEx.range['C2:C5'].NumberFormat := '0.000000';

  HojaEx.cells[2,1].value := '1.1';   HojaEx.cells[2,2].value := 'CDP'; HojaEx.cells[2,3].formula := '=SUMIF(H9:H17,A2,K9:K17)';
  HojaEx.cells[3,1].value := '1.2';   HojaEx.cells[3,2].value := 'SDP'; HojaEx.cells[3,3].formula := '=SUMIF(H9:H17,A3,K9:K17)';
  HojaEx.cells[4,1].value := '1.3';   HojaEx.cells[4,2].value := 'ESP'; HojaEx.cells[4,3].formula := '=SUMIF(H9:H17,A4,K9:K17)';
  HojaEx.cells[5,1].value := 'S/N';   HojaEx.cells[5,2].value := 'CIA'; HojaEx.cells[5,3].formula := '=SUMIF(H9:H17,A5,K9:K17)';

  HojaEx.range['A2:A2'].AddComment('Referencias de movimientos.');
  HojaEx.range['A2:A2'].Comment.Visible := False;
  {$ENDREGION}

  {$REGION 'Cuadro superior derecho'}
  HojaEx.Range['H2:J2'].MergeCells := True;
  HojaEx.Range['H3:I3'].MergeCells := True;
  HojaEx.Range['K3:L3'].MergeCells := True;
  HojaEx.Range['H4:I4'].MergeCells := True;
  HojaEx.Range['K4:L4'].MergeCells := True;
  HojaEx.Range['H5:L5'].MergeCells := True;

  HojaEx.range['G2:L5'].Borders.LineStyle := xlContinuous;
  HojaEx.range['G2:L5'].HorizontalAlignment:=-4108;
  HojaEx.range['G2:L5'].Font.Bold:=True;
  HojaEx.range['G2:L5'].Font.size := 11;

  HojaEx.range['G2:G5'].interior.colorindex := Clgris;
  HojaEx.range['J3:J4'].interior.colorindex := Clgris;
  HojaEx.range['K2:K2'].interior.colorindex := Clgris;

  HojaEx.cells[2,8].NumberFormat := 'dddd"," dd "de" mmmm "de" AAA';
  HojaEx.cells[2,12].NumberFormat := '[hh]:mm';
  HojaEx.cells[3,11].NumberFormat := '@';
  HojaEx.cells[4,11].NumberFormat := '@';


  HojaEx.cells[2,7].value := 'FECHA';       HojaEx.cells[2,8].value := now;              HojaEx.cells[2,11].value := 'HORA';
  HojaEx.cells[3,7].value := 'OLAS';        HojaEx.cells[3,10].value := 'PIES';
  HojaEx.cells[4,7].value := 'VIENTOS';     HojaEx.cells[4,10].value := 'NUDOS';
  HojaEx.cells[5,7].value := 'LOCALIZACION';

  HojaEx.range['H2:H2'].AddComment('Fecha a importar.');
  HojaEx.range['H2:H2'].Comment.Visible := False;
  HojaEx.range['L2:L2'].AddComment('Horario de condiciones climatológicas en formato 00:00 rango [00:00-24:00].');
  HojaEx.range['L2:L2'].Comment.Visible := False;

  HojaEx.range['H3:H3'].AddComment('Dirección con comillas ej: "NE", "NO", verificar'+#10+' que la direcion exista en el sistema');
  HojaEx.range['H3:H3'].Comment.Visible := False;
  HojaEx.range['H4:H4'].AddComment('Dirección con comillas ej: "NE", "NO", verificar'+#10+' que la direcion exista en el sistema');
  HojaEx.range['H4:H4'].Comment.Visible := False;
  {$ENDREGION}

  {$REGION 'Cuadro movimientos de embarcacion'}
  //Encabezado
  HojaEx.Range['A7:G7'].MergeCells := True;
  HojaEx.Range['A8:D8'].MergeCells := True;
  HojaEx.Range['E8:G8'].MergeCells := True;
  HojaEx.range['A7:G7'].interior.colorindex := clgris;

  HojaEx.range['A7:G7'].Borders.LineStyle := xlContinuous;
  HojaEx.range['A7:G7'].HorizontalAlignment:=-4108;
  HojaEx.range['A7:G7'].Font.Bold:=True;
  HojaEx.range['A7:G7'].Font.size := 14;
  HojaEx.range['A8:L8'].Borders.LineStyle := xlContinuous;
  HojaEx.range['A8:L8'].HorizontalAlignment:=-4108;
  HojaEx.range['A8:L8'].Font.Bold:=True;
  HojaEx.range['A8:L8'].Font.size := 11;
  HojaEx.range['A8:L8'].interior.colorindex := clgris;

  HojaEx.cells[7,1].value := 'MOVIMIENTOS DE EMBARCACION';
  HojaEx.cells[8,1].value := 'HORARIO';
  HojaEx.cells[8,5].value := 'DESCRIPCION';
  HojaEx.cells[8,8].value := 'PART';
  HojaEx.cells[8,9].value := 'FOLIO';
  HojaEx.cells[8,10].value := 'OT';
  HojaEx.cells[8,11].value := 'TIEMPO';
  HojaEx.cells[8,12].value := 'CLAS.';

  HojaEx.range['A8:A8'].AddComment('Verifique que los horarios sigan el formato 00:00 - 24:00,'+#10+'Si el horario 24:00 muestra un valor de fecha desconocido '+#10+'no importa ya que sólo se toma el horario mostrado.');
  HojaEx.range['A8:A8'].Comment.Visible := False;

  HojaEx.range['E8:E8'].AddComment('Descripción del movimiento.');
  HojaEx.range['E8:E8'].Comment.Visible := False;

  HojaEx.range['H8:H8'].AddComment('La partida debe regirse de acuerdo al cuadro superior izquierdo'+#10'Sólo debe cargar los identificadores ejemplo: 1.1 o 1.2 etc...'+#10+' y pulsar enter para que la columna L muestre el valor corrrecto.');
  HojaEx.range['H8:H8'].Comment.Visible := False;

  HojaEx.range['I8:I8'].AddComment('Puede importar folios separados por coma,'+#10+' sin embargo verifique que esos fólios esten dados de alta en la OT.');
  HojaEx.range['I8:I8'].Comment.Visible := False;

  HojaEx.range['J8:J8'].AddComment('Ingresar orden de trabajo sin espacios tomando en cuenta'+#10+' que los fólios ingresados en la columna I pertenecen a la OT.');
  HojaEx.range['J8:J8'].Comment.Visible := False;

  HojaEx.range['K8:K8'].AddComment('Calculado sólo con fines informativos.');
  HojaEx.range['K8:K8'].Comment.Visible := False;

  //Cuerpo
  HojaEx.range['A9:L18'].Borders.LineStyle := xlContinuous;
  HojaEx.range['A9:L18'].HorizontalAlignment:=-4108;
  HojaEx.range['A9:L18'].Font.Bold:=false;
  HojaEx.range['A9:L18'].Font.size := 9;
  for Cont := 9 to 18 do
  begin
    HojaEx.Range['A'+inttostr(Cont)+':B'+inttostr(Cont)].MergeCells := True;
    HojaEx.Range['C'+inttostr(Cont)+':D'+inttostr(Cont)].MergeCells := True;
    HojaEx.Range['E'+inttostr(Cont)+':G'+inttostr(Cont)].MergeCells := True;

    HojaEx.cells[cont,1].NumberFormat := '[hh]:mm';
    HojaEx.cells[cont,3].NumberFormat := '[hh]:mm';

    HojaEx.cells[cont,Cpartida].NumberFormat := '@';
    HojaEx.cells[cont,CDescripcion].NumberFormat := '@';
    HojaEx.cells[cont,CFolio].NumberFormat := '@';
    HojaEx.cells[cont,COt].NumberFormat := '@';
    HojaEx.cells[cont,11].NumberFormat := '0.000000';

    HojaEx.cells[cont,11].formula := '=+ROUND(C'+inttostr(cont)+'-A'+inttostr(cont)+',6)';
    HojaEx.cells[cont,12].formula := '=VLOOKUP(H'+inttostr(cont)+',$A$2:$B$5,2)';

  end;
  HojaEx.range['A18:L18'].interior.colorindex := 41;
  HojaEx.Rows[18].RowHeight:= 6;

  HojaEx.range['K19:K19'].Borders.LineStyle := xlContinuous;
  HojaEx.cells[19,11].formula := '=SUM(K9:K17)';
  HojaEx.cells[19,11].NumberFormat := '0.000000';
  {$ENDREGION}

  {$REGION 'Cuadro embarcaciones'}
  //Encabezado
  HojaEx.Range['A20:G20'].MergeCells := True;
  HojaEx.Range['A21:D21'].MergeCells := True;
  HojaEx.Range['E21:G21'].MergeCells := True;
  HojaEx.range['A20:G21'].interior.colorindex := clgris;
  HojaEx.range['A20:G21'].Borders.LineStyle := xlContinuous;
  HojaEx.range['A20:G21'].HorizontalAlignment:=-4108;
  HojaEx.range['A20:G21'].Font.Bold:=True;
  HojaEx.range['A20:G20'].Font.size := 14;
  HojaEx.range['A21:G21'].Font.size := 11;

  HojaEx.cells[20,1].value := '<*** EMBARCACIONES ***>';
  HojaEx.cells[21,1].value := 'HORARIO';
  HojaEx.cells[21,5].value := 'DESCRIPCION';

  HojaEx.range['A21:A21'].AddComment('Verifique que los horarios sigan el formato 00:00 - 24:00,'+#10+'Si el horario 24:00 muestra un valor de fecha desconocido '+#10+'no importa ya que sólo se toma el horario mostrado.');
  HojaEx.range['A21:A21'].Comment.Visible := False;

  HojaEx.range['E21:E21'].AddComment('Descripción del movimiento.');
  HojaEx.range['E21:E21'].Comment.Visible := False;

  //Cuerpo
  HojaEx.range['A22:G25'].Borders.LineStyle := xlContinuous;
  HojaEx.range['A22:G25'].HorizontalAlignment:=-4108;
  HojaEx.range['A22:G25'].Font.Bold:=false;
  HojaEx.range['A22:G25'].Font.size := 9;
  for Cont := 22 to 25 do
  begin
    HojaEx.Range['A'+inttostr(Cont)+':B'+inttostr(Cont)].MergeCells := True;
    HojaEx.Range['C'+inttostr(Cont)+':D'+inttostr(Cont)].MergeCells := True;
    HojaEx.Range['E'+inttostr(Cont)+':G'+inttostr(Cont)].MergeCells := True;
    HojaEx.cells[cont,1].NumberFormat := '[hh]:mm';
    HojaEx.cells[cont,3].NumberFormat := '[hh]:mm';
    HojaEx.cells[cont,CDescripcion].NumberFormat := '@';
  end;
  HojaEx.range['A25:G25'].interior.colorindex := 41;
  {$ENDREGION}

  {$REGION 'Cuadro hELICOPTEROS'}
  //Encabezado
  HojaEx.Range['A27:G27'].MergeCells := True;
  HojaEx.Range['A28:D28'].MergeCells := True;
  HojaEx.Range['E28:G28'].MergeCells := True;
  HojaEx.range['A27:G28'].interior.colorindex := clgris;
  HojaEx.range['A27:G28'].Borders.LineStyle := xlContinuous;
  HojaEx.range['A27:G28'].HorizontalAlignment:=-4108;
  HojaEx.range['A27:G28'].Font.Bold:=True;
  HojaEx.range['A27:G27'].Font.size := 14;
  HojaEx.range['A28:G28'].Font.size := 11;

  HojaEx.cells[27,1].value := '<*** HELICÓPTEROS ***>';
  HojaEx.cells[28,1].value := 'HORARIO';
  HojaEx.cells[28,5].value := 'DESCRIPCION';

  HojaEx.range['A28:A28'].AddComment('Verifique que los horarios sigan el formato 00:00 - 24:00,'+#10+'Si el horario 24:00 muestra un valor de fecha desconocido '+#10+'no importa ya que sólo se toma el horario mostrado.');
  HojaEx.range['A28:A28'].Comment.Visible := False;

  HojaEx.range['E28:E28'].AddComment('Descripción del movimiento.');
  HojaEx.range['E28:E28'].Comment.Visible := False;

  //Cuerpo
  HojaEx.range['A29:G34'].Borders.LineStyle := xlContinuous;
  HojaEx.range['A29:G34'].HorizontalAlignment:=-4108;
  HojaEx.range['A29:G34'].Font.Bold:=false;
  HojaEx.range['A29:G34'].Font.size := 9;
  for Cont := 29 to 34 do
  begin
    HojaEx.Range['A'+inttostr(Cont)+':B'+inttostr(Cont)].MergeCells := True;
    HojaEx.Range['C'+inttostr(Cont)+':D'+inttostr(Cont)].MergeCells := True;
    HojaEx.Range['E'+inttostr(Cont)+':G'+inttostr(Cont)].MergeCells := True;
    HojaEx.cells[cont,1].NumberFormat := '[hh]:mm';
    HojaEx.cells[cont,3].NumberFormat := '[hh]:mm';
    HojaEx.cells[cont,CDescripcion].NumberFormat := '@';
  end;
  HojaEx.range['A34:G34'].interior.colorindex := 41;
  {$ENDREGION}

  {$REGION 'Cuadro posterior derecho'}
  HojaEx.Range['K36:K38'].MergeCells := True;

  HojaEx.range['I36:L38'].Borders.LineStyle := xlContinuous;
  HojaEx.range['I36:L38'].HorizontalAlignment:=-4108;
  HojaEx.range['I36:L38'].Font.size := 10;
  HojaEx.range['K36:K38'].VerticalAlignment:= -4108;

  HojaEx.cells[36,9].value := 'MAQ';  HojaEx.cells[36,11].value := '24'; HojaEx.cells[36,12].formula := '=+J36/24*K36';
  HojaEx.cells[37,9].value := 'COMP';                                    HojaEx.cells[37,12].formula := '=+J37/24*K36';
  HojaEx.cells[38,9].value := 'HIDRO';                                   HojaEx.cells[38,12].formula := '=+J38/24*K36';
  {$ENDREGION}

end;

procedure TfrmAdmonyTiempos.BtnImpPlantillaClick(Sender: TObject);
var
  CurOld :TCursor;
begin
  CurOld := Screen.Cursor;
  Screen.Cursor := crAppStart;
  try
    if not ZMovtos.Active then
      ZMovtos.Open;
    if not ZFolios.Active then
      ZFolios.Open;
  finally
    screen.cursor := curold;
  end;
  CurOld := Screen.Cursor;
  Screen.Cursor := crAppStart;
  try
    ImportarMovtosBarco(global_Contrato_Barco,Global_nombre_Embarcacion);
  finally
    screen.cursor := curold;
  end;
end;

procedure TfrmAdmonyTiempos.BtnImprimirClick(Sender: TObject);
var RFi,Rff:TDateTime;
    CFi,CFf:string;
    annos,mess,diass:string;
begin

  annos := CmbAnno.Text;
  Mess := inttostr(CmbMeses.ItemIndex+1);
  if Length(mess) = 1 then
    mess := '0'+mess;
  RFi := StrToDate('01/'+mess+'/'+annos);

  diass := vartostr(DaysInMonth(Rfi));
  if Length(diass) = 1 then
    diass := '0'+diass;
  RFf := StrToDate(diass+'/'+mess+'/'+annos);
  ImprimirExistenciasConsumo(RFi,Rff
  ,global_Contrato_Barco);
end;

procedure TfrmAdmonyTiempos.ImprimirExistenciasConsumo(Finicial,Ffinal:TDateTime;ContBarco:string);
const
  ColIni = 2;
  RowIni = 6;
  xlCenter = -4108;  xlContext= -5002;

var
  ZqConsulta,ZqRecursos:TZReadOnlyQuery;
  EmbA:string;
  ExcelAp,Alibro,AHoja:Variant;
  Correcto :Boolean;

  CFecha:TDateTime;
  CCol,CRow,AgrColGen,AgrRecCol,CRecurso,x,UColumna:Integer;

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

    Fcl := 15;
    ExcelaP.Range[ColumnaNombre(Icl)+IntToStr(iFl)+':'+ColumnaNombre(FcL)+IntToStr(iFl)].Select;
    PFormatosExcel_H2(ExcelaP, 16, True, 12, clBlack, 'Arial');
    ExcelaP.Selection.HorizontalAlignment := xlCenter;
    ExcelAp.Selection.Value := 'EXISTENCIAS Y CONSUMOS.';
    Inc (iFl);

    ExcelaP.Range[ColumnaNombre(Icl)+IntToStr(iFl)+':'+ColumnaNombre(FcL)+IntToStr(iFl)].Select;
    PFormatosExcel_H2(ExcelaP, 38, True, 8, clBlack, 'Arial');
    ExcelaP.Selection.HorizontalAlignment := xlCenter;
    ExcelaP.Selection.Value := global_contrato;
    ExcelaP.Selection.ReadingOrder := xlContext;
    ExcelaP.Selection.WrapText := True;
  end;

  procedure EncabezadoImagen(Izquierda,Derecha:boolean;FcL:Integer;Modo:integer = 1);
  VAR tMPNAME,TempPath :string;
  imgAux : TImage;
  Pic:TJpegImage;
  fs:TStream;
  begin

      Fcl := 15;

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

          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True,900, 1, 90, 85);
        if Modo = 2 then
          ExcelaP.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 490, 1, 75, 70);
      end;
    end;
  end;
begin
  //
    Correcto := False;
    ZqConsulta:=TZReadOnlyQuery.Create(nil);
    try
      ZqConsulta.Connection := connection.zConnection;
      //Localizar embarcacion activa
      ZqConsulta.Active := False;
      ZqConsulta.SQL.Text :=
      'select max(dfechainicio),sidembarcacion, ifnull(LENGTH(dfechainicio+sidembarcacion),0) as ancho '+
      'from embarcacion_vigencia where sContrato = :ContratoBarco and dFechaInicio < :fecha ';
      ZqConsulta.ParamByName('ContratoBarco').AsString:= ContBarco;
      ZqConsulta.ParamByName('fecha').AsDate:=Now;
      ZqConsulta.Open;

      if ZqConsulta.FieldByName('ancho').AsInteger = 0 then
        raise Exception.Create('No existe una embarcación activa.'+#10+
        ' puede ir a panel de control, administración de catálogos, embarcación vigencias y dar de alta uno.');

      if (CmbEmb.Items.Count = 0) or (CmbEmb.ItemIndex < 0) then
        raise Exception.Create('Es necesario que seleccione una embarcación.');

      EmbA := ZqConsulta.FieldByName('sidembarcacion').AsString;

      ZqConsulta.Active := False;
      ZqConsulta.SQL.Text :=
      'select iidrecursoexistencia,sdescripcion from recursosdeexistencias order by iorden ';
      ZqConsulta.Open;

      //llenar encabezado

      ZqConsulta.Active := False;
      ZqConsulta.SQL.Text :=
      'select r.didfecha,re.iidrecursoexistencia,r.sIdEmbarcacion,r.dExistenciaAnterior,'+
      'r.dconsumo, r.drecibido,r.dproduccion, r.dconsumoequipos,r.dtrasiego,r.dprestamos,r.dExistenciaActual '+
      'from recursos r '+
      'inner join recursosdeexistencias re '+
      'on (r.iIdRecursoExistencia = re.iIdRecursoExistencia) '+
      'where (r.dIdFecha between :FechaInicial and :FechaFinal) and r.sContrato = :ContratoBarco and r.sIdEmbarcacion = :sidembarcacion '+
      'group by r.dIdFecha,r.sidembarcacion,r.iIdRecursoExistencia '+
      'order by r.sIdEmbarcacion';
      ZqConsulta.ParamByName('ContratoBarco').AsString:= ContBarco;
      ZqConsulta.ParamByName('sidembarcacion').AsString:=Tembarcacion(CmbEmb.Items.Objects[CmbEmb.ItemIndex]).Identificador;
      ZqConsulta.ParamByName('FechaInicial').AsDate:=Finicial;
      ZqConsulta.ParamByName('FechaFinal').AsDate:=Ffinal;
      ZqConsulta.Open;

      if ZqConsulta.RecordCount = 0 then
        raise Exception.Create('No hay registros cargados en ese rango de fechas.');

      if not GuardaExcel.Execute then
        raise Exception.Create('Proceso de generacion de archivo excel cancelado por el usuario.');


      ZqRecursos := TZReadOnlyQuery.Create(nil);
      try
        ZqRecursos.Connection := connection.zConnection;
        ZqRecursos.Active := False;
        ZqRecursos.SQL.Text:='select iidrecursoexistencia,sdescripcion from recursosdeexistencias order by iorden';
        ZqRecursos.Open;

        Try
          ExcelAp := CreateOleObject('Excel.Application');
        Except
          On E: Exception do
          begin
            raise Exception.Create('No se puede iniciar la aplicación excel, verifique que tenga instalado la paquetería office.');
          end;
        End;
        ExcelAp.Visible := True;
        ExcelAp.DisplayAlerts:= False;
        ALibro := ExcelAp.Workbooks.Add;

        //IndiceLibro := ExcelAp.Workbooks.count;
        AHoja := ALibro.Sheets.Add;
        AHoja.Name := CmbMeses.Text;
        ExcelAp.activeWindow.DisplayGridlines := false;
        //ALibro.Sheets[IndiceHoja].Name := MesesDA[MonthOf(cfechas)] +FormatDateTime('-yyyy',CFechas);

        EncabezadoTexto(2,3,1);
        EncabezadoImagen(True,True,1,1);


        CFecha := Finicial;
        Crow := RowIni;

        //Encabezado
        CCol := ColIni;
        ExcelAp.Range[ColumnaNombre(CCol-1)+IntToStr(Crow)+':'+ColumnaNombre(CCol-1)+IntToStr(Crow+1)].Select;
        ExcelAp.Selection.Value := 'FECHA';
        ExcelAp.Selection.MergeCells := True;
        ExcelAp.selection.interior.colorindex := 15;

        ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol+Zqrecursos.RecordCount-1)+IntToStr(Crow)].Select;
        ExcelAp.Selection.Value := 'APERTURA';
        ExcelAp.Selection.MergeCells := True;
        ExcelAp.selection.interior.colorindex := 15;

        CCol := CCol + Zqrecursos.RecordCount;
        ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol+Zqrecursos.RecordCount-1)+IntToStr(Crow)].Select;
        ExcelAp.Selection.Value := 'CONSUMO';
        ExcelAp.Selection.MergeCells := True;
        ExcelAp.selection.interior.colorindex := 15;

        CCol := CCol + Zqrecursos.RecordCount;
        ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol+Zqrecursos.RecordCount-2)+IntToStr(Crow)].Select;
        ExcelAp.Selection.Value := 'RECIBIDO';
        ExcelAp.Selection.MergeCells := True;
        ExcelAp.selection.interior.colorindex := 40;

        CCol := CCol + Zqrecursos.RecordCount-1;

        ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol+1)+IntToStr(Crow)].Select;
        ExcelAp.Selection.Value := 'RECIBIDO/PRODUCIDO';
        ExcelAp.Selection.MergeCells := True;
        ExcelAp.selection.interior.colorindex := 36;

        CCol := CCol + 2;

        ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol+Zqrecursos.RecordCount-1)+IntToStr(Crow)].Select;
        ExcelAp.Selection.Value := 'TRASEGADO';
        ExcelAp.Selection.MergeCells := True;
        ExcelAp.selection.interior.colorindex := 15;

        CCol := CCol + Zqrecursos.RecordCount;

        ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol+Zqrecursos.RecordCount-1)+IntToStr(Crow)].Select;
        ExcelAp.Selection.Value := 'CIERRE';
        ExcelAp.selection.interior.colorindex := 15;
        ExcelAp.Selection.MergeCells := True;
        UColumna := CCol+Zqrecursos.RecordCount-1;
        CRow := CRow+1;
        CCol := ColIni;
        for x := 0 to 1 do
        begin
          ZqRecursos.First;
          while not ZqRecursos.eof do
          begin
            ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
            ExcelAp.Selection.Value := ZqRecursos.FieldByName('sdescripcion').AsString;
            ExcelAp.selection.interior.colorindex := 15;
            CCol := Ccol+1;
            ZqRecursos.Next;
          end;
        end;

        ZqRecursos.First;
        while not ZqRecursos.eof do
        begin
          if LowerCase(ZqRecursos.FieldByName('sdescripcion').AsString) <> 'agua' then
          begin
            ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
            ExcelAp.Selection.Value := ZqRecursos.FieldByName('sdescripcion').AsString;
            ExcelAp.selection.interior.colorindex := 40;
            CCol := Ccol+1;
          end;
          ZqRecursos.Next;
        end;
        ZqRecursos.First;
        while not ZqRecursos.eof do
        begin
          if LowerCase(ZqRecursos.FieldByName('sdescripcion').AsString) = 'agua' then
          begin
            ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol+1)+IntToStr(Crow)].Select;
            ExcelAp.Selection.Value := ZqRecursos.FieldByName('sdescripcion').AsString;
            ExcelAp.selection.interior.colorindex := 36;
            CCol := Ccol+1;
          end;
          ZqRecursos.Next;
        end;
        CCol := Ccol+1;
        for x := 0 to 1 do
        begin
          ZqRecursos.First;
          while not ZqRecursos.eof do
          begin
            ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
            ExcelAp.Selection.Value := ZqRecursos.FieldByName('sdescripcion').AsString;
            ExcelAp.selection.interior.colorindex := 15;
            CCol := Ccol+1;
            ZqRecursos.Next;
          end;
        end;

        CRow := CRow+1;
        while CFecha <= Ffinal do
        begin
          CCol := ColIni;
          ZqRecursos.First;
          ZqConsulta.Filtered := False;
          ZqConsulta.Filter := ' didfecha ='+quotedstr(FormatDateTime('YYYY/MM/DD',CFecha));
          ZqConsulta.Filtered := True;
          ExcelAp.Range[ColumnaNombre(CCol-1)+IntToStr(Crow)+':'+ColumnaNombre(CCol-1)+IntToStr(Crow)].Select;
          ExcelAp.Selection.NumberFormat := '@';
          ExcelAp.Selection.Value := FormatDateTime('dd/MM/YYYY',CFecha);//ZqConsulta.FieldByName('didfecha').asdatetime;
          //Apertura
          ZqRecursos.First;
          while not ZqRecursos.Eof do
          begin
            if ZqConsulta.Locate('iidrecursoexistencia',ZqRecursos.FieldByName('iidrecursoexistencia').AsInteger,[]) then
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := ZqConsulta.FieldByName('dexistenciaanterior').AsFloat;
            end
            else
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := 0;
            end;          
            CCol := Ccol+1;
            ZqRecursos.Next;
          end;
          //Consumo
          ZqRecursos.First;
          while not ZqRecursos.Eof do
          begin
            if ZqConsulta.Locate('iidrecursoexistencia',ZqRecursos.FieldByName('iidrecursoexistencia').AsInteger,[]) then
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := ZqConsulta.FieldByName('dconsumo').AsFloat;
            end
            else
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := 0;
            end;
            CCol := Ccol+1;
            ZqRecursos.Next;
          end;
          //recibido
          ZqRecursos.First;
          while not ZqRecursos.Eof do
          begin
            if ZqConsulta.Locate('iidrecursoexistencia',ZqRecursos.FieldByName('iidrecursoexistencia').AsInteger,[]) then
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := ZqConsulta.FieldByName('drecibido').AsFloat;
            end
            else
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := 0;
            end;
            CCol := Ccol+1;

            ZqRecursos.Next;
          end;
          ZqRecursos.First;
          while not ZqRecursos.Eof do
          begin
            if lowercase(ZqRecursos.FieldByName('sdescripcion').AsString) = 'agua' then
            begin
              if ZqConsulta.Locate('iidrecursoexistencia',ZqRecursos.FieldByName('iidrecursoexistencia').AsInteger,[]) then
              begin
                ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
                ExcelAp.Selection.Value := ZqConsulta.FieldByName('dproduccion').AsFloat;
              end
              else
              begin
                ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
                ExcelAp.Selection.Value := 0;
              end;
              CCol := Ccol+1;
            end;
            ZqRecursos.Next;
          end;
          //trasiego
          ZqRecursos.First;
          while not ZqRecursos.Eof do
          begin
            if ZqConsulta.Locate('iidrecursoexistencia',ZqRecursos.FieldByName('iidrecursoexistencia').AsInteger,[]) then
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := ZqConsulta.FieldByName('dtrasiego').AsFloat;
            end
            else
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := 0;
            end;
            CCol := Ccol+1;
            ZqRecursos.Next;
          end;
          ZqRecursos.First;
          while not ZqRecursos.Eof do
          begin
            if ZqConsulta.Locate('iidrecursoexistencia',ZqRecursos.FieldByName('iidrecursoexistencia').AsInteger,[]) then
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := ZqConsulta.FieldByName('dexistenciaactual').AsFloat;
            end
            else
            begin
              ExcelAp.Range[ColumnaNombre(CCol)+IntToStr(Crow)+':'+ColumnaNombre(CCol)+IntToStr(Crow)].Select;
              ExcelAp.Selection.Value := 0;
            end;
            CCol := Ccol+1;
            ZqRecursos.Next;
          end;

          CFecha := IncDay(CFecha,1);
          CRow := CRow+1;
        end;

        ExcelAp.Range[ColumnaNombre(ColIni-1)+IntToStr(CRow)+':'+ColumnaNombre(Colini+ZqRecursos.recordcount-1)+IntToStr(crow)].Select;
        ExcelAp.selection.interior.colorindex := 15;
        ExcelAp.Selection.MergeCells := True;
        ExcelaP.Selection.HorizontalAlignment := xlCenter;
        ExcelAp.Selection.Value := 'Totales del mes.';

        for x := (ColIni +ZqRecursos.recordcount)  to ((ZqRecursos.recordcount * ZqRecursos.recordcount)+ColIni +ZqRecursos.recordcount) do
        begin
          ExcelAp.Range[ColumnaNombre(x)+IntToStr(crow)+':'+ColumnaNombre(x)+IntToStr(crow)].Select;
          ExcelAp.Selection.Formula := '= SUM(R[-'+inttostr(Crow-1)+']C:R[-1]C)';
          ExcelaP.Selection.HorizontalAlignment := xlCenter;
          ExcelAp.selection.interior.colorindex := 36;
          ExcelAp.Selection.NumberFormat := '0.000';
          ExcelAp.Selection.Borders.LineStyle := xlContinuous;
        end;

        ExcelAp.Range[ColumnaNombre((ZqRecursos.recordcount * ZqRecursos.recordcount)+ColIni +ZqRecursos.recordcount+1)+IntToStr(cRow)+':'+ColumnaNombre(UColumna)+IntToStr(crow)].Select;
        ExcelAp.Selection.Borders.LineStyle := xlContinuous;
        ExcelAp.Selection.MergeCells := True;
        ExcelAp.selection.interior.colorindex := 15;

        ExcelAp.Range[ColumnaNombre(ColIni-1)+IntToStr(RowIni)+':'+ColumnaNombre(UColumna)+IntToStr(crow-1)].Select;
        ExcelaP.Selection.HorizontalAlignment := xlCenter;
        ExcelAp.Selection.Borders.LineStyle := xlContinuous;


        ExcelAp.Range[ColumnaNombre(ColIni)+IntToStr(RowIni)+':'+ColumnaNombre(UColumna)+IntToStr(crow-1)].Select;
        ExcelAp.Selection.NumberFormat := '0.000';

        ExcelAp.Range[ColumnaNombre(ColIni-1)+IntToStr(RowIni)+':'+ColumnaNombre(UColumna)+IntToStr(Rowini+1)].Select;
        ExcelAp.Selection.Font.Bold  := True;


        //Formato de pagina
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

        Correcto := True;
      finally
        ZqRecursos.Free;
        if (Correcto)  then
        begin
          ALibro.SaveAs(guardaexcel.FileName);
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

    finally
      ZqConsulta.Free;
    end;
end;

procedure TfrmAdmonyTiempos.arriboembarcacionesEnter(Sender: TObject);
begin
  Embarcaciones.Active := False;
  Embarcaciones.sql.Clear;
  Embarcaciones.sql.Add('select distinct sIdEmbarcacion, sDescripcion, sTipo from embarcaciones ' +
                        'order by sDescripcion');
  Embarcaciones.Open;
end;

procedure TfrmAdmonyTiempos.ImportarMovtosBarco(Contrat,Emb:String);
var PathExl:string;
  ImpForzada:Boolean;

  {$REGION 'CERRAR EL EXCEL'}
  Function ExcelCloseWorkBooks(Excel : Variant; SaveAll: Boolean): Boolean;
  var
    loop: byte;
  Begin
    Result := True;
    Try
      For loop := 1 to Excel.Workbooks.Count Do
        Excel.Workbooks[1].Close[SaveAll];
    Except
      Result := False;
    End;
  End;

  Function ExcelClose(Excel : Variant; SaveAll: Boolean): Boolean;
  Begin
    Result := True;
    Try
      ExcelCloseWorkBooks(Excel, SaveAll);
      Excel.Quit;
    Except
      MessageDlg('Unable to Close Excel', mtError, [mbOK], 0);
      Result := False;
    End;
  End;
  {$ENDREGION}

  {$Region'Direccion de archivo'}
  function DireccionPlantilla :Boolean;
  var ResDP:Boolean;
    DlgPat:TOpenDialog;
    Cancelar :Boolean;
  begin
    ResDP := False;
    Cancelar := False;
    try
      DlgPat := TOpenDialog.Create(nil);
      try
        DlgPat.Filter :=  'Archivo Excel  (*.xls,*xlsx)|*.XLS;*.XLSX';
        DlgPat.FilterIndex := 0;
        repeat
          if DlgPat.Execute then
          begin
            if (AnsiEndsText('.xls',lowercase(DlgPat.FileName))) or (AnsiEndsText('.xlsx',lowercase(DlgPat.FileName))) then
            begin
              PathExl := DlgPat.FileName;
              ResDP := True
            end
            else
              ShowMessage('El archivo seleccionado no corresponde al formato excel requerido por el sistema.'+#10+'Intente de nuevo o cancele el proceso porfavor.')
          end
          else
            Cancelar := True;
        until ResDP or Cancelar;
      finally
        DlgPat.Free;
      end;
    finally
      Result := ResDP;
    end;
  end;
  {$ENDREGION}

  function ImportarPlantilla(Direccion,Cont,Embarq:string;Guardar:Boolean):Boolean;


  var
    MostrarExcel:Boolean;
    Excel,Libro,Hoja:variant;
    ILibro,IPagina:integer;

    ValsFecha,                //valores de celdas
    ValsHClima,
    ValsOlas,
    ValsPies,
    ValsVientos,
    ValsNudos,
    ValsLocalizacion,
    ValsHoraInicio,
    ValsHoraFin,
    ValsDescripcion,
    ValsPartida,
    ValsFolio,
    ValsOt,
    ValsFactor,
    ValsEmbActiva, //almacenara el id de la embarcacion
    SValorA:String;

    FilaInicio :Integer;
    FilaFecha:Integer;
    CurFila: Integer;           //Recorrer las filas
    FilasVacias: Integer;
    FilaOlas:Integer;
    FilaVientos:Integer;
    FilaIniPartidas:Integer;    //Apartir de q filas estan las partidas
    FilaIniEmbarcaciones:Integer; //Apartir de q filas estan las embarcaciones
    FilaIniHelicopteros:Integer;

    CurCol: Integer;            //Recorrer las columnas
    AuxCont:Integer;

    FSig:Boolean;// Me dira si es necesario pasar a linea siguiente
    i,
    UIdDiario: Integer;//ultimo iddiario

    ZAux,ZDireccion:TZReadOnlyQuery;
    ZUpdt: TZQuery;

    LstFol,LsTemp:tstringlist;

    ImportaOlas,ImportaVientos:Boolean;

    {$REGION 'Buscar hoja y eliminaacentos'}
    Function BuscaHoja(Nomb:String):Integer;
    var ResBuscaHoja,x:Integer;
    begin
      ResBuscaHoja := -1;
      try
        for x := 1 to Excel.WorkBooks[Ilibro].Sheets.Count  do
        begin
          if LowerCase(Excel.workbooks[iLibro].workSheets[x].Name) = LowerCase(Nomb) then
            ResBuscaHoja := x;
        end;
      finally
        Result := ResBuscaHoja;
      end;
    end;

    Function EliminaAcentos(Texto:string):string;
    const Acentos = 'áéíóúÁÉÍÓÚ'; NoAcentos = 'aeiouAEIOU';
    var i: integer;
    begin
      for i:= 1 to length(Texto) do
        begin
          if pos(Texto[i], Acentos) <> 0 then
          Texto:= StringReplace(Texto, Acentos[pos(Texto[i], Acentos)], NoAcentos[pos(Texto[i], Acentos)], [rfReplaceAll, rfIgnoreCase]);
        end;
      Result:= Texto;
    end;
    {$ENDREGION}

  begin
    try
      MostrarExcel := False;
      LstFol := Tstringlist.create;
      LsTemp := Tstringlist.create;
      try
        //iniciando excel
        try
          Excel:=CreateOleObject('Excel.Application');
          Excel.Visible := MostrarExcel;
        except
          raise Exception.Create('No tiene instalado microsoft excel o bién ocurre un problema con el mismo.');
        end;

        Excel.Workbooks.Open(Direccion);
        Libro := Excel.Workbooks[Excel.Workbooks.count];
        ILibro := Excel.Workbooks.count;

        IPagina:= BuscaHoja('CARATULA');
        if IPagina = -1 then
          raise Exception.Create('La hoja con nombre caratula o CARATULA no existe en el libro excel.');

        Hoja := Libro.worksheets[IPagina];

        //Inicia la validacion

        {$REGION 'Validacion de formato'}
        //Obtenemos y asignamos fecha
        FilaFecha := -1;
        CurFila := 1;
        FilasVacias := 0;
        AuxCont := 10;
        while (FilaFecha = -1) and (FilasVacias < AuxCont) do
        begin
          try
            ValsFecha := Hoja.Cells[CurFila,CFecha].value;
          except
            ValsFecha := '';
          end;

          ValsFecha := Trim(ValsFecha);
          ValsFecha := AnsiLeftStr(ValsFecha,10);

          try
            strtodate(ValsFecha);
            FilaFecha := Curfila
          except
            inc(filasvacias);
          end;
          Inc(CurFila);
        end;
        if FilaFecha = -1 then
        begin
          raise Exception.Create('No se ha podido localizar el renglón donde se cargará el valor de la fecha '+#10+'Verifique su formato porfavor.');
        end;

        SValorA := Hoja.cells[FilaFecha,CHClima-1].text;
        if SValorA <> 'HORA' then
          raise Exception.Create('Debido a un cambio en la plantilla, se requiere regenerarla'+#10+' ya que incluye una nueva celda de horario de condiciones climatologicas, sirvase a generarla desde el sistema porfavor.');

        ValsHClima := Hoja.cells[FilaFecha,CHClima].text;
        ValsHClima := StringReplace(ValsHClima, ' ', '', [rfReplaceAll]);
        ValsHClima := AnsiLeftStr(ValsHClima, 5 );
        if length(trim(ValsHClima)) > 0 then
        begin
          if ValsHClima <> '24:00' then
          begin
            try
              StrToTime(ValsHClima);
            finally
              ValsHClima := '00:00';
            end;
          end;
        end
        else
          ValsHClima := '00:00';

        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Text := 'select sidembarcacion from embarcacion_vigencia where :fecha >= dfechainicio and :fecha <= dfechafinal  order by dFechaFinal desc limit 1';
        connection.QryBusca.ParamByName('fecha').asdate := StrToDate(ValsFecha);
        connection.QryBusca.Open;

        if connection.QryBusca.RecordCount = 0 then
          raise Exception.Create('Para la fecha '+valsfecha+' No se encontro embarcacion activa'+#10+' Ingrese una embarcacion que cubra esa fecha.');

        ValsEmbActiva := connection.QryBusca.FieldByName('sidembarcacion').AsString;

        //localizar y almacenar olas, vientos y localizacion
        FilaIniPartidas := -1;
        CurFila := 1;
        FilasVacias := 0;
        AuxCont := 10;
        while (FilaIniPartidas = -1) and (FilasVacias < AuxCont) do
        begin
          if trim(lowercase(Hoja.cells[CurFila,Cvar1].text)) = 'olas' then
            FilaIniPartidas := Curfila
          else
            inc(filasvacias);
          Inc(CurFila);
        end;
        if FilaIniPartidas = -1 then
          raise Exception.Create('No se ha podido localizar el renglón donde se cargará el valor para las olas, vientos y localización '+#10+'Verifique su formato porfavor.');

        FilaOlas := FilaIniPartidas;
        FilaVientos := FilaOlas+1;

        ValsOlas := Hoja.Cells[FilainiPArtidas,CVar1+1].text;

        ValsVientos := Hoja.Cells[FilainiPArtidas+1,CVar1+1].text;

        ValsLocalizacion := Hoja.Cells[FilainiPArtidas+2,CVar1+1].text;
        // almacenar pies y nudos
        ValsPies := Hoja.Cells[FilainiPArtidas,CVar2+1].text;

        ValsNudos := Hoja.Cells[FilainiPArtidas+1,CVar2+1].text;

        //localizar el inicio de fila de movimientos de embarcacion
        FilaIniPartidas := -1;
        CurFila := 1;
        FilasVacias := 0;
        AuxCont := 50;
        while (FilaIniPartidas = -1) and (FilasVacias < AuxCont) do
        begin
          if trim(lowercase(Hoja.cells[CurFila,Cpartida].text)) = 'part' then
            FilaIniPartidas := Curfila+1
          else
            inc(filasvacias);
          Inc(CurFila);
        end;
        if FilaIniPartidas = -1 then
          raise Exception.Create('No se ha podido localizar el renglón donde inician las partidas '+#10+'Verifique su formato porfavor.');

        //localizar el inicio de fila de embarcaciones
        FilaIniEmbarcaciones := -1;
        CurFila := 1;
        FilasVacias := 0;
        AuxCont := 100;
        while (FilaIniEmbarcaciones = -1) and (FilasVacias < AuxCont) do
        begin
          sValorA := Hoja.cells[CurFila,cembarcaciones].text;
          sValorA := lowercase(sValorA);
          sValorA := StringReplace(sValorA, ' ', '', [rfReplaceAll]);
          if sValorA = '<***embarcaciones***>' then
            FilaIniEmbarcaciones := Curfila+2
          else
            inc(filasvacias);
          Inc(CurFila);
        end;
        if FilaIniEmbarcaciones = -1 then
          raise Exception.Create('No se ha podido localizar el renglón donde inician las embarcaciones '+#10+'Verifique su formato porfavor.');

        //localizar el inicio de fila de Helicopteros
        FilaIniHelicopteros := -1;
        CurFila := 1;
        FilasVacias := 0;
        AuxCont := 150;
        while (FilaIniHelicopteros = -1) and (FilasVacias < AuxCont) do
        begin
          sValorA := Hoja.cells[CurFila,CHelicopteros].text;
          SValorA := EliminaAcentos(sValorA);
          sValorA := lowercase(sValorA);
          sValorA := StringReplace(sValorA, ' ', '', [rfReplaceAll]);
          if sValorA = '<***helicopteros***>' then
            FilaIniHelicopteros := Curfila+2
          else
            inc(filasvacias);
          Inc(CurFila);
        end;
        if FilaIniHelicopteros = -1 then
          raise Exception.Create('No se ha podido localizar el renglón donde inician los helicópteros '+#10+'Verifique su formato porfavor.');

        {$ENDREGION}

        //*********************************Validacion de datos

        {$REGION 'Olas y vientos'}
        if (length(trim(ValsOlas)) > 0) or (length(trim(ValsVientos)) > 0) then
        begin
          ZDireccion := TZReadOnlyQuery.Create(nil);
          try
            ZDireccion.Active := False;
            ZDireccion.sql.Clear;
            ZDireccion.SQL.Text := 'select iIddireccion,lower(sdescripcion) as sdescripcion from direcciones';
            ZDireccion.Connection := connection.zConnection;
            ZDireccion.Open;

            //Importar olas?
            if (length(trim(ValsOlas)) > 0) then
            begin
              ValsOlas := StringReplace(ValsOlas, ' ', '', [rfReplaceAll]);
              ValsOlas := LowerCase(ValsOlas);
              if not ZDireccion.Locate('sdescripcion',ValsOlas,[]) then
              begin
                ImportaOlas := False;
                Hoja.cells[FilaOlas,CVar2+1].interior.colorindex := ClNoBd;
                MostrarExcel := True;
              end
              else
              begin
                ImportaOlas := True;
                ValsOlas := ZDireccion.FieldByName('iIddireccion').AsString;
                if length(trim(ValsPies)) = 0 then
                  ValsNudos := '0';
                Hoja.cells[FilaOlas,CVar1].interior.colorindex := ClOk;
              end;
            end;

            //Importar vientos?
            if (length(trim(ValsVientos)) > 0) then
            begin
              ValsVientos := StringReplace(ValsVientos, ' ', '', [rfReplaceAll]);
              ValsVientos := LowerCase(ValsVientos);
              if not ZDireccion.Locate('sdescripcion',ValsVientos,[]) then
              begin
                ImportaVientos := False;
                Hoja.cells[FilaVientos,CVar2+1].interior.colorindex := ClNoBd;
                MostrarExcel := True;
              end
              else
              begin
                ImportaVientos := True;
                ValsVientos := ZDireccion.FieldByName('iIddireccion').AsString;
                if length(trim(ValsPies)) = 0 then
                  ValsNudos := '0';
                Hoja.cells[FilaVientos,CVar1].interior.colorindex := ClOk;
              end;
            end;
          finally
            ZDireccion.Free;
          end;
        end;
        {$ENDREGION}

        {$REGION 'Validacion de apartado movimientos'}
        //-----------------------Movimientos
        FilasVacias := 0;
        for CurFila := FilaIniPartidas to FilaIniEmbarcaciones -3 do
        begin
          try
            Hoja.range[ColumnaNombre(1)+inttostr(Curfila)+':'+ColumnaNombre(10)+inttostr(Curfila)].select;
            SValorA := '';
            for i := 1 to cot do
            begin
              SValorA := SValorA + vartostr(Hoja.cells[curfila,i].value);
            end;

            SValorA := StringReplace(sValorA, ' ', '', [rfReplaceAll]);

            if Length(SValorA) = 0 then
            begin

              for i := 1 to cot do
              begin
                Hoja.cells[CurFila,i].interior.colorindex := ClFVacia;
              end;  
              raise exception.Create('pasar a siguiente linea');
            end;

            //nomas para saber q ya se cargaron
            {ValsFecha;
            ValsOlas;
            ValsVientos;
            ValsLocalizacion;  }
            //Cargamos valores necesarios
            ValsHoraInicio := Hoja.Cells[CurFila,CHInicio].text;
            ValsHoraFin := Hoja.Cells[CurFila,CHFin].text;
            ValsDescripcion := Hoja.Cells[CurFila,CDescripcion].text;
            ValsPartida := Hoja.Cells[CurFila,CPartida].text;
            ValsFolio := Hoja.Cells[CurFila,CFolio].text;
            ValsOt:= Hoja.Cells[CurFila,COt].text;

            ValsHoraInicio := StringReplace(ValsHoraInicio, ' ', '', [rfReplaceAll]);
            ValsHoraFin := StringReplace(ValsHoraFin, ' ', '', [rfReplaceAll]);
            ValsPartida := StringReplace(ValsPartida, ' ', '', [rfReplaceAll]);
            ValsFolio := StringReplace(ValsFolio, ' ', '', [rfReplaceAll]);
            ValsOt  := StringReplace(ValsOt, ' ', '', [rfReplaceAll]);

            //Inician las validaciones de existencia
            if length(trim(ValsHoraInicio)) = 0 then
            begin
              Hoja.cells[CurFila,CHInicio].interior.colorindex := ClCeldaVacia;
              MostrarExcel := True;
            end
            else
              Hoja.cells[CurFila,CHInicio].interior.colorindex := ClOk;

            if length(trim(ValsHoraFin)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,CHFin].interior.colorindex := ClCeldaVacia
            end
            else
              Hoja.cells[CurFila,CHFin].interior.colorindex := ClOk;
            try
              if (Length(Trim(ValsHoraInicio)) > 0) and (Length(Trim(ValsHoraFin)) > 0)then
                if ValsHoraFin <> '24:00' then
                begin
                  if StrToTime(ValsHoraInicio) > StrToTime(ValsHoraFin) then
                  begin
                    Hoja.cells[CurFila,CHInicio].interior.colorindex := ClHFinMenor;
                    Hoja.cells[CurFila,CHFin].interior.colorindex := ClHFinMenor;
                  end;
                end
                else
                begin
                  if StrToTime(ValsHoraInicio) > StrToTime('23:59') then
                  begin
                    MostrarExcel := True;
                    Hoja.cells[CurFila,CHInicio].interior.colorindex := ClHFinMenor;
                    Hoja.cells[CurFila,CHFin].interior.colorindex := ClHFinMenor;
                  end;
                end;

            except
              MostrarExcel := True;
              Hoja.cells[CurFila,CHInicio].interior.colorindex :=  ClCeldaVacia;
              Hoja.cells[CurFila,CHFin].interior.colorindex :=  ClCeldaVacia;
            end;

            if length(trim(ValsDescripcion)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,CDescripcion].interior.colorindex :=  ClCeldaVacia
            end
            else
              Hoja.cells[CurFila,CDescripcion].interior.colorindex := ClOk;

            if length(trim(ValsPartida)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,Cpartida].interior.colorindex :=  ClCeldaVacia
            end
            else
            begin
              if ZMovtos.Locate('sidtipomovimiento',ValsPartida,[]) then
                Hoja.cells[CurFila,Cpartida].interior.colorindex := ClOk
              else
              begin
                MostrarExcel := True;
                Hoja.cells[CurFila,Cpartida].interior.colorindex := ClNoBd;
              end;
            end;

            if length(trim(ValsFolio)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,CFolio].interior.colorindex := ClCeldaVacia
            end
            else
            begin
              {Hoja.Cells[CurFila,CFolio].interior.colorindex := ClOk ;
              LstFol.CommaText := LstFol.CommaText + ValsFolio;
              LstFol.Add(ValsFolio);  }
              LsTemp.CommaText := ValsFolio;
              for I := 0 to LsTemp.Count - 1 do
              begin
                if length(trim(LsTemp[i])) > 0 then
                begin
                   Hoja.Cells[CurFila,CFolio].interior.colorindex := ClOk;
                  if ZFolios.Locate('scontrato;snumeroorden',VarArrayOf([ValsOt,LsTemp[i]]),[]) then
                  begin
                    if LstFol.IndexOf(LsTemp[i]) = -1 then
                      LstFol.Add(LsTemp[i]);

                  end
                  else
                  begin
                    MostrarExcel := True;
                    Hoja.Cells[CurFila,CFolio].interior.colorindex := ClFNoOt;
                  end;

                end;


              end;
              {
              if ZFolios.Locate('scontrato;snumeroorden',VarArrayOf([ValsOt,ValsFolio]),[]) then
              begin
                if LstFol.IndexOf(ValsFolio) = -1 then
                  LstFol.Add(ValsFolio);
                Hoja.Cells[CurFila,CFolio].interior.colorindex := ClOk
              end
              else
              begin
                MostrarExcel := True;
                Hoja.Cells[CurFila,CFolio].interior.colorindex := ClFNoOt;
              end;  }
            end;

            if length(trim(ValsOt)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,COt].Interior.ColorIndex := ClCeldaVacia
            end
            else
            begin
              if ZFolios.Locate('scontrato',Valsot,[]) then
                Hoja.Cells[CurFila,COt].interior.colorindex := ClOk
              else
              begin
                MostrarExcel := True;
                Hoja.Cells[CurFila,COt].interior.colorindex := ClNoBd
              end;
            end;

          except
            ;
          end;
        end;
        {$ENDREGION}

        {$REGION 'Validacion de apartado Embarcaciones'}
        //-----------------------Embarcaciones
        FilasVacias := 0;
        for CurFila := FilaIniEmbarcaciones to FilaIniHelicopteros -3 do
        begin
          try
            Hoja.range[ColumnaNombre(1)+inttostr(Curfila)+':'+ColumnaNombre(5)+inttostr(Curfila)].select;
            SValorA := '';
            for i := 1 to CDescripcion do
            begin
              SValorA := SValorA + vartostr(Hoja.cells[curfila,i].value);
            end;

            SValorA := StringReplace(sValorA, ' ', '', [rfReplaceAll]);

            if Length(SValorA) = 0 then
            begin
              for i := 1 to CDescripcion do
              begin
                Hoja.cells[CurFila,i].interior.colorindex := ClFVacia;
              end;
              raise exception.Create('pasar a siguiente linea');
            end;

            //Cargamos valores necesarios
            ValsHoraInicio := Hoja.Cells[CurFila,CHInicio].text;
            ValsHoraFin := Hoja.Cells[CurFila,CHFin].text;
            ValsDescripcion := Hoja.Cells[CurFila,CDescripcion].value;
            ValsHoraInicio := StringReplace(ValsHoraInicio, ' ', '', [rfReplaceAll]);
            ValsHoraFin := StringReplace(ValsHoraFin, ' ', '', [rfReplaceAll]);
            if length(trim(ValsHoraInicio)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,CHInicio].interior.colorindex := ClCeldaVacia
            end
            else
              Hoja.cells[CurFila,CHInicio].interior.colorindex := ClOk;

            if length(trim(ValsHoraFin)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,CHFin].interior.colorindex := ClCeldaVacia
            end
            else
              Hoja.cells[CurFila,CHFin].interior.colorindex := ClOk;
            try
              if (Length(Trim(ValsHoraInicio)) > 0) and (Length(Trim(ValsHoraFin)) > 0)then
                if ValsHoraFin <> '24:00' then
                begin
                  if StrToTime(ValsHoraInicio) > StrToTime(ValsHoraFin) then
                  begin
                    MostrarExcel := True;
                    Hoja.cells[CurFila,CHInicio].interior.colorindex := ClHFinMenor;
                    Hoja.cells[CurFila,CHFin].interior.colorindex := ClHFinMenor;
                  end;
                end
                else
                begin
                  if StrToTime(ValsHoraInicio) > StrToTime('23:59') then
                  begin
                    MostrarExcel := True;
                    Hoja.cells[CurFila,CHInicio].interior.colorindex := ClHFinMenor;
                    Hoja.cells[CurFila,CHFin].interior.colorindex := ClHFinMenor;
                  end;
                end;

            except
              MostrarExcel := True;
              Hoja.cells[CurFila,CHInicio].interior.colorindex :=  ClCeldaVacia;
              Hoja.cells[CurFila,CHFin].interior.colorindex :=  ClCeldaVacia;
            end;

            if length(trim(ValsDescripcion)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,CDescripcion].interior.colorindex :=  ClCeldaVacia
            end
            else
              Hoja.cells[CurFila,CDescripcion].interior.colorindex := ClOk;

          except
            ;
          end;

        end;
        {$ENDREGION}

        {$REGION 'Validacion de apartado Helicopteros'}
        //----------------------Helicopteros
        CurFila := FilaIniHelicopteros;
        FilasVacias := 0;
        while FilasVacias < MaxVacias do
        begin
          try
            Hoja.range[ColumnaNombre(1)+inttostr(Curfila)+':'+ColumnaNombre(5)+inttostr(Curfila)].select;
            SValorA := '';
            for i := 1 to CDescripcion do
            begin
              SValorA := SValorA + vartostr(Hoja.cells[curfila,i].value);
            end;

            SValorA := StringReplace(sValorA, ' ', '', [rfReplaceAll]);

            if Length(SValorA) = 0 then
            begin
              inc(FilasVacias);
              try
                for i := 1 to CDescripcion do
                begin
                  Hoja.cells[CurFila,i].interior.colorindex := ClFVacia;
                end;
              except
                ;
              end;
              raise exception.Create('pasar a siguiente linea');
            end;

            //Cargamos valores necesarios
            ValsHoraInicio := Hoja.Cells[CurFila,CHInicio].text;
            ValsHoraFin := Hoja.Cells[CurFila,CHFin].text;
            ValsDescripcion := Hoja.Cells[CurFila,CDescripcion].text;
            ValsHoraInicio := StringReplace(ValsHoraInicio, ' ', '', [rfReplaceAll]);
            ValsHoraFin := StringReplace(ValsHoraFin, ' ', '', [rfReplaceAll]);
            if length(trim(ValsHoraInicio)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,CHInicio].interior.colorindex := ClCeldaVacia
            end
            else
              Hoja.cells[CurFila,CHInicio].interior.colorindex := ClOk;

            if length(trim(ValsHoraFin)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,CHFin].interior.colorindex := ClCeldaVacia
            end
            else
              Hoja.cells[CurFila,CHFin].interior.colorindex := ClOk;
            try
              if (Length(Trim(ValsHoraInicio)) > 0) and (Length(Trim(ValsHoraFin)) > 0)then
                if ValsHoraFin <> '24:00' then
                begin
                  if StrToTime(ValsHoraInicio) > StrToTime(ValsHoraFin) then
                  begin
                    MostrarExcel := True;
                    Hoja.cells[CurFila,CHInicio].interior.colorindex := ClHFinMenor;
                    Hoja.cells[CurFila,CHFin].interior.colorindex := ClHFinMenor;
                  end;
                end
                else
                begin
                  if StrToTime(ValsHoraInicio) > StrToTime('23:59') then
                  begin
                    MostrarExcel := True;
                    Hoja.cells[CurFila,CHInicio].interior.colorindex := ClHFinMenor;
                    Hoja.cells[CurFila,CHFin].interior.colorindex := ClHFinMenor;
                  end;
                end;

            except
              MostrarExcel := True;
              Hoja.cells[CurFila,CHInicio].interior.colorindex :=  ClCeldaVacia;
              Hoja.cells[CurFila,CHFin].interior.colorindex :=  ClCeldaVacia;
            end;

            if length(trim(ValsDescripcion)) = 0 then
            begin
              MostrarExcel := True;
              Hoja.cells[CurFila,CDescripcion].interior.colorindex :=  ClCeldaVacia
            end
            else
              Hoja.cells[CurFila,CDescripcion].interior.colorindex := ClOk;
          except
            ;
          end;
          Inc(CurFila);
        end;
        {$ENDREGION}

        if MostrarExcel then
          raise Exception.Create('Se detectaron errores en la plantilla, favor de corregir de acuerdo al codigo de colores.');
        //**********************************Importacion de datos

        if not Assigned(ZUpdt) then
        begin
          ZUpdt := TZQuery.Create(self);
          ZUpdt.Connection := connection.zConnection;
        end;

        ZUpdt.Connection.StartTransaction;

        {$REGION 'Importacion de olas y vientos'}
        if BorrarEx.Checked then
        begin
          if ImportaOlas or ImportaVientos then
          begin
            ZUpdt.Active := False;
            ZUpdt.SQL.Clear;
            ZUpdt.SQL.Text := 'DELETE FROM condicionesclimatologicas WHERE scontrato = :Contrato AND didfecha =  :fecha ';
            ZUpdt.ParamByName('Contrato').AsString := Cont;
            ZUpdt.ParamByName('Fecha').asdate := strtodate(ValsFecha);
            ZUpdt.ExecSQL;
          end;
        end;

        if ImportaOlas then
        begin        //verificar horario
          ZUpdt.Active := False;
          ZUpdt.SQL.Clear;
          ZUpdt.SQL.Text := 'insert into condicionesclimatologicas (sContrato,didfecha,sidturno,iidcondicion,slocalizacion,mpronostico,shorario,iIdDireccion,scantidad) '+
                            'values (:contrato,:fecha,"A",1,:localizacion,:pronostico,:horario,:direccion,:cantidad)';
//                            'values (:contrato,:fecha,"A",(select ifnull(max(cc.iidcondicion)+1,1) as maximo from condicionesclimatologicas cc where sContrato = :contrato and dIdFecha = :fecha),:localizacion,:pronostico,:horario,:direccion,:cantidad)';
          ZUpdt.ParamByName('Contrato').AsString := Cont;
          ZUpdt.ParamByName('Fecha').asdate := strtodate(ValsFecha);
          ZUpdt.ParamByName('localizacion').AsString := ValsLocalizacion;
          ZUpdt.ParamByName('pronostico').AsString := '*';
          ZUpdt.ParamByName('horario').AsString := ValsHClima;
          ZUpdt.ParamByName('direccion').AsString := ValsOlas;
          ZUpdt.ParamByName('cantidad').AsString := ValsPies;
          ZUpdt.ExecSQL;
        end;

        if ImportaVientos then
        begin
          ZUpdt.Active := False;
          ZUpdt.SQL.Clear;
          ZUpdt.SQL.Text := 'insert into condicionesclimatologicas (sContrato,didfecha,sidturno,iidcondicion,slocalizacion,mpronostico,shorario,iIdDireccion,scantidad) '+
                            'values (:contrato,:fecha,"A",2,:localizacion,:pronostico,:horario,:direccion,:cantidad)';
          ZUpdt.ParamByName('Contrato').AsString := Cont;
          ZUpdt.ParamByName('Fecha').asdate := strtodate(ValsFecha);
          ZUpdt.ParamByName('localizacion').AsString := ValsLocalizacion;
          ZUpdt.ParamByName('pronostico').AsString := '*';
          ZUpdt.ParamByName('horario').AsString := ValsHClima;
          ZUpdt.ParamByName('direccion').AsString := ValsVientos;
          ZUpdt.ParamByName('cantidad').AsString := ValsNudos;
          ZUpdt.ExecSQL;
        end;
        {$ENDREGION}

        {$REGION 'Importacion apartado movimientos'}
        //-----------------------Movimientos
        if BorrarEx.Checked then
        begin
          ZUpdt.Active := False;
          ZUpdt.SQL.Clear;
          ZUpdt.SQL.Text := 'DELETE FROM movimientosdeembarcacion WHERE scontrato = :Contrato AND didfecha =  :fecha';
          ZUpdt.ParamByName('Contrato').AsString := Cont;
          ZUpdt.ParamByName('Fecha').asdate := strtodate(ValsFecha);
          ZUpdt.ExecSQL;
        end;


        FilasVacias := 0;
        for CurFila := FilaIniPartidas to FilaIniEmbarcaciones -3 do
        begin

          try
            Hoja.range[ColumnaNombre(1)+inttostr(Curfila)+':'+ColumnaNombre(10)+inttostr(Curfila)].select;
            SValorA := '';
            for i := 1 to cot do
            begin
              SValorA := SValorA + vartostr(Hoja.cells[curfila,i].value);
            end;

            SValorA := StringReplace(sValorA, ' ', '', [rfReplaceAll]);

            if Length(SValorA) = 0 then
            begin

              for i := 1 to cot do
              begin
                Hoja.cells[CurFila,i].interior.colorindex := ClFVacia;
              end;
              raise exception.Create('pasar a siguiente linea');
            end;

            //nomas para saber q ya se cargaron
            {ValsFecha;
            ValsOlas;
            ValsVientos;
            ValsLocalizacion;  }
            //Cargamos valores necesarios
            ValsHoraInicio := Hoja.Cells[CurFila,CHInicio].text;
            ValsHoraInicio := AnsiLeftStr(ValsHoraInicio, 5 );

            ValsHoraFin := Hoja.Cells[CurFila,CHFin].text;
            ValsHoraFin := AnsiLeftStr(ValsHoraFin, 5 );

            ValsDescripcion := Hoja.Cells[CurFila,CDescripcion].text;
            ValsPartida := Hoja.Cells[CurFila,CPartida].text;
            ValsFolio := Hoja.Cells[CurFila,CFolio].text;
            ValsOt:= Hoja.Cells[CurFila,COt].text;

            ValsHoraInicio := StringReplace(ValsHoraInicio, ' ', '', [rfReplaceAll]);
            ValsHoraFin := StringReplace(ValsHoraFin, ' ', '', [rfReplaceAll]);
            ValsPartida := StringReplace(ValsPartida, ' ', '', [rfReplaceAll]);
            ValsFolio := StringReplace(ValsFolio, ' ', '', [rfReplaceAll]);
            ValsOt  := StringReplace(ValsOt, ' ', '', [rfReplaceAll]);

            //trabajando query
            connection.QryBusca.Active := False;
            connection.QryBusca.SQL.Text := 'Select max(iiddiario)+1 as maximo from movimientosdeembarcacion';
            connection.QryBusca.Open;
            if connection.QryBusca.RecordCount = 1 then
              UIdDiario :=  connection.QryBusca.FieldByName('maximo').asinteger
            else
              UIdDiario := 0;

            ZUpdt.Active := False;
            ZUpdt.SQL.Clear;
            ZUpdt.SQL.Text := 'INSERT INTO movimientosdeembarcacion (sContrato, dIdFecha, iIdDiario, sTipoEmbarcacion, sClasificacion, sidembarcacion, sidfase,shorainicio,shorafinal,sfactor,mdescripcion,iddiario,stipo,lcontinuo,sorden,snumeroactividad) VALUES '+
                              '(:Contrato, :Fecha, :idiario,"", :PartBarco,:embarcacion,"OPER",:horainicio,:horafinal,:factor,:descripcion,0,:tipo,"No",:Ot,:PartBarco)';
            ZUpdt.ParamByName('Contrato').AsString := Cont;
            ZUpdt.ParamByName('embarcacion').AsString := ValsEmbActiva;
            ZUpdt.ParamByName('idiario').asinteger := UIdDiario;
            ZUpdt.ParamByName('Fecha').asdate := strtodate(ValsFecha);
            ZUpdt.ParamByName('PartBarco').AsString := ValsPartida;
            ZUpdt.ParamByName('horainicio').AsString := ValsHoraInicio;
            ZUpdt.ParamByName('horafinal').AsString := ValsHoraFin;

            if (ValsHoraInicio = '00:00') and (ValsHoraFin = '24:00') then
              ValsFactor  := '1'
            else
              ValsFactor := sfnFactor(ValsHoraInicio, ValsHoraFin, 24, 6);
            ZUpdt.ParamByName('factor').AsString := ValsFactor;
            ZUpdt.ParamByName('descripcion').AsString := ValsDescripcion;
            ZUpdt.ParamByName('tipo').AsString := (Hoja.cells[curfila,Cclasificacion].text);
            ZUpdt.ParamByName('Ot').AsString := ValsOt;
            ZUpdt.ExecSQL;



            ZUpdt.Active:=False;
            ZUpdt.SQL.Text:='delete from movimientosxfolios where sContrato=:Contrato and dIdFecha=:fecha and iiddiario = :iddiario' ;
            ZUpdt.ParamByName('Contrato').AsString := cont;
            ZUpdt.ParamByName('iddiario').AsInteger := UIdDiario;
            ZUpdt.ParamByName('Fecha').AsDate      := strtodate(valsfecha);
            ZUpdt.ExecSQL;

            ZUpdt.Active:=False;
            ZUpdt.SQL.Text:='INSERT INTO movimientosxfolios (sContrato, dIdFecha, iIdDiario, sNumeroOrden, sFolio, lFactor, sfactor) VALUES (:Contrato, :Fecha, :Diario, :Ot, :Folio, :Aplica,:factor)';

            LsTemp.CommaText := ValsFolio;

            for I := 0 to LsTemp.Count - 1 do
            begin
              if Length(Trim(LsTemp[i])) > 0 then
              begin
                ZUpdt.Active:=False;
                ZUpdt.ParamByName('Contrato').AsString  := Cont;
                ZUpdt.ParamByName('Fecha').AsDate       := StrToDate(ValsFecha);
                ZUpdt.ParamByName('Diario').AsInteger   := UIdDiario;
                ZUpdt.ParamByName('ot').AsString        := ValsOt;
                ZUpdt.ParamByName('Folio').AsString     := LsTemp[i];
                ZUpdt.ParamByName('Aplica').AsString := 'Si';
                ZUpdt.ParamByName('factor').AsString := ValsFactor;
                ZUpdt.ExecSQL;
              end;
            end;
            Hoja.cells[CurFila,1].interior.colorindex := ClLineaImp;
          except
            on e : exception do
              mostrarexcel := True;
          end;

        end;
        {$ENDREGION}

        {$REGION 'Importacion apartado embarcacion'}
        FilasVacias := 0;
        for CurFila := FilaIniEmbarcaciones to FilaIniHelicopteros -3 do
        begin

          try
            Hoja.range[ColumnaNombre(1)+inttostr(Curfila)+':'+ColumnaNombre(10)+inttostr(Curfila)].select;
            SValorA := '';
            for i := 1 to CDescripcion do
            begin
              SValorA := SValorA + vartostr(Hoja.cells[curfila,i].value);
            end;

            SValorA := StringReplace(sValorA, ' ', '', [rfReplaceAll]);

            if Length(SValorA) = 0 then
            begin

              for i := 1 to CDescripcion do
              begin
                Hoja.cells[CurFila,i].interior.colorindex := ClFVacia;
              end;  
              raise exception.Create('pasar a siguiente linea');
            end;

            //nomas para saber q ya se cargaron
            {ValsFecha;
            ValsOlas;
            ValsVientos;
            ValsLocalizacion;  }
            //Cargamos valores necesarios
            ValsHoraInicio := Hoja.Cells[CurFila,CHInicio].text;
            ValsHoraInicio := AnsiLeftStr(ValsHoraInicio, 5 );
            ValsHoraFin := Hoja.Cells[CurFila,CHFin].text;
            ValsHoraFin := AnsiLeftStr(ValsHoraFin, 5 );
            ValsDescripcion := Hoja.Cells[CurFila,CDescripcion].text;
            ValsPartida := Hoja.Cells[CurFila,CPartida].text;
            ValsFolio := Hoja.Cells[CurFila,CFolio].text;
            ValsOt:= '';

            ValsHoraInicio := StringReplace(ValsHoraInicio, ' ', '', [rfReplaceAll]);
            ValsHoraFin := StringReplace(ValsHoraFin, ' ', '', [rfReplaceAll]);
            ValsPartida := StringReplace(ValsPartida, ' ', '', [rfReplaceAll]);
            ValsFolio := StringReplace(ValsFolio, ' ', '', [rfReplaceAll]);
            ValsOt  := StringReplace(ValsOt, ' ', '', [rfReplaceAll]);

            //trabajando query
            ZUpdt.Active := False;
            ZUpdt.SQL.Clear;
            ZUpdt.SQL.Text := 'INSERT INTO movimientosdeembarcacion (sContrato, dIdFecha, iIdDiario, sTipoEmbarcacion, sClasificacion, sidembarcacion, sidfase,shorainicio,shorafinal,sfactor,mdescripcion,iddiario,stipo,lcontinuo,sorden,snumeroactividad) VALUES '+
                              '(:Contrato, :Fecha, 0,"", :PartBarco,:embarcacion,"",:horainicio,:horafinal,:factor,:descripcion,0,:tipo,"No",:Ot,:PartBarco)';
            ZUpdt.ParamByName('Contrato').AsString := Cont;
            ZUpdt.ParamByName('embarcacion').AsString := '038';
            ZUpdt.ParamByName('Fecha').asdate := strtodate(ValsFecha);
            ZUpdt.ParamByName('PartBarco').AsString := '';
            ZUpdt.ParamByName('horainicio').AsString := ValsHoraInicio;
            ZUpdt.ParamByName('horafinal').AsString := ValsHoraFin;


            ZUpdt.ParamByName('factor').AsString := '0';
            ZUpdt.ParamByName('descripcion').AsString := ValsDescripcion;
            ZUpdt.ParamByName('tipo').AsString := 'ARRIBO';
            ZUpdt.ParamByName('Ot').AsString := '';
            ZUpdt.ExecSQL;

            Hoja.cells[CurFila,1].interior.colorindex := ClLineaImp;
          except
            ;
          end;
        end;
        {$ENDREGION}

        {$REGION 'Importacion apartado Helicoptero'}
        CurFila := FilaIniHelicopteros;
        FilasVacias := 0;
        while FilasVacias < MaxVacias do
        begin
          try
            Hoja.range[ColumnaNombre(1)+inttostr(Curfila)+':'+ColumnaNombre(5)+inttostr(Curfila)].select;
            SValorA := '';
            for i := 1 to CDescripcion do
            begin
              SValorA := SValorA + vartostr(Hoja.cells[curfila,i].value);
            end;

            SValorA := StringReplace(sValorA, ' ', '', [rfReplaceAll]);

            if Length(SValorA) = 0 then
            begin
              Inc(FilasVacias);
              for i := 1 to CDescripcion do
              begin
                Hoja.cells[CurFila,i].interior.colorindex := ClFVacia;

              end;  
              raise exception.Create('pasar a siguiente linea');
            end;

            //Cargamos valores necesarios
            ValsHoraInicio := Hoja.Cells[CurFila,CHInicio].text;
            ValsHoraInicio := AnsiLeftStr(ValsHoraInicio, 5 );
            ValsHoraFin := Hoja.Cells[CurFila,CHFin].text;
            ValsHoraFin := AnsiLeftStr(ValsHoraFin, 5 );
            ValsDescripcion := Hoja.Cells[CurFila,CDescripcion].text;
            ValsHoraInicio := StringReplace(ValsHoraInicio, ' ', '', [rfReplaceAll]);
            ValsHoraFin := StringReplace(ValsHoraFin, ' ', '', [rfReplaceAll]);

            ValsHoraInicio := StringReplace(ValsHoraInicio, ' ', '', [rfReplaceAll]);
            ValsHoraFin := StringReplace(ValsHoraFin, ' ', '', [rfReplaceAll]);

            //trabajando query
            ZUpdt.Active := False;
            ZUpdt.SQL.Clear;
            ZUpdt.SQL.Text := 'INSERT INTO movimientosdeembarcacion (sContrato, dIdFecha, iIdDiario, sTipoEmbarcacion, sClasificacion, sidembarcacion, sidfase,shorainicio,shorafinal,sfactor,mdescripcion,iddiario,stipo,lcontinuo,sorden,snumeroactividad) VALUES '+
                              '(:Contrato, :Fecha, 0,"", :PartBarco,:embarcacion,"",:horainicio,:horafinal,:factor,:descripcion,0,:tipo,"No",:Ot,:PartBarco)';
            ZUpdt.ParamByName('Contrato').AsString := Cont;
            ZUpdt.ParamByName('embarcacion').AsString := '039';
            ZUpdt.ParamByName('Fecha').asdate := strtodate(ValsFecha);
            ZUpdt.ParamByName('PartBarco').AsString := '';
            ZUpdt.ParamByName('horainicio').AsString := ValsHoraInicio;
            ZUpdt.ParamByName('horafinal').AsString := ValsHoraFin;


            ZUpdt.ParamByName('factor').AsString := '0';
            ZUpdt.ParamByName('descripcion').AsString := ValsDescripcion;
            ZUpdt.ParamByName('tipo').AsString := 'ARRIBO';
            ZUpdt.ParamByName('Ot').AsString := '';
            ZUpdt.ExecSQL;

            Hoja.cells[CurFila,1].interior.colorindex := ClLineaImp;

          except
            ;
          end;
          Inc(CurFila);
        end;

        {$ENDREGION}

        ShowMessage('Se importaron varios registros, puede verificar de acuerdo al codigo de colores.');
        ZUpdt.Connection.Commit;
      finally
        lstFol.free;
        Lstemp.free;
        try
          Hoja.cells[FilaIniEmbarcaciones-1,10].value := 'Cod. Color';
          Hoja.cells[FilaIniEmbarcaciones-1,10].interior.colorindex :=ClGris;
          Hoja.cells[FilaIniEmbarcaciones,10].value := 'Fila vacía';
          Hoja.cells[FilaIniEmbarcaciones,10].interior.colorindex :=ClFVacia;
          Hoja.cells[FilaIniEmbarcaciones+1,10].value := 'Hinicio > Hfin';
          Hoja.cells[FilaIniEmbarcaciones+1,10].interior.colorindex := ClHFinMenor;
          Hoja.cells[FilaIniEmbarcaciones+2,10].value := 'Vacío';
          Hoja.cells[FilaIniEmbarcaciones+2,10].interior.colorindex := ClCeldaVacia;
          Hoja.cells[FilaIniEmbarcaciones+3,10].value := 'No en sistema';
          Hoja.cells[FilaIniEmbarcaciones+3,10].interior.colorindex := ClNoBd;
          Hoja.cells[FilaIniEmbarcaciones+4,10].value := 'Folio no en ot';
          Hoja.cells[FilaIniEmbarcaciones+4,10].interior.colorindex := ClFNoOt;
          Hoja.cells[FilaIniEmbarcaciones+5,10].value := 'Ok';
          Hoja.cells[FilaIniEmbarcaciones+5,10].interior.colorindex := ClOk;
          Hoja.cells[FilaIniEmbarcaciones+6,10].value := 'Lin Importado';
          Hoja.cells[FilaIniEmbarcaciones+6,10].interior.colorindex := ClLineaImp;
          hoja.range[ColumnaNombre(10)+inttostr(FilaIniEmbarcaciones-1)+':'+ColumnaNombre(10)+inttostr(FilaIniEmbarcaciones+6)].Borders.LineStyle := xlContinuous;
          hoja.range[ColumnaNombre(10)+inttostr(FilaIniEmbarcaciones-1)+':'+ColumnaNombre(10)+inttostr(FilaIniEmbarcaciones+6)].HorizontalAlignment:=-4108;
        except
          ;
        end;
      end;

      if MostrarExcel then
        Excel.Visible := MostrarExcel
      else
        ExcelClose(Excel,False);
    except
      on e:Exception do
      begin
        if Assigned(ZUpdt) then
          if ZUpdt.Connection.InTransaction then
            ZUpdt.Connection.Rollback;
        ShowMessage('No se puede importar la plantilla por el siguiente motivo:'+#10+e.Message);
        if MostrarExcel then
          Excel.Visible := MostrarExcel
        else
          ExcelClose(Excel,False);
      end;
    end;
  end;

begin
  PathExl := '';
  if DireccionPlantilla then
    ImportarPlantilla(PathExl,Contrat,Emb,False)
end;

procedure TfrmAdmonyTiempos.frmBarra2btnDeleteClick(Sender: TObject);
begin
  if ReporteLock then
  begin
    messageDLg('El Reporte Dairio se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;
  if movimientosdeembarcacion.RecordCount > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
       { lModificar := lfnVerificarHorasEmbarcacion( global_contrato, FloatToStr( dtFecha2.Date ), tsHoraInicio.Text, '' ); }
       // if lModificar then
        movimientosdeembarcacion.Delete;
      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al eliminar registro en arribo de embarcaciones', 0);
        end;
      end
    end

end;

procedure TfrmAdmonyTiempos.tsHoraInicioEnter(Sender: TObject);
begin
  tsHoraInicio.Color := Global_Color_Entrada;
end;

procedure TfrmAdmonyTiempos.tsHoraFinalEnter(Sender: TObject);
begin
  tsHoraFinal.Color := Global_Color_Entrada;
end;

procedure TfrmAdmonyTiempos.tsHoraInicioExit(Sender: TObject);
begin
    if ValidaHorario(tsHoraInicio.Text) = false then
       tsHoraInicio.SetFocus
    else
       tsHoraInicio.Color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos.tsHoraInicioKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsHoraFinal.SetFocus;
end;

procedure TfrmAdmonyTiempos.tsHoraFinalExit(Sender: TObject);
begin
    if ValidaHorario(tsHoraFinal.Text) = false then
       tsHoraFinal.SetFocus
    else
       tsHoraFinal.Color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos.tsHoraFinalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tmDescripcion.SetFocus;
end;

procedure TfrmAdmonyTiempos.frmBarra2btnEditClick(Sender: TObject);
begin
  if movimientosdeembarcacion.RecordCount < 1 then
    exit;
  if ReporteLock then
  begin
    messageDLg('El Reporte Dairio se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;
  dbMovtosEmbarcacion.Enabled := False;
    //manejo de tab carmen
  tabexistencias.PageControl.Pages[0].TabVisible := false;
  tabexistencias.PageControl.Pages[1].TabVisible := false;
  tabexistencias.PageControl.Pages[3].TabVisible := false;

  try
    PanelArribo.Enabled := True;
    frmBarra2.btnEditClick(Sender);
    dBEmbarcaciones.Enabled := True;
    tsHoraInicio.ReadOnly := false;
    tsHoraFinal.ReadOnly := False;
    tmDescripcion.ReadOnly := False;
    sOpcionEmb := 'Edit';
    movimientosdeembarcacion.Edit;
    movimientosdeEmbarcacion.FieldValues['sOrden'] := '';
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al editar registro en arribo de embarcaciones', 0);
    end;
  end;
  //activapop2(arriboembarcaciones, popupprincipal);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmAdmonyTiempos.BitBtn1Click(Sender: TObject);
begin
  PanelMovimientosxfolio.Visible := False;
end;

procedure TfrmAdmonyTiempos.btnAjustaClick(Sender: TObject);
var
  qAjuste: TZReadOnlyQuery;
begin
  try
      if MovimientosdeBarco.RecordCount > 0 then
      begin
          MovimientosdeBarco.First ;
          while NOT MovimientosdeBarco.Eof do
          begin
              GrabarCheckCombo(CmbFolios);
              if (MovimientosdeBarco.FieldByName('sActividades').AsString <> 'MOV') and (MovimientosdeBarco.FieldByName('sIdFase').AsString = 'OPER' ) then
                 TdProrrateoFolio(MovimientosdeBarco.FieldByName('sContrato').AsString,tdFecha.Date,MovimientosdeBarco.FieldByName('iIdDiario').AsInteger);

              MovimientosdeBarco.Next
          end;
          MovimientosdeBarco.Refresh ;

      end;

  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Calculos de Prorrateos en BASE A PERNOCTAS', 'Al ajustar ordenes', 0);
    end;
  end;
end;

procedure TfrmAdmonyTiempos.btnAplicarClick(Sender: TObject);
var
   dCantidadMov : double;
    zqReporte: TZQuery;
begin
    dCantidadMov := 0;
    zqMovimientosxfolio.First;
    while not zqMovimientosxfolio.Eof do
    begin
        //Sumamos los movmimientos de barco si pasan de la unidad..
        dCantidadMov := dCantidadMov +  zqMovimientosxfolio.FieldValues['sFactor'];
        zqMovimientosxfolio.Next;
    end;

    if dCantidadMov > 1.0 then
       messageDLG('La suma de los Factores es mayor a 1.00000 ', mtInformation, [mbOk], 0)
    else
       PanelMovimientosxfolio.Visible := False;

    //Guardamos en el reporte diario si se indico recalcular el factor de barco x folio..
    zqReporte := TZQuery.Create(Self);
    zqReporte.Connection := Connection.zConnection;

    zqReporte.Active := False;
    zqReporte.SQL.Clear;
    zqReporte.SQL.Add('Update reportediario set lAjustaFactorxFolio =:Aplica where sContrato =:Contrato and dIdFecha =:Fecha ');
    zqReporte.ParamByName('Contrato').AsString := global_contrato_barco;
    zqReporte.ParamByName('Fecha').AsDate      := tdFecha.Date;
    if chkCalcula.Checked then
       zqReporte.ParamByName('Aplica').AsString   := 'Si'
    else
       zqReporte.ParamByName('Aplica').AsString   := 'No';
    zqReporte.ExecSQL;
    zqReporte.Destroy;
end;

procedure TfrmAdmonyTiempos.btnDownClick(Sender: TObject);
begin
    OrdenarOrdenes('Abajo');
end;

procedure TfrmAdmonyTiempos.btnExitClick(Sender: TObject);
begin
  PnlExistenciasC.Visible := False;
end;

procedure TfrmAdmonyTiempos.btnOkClick(Sender: TObject);
begin
    PanelOrdena.Visible := False;
end;

procedure TfrmAdmonyTiempos.btnUpClick(Sender: TObject);
begin
    OrdenarOrdenes('Arriba');
end;

procedure TfrmAdmonyTiempos.Split
  (const Delimiter: Char;
  Input: string;
  const Strings: TStrings);
begin
  Assert(Assigned(Strings));
  Strings.Clear;
  Strings.Delimiter := Delimiter;
  Strings.DelimitedText := Input;
end;

procedure TfrmAdmonyTiempos.sPronosticoEnter(Sender: TObject);
begin
   sPronostico.color := Global_Color_Entrada;
   lCopiaObjeto := True;
end;

procedure TfrmAdmonyTiempos.sPronosticoExit(Sender: TObject);
begin
    sPronostico.color := Global_Color_Salida;
    lCopiaObjeto := False;
end;

procedure TfrmAdmonyTiempos.Button1Click(Sender: TObject);
var
  qry: TZQuery;
begin

  qry := TZQuery.Create(Self);
  qry.Connection := Connection.zConnection;
  qry.sql.clear;
  qry.sql.add('update movimientosdeembarcacion set sIdPlataforma = :sIdPlataforma ' +
    'where sContrato = :contrato and dIdFecha between :fechai and :fechaf');
  qry.Params.ParamByName('sIdPlataforma').Value := Plataformas.FieldValues['sIdPlataforma'];
  qry.Params.ParamByName('contrato').Value := global_contrato_Barco;
  qry.Params.ParamByName('fechai').DataType := ftDate;
  qry.Params.ParamByName('fechai').Value := Fechai.Date;
  qry.Params.ParamByName('fechaf').DataType := ftDate;
  qry.Params.ParamByName('fechaf').Value := Fechaf.Date;
  qry.ExecSQL;

  qryTiempoCia.Active := false;
  qryTiempoCia.SQL.Clear;
  qryTiempoCia.SQL.Add('call rptTiempoxCompania(:contrato,:fechai,:fechaf)');
  qryTiempoCia.Params.ParamByName('contrato').DataType := ftString;
  qryTiempoCia.Params.ParamByName('contrato').value := global_contrato_Barco;
  qryTiempoCia.Params.ParamByName('fechai').DataType := ftDate;
  qryTiempoCia.Params.ParamByName('fechai').value := Fechai.Date;
  qryTiempoCia.Params.ParamByName('fechaf').DataType := ftDate;
  qryTiempoCia.Params.ParamByName('fechaf').value := Fechaf.Date;
  qryTiempoCia.Open;

  fReporte.LoadFromFile(global_files + 'rptTiempoCia.fr3');
  fReporte.ShowReport();
  Panel1.Visible := false;

  Connection.zConnection.Reconnect;
  close;
end;

procedure TfrmAdmonyTiempos.Can1Click(Sender: TObject);
begin
  if tabexistencias.PageControl.Pages[0].TabVisible = true then
    frmBarra1.btnCancel.Click;
  if tabexistencias.PageControl.Pages[1].TabVisible = true then
    frmBarra3.btnCancel.Click;
  if tabexistencias.PageControl.Pages[2].TabVisible = true then
    frmBarra2.btnCancel.Click;
  if tabexistencias.PageControl.Pages[3].TabVisible = true then
    frmBarra4.btnCancel.Click;
end;

procedure TfrmAdmonyTiempos.chkDescuentoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    frmBarra4.btnPost.SetFocus;
end;
function TfrmAdmonyTiempos.OrdenarCadena(CadOriginal:String):string;
var CadORg,minimo:string;
    Item,ItemSig:Integer;
    Lista,ListaOrd:TStringList;
function EsMayor(Actual,Siguiente:String):Boolean;
var CActual,CSiguiente:string;
    TActual,TSiguiente,P:Integer;
    RMAyor,Continue:Boolean;
begin
  TActual := Length(actual);
  TSiguiente := Length(Siguiente);
  RMAyor := False;
  for P:= 1 to Tactual do
  begin
    Continue := True;
    if (TSiguiente >= P) and Continue then
    begin
      CActual :=  AnsiMidStr( Actual,p, 1 );
      CSiguiente := AnsiMidStr( Siguiente,p, 1 );
      if (CActual < CSiguiente) or ((p = tactual) and (tsiguiente > tactual))  then
      begin
        RMAyor := True;
        Continue := False;
      end;
    end;
  end;
  Result := RMAyor;
end;

begin
  Lista := TStringList.Create;
  ListaOrd := TStringList.Create;

  Lista.Sorted := False;
  ListaOrd.Sorted := False;
  try
    CadOrg := Trim(CadOriginal);
    Lista.CommaText := CadOrg;

    for Item := Lista.Count-1 downto 0 do
    begin
      if Length(Trim(Lista[item])) = 0 then
        Lista.Delete(Item);
    end;
    for Item := Lista.Count-1 downto 0 do
    begin
      Minimo := Lista[Item];
      for ItemSig := Lista.Count-1 downto 0 do
      begin
        if not EsMayor(minimo,Lista[Itemsig]) then
          minimo := Lista[ItemSig];
      end;
      ListaOrd.Add(Minimo);
      Lista.Delete(Lista.IndexOf(Minimo));
    end;
    Lista.CommaText := ListaOrd.CommaText;

    Result := Lista.CommaText;
  finally
    Lista.Free;
    ListaOrd.Free;
  end;
end;
procedure TfrmAdmonyTiempos.CmbAnnoChange(Sender: TObject);
begin
  CargaEmb(CmbAnno.Text,inttostr(CmbMeses.ItemIndex+1));
end;

procedure TfrmAdmonyTiempos.CmbFoliosExit(Sender: TObject);

begin

  zqPartida.Active := False;
  zqPartida.ParamByName('Contrato').AsString  := zqOrdenes.FieldValues['sContrato'];
  zqPartida.ParamByName('Convenio').AsString  := global_convenio;
  zqPartida.ParamByName('Orden').AsString     := cmbFolios.Text;
  zqPartida.Open;
  zqPartida.First;
  CbPartidas.Items.Clear;
  while not zqPartida.Eof do
  begin
    CbPartidas.Items.Add(zqPartida.FieldByName('snumeroactividad').asstring);
    zqPartida.Next;
  end;
  EstableceCheckPArtidas(MovimientosdeBarco.FieldByName('sactividades').asstring,CbPartidas);

end;

procedure TfrmAdmonyTiempos.CmbMesesChange(Sender: TObject);
begin
  CargaEmb(CmbAnno.Text,inttostr(CmbMeses.ItemIndex+1));
end;

procedure TfrmAdmonyTiempos.cmdEmbarcacionesClick(Sender: TObject);
begin
    global_frmActivo := 'frm_AdmonyTiempos';
    Application.CreateForm(TfrmEmbarcaciones, frmEmbarcaciones);
    frmEmbarcaciones.FormStyle := fsNormal;
    frmEmbarcaciones.show;
end;

procedure TfrmAdmonyTiempos.Copy1Click(Sender: TObject);
begin
  if lCopiaObjeto then
  begin
     if tabexistencias.PageControl.Pages[0].TabVisible = true then
        tmDescripcion2.CopyToClipboard;
     If tabexistencias.PageControl.Pages[1].TabVisible = true then
        tmDescripcion.CopyToClipboard;

  end
  else
  begin
      if tabexistencias.PageControl.Pages[0].TabVisible = true then
      begin
        UtGrid.CopyRowsToClip;
      end;
      if tabexistencias.PageControl.Pages[1].TabVisible = true then
      begin
        UtGrid2.CopyRowsToClip;
      end;
      if tabexistencias.PageControl.Pages[2].TabVisible = true then
      begin
        UtGrid3.CopyRowsToClip;
      end;
      if tabexistencias.PageControl.Pages[3].TabVisible = true then
      begin
        UtGrid4.CopyRowsToClip;
      end;
  end;
end;

procedure TfrmAdmonyTiempos.frmBarra3btnAddClick(Sender: TObject);
begin

  if ReporteLock then
  begin
    messageDLg('El Reporte Diario se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;
  dbCondiciones.Enabled := False;
  tabexistencias.PageControl.Pages[0].TabVisible := false;
  tabexistencias.PageControl.Pages[2].TabVisible := false;
  tabexistencias.PageControl.Pages[3].TabVisible := false;

  frmBarra3.btnAddClick(Sender);
  //des-habilitar componentes para editar
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;

  sHorario.Enabled := true;
  iIdDireccion.Enabled := true;
  iIdCondiciones.Enabled := true;
  sPronostico.Enabled := true;
  dbedtsCantidad.Enabled := True;

  //Abrir componente para agregar
  PanelCondiciones.Enabled := true;
  qryCondicionesClimatologicas.Append;
  qryCondicionesClimatologicas.FieldValues['sContrato']   := Global_Contrato;
  qryCondicionesClimatologicas.FieldValues['dIdFecha']    := tdFecha.Date;
  qryCondicionesClimatologicas.FieldValues['dcantidad']   := 0;
  qryCondicionesClimatologicas.FieldValues['mPronostico'] := '*';
  qryCondicionesClimatologicas.FieldValues['sHorario']    := '24:00';
  qryCondicionesClimatologicas.FieldValues['sLocalizacion']  := '*';
  qryCondicionesClimatologicas.FieldValues['iiddireccion']   := qryDirecciones.FieldByName('iIddireccion').AsInteger;
  iIdCondiciones.SetFocus;

  BotonPermiso.permisosBotones(frmBarra3);

end;

procedure TfrmAdmonyTiempos.frmBarra3btnCancelClick(Sender: TObject);
begin
  tabexistencias.PageControl.Pages[0].TabVisible := true;
  tabexistencias.PageControl.Pages[2].TabVisible := true;
  tabexistencias.PageControl.Pages[3].TabVisible := true;

  frmBarra3.btnCancelClick(Sender);
  qryCondicionesClimatologicas.Cancel;
  desactivapop(popupprincipal);
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  mkHora1.ReadOnly := False;
  BotonPermiso.permisosBotones(frmBarra3);
  dbCondiciones.Enabled := True;
  PanelCondiciones.Enabled := false;

  lblBusca.Text := '';
  QryDirecciones.Filtered := False;

end;

procedure TfrmAdmonyTiempos.frmBarra3btnRefreshClick(Sender: TObject);
begin
   qryCondicionesClimatologicas.Refresh;
   PanelCondiciones.Enabled := false;
end;

procedure TfrmAdmonyTiempos.frmBarra3btnDeleteClick(Sender: TObject);
begin
  if ReporteLock then
  begin
    messageDLg('El Reporte Dairio se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;

  try
    if qryCondicionesClimatologicas.RecordCount > 0 then
      if MessageDlg('Desea eliminar el Registro Activo?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        qryCondicionesClimatologicas.Delete;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al eliminar registro en condiciones meteorologicas.', 0);
    end;
  end;
end;

procedure TfrmAdmonyTiempos.frmBarra3btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
  {Validacion de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('Condicion');  nombres.Add('Velocidad');
  cadenas.Add(iIdCondiciones.Text);    cadenas.Add(dbedtsCantidad.Text);
  if not validaTexto(nombres, cadenas, '', '') then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
  {Continua insercion de datos}
  tabexistencias.PageControl.Pages[0].TabVisible := true;
  tabexistencias.PageControl.Pages[2].TabVisible := true;
  tabexistencias.PageControl.Pages[3].TabVisible := true;
  if (tsLocalizacion.Text <> '*') and (tsLocalizacion.Text <> '') then
     sLocalizacion := tsLocalizacion.Text;

  try
    //Actualizar o Insertar elregistro
    qryCondicionesClimatologicas.FieldValues['sHorario']    := '24:00';
    qryCondicionesClimatologicas.Post;

    if tsLocalizacion.Text <> '' then
    begin
        //Actualizamos la ubiacion para el mismo horario..
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Update condicionesclimatologicas set sLocalizacion =:Localizacion where sContrato =:contrato and dIdFecha =:Fecha and sHorario =:hora');
        connection.zCommand.ParamByName('Contrato').AsString     := global_Contrato_Barco;
        connection.zCommand.ParamByName('Fecha').AsDate          := tdFecha.Date;
        connection.zCommand.ParamByName('Hora').AsString         := qryCondicionesClimatologicas.FieldValues['sHorario'];
        connection.zCommand.ParamByName('Localizacion').AsString := sLocalizacion;
        connection.zCommand.ExecSQL;

        qryCondicionesClimatologicas.Refresh;
    end;

    PanelCondiciones.Enabled := false;
    Insertar1.Enabled := True;
    Editar1.Enabled := True;
    Registrar1.Enabled := False;
    Can1.Enabled := False;
    Eliminar1.Enabled := True;
    Refresh1.Enabled := True;
    Salir1.Enabled := True;
    dbCondiciones.Enabled := True;
    frmBarra3.btnPostClick(Sender);
    desactivapop(popupprincipal);
    
    lblBusca.Text := '';
    QryDirecciones.Filtered := False;

  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al salvar registro en condiciones meteorologicas', 0);
      frmbarra3.btnCancel.Click;
    end;
  end;
  BotonPermiso.permisosBotones(frmBarra3);
end;

procedure TfrmAdmonyTiempos.frmBarra3btnEditClick(Sender: TObject);
begin

  if qryCondicionesClimatologicas.RecordCount < 1 then
    exit;
  if ReporteLock then
  begin
    messageDLg('El Reporte Dairio se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;
  dbCondiciones.Enabled := False;
  tabexistencias.PageControl.Pages[0].TabVisible := false;
  tabexistencias.PageControl.Pages[2].TabVisible := false;
  tabexistencias.PageControl.Pages[3].TabVisible := false;

  frmBarra3.btnEditClick(Sender);
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;

  PanelCondiciones.Enabled := true;
  iIdDireccion.Enabled := true;
  iIdCondiciones.Enabled := True;
  sPronostico.Enabled := true;
  sHorario.Enabled := true;
  dbedtsCantidad.Enabled := True;
  qryCondicionesClimatologicas.Edit;
  //activapop2(ConClimatologicas, popupprincipal);
  iIdCondiciones.SetFocus;
  BotonPermiso.permisosBotones(frmBarra3);

end;

procedure TfrmAdmonyTiempos.frmBarra3btnExitClick(Sender: TObject);
begin
  frmBarra3.btnExitClick(Sender);
  frmBarra3.btnCancel.Click;
  close;
end;

procedure TfrmAdmonyTiempos.dIdFechaKeyPress(Sender: TObject;
  var Key: Char);
begin
 { if key = #13 then
    iIdCondiciones.SetFocus;}
end;

procedure TfrmAdmonyTiempos.TabExistenciasEnter(Sender: TObject);
var
   indice : integer;
   lEncuentra : boolean;
begin
     Embarcaciones.Active := False;
     Embarcaciones.sql.Clear;
     Embarcaciones.sql.Add('select distinct sIdEmbarcacion, sDescripcion, sTipo from embarcaciones ' +
                           'Where sContrato=:Contrato order by sDescripcion') ;
     Embarcaciones.Params.ParamByName('Contrato').DataType := ftString ;
     Embarcaciones.Params.ParamByName('Contrato').Value    := Global_Contrato_barco;
     Embarcaciones.Open;

     //Primero revisamos cuales son las ordenes dadas de alta en los movimintos de barco..
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('select sContrato as sOrden from consumosdecombustibleporequipo where dIdFecha =:fecha '+
                                'group by sContrato '+
                                'union '+
                                'select sOrden from movimientosdeembarcacion where sContrato = :contrato and dIdFecha = :fecha '+
                                'and sClasificacion <> "" group by sOrden ');
    connection.QryBusca.ParamByName('Contrato').AsString := global_contrato_barco;
    connection.QryBusca.ParamByName('Fecha').AsDate      := tdFecha.Date;
    connection.QryBusca.Open;

    zqOrdenaOrden.Active := False;
    zqOrdenaOrden.ParamByName('Fecha').AsDate := tdFecha.Date;
    zqOrdenaOrden.Open;

    indice := 0;
    while not connection.QryBusca.Eof do
    begin
        lEncuentra := False;
        zqOrdenaOrden.First;
        while not zqOrdenaOrden.Eof do
        begin
            if zqOrdenaOrden.FieldValues['sOrden'] = connection.QryBusca.FieldValues['sOrden'] then
               lEncuentra := True;
            indice := zqOrdenaOrden.FieldValues['iIdOrden'];
            zqOrdenaOrden.Next;
        end;

        if lEncuentra = False then
        begin
            if connection.QryBusca.FieldByName('sOrden').AsString <> '' then
            begin
                inc(indice);
                connection.QryBusca2.Active := False;
                connection.QryBusca2.SQL.Clear;
                connection.QryBusca2.SQL.Add('insert into recursosordenados_orden (iIdOrden, dIdFecha, sOrden) values (:Id, :fecha, :Orden)');
                connection.QryBusca2.ParamByName('Id').AsInteger   := indice;
                connection.QryBusca2.ParamByName('Fecha').AsDate   := tdFecha.Date;
                connection.QryBusca2.ParamByName('Orden').AsString := connection.QryBusca.FieldByName('sOrden').AsString;
                connection.QryBusca2.ExecSQL;
                zqOrdenaOrden.Refresh;
            end;
        end;

        connection.QryBusca.Next;
    end;

end;

procedure TfrmAdmonyTiempos.iIdCondicionesEnter(Sender: TObject);
begin
    iIdCondiciones.color := Global_Color_Entrada; 
end;

procedure TfrmAdmonyTiempos.iIdCondicionesExit(Sender: TObject);
begin
     iIdCondiciones.color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos.iIdCondicionesKeyPress(Sender: TObject;
  var Key: Char);
begin
   if key = #13 then
      iIdDireccion.SetFocus;
end;

procedure TfrmAdmonyTiempos.iIdDireccionEnter(Sender: TObject);
begin
    iIdDireccion.color := Global_Color_Entrada;
end;

procedure TfrmAdmonyTiempos.iIdDireccionExit(Sender: TObject);
begin
    iIdDireccion.color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos.iIdDireccionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    sHorario.SetFocus
  else
     if key = #8 then
        lblBusca.text := ''
     else
         lblBusca.text := lblBusca.text + char(key);
end;

procedure TfrmAdmonyTiempos.iIdDireccionKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    filtracondicion;
end;

procedure TfrmAdmonyTiempos.qryCondicionesClimatologicasAfterScroll(
  DataSet: TDataSet);
begin
   if qryCondicionesClimatologicas.RecordCount > 0 then
    if qryCondicionesClimatologicas.Eof = False then
    begin
      iIdCondiciones.KeyValue := qryCondicionesClimatologicas.FieldValues['iIdCondicion'];
      iIdDireccion.KeyValue := qryCondicionesClimatologicas.FieldValues['iIdDireccion'];
      sPronostico.Text := qryCondicionesClimatologicas.FieldValues['mPronostico'];
      sHorario.Text := qryCondicionesClimatologicas.FieldValues['sHorario'];
      dCantidad.Text := qryCondicionesClimatologicas.FieldValues['dCantidad'];
    end;
end;

procedure TfrmAdmonyTiempos.qryCondicionesClimatologicasBeforePost(
  DataSet: TDataSet);
begin
    qryCondicionesClimatologicas.FieldValues['sIdTurno'] := global_turno;
    qryCondicionesClimatologicas.FieldValues['sHorario'] := sHorario.Text;
    if dCantidad.Text = '' then
      qryCondicionesClimatologicas.FieldValues['dCantidad'] := 0
    else
      qryCondicionesClimatologicas.FieldValues['dCantidad'] := dCantidad.Text;
    if sPronostico.Text = '' then
      qryCondicionesClimatologicas.FieldValues['mPronostico'] := '*';
end;

procedure TfrmAdmonyTiempos.qryCondicionesClimatologicasCalcFields(
  DataSet: TDataSet);
begin
  //Cargar los comentarios de la tabla codiciones en el grid principal
  connection.QryBusca.Active := false;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select sDescripcion from condiciones where iIdCondicion=:iIdCondicion');
  connection.QryBusca.Params.ParamByName('iIdCondicion').Value := qryCondicionesClimatologicas.FieldValues['iIdCondicion'];
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
    qryCondicionesClimatologicasCalcCondiciones.Text := connection.QryBusca.FieldValues['sDescripcion'];

  //Cargar los comentarios de la tabla direcciones en el grid principal
  connection.QryBusca.Active := false;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select sDescripcion from direcciones where iIdDireccion=:iIdDireccion');
  connection.QryBusca.Params.ParamByName('iIdDireccion').Value := qryCondicionesClimatologicas.FieldValues['iIdDireccion'];
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
    qryCondicionesClimatologicasCalcDireccion.Text := connection.QryBusca.FieldValues['sDescripcion'];

  //Cargar los comentarios de la tabla codiciones en el grid principal
  connection.QryBusca.Active := false;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select sMedida from condiciones where iIdCondicion=:iIdCondicion');
  connection.QryBusca.Params.ParamByName('iIdCondicion').Value := qryCondicionesClimatologicas.FieldValues['iIdCondicion'];
  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount > 0 then
    qryCondicionesClimatologicasCalcCMedida.Text := connection.QryBusca.FieldValues['sMedida'];
end;


procedure TfrmAdmonyTiempos.Editar1Click(Sender: TObject);
begin
  if tabexistencias.PageControl.Pages[0].TabVisible = true then
    frmBarra1.btnEdit.Click;
  if tabexistencias.PageControl.Pages[1].TabVisible = true then
    frmBarra3.btnEdit.Click;
  if tabexistencias.PageControl.Pages[2].TabVisible = true then
    frmBarra2.btnEdit.Click;
  if tabexistencias.PageControl.Pages[3].TabVisible = true then
    frmBarra4.btnEdit.Click;
end;

procedure TfrmAdmonyTiempos.Eliminar1Click(Sender: TObject);
begin
  if tabexistencias.PageControl.Pages[0].TabVisible = true then
    frmBarra1.btnDelete.Click;
  if tabexistencias.PageControl.Pages[1].TabVisible = true then
    frmBarra3.btnDelete.Click;
  if tabexistencias.PageControl.Pages[2].TabVisible = true then
    frmBarra2.btnDelete.Click;
  if tabexistencias.PageControl.Pages[3].TabVisible = true then
    frmBarra4.btnDelete.Click;
end;

procedure TfrmAdmonyTiempos.Embarcaciones2AfterScroll(DataSet: TDataSet);
begin
  //********************BRITO 22/11/10****************************************
  if Embarcaciones2.RecordCount > 0 then
    if Embarcaciones2.Eof = False then
      chkDescuento.Visible := not (Embarcaciones2.FieldByName('sTipo').AsString = 'Principal');
 //********************BRITO 22/11/10****************************************
end;

procedure TfrmAdmonyTiempos.frmBarra2btnCancelClick(Sender: TObject);
begin
//manejo de tab carmen
  tabexistencias.PageControl.Pages[0].TabVisible := true;
  tabexistencias.PageControl.Pages[1].TabVisible := true;
  tabexistencias.PageControl.Pages[3].TabVisible := true;
  sOpcionEmb := '';
  frmBarra2.btnCancelClick(Sender);
  tsHoraInicio.ReadOnly := True;
  tsHoraFinal.ReadOnly := True;
  tmDescripcion.ReadOnly := True;
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  mkHora1.ReadOnly := False;
  PanelArribo.Enabled := False;
        dbMovtosEmbarcacion.Enabled := True;
 // tsHoraInicio.ReadOnly := true;
  try
    movimientosdeembarcacion.Cancel;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al agregar registro en arribo de embarcaciones', 0);
    end;
  end;
  //desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra2);

  lblBusca.Text := '';
  embarcaciones.Filtered := False;

end;

procedure TfrmAdmonyTiempos.frmBarra2btnRefreshClick(Sender: TObject);
begin
  try      
    movimientosdeembarcacion.Refresh;
    PanelArribo.Enabled := False;
  except
  end
end;

procedure TfrmAdmonyTiempos.tdEmbarcacionExistEnter(Sender: TObject);
begin
  tdEmbarcacionExist.Color := global_color_entrada
end;

procedure TfrmAdmonyTiempos.tdEmbarcacionExistExit(Sender: TObject);
begin
    tdEmbarcacionExist.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.tdEmbarcacionExistKeyPress(Sender: TObject;
  var Key: Char);
begin
  if iIdrecursoExistencia.Enabled = true then
    if key = #13 then
      iIdRecursoExistencia.SetFocus;
end;

procedure TfrmAdmonyTiempos.tdFechaEnter(Sender: TObject);
begin
  tdFecha.Color := global_color_entrada
end;

procedure TfrmAdmonyTiempos.tdFechaExit(Sender: TObject);
var
  sFactor: string;
  dProrrateo: Double;
  dAjuste: Double;
  iMultiplo: Integer;
  iDecimales: Integer;
begin

    global_barco :=  connection.configuracion.FieldValues['sIdEmbarcacion'] ;

    QryBarcoVigencia.Active := False;
    QryBarcoVigencia.SQL.Clear;
    QryBarcoVigencia.SQL.Add('select sIdEmbarcacion from embarcacion_vigencia '+
                       'where sContrato =:Contrato and dFechaInicio <= :Fecha and dFechaFinal >=:Fecha order by dFechaInicio');
    QryBarcoVigencia.ParamByName('Contrato').AsString := global_contrato_barco;
    QryBarcoVigencia.ParamByName('Fecha').AsDate      := tdFecha.Date;
    QryBarcoVigencia.Open;

    if QryBarcoVigencia.RecordCount > 0 then
       global_barco := QryBarcoVigencia.FieldValues['sIdEmbarcacion']
    else
       messageDLG('No existe una Vigencia de Embarcacion Principal', mtInformation, [mbOk], 0);



  tdFecha.Color := global_color_salida;

  movimientosdebarco.Active := False;
  movimientosdebarco.Params.ParamByName('Contrato').DataType := ftString;
  movimientosdebarco.Params.ParamByName('Contrato').Value    := Global_Contrato_Barco;
  movimientosdebarco.Params.ParamByName('Fecha').DataType := ftDate;
  movimientosdebarco.Params.ParamByName('Fecha').Value := tdFecha.Date;
  movimientosdebarco.Params.ParamByName('ContratoNormal').DataType := ftString;
  if tsOrdenesSeleccion.Text = '<<Todas>>' then
     movimientosdebarco.Params.ParamByName('ContratoNormal').Value := '%'
  else
     movimientosdebarco.Params.ParamByName('ContratoNormal').Value := tsOrdenesSeleccion.Text;
  movimientosdebarco.Open;

  qryCondiciones.Active := True;
  qryDirecciones.Active := True;

  qryCondicionesClimatologicas.Active := False;
  qryCondicionesClimatologicas.ParamByName('Fecha').AsDateTime := tdFecha.DateTime;
  qryCondicionesClimatologicas.ParamByName('Contrato').AsString := global_contrato;
  qryCondicionesClimatologicas.Open;

  movimientosdeEmbarcacion.Active := False;
  movimientosdeEmbarcacion.Params.ParamByName('Contrato').DataType := ftString;
  movimientosdeEmbarcacion.Params.ParamByName('Contrato').Value    := Global_Contrato_barco;
  movimientosdeEmbarcacion.Params.ParamByName('Fecha').DataType    := ftDate;
  movimientosdeEmbarcacion.Params.ParamByName('Fecha').Value       := tdFecha.date;
  movimientosdeEmbarcacion.Open;

  qryRecursos.Active := False;
  qryRecursos.Params.ParamByName('Fecha').DataType := ftDate;
  qryRecursos.Params.ParamByName('Fecha').Value := tdFecha.Date;
  qryRecursos.Open;

  qryMezclas.Active := False;
  qryMezclas.Open;

  CargarCheckCombo(CmbFolios);

  try
    if movimientosdebarco.recordcount > 0 then

      tdJornada.Text := sFnSumaBarco(movimientosdebarco.FieldValues['dIdFecha'], movimientosdebarco.FieldValues['sIdEmbarcacion'], frmAdmonyTiempos, -1)

    else
      tdJornada.Text := '0.000000';
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al seleccionar fecha', 0);
    end;
  end;
end;

procedure TfrmAdmonyTiempos.tdFechaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
  begin
    if pgAdmon.ActivePageIndex = 0 then
       dbMovBarco.SetFocus;
    if pgAdmon.ActivePageIndex = 1 then
       dbCondiciones.SetFocus;
    if pgAdmon.ActivePageIndex = 2 then
       dbMovtosEmbarcacion.SetFocus;
    if pgAdmon.ActivePageIndex = 3 then
       dbExistencias.SetFocus;
  end;
end;

procedure TfrmAdmonyTiempos.tmDescripcion2Enter(Sender: TObject);
begin
    tmDescripcion2.Color := global_color_entrada;
    lCopiaObjeto := True;
end;

procedure TfrmAdmonyTiempos.tmDescripcion2Exit(Sender: TObject);
begin
    tmDescripcion2.Color := global_color_salida;
    lcopiaObjeto := False;
end;

procedure TfrmAdmonyTiempos.tmDescripcionEnter(Sender: TObject);
begin
   tmDescripcion.Color := global_Color_Entrada;
   lCopiaObjeto := True;

   
end;

procedure TfrmAdmonyTiempos.tmDescripcionExit(Sender: TObject);
begin
    tmDescripcion.Color := global_Color_Salida;
    lCopiaObjeto := False;
end;

procedure TfrmAdmonyTiempos.qryRecursosCalcFields(DataSet: TDataSet);
begin
  if (qryRecursos.State <> dsEdit) and (qryRecursos.State <> dsInsert) then
  begin
      //Cargar los comentarios de la tabla codiciones en el grid principal
    connection.QryBusca.Active := false;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('select sDescripcion,sMedida from recursosdeexistencias where iIdRecursoExistencia= :RecursoExistencia');
    connection.QryBusca.params.paramByName('RecursoExistencia').DataType := ftInteger;
    connection.QryBusca.Params.ParamByName('RecursoExistencia').Value := qryRecursos.FieldValues['iIdRecursoExistencia'];
    connection.QryBusca.Open;

    if connection.QryBusca.RecordCount > 0 then
    begin
      qryRecursosCalcMezclas.Text := connection.QryBusca.FieldValues['sDescripcion'];
      qryRecursosCalcMezclasMedidas.Text := connection.QryBusca.FieldValues['sMedida'];
    end;

    connection.QryBusca.Active := false;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('select sDescripcion from embarcaciones where sIdEmbarcacion =:Embarcacion ');
    connection.QryBusca.params.paramByName('Embarcacion').DataType := ftString;
    connection.QryBusca.Params.ParamByName('Embarcacion').Value := QryRecursos.FieldValues['sIdEmbarcacion'];
    connection.QryBusca.Open;

    if connection.QryBusca.RecordCount > 0 then
      qryRecursossEmbarcacion.Text := connection.QryBusca.FieldValues['sDescripcion'];
  end;

end;

procedure TfrmAdmonyTiempos.rbArriboEnter(Sender: TObject);
begin
    embarcaciones.Locate('sIdEmbarcacion', 'N/A', [loCaseInsensitive]);
    dbEmbarcaciones.KeyValue := embarcaciones.FieldByName('sIdEmbarcacion').AsString;
    label40.Visible := False;
    tsMovimiento.Visible := False;
end;

procedure TfrmAdmonyTiempos.rbMovimientoEnter(Sender: TObject);
begin
    embarcaciones.Locate('sTipo', 'Principal', [loCaseInsensitive]);
    dbEmbarcaciones.KeyValue := embarcaciones.FieldByName('sIdEmbarcacion').AsString;
    label40.Visible := True;
    tsMovimiento.Visible := True;
end;

procedure TfrmAdmonyTiempos.Refresh1Click(Sender: TObject);
begin
  if tabexistencias.PageControl.Pages[0].TabVisible = true then
    frmBarra1.btnRefresh.Click;
  if tabexistencias.PageControl.Pages[1].TabVisible = true then
    frmBarra3.btnRefresh.Click;
  if tabexistencias.PageControl.Pages[2].TabVisible = true then
    frmBarra2.btnRefresh.Click;
  if tabexistencias.PageControl.Pages[3].TabVisible = true then
    frmBarra4.btnRefresh.Click;
end;

procedure TfrmAdmonyTiempos.Registrar1Click(Sender: TObject);
begin
  if tabexistencias.PageControl.Pages[0].TabVisible = true then
    frmBarra1.btnPost.Click;
  if tabexistencias.PageControl.Pages[1].TabVisible = true then
    frmBarra3.btnPost.Click;
  if tabexistencias.PageControl.Pages[2].TabVisible = true then
    frmBarra2.btnPost.Click;
  if tabexistencias.PageControl.Pages[3].TabVisible = true then
    frmBarra4.btnPost.Click;
end;

procedure TfrmAdmonyTiempos.Salir1Click(Sender: TObject);
begin
  if tabexistencias.PageControl.Pages[0].TabVisible = true then
    frmBarra1.btnExit.Click;
  if tabexistencias.PageControl.Pages[1].TabVisible = true then
    frmBarra3.btnExit.Click;
  if tabexistencias.PageControl.Pages[2].TabVisible = true then
    frmBarra2.btnExit.Click;
  if tabexistencias.PageControl.Pages[3].TabVisible = true then
    frmBarra4.btnExit.Click;
end;

procedure TfrmAdmonyTiempos.sHorarioEnter(Sender: TObject);
begin
    sHorario.color := Global_Color_Entrada;
end;

procedure TfrmAdmonyTiempos.sHorarioExit(Sender: TObject);
begin
//    if ValidaHorario(sHorario.Text) = false then
//       sHorario.SetFocus
//    else
//       sHorario.Color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos.sHorarioKeyPress(Sender: TObject; var Key: Char);
begin
    if key = #13 then
       dbedtsCantidad.SetFocus;
end;

procedure TfrmAdmonyTiempos.frmBarra4btnAddClick(Sender: TObject);
begin
  if ReporteLock then
  begin
    messageDLg('El Reporte Dairio se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;
  dbExistencias.Enabled := False;
  //manejo de tab carmen
  tabexistencias.PageControl.Pages[0].TabVisible := false;
  tabexistencias.PageControl.Pages[1].TabVisible := false;
  tabexistencias.PageControl.Pages[2].TabVisible := false;

  Entre := False;
  frmBarra4.btnAddClick(Sender);
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;
  sOpcionEmb := 'Insert';

  PanelExistencias.Enabled := True;

  dConsumo.value := 0;
  dConsumoEquipos.value := 0;
  dPrestamos.value := 0;
  dProduccion.value := 0;
  dAjuste.value := 0;
  dRecibido.value := 0;
  dTrasiego.value := 0;
  dExistenciaActual.Value := 0;
  dExistenciaAnterior.Value := 0;
  dGalones.Value  := 0;

  dConsumos := dConsumo.Value;
  dEquipos  := dConsumoEquipos.Value;

  iIdRecursoExistencia.Enabled := True;
  dProduccion.Enabled := True;
  dAjuste.Enabled := True; //*****************BRITO 17/12/10************
  dRecibido.Enabled := True;
  dConsumo.Enabled := True;
  dConsumoEquipos.Enabled := True;
  dPrestamos.Enabled := True;
  dTrasiego.Enabled := True;
  try
    qryRecursos.Append;
    qryRecursos.FieldValues['sContrato'] := Global_Contrato_Barco;
    qryRecursos.FieldValues['dIdFecha'] := tdFecha.Date;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al agregar registro en existencias', 0);
      frmbarra4.btnCancel.Click;
    end;
  end;
  tdEmbarcacionExist.ListFieldIndex := 0;
  tdEmbarcacionExist.SetFocus;

  ExistenciaAnterior; //<-Procedimiento

  BotonPermiso.permisosBotones(frmBarra4);
  PanelExistencias.Enabled := True;

  label25.Visible := True;
  dConsumoEquipos.Visible := True;

  embarcaciones2.Locate('sTipo', 'Principal', [loCaseInsensitive]);
  tdEmbarcacionExist.KeyValue := embarcaciones2.FieldByName('sIdEmbarcacion').AsString;
  qryRecursos.FieldValues['sIdEmbarcacion'] := Embarcaciones2.FieldValues['sIdEmbarcacion'];

end;

procedure TfrmAdmonyTiempos.dAjusteChange(Sender: TObject);
begin
//TRxCalcEditChangef(dAjuste,'Ajuste');
end;

procedure TfrmAdmonyTiempos.dAjusteEnter(Sender: TObject);
begin
  dAjuste.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dAjusteExit(Sender: TObject);
begin
  dAjuste.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dAjusteKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(dAjuste, key) then
    key := #0;
  if key = #13 then
    chkDescuento.SetFocus;
end;

procedure TfrmAdmonyTiempos.dbCondicionesMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid2.dbGridMouseMoveCoord(x, y);
end;

procedure TfrmAdmonyTiempos.dbCondicionesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid2.DbGridMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TfrmAdmonyTiempos.dbCondicionesTitleClick(Column: TColumn);
begin
  UtGrid2.DbGridTitleClick(Column);
end;

procedure TfrmAdmonyTiempos.dbedtsCantidadEnter(Sender: TObject);
begin
    dbedtsCantidad.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dbedtsCantidadExit(Sender: TObject);
begin
    dbedtsCantidad.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dbedtsCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
     if key = #13 then
       tsLocalizacion.SetFocus;
end;

procedure TfrmAdmonyTiempos.dbEmbarcacionesEnter(Sender: TObject);
begin
  dbembarcaciones.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dbEmbarcacionesExit(Sender: TObject);
begin
  dbembarcaciones.Color := global_color_salida;
  if embarcaciones.FieldByName('sTipo').AsString = 'Principal' then
  begin
      rbArribo.Checked := False;
      rbDisposicion.Checked := False;
      rbDos.Checked := False;
      rbMovimiento.Checked := True;
  end
  else
  begin
      rbArribo.Checked := True;
      rbDisposicion.Checked := False;
      rbDos.Checked := False;
      rbMovimiento.Checked := False;
  end;

end;

procedure TfrmAdmonyTiempos.dbEmbarcacionesKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsHorainicio.SetFocus
  else
     if key = #8 then
        lblBusca.text := ''
     else
         lblBusca.text := lblBusca.text + char(key);

end;

procedure TfrmAdmonyTiempos.dbEmbarcacionesKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     filtra;
end;

procedure TfrmAdmonyTiempos.dbExistenciasDblClick(Sender: TObject);
begin    
    PanelOrdena.Visible := True;
end;

procedure TfrmAdmonyTiempos.dbExistenciasEnter(Sender: TObject);
begin
  if iIdRecursoExistencia.Enabled = True then
//         frmBarra4.btnCancel.Click;
end;

procedure TfrmAdmonyTiempos.dbExistenciasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid4.dbGridMouseMoveCoord(x, y);
end;

procedure TfrmAdmonyTiempos.dbExistenciasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid4.DbGridMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TfrmAdmonyTiempos.dbExistenciasTitleClick(Column: TColumn);
begin
  UtGrid4.DbGridTitleClick(Column);
end;

procedure TfrmAdmonyTiempos.DBMemo1Enter(Sender: TObject);
begin
    dbmemo1.Color :=  global_color_entrada;
end;

procedure TfrmAdmonyTiempos.DBMemo1Exit(Sender: TObject);
begin
    dbmemo1.Color :=  global_color_salida;
end;

procedure TfrmAdmonyTiempos.dbMovBarcoCellClick(Column: TColumn);
begin
  if MovimientosdeBarco.RecordCount > 0 then
  begin
    self.tsIdFase.Field.Index := self.dbMovBarco.Fields[4].Value;
    self.mkHora1.Text := self.dbMovBarco.Fields[2].Value;
    self.mkHora2.Text := self.dbMovBarco.Fields[3].Value;
    self.sSuma.Text := self.dbMovBarco.Fields[6].Value;
     //self.tmDescripcion2.Text := self.dbMovBarco.Fields[5].Value;
  end;
end;


procedure TfrmAdmonyTiempos.dbMovBarcoGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
    if movimientosdebarco.RecordCount > 0 then
    begin
        if movimientosdebarco.FieldValues['sHoraFinal'] <= movimientosdebarco.FieldValues['sHoraInicio'] then
           AFont.Color := clRed
        else
           AFont.Color := clBlack
    end;
end;

procedure TfrmAdmonyTiempos.dbMovBarcoKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (Key = 38) or (Key = 40) then
  begin
    self.tsIdFase.Field.Index := self.dbMovBarco.Fields[4].Value;
    self.mkHora1.Text := self.dbMovBarco.Fields[2].Value;
    self.mkHora2.Text := self.dbMovBarco.Fields[3].Value;
    self.sSuma.Text := self.dbMovBarco.Fields[6].Value;
     //self.tmDescripcion2.Text := self.dbMovBarco.Fields[5].Value;
  end;
end;

procedure TfrmAdmonyTiempos.dbMovBarcoMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x, y);
end;

procedure TfrmAdmonyTiempos.dbMovBarcoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TfrmAdmonyTiempos.dbMovBarcoMouseWheel(Sender: TObject;
  Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint;
  var Handled: Boolean);
begin
  self.tsIdFase.Field.Index := self.dbMovBarco.Fields[4].Value;
  self.mkHora1.Text := self.dbMovBarco.Fields[2].Value;
  self.mkHora2.Text := self.dbMovBarco.Fields[3].Value;
  self.sSuma.Text := self.dbMovBarco.Fields[6].Value;
end;

procedure TfrmAdmonyTiempos.dbMovBarcoTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmAdmonyTiempos.dbMovtosEmbarcacionGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
    if movimientosdeembarcacion.RecordCount > 0 then
    begin
        if movimientosdeembarcacion.FieldValues['sHoraFinal'] <= movimientosdeembarcacion.FieldValues['sHoraInicio'] then
           AFont.Color := clRed
        else
           AFont.Color := clBlack;

        if ( movimientosdeembarcacion.FieldValues['sOrden'] = Null) or (movimientosdeembarcacion.FieldValues['sOrden'] = '') then
           AFont.Color := clRed
        else
           AFont.Color := clBlack
    end;
end;

procedure TfrmAdmonyTiempos.dbMovtosEmbarcacionKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (Key = 38) or (Key = 40) then
  begin
    self.tsHoraInicio.Text := self.dbMovtosEmbarcacion.Fields[0].Value;
    self.tsHoraFinal.Text := self.dbMovtosEmbarcacion.Fields[1].Value;
    self.tmDescripcion.Text := self.dbMovtosEmbarcacion.Fields[2].Value;

    self.rbArribo.Checked := false;
    self.rbDisposicion.Checked := false;
    self.rbDos.Checked := false;

    if dbMovtosEmbarcacion.Fields[3].Value = 'ARRIBO' then
      self.rbArribo.Checked := true;

    if dbMovtosEmbarcacion.Fields[3].Value = 'DISPOSICION' then
      self.rbDisposicion.Checked := true;

    if dbMovtosEmbarcacion.Fields[3].Value = 'ARRIBO Y DISPOSICION' then
      self.rbDos.Checked := true;

  end;
end;

procedure TfrmAdmonyTiempos.dbMovtosEmbarcacionMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid3.dbGridMouseMoveCoord(x, y);
end;

procedure TfrmAdmonyTiempos.dbMovtosEmbarcacionMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid3.DbGridMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TfrmAdmonyTiempos.dbMovtosEmbarcacionMouseWheel(Sender: TObject;
  Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint;
  var Handled: Boolean);
begin
  self.tsHoraInicio.Text := self.dbMovtosEmbarcacion.Fields[0].Value;
  self.tsHoraFinal.Text := self.dbMovtosEmbarcacion.Fields[1].Value;
  self.tmDescripcion.Text := self.dbMovtosEmbarcacion.Fields[2].Value;

  self.rbArribo.Checked := false;
  self.rbDisposicion.Checked := false;
  self.rbDos.Checked := false;

  if dbMovtosEmbarcacion.Fields[3].Value = 'ARRIBO' then
    self.rbArribo.Checked := true;

  if dbMovtosEmbarcacion.Fields[3].Value = 'DISPOSICION' then
    self.rbDisposicion.Checked := true;

  if dbMovtosEmbarcacion.Fields[3].Value = 'ARRIBO Y DISPOSICION' then
    self.rbDos.Checked := true;
end;

procedure TfrmAdmonyTiempos.dbMovtosEmbarcacionTitleClick(Column: TColumn);
begin
  UtGrid3.DbGridTitleClick(Column);
end;

procedure TfrmAdmonyTiempos.frmBarra4btnEditClick(Sender: TObject);
begin
  if qryRecursos.RecordCount < 1 then
    exit;
  if ReporteLock then
  begin
    messageDLg('El Reporte Dairio se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;
  PanelExistencias.Enabled := True;
  dbExistencias.Enabled := False;
  //manejo de tab carmen
  tabexistencias.PageControl.Pages[0].TabVisible := false;
  tabexistencias.PageControl.Pages[1].TabVisible := false;
  tabexistencias.PageControl.Pages[2].TabVisible := false;

  frmBarra4.btnEditClick(Sender);
  Entre := False;
  sOpcionEmb := 'Edit';
  iIdRecursoExistencia.Enabled := True;
  dProduccion.Enabled := True;
  dAjuste.Enabled := True; //*****************BRITO 17/12/10************
  dRecibido.Enabled := True;
  dConsumo.Enabled := True;
  dConsumoEquipos.Enabled := True;
  dPrestamos.Enabled := True;
  dTrasiego.Enabled := True;

  dConsumos := dConsumo.Value;
  dEquipos  := dConsumoEquipos.Value;

  try
    qryRecursos.Edit;
 //   activapop2(TabExistencias, popupprincipal);
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al editar registro en existencias', 0);
    end;
  end;

  iIdRecursoExistencia.SetFocus;
  BotonPermiso.permisosBotones(frmBarra4);
end;

function TfrmAdmonyTiempos.existeReporte: boolean;
begin
  result := false;
  connection.QryBusca.Active := false;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('SELECT dIdFecha FROM reportediario WHERE ' +
    'sContrato = :Contrato AND dIdFecha = :Fecha AND sIdTurno = :Turno');
  connection.QryBusca.ParamByName('Contrato').AsString := qryRecursos.FieldByName('sContrato').AsString;
  connection.QryBusca.ParamByName('Fecha').AsDateTime  := qryRecursos.FieldByName('dIdFecha').AsDateTime;
  connection.QryBusca.ParamByName('Turno').AsString    := global_turno;
  //try
    connection.QryBusca.Open;
    if connection.QryBusca.RecordCount > 0 then
       result := true
    else
    begin
       connection.zCommand.Active := False ;
       connection.zCommand.SQL.Clear ;
       connection.zcommand.SQL.Add ('INSERT INTO reportediario ( sContrato , sOrden, dIdFecha, sNumeroOrden, sIdTurno, sIdConvenio, sNumeroReporte, iPersonal, ' +
                                    'sOperacionInicio, sOperacionFinal, sTiempoEfectivo, sTiempoMuerto, sTiempoMuertoReal, sTiempo, sTransporte, sTema, ' +
                                    'sInicioPlatica, sFinalPlatica, lStatus, sIdUsuario, '+
                                    'lAplicaComida, dAvProgAnteriorOrden, dAvProgActualOrden, dAvProgAnteriorContrato, dAvProgActualContrato, dAvRealAnteriorOrden, ' +
                                    'dAvRealActualOrden, dAvRealAnteriorContrato, dAvRealActualContrato, TipoAjuste, iFactorAjuste )' +
                                    'VALUES (:Contrato, :Orden, :Fecha, :Folio, :Turno, :Convenio, :Reporte, 0, "00:00", "00:00", "00:00", "00:00", "00:00", ' +
                                    '"00:00","", "", "00:00", "00:00", :Status, :Usuario,"No", 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 0, -1 )') ;
       connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
       connection.zcommand.Params.ParamByName('Contrato').value    := global_Contrato_Barco;
       connection.zcommand.Params.ParamByName('Orden').DataType    := ftString ;
       connection.zcommand.Params.ParamByName('Orden').value       := global_Contrato_Barco;
       connection.zcommand.Params.ParamByName('Fecha').DataType    := ftDate ;
       connection.zcommand.Params.ParamByName('Fecha').value       := tdFecha.Date ;
       connection.zcommand.Params.ParamByName('Folio').DataType    := ftString ;
       connection.zcommand.Params.ParamByName('Folio').value       := global_Contrato_Barco ;
       connection.zcommand.Params.ParamByName('Turno').DataType    := ftString ;
       connection.zcommand.Params.ParamByName('Turno').value       := global_turno ;
       connection.zcommand.Params.ParamByName('Convenio').DataType := ftString ;
       connection.zcommand.Params.ParamByName('Convenio').value    := global_convenio ;
       connection.zcommand.Params.ParamByName('Reporte').DataType  := ftString ;
       connection.zcommand.Params.ParamByName('Reporte').value     := 'S/N' ;
       connection.zcommand.Params.ParamByName('Status').DataType   := ftString ;
       connection.zcommand.Params.ParamByName('Status').value      := 'Pendiente' ;
       connection.zcommand.Params.ParamByName('Usuario').DataType  := ftString ;
       connection.zcommand.Params.ParamByName('Usuario').value     := global_usuario ;
       connection.zCommand.ExecSQL ;
       result := true;
    end;
//  except
//    result := false;
//  end;
end;

procedure TfrmAdmonyTiempos.frmBarra4btnPostClick(Sender: TObject);
var
  vdExistenciaActual: Double;
  sConsulta, sCondicion: string;
  nombres, cadenas: TStringList;
  lEncuentra : Boolean;
  zAgua:tzquery;
begin
     {Validacion de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('Embarcacion'); nombres.Add('Existencia'); cadenas.Add(tdEmbarcacionExist.Text); cadenas.Add(iIdRecursoExistencia.Text);
  if not validaTexto(nombres, cadenas, '', '') then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;

  if existeReporte then

  {Continua insercion de datos}
     //manejo de tab carmen
  tabexistencias.PageControl.Pages[0].TabVisible := true;
  tabexistencias.PageControl.Pages[1].TabVisible := true;
  tabexistencias.PageControl.Pages[2].TabVisible := true;

  if QryRecursos.FieldValues['sIdEmbarcacion'] = '' then
  begin
    messageDLG('Seleccione una Embarcacion!', mtInformation, [mbOk], 0);
    exit;
  end;
  inciandoAgua := False;
  if sOpcionEmb = 'Insert' then
  begin
    //A buscar agua inicial si no hay registros de este tipo y embarcacion
    if lowercase(trim(iIdRecursoExistencia.Text)) = 'agua' then
    begin
      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Text := 'Select * from recursos where didfecha <= :Fecha and iIdRecursoExistencia = :Recurso and scontrato = :contrato and sIdEmbarcacion =:Embarcacion';
      connection.QryBusca2.ParamByName('Contrato').AsString :=  global_contrato_Barco;
      connection.QryBusca2.ParamByName('Embarcacion').AsString := tdEmbarcacionExist.KeyValue;
      connection.QryBusca2.ParamByName('Recurso').AsString := iIdRecursoExistencia.KeyValue;
      connection.QryBusca2.Open;

      if connection.QryBusca2.RecordCount = 0 then
      begin
        connection.QryBusca2.Active := False;
        connection.QryBusca2.SQL.Text := 'Select * from embarcaciones where  scontrato = :contrato and sTipo = "Principal" and sIdEmbarcacion =:Embarcacion and lIniciaAgua = "Si"';
        connection.QryBusca2.ParamByName('Contrato').AsString :=  global_contrato_Barco;
        connection.QryBusca2.ParamByName('Embarcacion').AsString := tdEmbarcacionExist.KeyValue;
        connection.QryBusca2.Open;

        if connection.QryBusca2.RecordCount > 0 then
        begin
          qryRecursos.FieldByName('dExistenciaAnterior').AsFloat := connection.QryBusca2.FieldByName('dCantidadInicialAgua').AsFloat;
          qryRecursos.FieldByName('dacumulado').AsFloat          := qryRecursos.FieldByName('dExistenciaAnterior').AsFloat +qryRecursos.FieldByName('dExistenciaActual').AsFloat;
          qryRecursos.FieldByName('dConsumo').AsFloat            := dConsumo.Value;
          qryRecursos.FieldByName('dProduccion').AsFloat         := dProduccion.Value;
          qryRecursos.FieldByName('dRecibido').AsFloat           := dRecibido.Value;

          if qryRecursos.FieldByName('dExistenciaActual').AsFloat = 0 then
             qryRecursos.FieldByName('dExistenciaActual').AsFloat := (qryRecursos.FieldByName('dExistenciaAnterior').AsFloat - dConsumo.Value) + dProduccion.Value + dRecibido.Value ;
          qryRecursos.FieldByName('sidturno').asstring            := global_turno;
          qryRecursos.FieldByName('iIdRecursoExistencia').asstring:= iIdRecursoExistencia.KeyValue;

          connection.zCommand.Active := False;
          connection.zCommand.sql.Text := 'update embarcaciones set lIniciaAgua = "No" where scontrato = :contrato and sIdEmbarcacion = :embarcacion;';
          connection.zCommand.ParamByName('contrato').AsString    := global_Contrato_Barco;
          connection.zCommand.ParamByName('embarcacion').AsString := tdEmbarcacionExist.KeyValue;
          connection.zCommand.ExecSQL;
          inciandoAgua := True;
        end;
      end;
    end;


          //Verificamos que si exista el consumooo..
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select * from recursos where sContrato =:Contrato and dIdFecha =:Fecha ' +
      'and sIdTurno =:Turno and sIdEmbarcacion =:Embarcacion and iIdRecursoExistencia =:Recurso ');
    connection.zCommand.ParamByName('Contrato').AsString    := global_contrato_Barco;
    connection.zCommand.ParamByName('Fecha').AsDate         := tdFecha.Date;
    connection.zCommand.ParamByName('Turno').AsString       := global_turno;
    connection.zCommand.ParamByName('Embarcacion').AsString := tdEmbarcacionExist.KeyValue;
    connection.zCommand.ParamByName('Recurso').AsInteger    := iIdRecursoExistencia.KeyValue;
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
    begin
      messageDLG('La existencia de ' + iIdRecursoExistencia.Text + ' para ' + tdEmbarcacionExist.Text + ' ya existe. Favor de verificar!', mtInformation, [mbOk], 0);
      frmbarra4.btnCancel.Click;
      exit;
    end
    else
      qryRecursos.Post;
  end
  else
    qryRecursos.Post;

  if (qryMezclas.FieldValues['lCombustible'] = 'Si') and (qryRecursos.FieldValues['dRecibido'] > 0) then
  begin
      if zqOrdenaOrden.RecordCount > 0 then
      begin
          lEncuentra := False;
          zqOrdenaOrden.First;
          while not zqOrdenaOrden.Eof do
          begin
              if zqOrdenaOrden.FieldValues['lAplicaRecibidoDiesel'] = 'Si' then
                 lEncuentra := True;
             zqOrdenaOrden.Next;
          end;

          if lEncuentra = False then
          begin
              zqOrdenaOrden.First;
              zqOrdenaOrden.Edit;
              zqOrdenaOrden.FieldValues['lAplicaRecibidoDiesel'] := 'Si';
              zqOrdenaOrden.Post;
          end;
      end;
  end;

  if (qryMezclas.FieldValues['lCombustible'] = 'No') and (qryRecursos.FieldValues['dRecibido'] > 0) then
  begin
      if zqOrdenaOrden.RecordCount > 0 then
      begin
          lEncuentra := False;
          zqOrdenaOrden.First;
          while not zqOrdenaOrden.Eof do
          begin
              if zqOrdenaOrden.FieldValues['lAplicaRecibidoAgua'] = 'Si' then
                 lEncuentra := True;
             zqOrdenaOrden.Next;
          end;

          if lEncuentra = False then
          begin
              zqOrdenaOrden.First;
              zqOrdenaOrden.Edit;
              zqOrdenaOrden.FieldValues['lAplicaRecibidoAgua'] := 'Si';
              zqOrdenaOrden.Post;
          end;
      end;
  end;

  //Sumamos lo recibido de las embarcaciones de apoyo
  if chkDescuento.Checked then
  begin //si esta marcado como descontar
      sConsulta :=  'select sum(r.dRecibido) as Recibido from recursos r ' +
      'inner join embarcaciones e on (r.sIdEmbarcacion = e.sIdEmbarcacion and e.sTipo = "Secundario" )' +
      'where r.sContrato =:Contrato and r.dIdFecha =:Fecha and r.sIdTurno =:Turno and r.iIdRecursoExistencia =:Recurso';

      //Asumir que este registro no se contara en el descuento
      sCondicion := ' and r.sIdEmbarcacion <> :embarcacion';

      //El registro se contara en el descuento
      sCondicion := '';

      //añadir el group by
      sCondicion := sCondicion + ' group by r.sContrato ';
      //agregar el final de la consulta
      sConsulta := sConsulta + sCondicion;

      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add(sConsulta);
      connection.QryBusca.ParamByName('Contrato').AsString := global_contrato_Barco;
      connection.QryBusca.ParamByName('Fecha').AsDate      := tdFecha.Date;
      connection.QryBusca.ParamByName('Turno').AsString    := global_turno;
      connection.QryBusca.ParamByName('Recurso').AsInteger := iIdRecursoExistencia.KeyValue;

     //añadir el parametro si el registro no aplica al descuento
     if not chkDescuento.Checked then
        connection.QryBusca.ParamByName('embarcacion').AsString := tdEmbarcacionExist.KeyValue;
     connection.QryBusca.Open;

     if connection.QryBusca.RecordCount > 0 then
     begin
        //Buscamos embarcacion principal..
        connection.QryBusca2.Active := False;
        connection.QryBusca2.SQL.Clear;
        connection.QryBusca2.SQL.Add('select sIdEmbarcacion from embarcaciones where sTipo = "Principal" ');
        connection.QryBusca2.Open;

        if connection.QryBusca2.RecordCount > 0 then
        begin
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select * from recursos where sContrato =:Contrato and dIdFecha =:Fecha ' +
              'and sIdTurno =:Turno and sIdEmbarcacion =:Embarcacion and iIdRecursoExistencia =:Recurso ');
            connection.zCommand.ParamByName('Contrato').AsString    := global_contrato_Barco;
            connection.zCommand.ParamByName('Fecha').AsDate         := tdFecha.Date;
            connection.zCommand.ParamByName('Turno').AsString       := global_turno;
            connection.zCommand.ParamByName('Embarcacion').AsString := connection.QryBusca2.FieldValues['sIdEmbarcacion'];
            connection.zCommand.ParamByName('Recurso').AsInteger    := iIdRecursoExistencia.KeyValue;
            connection.zCommand.Open;

            if connection.zCommand.RecordCount > 0 then
            begin
                vdExistenciaActual := connection.zCommand.FieldValues['dProduccion'] + connection.zCommand.FieldValues['dExistenciaAnterior'] + connection.zCommand.FieldValues['dRecibido'];
                vdExistenciaActual := (vdExistenciaActual) - (connection.zCommand.FieldValues['dConsumo'] + connection.zCommand.FieldValues['dConsumoEquipos'] + connection.QryBusca.FieldValues['Recibido']);

                          //Actualizamos la esxitencia embarcacion principal
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('update recursos set dPrestamos =:Prestamo, dExistenciaActual =:Existencia where sContrato =:Contrato and dIdFecha =:Fecha ' +
                  'and sIdTurno =:Turno and sIdEmbarcacion =:Embarcacion and iIdRecursoExistencia =:Recurso ');
                connection.zCommand.ParamByName('Contrato').AsString    := global_contrato_Barco;
                connection.zCommand.ParamByName('Fecha').AsDate         := tdFecha.Date;
                connection.zCommand.ParamByName('Turno').AsString       := global_turno;
                connection.zCommand.ParamByName('Embarcacion').AsString := connection.QryBusca2.FieldValues['sIdEmbarcacion'];
                connection.zCommand.ParamByName('Prestamo').AsFloat     := connection.QryBusca.FieldValues['Recibido'];
                connection.zCommand.ParamByName('Existencia').AsFloat   := vdExistenciaActual;
                connection.zCommand.ParamByName('Recurso').AsInteger    := iIdRecursoExistencia.KeyValue;
                connection.zCommand.ExecSQL;
            end;
         end;
      end;
  end;

  if sOpcionEmb = 'Edit' then
  begin
      //Actualizamos la existencia anterior del dia siguiente si existe..
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('update recursos set dExistenciaAnterior =:Anterior where sContrato =:Contrato and dIdFecha =:Fecha ' +
        'and sIdTurno =:Turno and sIdEmbarcacion =:Embarcacion and iIdRecursoExistencia =:Recurso ');
      connection.zCommand.ParamByName('Contrato').AsString    := global_contrato_Barco;
      connection.zCommand.ParamByName('Fecha').AsDate         := incday(tdFecha.Date);
      connection.zCommand.ParamByName('Turno').AsString       := global_turno;
      connection.zCommand.ParamByName('Embarcacion').AsString := tdEmbarcacionExist.KeyValue;
      connection.zCommand.ParamByName('Anterior').AsFloat     := dExistenciaActual.Value;
      connection.zCommand.ParamByName('Recurso').AsInteger    := iIdRecursoExistencia.KeyValue;
      connection.zCommand.ExecSQL;
  end;

  label34.Caption          := '';
  tdEmbarcacionExist.SetFocus;
  frmBarra4.btnCancelClick(Sender);
  PanelExistencias.Enabled := False;
  chkDescuento.Visible     := False;
  dbExistencias.Enabled    := True;

  if sOpcionEmb = 'Insert' then
  begin
    qryRecursos.Refresh;
    frmBarra4.btnAdd.Click
  end
  else
    sOpcionEmb := '';
  BotonPermiso.permisosBotones(frmBarra4);

end;

procedure TfrmAdmonyTiempos.frmBarra4btnPrinterClick(Sender: TObject);
begin
  Application.CreateForm(TfrmOpcionesBarco, frmOpcionesBarco);
  frmOpcionesBarco.showModal;
end;

procedure TfrmAdmonyTiempos.frmBarra4btnCancelClick(Sender: TObject);
begin
//manejo de tab carmen
  tabexistencias.PageControl.Pages[0].TabVisible := true;
  tabexistencias.PageControl.Pages[1].TabVisible := true;
  tabexistencias.PageControl.Pages[2].TabVisible := true;
  frmBarra4.btnCancelClick(Sender);
  try
    qryRecursos.Cancel;
    iIdRecursoExistencia.Enabled := False;
    dProduccion.Enabled := False;
    dAjuste.Enabled := False; //*****************BRITO 17/12/10************
    dRecibido.Enabled := False;
    dConsumo.Enabled := False;
    dConsumoEquipos.Enabled := False;
    dPrestamos.Enabled := False;
    dTrasiego.Enabled := False;
    dExistenciaActual.Enabled := False;
    dExistenciaAnterior.Enabled := False;
    label34.Caption := '';
    chkDescuento.Visible := False;
    Insertar1.Enabled := True;
    Editar1.Enabled := True;
    Registrar1.Enabled := False;
    Can1.Enabled := False;
    Eliminar1.Enabled := True;
    Refresh1.Enabled := True;
    Salir1.Enabled := True;
    mkHora1.ReadOnly := False;
    label25.Visible := True;
    dConsumoEquipos.Visible := True;
  except
  end;
  //desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra4);
  dbExistencias.Enabled := True;
  PanelExistencias.Enabled := False;
end;

procedure TfrmAdmonyTiempos.frmBarra4btnDeleteClick(Sender: TObject);
var
  IdRecurso: integer;
  vdExistenciaActual: Double;
begin

  if ReporteLock then
  begin
    messageDLg('El Reporte Dairio se encuentra Validado/Autorizado. Favor de verificar!', mtInformation, [mbOk], 0);
    exit;
  end;

  IdRecurso := iIdRecursoExistencia.KeyValue;
  Entre := True;
  if qryRecursos.RecordCount > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        qryRecursos.Delete;
         //Sumamos lo recibido de las embarcaciones de apoyoo
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('select sum(r.dRecibido) as Recibido from recursos r ' +
          'inner join embarcaciones e on (r.sIdEmbarcacion = e.sIdEmbarcacion and e.sTipo = "Secundario" )' +
          'where r.sContrato =:Contrato and r.dIdFecha =:Fecha and r.sIdTurno =:Turno and r.iIdRecursoExistencia =:Recurso group by r.sContrato ');
        connection.QryBusca.ParamByName('Contrato').AsString := global_contrato_Barco;
        connection.QryBusca.ParamByName('Fecha').AsDate      := tdFecha.Date;
        connection.QryBusca.ParamByName('Turno').AsString    := global_turno;
        connection.QryBusca.ParamByName('Recurso').AsInteger := IdRecurso;
        connection.QryBusca.Open;

        if connection.QryBusca.RecordCount > 0 then
        begin
             //Buscamos embarcacion principal..
          connection.QryBusca2.Active := False;
          connection.QryBusca2.SQL.Clear;
          connection.QryBusca2.SQL.Add('select sIdEmbarcacion from embarcaciones where sTipo = "Principal" ');
          connection.QryBusca2.Open;

          if connection.QryBusca2.RecordCount > 0 then
          begin
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('select * from recursos where sContrato =:Contrato and dIdFecha =:Fecha ' +
                'and sIdTurno =:Turno and sIdEmbarcacion =:Embarcacion and iIdRecursoExistencia =:Recurso ');
              connection.zCommand.ParamByName('Contrato').AsString    := global_contrato_Barco;
              connection.zCommand.ParamByName('Fecha').AsDate         := tdFecha.Date;
              connection.zCommand.ParamByName('Turno').AsString       := global_turno;
              connection.zCommand.ParamByName('Embarcacion').AsString := connection.QryBusca2.FieldValues['sIdEmbarcacion'];
              connection.zCommand.ParamByName('Recurso').AsInteger    := IdRecurso;
              connection.zCommand.Open;

              vdExistenciaActual := connection.zCommand.FieldValues['dProduccion'] + connection.zCommand.FieldValues['dExistenciaAnterior'] + connection.zCommand.FieldValues['dRecibido'];
              vdExistenciaActual := (vdExistenciaActual) - (connection.zCommand.FieldValues['dConsumo'] + connection.zCommand.FieldValues['dConsumoEquipos'] + connection.QryBusca.FieldValues['Recibido']);

                  //Actualizamos la embarcacion principal
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('update recursos set dPrestamos =:Prestamo, dExistenciaActual =:Existencia where sContrato =:Contrato and dIdFecha =:Fecha ' +
                'and sIdTurno =:Turno and sIdEmbarcacion =:Embarcacion and iIdRecursoExistencia =:Recurso ');
              connection.zCommand.ParamByName('Contrato').AsString    := global_contrato_Barco;
              connection.zCommand.ParamByName('Fecha').AsDate         := tdFecha.Date;
              connection.zCommand.ParamByName('Turno').AsString       := global_turno;
              connection.zCommand.ParamByName('Embarcacion').AsString := connection.QryBusca2.FieldValues['sIdEmbarcacion'];
              connection.zCommand.ParamByName('Prestamo').AsFloat     := connection.QryBusca.FieldValues['Recibido'];
              connection.zCommand.ParamByName('Existencia').AsFloat   := vdExistenciaActual;
              connection.zCommand.ParamByName('Recurso').AsInteger    := IdRecurso;
              connection.zCommand.ExecSQL;
          end;
        end;

      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Al elimina registro en existencias', 0);
        end;
      end
    end

end;

procedure TfrmAdmonyTiempos.frmBarra4btnRefreshClick(Sender: TObject);
begin
  qryRecursos.Refresh;
  PanelExistencias.Enabled := False;
end;

procedure TfrmAdmonyTiempos.frmBarra4btnExitClick(Sender: TObject);
begin
  frmBarra4.btnExitClick(Sender);
  close;
end;

procedure TfrmAdmonyTiempos.qryRecursosBeforePost(DataSet: TDataSet);
var
  vdAcumuladoMezclas: Double;
  vdExistenciaAnterior: Double;
  vdExistenciaActual: Double;
  viIdRecursoExistencia: Integer;
  dConsumov, dConsumoEquiposv, dPrestamov,
    dProduccionv, dRecibidov, InicialAgua: Double;
  QryIniciaDiesel: TZReadOnlyQuery;
  FechaUltima: TDateTime;
begin
  if not inciandoAgua then
  begin
    QryIniciaDiesel := TZReadOnlyQuery.Create(self);
    QryIniciaDiesel.Connection := connection.zConnection;

    qryGetAcumulado := TZReadOnlyQuery.Create(self);
    qryGetAcumulado.Connection := connection.zConnection;
    qryGetAcumulado.Active := False;
    viIdRecursoExistencia := iIdRecursoExistencia.KeyValue;

    qryRecursos.FieldValues['sIdTurno'] := global_turno;
    qryRecursos.FieldValues['iIdRecursoExistencia'] := viIdRecursoExistencia;
    if dProduccion.Text = '' then
      qryRecursos.FieldValues['dProduccion'] := 0
    else
      qryRecursos.FieldValues['dProduccion'] := dProduccion.Value;

    //********************************BRITO 17/12/10******************************
    if dAjuste.Text = '' then
      qryRecursos.FieldValues['dAjuste'] := 0
    else
      qryRecursos.FieldValues['dAjuste'] := dAjuste.Value;
    //********************************BRITO 17/12/10******************************

    if dRecibido.Text = '' then
      qryRecursos.FieldValues['dRecibido'] := 0
    else
      qryRecursos.FieldValues['dRecibido'] := dRecibido.Value;

    if dConsumo.Text = '' then
      qryRecursos.FieldValues['dConsumo'] := 0
    else
      qryRecursos.FieldValues['dConsumo'] := dConsumo.Value;

    if dConsumoEquipos.Text = '' then
      qryRecursos.FieldValues['dConsumoEquipos'] := 0
    else
      qryRecursos.FieldValues['dConsumoEquipos'] := dConsumoEquipos.Value;

    if dPrestamos.Text = '' then
      qryRecursos.FieldValues['dPrestamos'] := 0
    else
      qryRecursos.FieldValues['dPrestamos'] := dPrestamos.Value;

    if dGalones.Text = '' then
      qryRecursos.FieldValues['dGalones'] := 0
    else
      qryRecursos.FieldValues['dGalones'] := dGalones.Value;

    //Obtener la existencia anterior
    vdExistenciaAnterior := 0;
    qryGetAcumulado.Active := False;
    qryGetAcumulado.SQL.Clear;
    qryGetAcumulado.SQL.Add('select dExistenciaActual as dExistenciaAnterior from recursos where ' +
      'sContrato=:sContrato and dIdFecha<:dIdFecha and ' +
      'sIdTurno=:sIdTurno and iIdRecursoExistencia=:iIdRecursoExistencia and sIdEmbarcacion =:Embarcacion ' +
      'Order By dIdFecha desc');
    qryGetAcumulado.Params.ParamByName('sContrato').DataType   := ftString;
    qryGetAcumulado.Params.ParamByName('sContrato').Value      := global_contrato_Barco;
    qryGetAcumulado.Params.ParamByName('dIdFecha').DataType    := ftDate;
    qryGetAcumulado.Params.ParamByName('dIdFecha').Value       := tdFecha.Date;
    qryGetAcumulado.Params.ParamByName('sIdTurno').DataType    := ftString;
    qryGetAcumulado.Params.ParamByName('sIdTurno').Value       := global_turno;
    qryGetAcumulado.Params.ParamByName('iIdRecursoExistencia').DataType := ftInteger;
    qryGetAcumulado.Params.ParamByName('iIdRecursoExistencia').Value    := IntToStr(viIdRecursoExistencia);
    qryGetAcumulado.Params.ParamByName('Embarcacion').DataType := ftString;
    qryGetAcumulado.Params.ParamByName('Embarcacion').Value    := tdEmbarcacionExist.KeyValue;

    qryGetAcumulado.Open;

    try
      if qryGetAcumulado.RecordCount > 0 then
        if qryGetAcumulado.FieldValues['dExistenciaAnterior'] <> 0 then
          vdExistenciaAnterior := qryGetAcumulado.FieldValues['dExistenciaAnterior']
        else
          vdExistenciaAnterior := 0
      else
        vdExistenciaAnterior := 0;
    except
      vdExistenciaAnterior := 0;
    end;

    if dAjuste.Text <> '' then
      vdExistenciaAnterior := vdExistenciaAnterior + strtofloat(dAjuste.Text);

    //Obtener el acumulado anterior
    vdAcumuladoMezclas := 0;
    qryGetAcumulado.Active := False;
    qryGetAcumulado.SQL.Clear;
    qryGetAcumulado.SQL.Add(' select sum(dExistenciaActual) as dAcumulado from recursos where ' +
      ' sContrato=:sContrato and dIdFecha<:dIdFecha and' +
      ' sIdTurno=:sIdTurno and iIdRecursoExistencia=:iIdRecursoExistencia and sIdEmbarcacion =:Embarcacion ');
    qryGetAcumulado.Params.ParamByName('sContrato').DataType     := ftString;
    qryGetAcumulado.Params.ParamByName('sContrato').Value        := global_contrato_Barco;
    qryGetAcumulado.Params.ParamByName('dIdFecha').DataType      := ftDate;
    qryGetAcumulado.Params.ParamByName('dIdFecha').Value         := tdFecha.Date;
    qryGetAcumulado.Params.ParamByName('sIdTurno').DataType      := ftString;
    qryGetAcumulado.Params.ParamByName('sIdTurno').Value         := global_turno;
    qryGetAcumulado.Params.ParamByName('iIdRecursoExistencia').DataType := ftInteger;
    qryGetAcumulado.Params.ParamByName('iIdRecursoExistencia').Value := IntToStr(viIdRecursoExistencia);
    qryGetAcumulado.Params.ParamByName('Embarcacion').DataType   := ftString;
    qryGetAcumulado.Params.ParamByName('Embarcacion').Value      := tdEmbarcacionExist.KeyValue;
    qryGetAcumulado.Open;

    try
      if qryGetAcumulado.RecordCount > 0 then
        if qryGetAcumulado.FieldValues['dAcumulado'] <> 0 then
          vdAcumuladoMezclas := qryGetAcumulado.FieldValues['dAcumulado']
        else
          vdAcumuladoMezclas := 0
      else
        vdAcumuladoMezclas := 0;
    except
      vdAcumuladoMezclas := 0;
    end;

    //Obtener la existencia actual verdadera
    if dConsumo.Text = '' then
      dConsumov := 0
    else
      dConsumov := dConsumo.value;

    if dConsumoEquipos.Text = '' then
      dConsumoEquiposv := 0
    else
      dConsumoEquiposV := dConsumoEquipos.value;

    if dPrestamos.Text = '' then
      dPrestamov := 0
    else
      dPrestamov := dPrestamos.value;

    if dProduccion.Text = '' then
      dProduccionv := 0
    else
      dProduccionv := dProduccion.value;
    if dRecibido.Text = '' then
      dRecibidov := 0
    else
      dRecibidov := dRecibido.value;

    //Consultamos si la embarcacion tiene diesel inicial.
    QryIniciaDiesel.Active := False;
    QryIniciaDiesel.SQL.clear;
    QryIniciaDiesel.SQL.Add('Select dCantidadInicial from embarcaciones where sIdEmbarcacion =:Embarcacion and sIniciaDiesel ="Si" ');
    QryIniciaDiesel.ParamByName('Embarcacion').AsString := tdEmbarcacionExist.KeyValue;
    QryIniciaDiesel.Open;

    if QryIniciaDiesel.RecordCount > 0 then
    begin
      vdExistenciaAnterior := QryIniciaDiesel.FieldValues['dCantidadInicial'];

        //Qutiamos el indicador de inicio de Disel..
      QryIniciaDiesel.Active := False;
      QryIniciaDiesel.SQL.clear;
      QryIniciaDiesel.SQL.Add('Update embarcaciones set sIniciaDiesel = "No" where sIdEmbarcacion =:Embarcacion ');
      QryIniciaDiesel.ParamByName('Embarcacion').AsString := tdEmbarcacionExist.KeyValue;
      QryIniciaDiesel.ExecSQL;

    end;

    vdExistenciaActual := dProduccionv + vdExistenciaAnterior + dRecibidov;
    vdExistenciaActual := (vdExistenciaActual) - (dConsumov + dConsumoEquiposv + dPrestamos.value);
    //Obtener el acumulado actual
    vdAcumuladoMezclas := vdAcumuladoMezclas + vdExistenciaActual;

    if vdExistenciaActual < 0 then
      vdExistenciaActual := 0;

    if vdExistenciaAnterior < 0 then
      vdExistenciaAnterior := 0;

    if vdAcumuladoMezclas < 0 then
      vdAcumuladoMezclas := 0;

    qryRecursos.FieldValues['dExistenciaActual'] := vdExistenciaActual;
    qryRecursos.FieldValues['dExistenciaAnterior'] := vdExistenciaAnterior;
    qryRecursos.FieldValues['dAcumulado'] := vdAcumuladoMezclas;

    qryGetAcumulado.Destroy;
  end;

end;

procedure TfrmAdmonyTiempos.qryRecursosAfterPost(DataSet: TDataSet);
begin
  try
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('SELECT iIdRecursoExistencia FROM recursosdeexistencias ');
    connection.QryBusca.Open;
    if Entre = true then
    begin
      while connection.QryBusca.Eof = false do
      begin
        procActualizaExistencias(global_contrato_Barco, DateToStr(tdFecha.Date), global_turno, connection.QryBusca.FieldValues['iIdRecursoExistencia'], self);
        connection.QryBusca.Next;
      end;
    end;
    qryRecursos.Refresh;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Despues de salvar registro', 0);
    end;
  end;
end;

procedure TfrmAdmonyTiempos.qryRecursosAfterScroll(DataSet: TDataSet);
begin
  if qryRecursos.RecordCount > 0 then
    if qryRecursos.Eof = False then
    begin
      iIdRecursoExistencia.KeyValue := qryRecursos.FieldValues['iIdRecursoExistencia'];
      dProduccion.Text := FloatToStr(qryRecursos.FieldValues['dProduccion']);
      dAjuste.Text := FloatToStr(qryRecursos.FieldValues['dAjuste']); //****************BRITO 17/12/10*************
      dRecibido.Text := FloatToStr(qryRecursos.FieldValues['dRecibido']);
      dConsumo.Text := FloatToStr(qryRecursos.FieldValues['dConsumo']);
      dConsumoEquipos.Text := FloatToStr(qryRecursos.FieldValues['dConsumoEquipos']);
      dPrestamos.Text := FloatToStr(qryRecursos.FieldValues['dPrestamos']);
      dGalones.Text   := FloatToStr(qryRecursos.FieldValues['dGalones']);
      dExistenciaActual.Text := FloatToStr(qryRecursos.FieldValues['dExistenciaActual']);
      dExistenciaAnterior.Text := FloatToStr(qryRecursos.FieldValues['dExistenciaAnterior']);
      //chkDescuento.Visible := not (Embarcaciones2.FieldByName('sTipo').AsString = 'Principal'); //********************BRITO 22/11/10****************************************
      label34.Caption := qryRecursos.FieldValues['CalcMezclasMedidas'];
    end
end;

procedure TfrmAdmonyTiempos.qryRecursosAfterDelete(DataSet: TDataSet);
begin
  try
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('SELECT iIdRecursoExistencia FROM recursosdeexistencias ');
    connection.QryBusca.Open;
    if Entre = False then
    begin
      while connection.QryBusca.Eof = false do
      begin
        procActualizaExistencias(global_contrato_Barco, DateToStr(tdFecha.Date), global_turno, connection.QryBusca.FieldValues['iIdRecursoExistencia'], self);
        connection.QryBusca.Next;
      end;
    end;
    qryRecursos.Refresh;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administracion y Tiempos', 'Despues de eliminar registro', 0);
    end;
  end;
end;

procedure TfrmAdmonyTiempos.dPrestamosKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(dPrestamos, key) then
    key := #0;
  if key = #13 then
    if dTrasiego.Visible then
      dTrasiego.SetFocus
    else
      dAjuste.SetFocus;
end;

procedure TfrmAdmonyTiempos.dProduccionEnter(Sender: TObject);
begin
  dProduccion.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dProduccionExit(Sender: TObject);
begin
  dProduccion.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dRecibidoEnter(Sender: TObject);
begin
  dRecibido.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dRecibidoExit(Sender: TObject);
begin
  dRecibido.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dConsumoEnter(Sender: TObject);
begin
  dConsumo.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dConsumoExit(Sender: TObject);
begin
     dConsumo.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dConsumoEquiposEnter(Sender: TObject);
begin
  dConsumoEquipos.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dConsumoEquiposExit(Sender: TObject);
begin
    dConsumoEquipos.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dTrasiegoEnter(Sender: TObject);
begin
  dPrestamos.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dTrasiegoExit(Sender: TObject);
begin
  dPrestamos.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dExistenciaActualChange(Sender: TObject);
begin
  TRxCalcEditChangef(dExistenciaActual, 'Existencia Actual');
end;

procedure TfrmAdmonyTiempos.dExistenciaActualEnter(Sender: TObject);
begin
  dExistenciaActual.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dExistenciaActualExit(Sender: TObject);
begin
  dExistenciaActual.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dExistenciaAnteriorExit(Sender: TObject);
begin
  dExistenciaAnterior.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dExistenciaAnteriorChange(Sender: TObject);
begin
  TRxCalcEditChangef(dExistenciaAnterior, 'Existencia Anterior');
end;

procedure TfrmAdmonyTiempos.dExistenciaAnteriorEnter(Sender: TObject);
begin
  dExistenciaAnterior.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.iIdRecursoExistenciaEnter(Sender: TObject);
begin
  iIdRecursoExistencia.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.iIdRecursoExistenciaExit(Sender: TObject);
begin
     iIdRecursoExistencia.Color := global_color_salida;
     label34.Caption := qryMezclas.FieldValues['sMedida'];
     label25.Visible := True;
     dConsumoEquipos.Visible := True;
     if iIdRecursoExistencia.Text = 'AGUA' then
     begin
         label25.Visible := False;
         dConsumoEquipos.Visible := False;
     end;
     ExistenciaAnterior; //<-Procedimiento
end;

procedure TfrmAdmonyTiempos.dIdFechaExistenciaExit(Sender: TObject);
begin
  iIdRecursoExistencia.Color := global_color_salida;
end;

procedure TfrmAdmonyTiempos.dIdFechaExistenciaKeyPress(Sender: TObject;
  var Key: Char);
begin
 { if key = #13 then
    iIdRecursoExistencia.SetFocus;   }
end;

procedure TfrmAdmonyTiempos.iIdRecursoExistenciaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    dConsumo.SetFocus;
end;

procedure TfrmAdmonyTiempos.Imprimir1Click(Sender: TObject);
begin
  if tabexistencias.PageControl.Pages[0].TabVisible = true then
    frmBarra1.btnPrinter.Click;
  if tabexistencias.PageControl.Pages[1].TabVisible = true then
    frmBarra3.btnPrinter.Click;
  if tabexistencias.PageControl.Pages[2].TabVisible = true then
    frmBarra2.btnPrinter.Click;
  if tabexistencias.PageControl.Pages[3].TabVisible = true then
    frmBarra4.btnPrinter.Click;
end;

procedure TfrmAdmonyTiempos.ImprimirTiempoCia1Click(Sender: TObject);
begin
  Panel1.Visible := true;
end;

procedure TfrmAdmonyTiempos.Insertar1Click(Sender: TObject);
begin
  if dbmovbarco.Focused then
    frmBarra1.btnAdd.Click;
  if dbcondiciones.Focused then
    frmBarra3.btnAdd.Click;
  if dbmovtosembarcacion.Focused then
    frmBarra2.btnAdd.Click;
  if dbexistencias.Focused then
    frmBarra4.btnAdd.Click;
end;

procedure TfrmAdmonyTiempos.dProduccionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(dProduccion, key) then
    key := #0;
  if key = #13 then
    dRecibido.SetFocus;
end;

procedure TfrmAdmonyTiempos.dRecibidoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(dRecibido, key) then
    key := #0;
  if key = #13 then
    dConsumoEquipos.SetFocus;
end;

procedure TfrmAdmonyTiempos.dConsumoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(dConsumo, key) then
    key := #0;
  if key = #13 then
    dProduccion.SetFocus;
end;

procedure TfrmAdmonyTiempos.dConsumoEquiposKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(dConsumoEquipos, key) then
    key := #0;

  if key = #13 then
    dPrestamos.SetFocus;
end;

procedure TfrmAdmonyTiempos.dTrasiegoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(dTrasiego, key) then
    key := #0;
  if key = #13 then
    dAjuste.SetFocus;
end;

procedure TfrmAdmonyTiempos.dExistenciaActualKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(dExistenciaActual, key) then
    key := #0;
  if key = #13 then
    dExistenciaAnterior.SetFocus;
end;

procedure TfrmAdmonyTiempos.dExistenciaAnteriorKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(dExistenciaAnterior, key) then
    key := #0;
  if key = #13 then
    frmBarra4.btnPost.SetFocus;
end;

procedure TfrmAdmonyTiempos.dGalonesEnter(Sender: TObject);
begin
    dGalones.Color := global_color_entrada;
end;

procedure TfrmAdmonyTiempos.dGalonesExit(Sender: TObject);
begin
     dGalones.Color := global_color_salida;
     //consultamos la equivalencia de Galones a M3
     connection.QryBusca.Active := False;
     connection.QryBusca.SQL.Clear;
     connection.QryBusca.SQL.Add('select dGalones from configuracion where sContrato =:contrato ');
     connection.QryBusca.ParamByName('contrato').AsString := global_contrato_Barco;
     connection.QryBusca.Open;

     dConsumo.Value :=  dGalones.Value * connection.QryBusca.FieldByName('dGalones').AsFloat;
end;

procedure TfrmAdmonyTiempos.dGalonesKeyPress(Sender: TObject; var Key: Char);
begin
    if key =#13 then
       dProduccion.SetFocus;
end;

procedure TfrmAdmonyTiempos.MovimientosdeBarcoAfterScroll(DataSet: TDataSet);
begin
  if sOpcionEmb = '' then
  begin
    if movimientosdebarco.RecordCount > 0 then
    begin
      mkHora1.Text := movimientosdebarco.FieldValues['sHoraInicio'];
      mkHora2.Text := movimientosdebarco.FieldValues['sHoraFinal'];
      sSuma.Text := movimientosdebarco.FieldValues['sFactor'];
      tmDescripcion2.Text := movimientosdebarco.FieldValues['mDescripcion'];

      if movimientosdebarco.FieldValues['lContinuo'] = 'Si' then
         chkContinuaMov.Checked := True
      else
         chkContinuaMov.Checked := False;

      if movimientosdebarco.FieldValues['sActividades'] = 'MOV' then
         chkAplicaFactor.Checked := False
      else
         chkAplicaFactor.Checked := True;
      CbPartidas.Text := MovimientosdeBarco.FieldByName('sActividades').AsString;
      InicializarCheckCombo(CmbFolios);
      CmbFoliosExit(CmbFolios);
    end

  end;
end;

procedure TfrmAdmonyTiempos.MovimientosdeBarcoCalcFields(DataSet: TDataSet);
begin
    MovimientosdeBarcosDescripcion.Text := MovimientosdeBarco.FieldValues['mDescripcion'];
end;

procedure TfrmAdmonyTiempos.movimientosdeembarcacionAfterScroll(
  DataSet: TDataSet);
begin
  if sOpcionEmb = '' then
    if movimientosdeembarcacion.RecordCount > 0 then
    begin
      tsHoraInicio.Text := movimientosdeembarcacion.FieldValues['sHoraInicio'];
      tsHoraFinal.Text := movimientosdeembarcacion.FieldValues['sHoraFinal'];
      tmDescripcion.Text := movimientosdeembarcacion.FieldValues['mDescripcion'];
      tsMovimiento.KeyValue := movimientosdeembarcacion.FieldValues['sNumeroActividad'];
      label40.Visible := False;
      tsMovimiento.Visible := False;
      if movimientosdeembarcacion.FieldValues['sTipo'] = 'ARRIBO' then
        rbArribo.Checked := True
      else
        if movimientosdeembarcacion.FieldValues['sTipo'] = 'DISPOSICION' then
          rbDisposicion.Checked := True
        else
           if movimientosdeembarcacion.FieldValues['sTipo'] = 'MOVIMIENTO' then
           begin
              rbMovimiento.Checked := True;
              label40.Visible := True;
              tsMovimiento.Visible := True;
           end;

      if movimientosdeembarcacion.FieldValues['lContinuo'] = 'Si' then
         chkContinuaArribo.Checked := True
      else
         chkContinuaArribo.Checked := False;
    end;
end;

procedure TfrmAdmonyTiempos.movimientosdeembarcacionCalcFields(
  DataSet: TDataSet);
begin
   movimientosdeembarcacionsDescripcion.Text := MovimientosdeEmbarcacion.FieldValues['mDescripcion'];
end;

procedure TfrmAdmonyTiempos.pgAdmonChange(Sender: TObject);
begin
    lblBusca.Text := '';
    label34.Caption := '';
end;

procedure TfrmAdmonyTiempos.tsClasificacionesEnter(Sender: TObject);
begin
  tsClasificaciones.Color := Global_Color_Entrada;
end;

procedure TfrmAdmonyTiempos.tsClasificacionesExit(Sender: TObject);
begin
  tsClasificaciones.Color := Global_Color_Salida;
end;

procedure TfrmAdmonyTiempos.tsClasificacionesKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    mkHora1.SetFocus;
end;

function TfrmAdmonyTiempos.ReporteLock;
begin
  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('select * from reportediario where sContrato =:Contrato and sIdConvenio =:Convenio and dIdFecha =:Fecha and sIdTurno =:Turno and lStatus <> "Pendiente"');
  connection.zCommand.ParamByName('Contrato').AsString := global_contrato_Barco;
  connection.zCommand.ParamByName('Fecha').AsDate      := tdFecha.Date;
  connection.zCommand.ParamByName('Turno').AsString    := global_turno;
  connection.zCommand.ParamByName('Convenio').AsString := global_convenio;
  connection.zCommand.Open;

  if connection.zCommand.RecordCount > 0 then
    ReporteLock := True
  else
    ReporteLock := False;
end;

procedure TfrmAdmonyTiempos.filtra;
var
   filtro, buscar : string;
begin
    filtro := '';
    if length(trim(lblBusca.Text)) > 0 then
    begin
      buscar := lblBusca.Text;
      buscar := '*'+buscar + '*';

      filtro := 'sDescripcion like ' + QuotedStr(buscar);
    end;

    if filtro <> '' then
    begin
        embarcaciones.Filtered := False;
        embarcaciones.Filter   := filtro;
        embarcaciones.Filtered := True;
    end
    else
    begin
        embarcaciones.Filtered := False;
    end;
end;

procedure TfrmAdmonyTiempos.filtraCondicion;
var
   filtro, buscar : string;
begin
    filtro := '';
    if length(trim(lblBusca.Text)) > 0 then
    begin
      buscar := lblBusca.Text;
      buscar := '*'+buscar + '*';

      filtro := 'sDescripcion like ' + QuotedStr(buscar);
    end;

    if filtro <> '' then
    begin
        Qrydirecciones.Filtered := False;
        Qrydirecciones.Filter   := filtro;
        Qrydirecciones.Filtered := True;
    end
    else
    begin
        Qrydirecciones.Filtered := False;
    end;
end;


procedure TfrmAdmonyTiempos.OrdenarOrdenes(sParamOrden: string);
var
   idAuxiliar, idAuxiliar2 : integer;
   SavePlace   : TBookmark;
begin
    if zqOrdenaOrden.RecordCount > 0 then
    begin
        if sParamOrden = 'Arriba' then
        begin
            idAuxiliar2 := zqOrdenaOrden.FieldValues['iIdOrden'];
            zqOrdenaOrden.Prior;

            idAuxiliar  := zqOrdenaOrden.FieldValues['iIdOrden'];
            zqOrdenaOrden.Next;
        end;

        if sParamOrden = 'Abajo' then
        begin
            idAuxiliar2 := zqOrdenaOrden.FieldValues['iIdOrden'];
            zqOrdenaOrden.Next;

            idAuxiliar  := zqOrdenaOrden.FieldValues['iIdOrden'];
            zqOrdenaOrden.Prior;
        end;
        //Colocamos un id mayor para evitar duplicidad..
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE recursosordenados_orden SET iIdOrden = :DiarioNuevo ' +
                                    'where dIdFecha = :fecha And iIdOrden = :diario ');
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := idAuxiliar2;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').DataType := ftInteger;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').value    := idAuxiliar + 500;
        connection.zCommand.ExecSQL;

        //Ahora actualizamos el item mayor
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE recursosordenados_orden SET iIdOrden = :DiarioNuevo ' +
                                    'where dIdFecha = :fecha And iIdOrden = :diario ');
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := idAuxiliar;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').DataType := ftInteger;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').value    := idAuxiliar2;
        connection.zCommand.ExecSQL;

         //Ahora actualizamos el item alterado
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE recursosordenados_orden SET iIdOrden = :DiarioNuevo ' +
                                    'where dIdFecha = :fecha And iIdOrden = :diario ');
        Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
        Connection.zCommand.Params.ParamByName('fecha').value       := tdFecha.Date;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := idAuxiliar + 500;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').DataType := ftInteger;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').value    := idAuxiliar;
        connection.zCommand.ExecSQL;

        if sParamOrden = 'Arriba' then
           zqOrdenaOrden.Prior
        else
           zqOrdenaOrden.Next;

        SavePlace := zqOrdenaOrden.GetBookmark;
        zqOrdenaOrden.Refresh;
        zqOrdenaOrden.GotoBookmark(SavePlace);
        zqOrdenaOrden.FreeBookmark(SavePlace);

    end;
end;

procedure TfrmAdmonyTiempos.Paste1Click(Sender: TObject);
begin
  if lCopiaObjeto then
  begin
     if tabexistencias.PageControl.Pages[0].TabVisible = true then
        tmDescripcion2.PasteFromClipboard;
     If tabexistencias.PageControl.Pages[1].TabVisible = true then
        sPronostico.PasteFromClipboard;
     If tabexistencias.PageControl.Pages[2].TabVisible = true then
        tmDescripcion.PasteFromClipboard;
  end
end;

procedure TfrmAdmonyTiempos.EstableceCheckPArtidas(Cadena:string;Combo:tjvcheckedcombobox);
var   LstPrt:TstringList;
  c:Integer;
begin
  LstPrt := Tstringlist.create;
  try

    LstPrt.commatext := Cadena;
    try
      if length(trim(Cadena)) = 0 then
        raise exception.create('*');

      for c := 0 to combo.items.Count - 1 do
      begin
        if LstPrt.indexof(combo.items[c]) >= 0 then
          combo.Checked[c]:=True;
      end;

    except
      on e:exception do
        if e.message = '*' then
          ;
    end;
  finally
    LstPrt.free;
  end;

end;
procedure TfrmAdmonyTiempos.Existenciasycosumos1Click(Sender: TObject);
var iin,ifi:Integer;
begin
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Text := 'select year(min(didfecha)) as minimo,year(max(didfecha)) as maximo from recursos where sContrato = :contratobarco';
  connection.QryBusca.ParamByName('contratobarco').AsString := global_Contrato_Barco;
  connection.QryBusca.Open;
  CmbAnno.Items.Clear;
  if (connection.QryBusca.RecordCount = 0) or (VarIsNull(connection.QryBusca.FieldValues['minimo'])) then
    CmbAnno.Items.Add(vartostr(yearof(now)))
  else
  begin
    iin := connection.QryBusca.FieldByName('minimo').AsInteger;
    ifi := connection.QryBusca.FieldByName('maximo').AsInteger;
    while iin <= ifi do
    begin
      CmbAnno.Items.Add(IntToStr(iin));
      Inc(iin);
    end;
  end;

  CmbAnno.ItemIndex := CmbAnno.Items.Count-1;
  PnlExistenciasC.Visible := True;
  CargaEmb(CmbAnno.Text,inttostr(CmbMeses.ItemIndex+1));
end;

procedure TfrmAdmonyTiempos.CargaEmb(saño,smes:string);
var ztemp:TZReadOnlyQuery;
Barc:Tembarcacion;
fi,ff:TDateTime;
begin
  if length(smes) < 2 then
    smes := '0'+smes;

  fi := StrToDate('01/'+smes+'/'+saño);
  ff := StrToDate(vartostr(daysinmonth(fi))+'/'+smes+'/'+saño);

  cmbemb.items.Clear;
  ztemp:=TZReadOnlyQuery.Create(nil);
  try
    ztemp.Connection := connection.zConnection;
    ztemp.Active := False;
    ztemp.SQL.Text :=
    'select e.sIdEmbarcacion,e.sDescripcion from embarcaciones e  '+
    'inner join embarcacion_vigencia ev  '+
    'on (e.sContrato = ev.sContrato and e.sIdEmbarcacion = ev.sIdEmbarcacion )  '+
    'where (ev.dFechaInicio between :FechaI and :fechaF) or (ev.dfechafinal between :FechaI '+
    'and :fechaF) or (ev.dFechaInicio = (select max(ev2.dfechainicio) from embarcacion_vigencia ev2 where ev2.dfechainicio < :FechaI)) '+
    'group by e.sIdEmbarcacion order by e.sDescripcion';
    ztemp.ParamByName('fechai').AsDate := fi;
    ztemp.ParamByName('fechaf').AsDate := ff;
    ztemp.Open;
    ztemp.First;
    while not ztemp.Eof do
    begin
      Barc:=Tembarcacion.Create;
      Barc.Identificador := ztemp.FieldByName('sIdEmbarcacion').AsString;
      Barc.Descrp := ztemp.FieldByName('sDescripcion').AsString;
      CmbEmb.AddItem(Barc.Identificador+' : '+Barc.Descrp,Barc);
      ztemp.Next;
    end;
    if ztemp.RecordCount > 0 then
       cmbemb.ItemIndex := 0;

  finally
    ztemp.Free;
  end;

end;

function TfrmAdmonyTiempos.ValidaHorario(sParamHorario: string) : boolean;
var
   sHora, sMinuto : string;
begin
    result := True;
    {Procedimiento para obtener la diferencia de horarios}
    sHora    := copy(sParamHorario,0,2);
    sMinuto := copy(sParamHorario,4,2);

    if sTrToInt(sHora) > 24 then
       result := False;

    if sTrToInt(sMinuto) > 59 then
       result := False;

    if (sTrToInt(sHora) = 24) and (sTrToInt(sMinuto) > 0) then
       result := False;

    if result = false then
       messageDLG('Horario Incorrecto!', mtWarning, [mbOk], 0);
end;

procedure TfrmAdmonyTiempos.AjusteProrrateo;
begin
  try
      if MovimientosdeBarco.RecordCount > 0 then
         TdProrrateoFolio(MovimientosdeBarco.FieldByName('sContrato').AsString,tdFecha.Date,MovimientosdeBarco.FieldByName('iIdDiario').AsInteger);

  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Calculos de Prorrateos de Folios', 'Al ajustar ordenes', 0);
    end;
  end;

end;

procedure TfrmAdmonyTiempos.ExistenciaAnterior;
begin
  //Consultamos el movimiento del día anterior
  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('select dExistenciaActual from recursos where sContrato =:Contrato and dIdFecha =:Fecha and sIdEmbarcacion =:Barco and iIdRecursoExistencia = :Id');
  connection.zCommand.ParamByName('Contrato').AsString := global_contrato_barco;
  connection.zCommand.ParamByName('Fecha').AsDate      := tdFecha.Date -1;
  connection.zCommand.ParamByName('Barco').AsString    := global_barco;
  connection.zCommand.ParamByName('Id').AsInteger      := qryMezclas.FieldByName('iIdRecursoExistencia').AsInteger;
  connection.zCommand.Open;

  if connection.zCommand.RecordCount > 0 then
     dExistenciaAnterior.Value := connection.zCommand.FieldByName('dExistenciaActual').AsFloat
  else
     dExistenciaAnterior.Value := 0;

end;

end.

