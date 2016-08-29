unit frm_EstimaInstalado;

interface

uses
  Windows, Messages, SysUtils, StrUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, DB, ADODB, DBCtrls, global,
  Mask, OleCtrls, Grids, DBGrids, frm_barra, ExtCtrls, Utilerias,
  Menus, frxClass, frxDBSet, RXDBCtrl, RxLookup, DateUtils, unitactivapop,
  RXCtrls, CheckLst, RxMemDS, HintComponent, Newpanel, DisPanel,
  PanelDown, ZAbstractRODataset, ZDataset, UnitValidaTexto,
  rxCurrEdit, rxToolEdit, UnitTBotonesPermisos, UnitExcepciones, UFunctionsGHH;

type
  TfrmEstimaInstalado = class(TForm)
    ds_estimaciones: TDataSource;
    ds_EstimacionxPartida: TDataSource;
    PgControl: TPageControl;
    TabSheet2: TTabSheet;
    Label11: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label17: TLabel;
    frmBarra1: TfrmBarra;
    tmComentarios: TMemo;
    ds_ordenesdetrabajo: TDataSource;
    PopGenerador: TPopupMenu;
    MenuItem9: TMenuItem;
    Salir2: TMenuItem;
    NumerosGeneradores1: TMenuItem;
    SemanalSImportes1: TMenuItem;
    lblIsometrico: TLabel;
    ds_isometricos: TDataSource;
    tlEstima: TCheckBox;
    GridPartidas: TRxDBGrid;
    tsIsometricoReferencia: TRxDBLookupCombo;
    tdCantidad: TRxCalcEdit;
    PopupPrincipal: TPopupMenu;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    N3: TMenuItem;
    ComentariosAdicionales: TMenuItem;
    N6: TMenuItem;
    imgNotas: TImage;
    TabSheet1: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label7: TLabel;
    Label3: TLabel;
    Label6: TLabel;
    tiNumeroEstimacion: TDBLookupComboBox;
    tsNumeroGenerador: TEdit;
    tiConsecutivo: TEdit;
    tdFechaInicial: TDateTimePicker;
    tdFechaFinal: TDateTimePicker;
    tsFaseObra: TEdit;
    tiSemana: TCurrencyEdit;
    tmComentariosGenerador: TMemo;
    GroupBox2: TGroupBox;
    tiFases: TRxCheckListBox;
    lblOrdenCambio: TLabel;
    ListadeVerificacin1: TMenuItem;
    NumerosGeneradoresCIA1: TMenuItem;
    ds_AnexoConvenio: TDataSource;
    Panel: TGroupBox;
    Grid_PartidasConvenios: TRxDBGrid;
    Historialdelapartidaanexo1: TMenuItem;
    tsPrefijo: TEdit;
    Label16: TLabel;
    ds_Prefijos: TDataSource;
    gbIsometricos: tNewGroupBox;
    Grid_Isometricos: TRxDBGrid;
    mnHistorial: TMenuItem;
    Label4: TLabel;
    tsBaseGeneracion: TLabel;
    tsNumeroOrden: TDBLookupComboBox;
    frmBarra2: TfrmBarra;
    lblInstalacion: TLabel;
    OrdenesdeTrabajo: TZReadOnlyQuery;
    QryPrefijos: TZReadOnlyQuery;
    Isometricos: TZReadOnlyQuery;
    AnexoConvenio: TZReadOnlyQuery;
    pdPaquete: TPanelDown;
    tsNumeroActividad: TRxDBLookupCombo;
    ds_actividadesiguales: TDataSource;
    ds_PartidasEfectivas: TDataSource;
    ActividadesIguales: TZReadOnlyQuery;
    QryPartidasEfectivas: TZReadOnlyQuery;
    grid_iguales: TRxDBGrid;
    ActividadesIgualessWbsAnterior: TStringField;
    ActividadesIgualessNumeroActividad: TStringField;
    ActividadesIgualesmDescripcion: TMemoField;
    ActividadesIgualesdCantidad: TFloatField;
    ActividadesIgualesdInstalado: TFloatField;
    ActividadesIgualesdPonderado: TFloatField;
    ActividadesIgualessMedida: TStringField;
    ActividadesIgualesdExcedente: TFloatField;
    tsInstalacion: TComboBox;
    Paquete: TZReadOnlyQuery;
    QryActividadesxOrden: TZReadOnlyQuery;
    QryActividadesxOrdensContrato: TStringField;
    QryActividadesxOrdensNumeroOrden: TStringField;
    QryActividadesxOrdeniNivel: TIntegerField;
    QryActividadesxOrdensWbs: TStringField;
    QryActividadesxOrdensWbsAnterior: TStringField;
    QryActividadesxOrdensNumeroActividad: TStringField;
    QryActividadesxOrdensTipoActividad: TStringField;
    QryActividadesxOrdenmDescripcion: TMemoField;
    QryActividadesxOrdendVentaMN: TFloatField;
    QryActividadesxOrdendVentaDLL: TFloatField;
    QryActividadesxOrdensMedida: TStringField;
    QryActividadesxOrdendCantidadAnexo: TFloatField;
    QryActividadesxOrdendTotal: TFloatField;
    dsQryActividadesxOrden: TfrxDBDataset;
    QryActividadesxOrdendCantidad: TFloatField;
    QryActividadesxOrdendCantidadAcumulado: TFloatField;
    mnResumen: TMenuItem;
    rxActividadesxOrden: TRxMemoryData;
    rxActividadesxOrdensContrato: TStringField;
    rxActividadesxOrdensNumeroOrden: TStringField;
    rxActividadesxOrdeniNivel: TIntegerField;
    rxActividadesxOrdensWbs: TStringField;
    rxActividadesxOrdensWbsAnterior: TStringField;
    rxActividadesxOrdensNumeroActividad: TStringField;
    rxActividadesxOrdensTipoActividad: TStringField;
    rxActividadesxOrdenmDescripcion: TMemoField;
    rxActividadesxOrdendVentaMN: TCurrencyField;
    rxActividadesxOrdendVentaDLL: TCurrencyField;
    rxActividadesxOrdendCantidadanexo: TFloatField;
    rxActividadesxOrdensMedida: TStringField;
    rxActividadesxOrdendCantidad: TFloatField;
    rxActividadesxOrdendTotal: TCurrencyField;
    rxActividadesxOrdendCantidadAcumulado: TFloatField;
    rxActividadesxOrdendTotalAcumulado: TCurrencyField;
    QryActividadesxOrdeniItemOrden: TStringField;
    rxActividadesxOrdeniItemOrden: TStringField;
    tiOrdenCambio: TComboBox;
    mnResumenGeneral: TMenuItem;
    mnConcentradoIsometricos: TMenuItem;
    mnConcentradoIsometricosGral: TMenuItem;
    QryActividadesxOrdendTotalAcumulado: TFloatField;
    estimaciones: TZReadOnlyQuery;
    Grid_Generadores: TDBGrid;
    EstimacionxPartida: TZReadOnlyQuery;
    EstimacionxPartidadMontoMN: TFloatField;
    EstimacionxPartidadVentaMN: TFloatField;
    EstimacionxPartidadGenerado: TFloatField;
    EstimacionxPartidaiItemOrden: TStringField;
    EstimacionxPartidasWbs: TStringField;
    EstimacionxPartidasNumeroActividad: TStringField;
    EstimacionxPartidasIsometrico: TStringField;
    EstimacionxPartidasPrefijo: TStringField;
    EstimacionxPartidasIsometricoReferencia: TStringField;
    EstimacionxPartidasInstalacion: TStringField;
    EstimacionxPartidamComentarios: TMemoField;
    EstimacionxPartidalEstima: TStringField;
    EstimacionxPartidaiOrdenCambio: TIntegerField;
    EstimacionxPartidasMedida: TStringField;
    EstimacionxPartidamDescripcion: TMemoField;
    estimacionessContrato: TStringField;
    estimacionesiNumeroEstimacion: TStringField;
    estimacionessNumeroOrden: TStringField;
    estimacionesiSemana: TIntegerField;
    estimacionesiConsecutivo: TIntegerField;
    estimacionesdFechaInicio: TDateField;
    estimacionesdFechaFinal: TDateField;
    estimacionesdBitacoraInicio: TDateField;
    estimacionesdBitacoraFinal: TDateField;
    estimacionessFaseObra: TStringField;
    estimacionesmComentarios: TMemoField;
    estimacioneslStatus: TStringField;
    estimacionesdMontoMN: TFloatField;
    estimacionesdMontoDLL: TFloatField;
    estimacionesdFinancieroGenerador: TFloatField;
    estimacionessIdUsuario: TStringField;
    estimacionessIdUsuarioValida: TStringField;
    estimacionessIdUsuarioAutoriza: TStringField;
    estimacionessIdUsuarioResidente: TStringField;
    SemanalCImportes1: TMenuItem;
    SemanalCImportesDLL1: TMenuItem;
    qryIsometricoReferencia: TZReadOnlyQuery;
    dsIsometricoReferencia: TfrxDBDataset;
    ConcentradodeIsometricosdeReferenciaxGenerador1: TMenuItem;
    panelIsometrico: TPanel;
    GroupBox1: TGroupBox;
    IsoReferencia: TRxDBLookupCombo;
    btnImprimirIsoRef: TButton;
    memoria: TRxMemoryData;
    memoriaiSemana: TIntegerField;
    memoriasContrato: TStringField;
    memoriasNumeroOrden: TStringField;
    memoriasNumeroActividad: TStringField;
    memoriasMedida: TStringField;
    memoriadVentaMN: TFloatField;
    memoriadVentaDLL: TFloatField;
    memoriasNumeroGenerador: TStringField;
    memoriadInstalado: TFloatField;
    memoriamDescripcion: TStringField;
    frxDBDataset1: TfrxDBDataset;
    panelSemanal: TPanel;
    gpTitulo: TGroupBox;
    cmdSemanal: TButton;
    Fi: TDateTimePicker;
    Ff: TDateTimePicker;
    DataSource1: TDataSource;
    R1: TMenuItem;
    memoriafi: TDateField;
    memoriaff: TDateField;
    fechas: TRxMemoryData;
    DateField1: TDateField;
    DateField2: TDateField;
    dsFechas: TfrxDBDataset;
    Label5: TLabel;
    Label8: TLabel;
    txtMarca: TEdit;
    txtSubMarca: TEdit;
    Label9: TLabel;
    Label10: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    EstimacionxPartidasMarcaRev: TStringField;
    EstimacionxPartidasSubMca: TStringField;
    EstimacionxPartidasLongArea: TStringField;
    EstimacionxPartidasLongAreaTotal: TStringField;
    EstimacionxPartidasPesoxUnidad: TStringField;
    NumGenDespiezados: TMenuItem;
    Label18: TLabel;
    EstimacionxPartidanPiezas: TIntegerField;
    EstimacionxPartidadPesoTotal: TFloatField;
    SemanalCImportes2: TMenuItem;
    frxReport1: TfrxReport;
    btSalir: TButton;
    frGenerador: TfrxReport;
    estimacionessNumeroGenerador: TStringField;
    txtLongArea: TRxCalcEdit;
    txtPesoxUnidad: TRxCalcEdit;
    txtLongTotal: TRxCalcEdit;
    txtPesoTotal: TRxCalcEdit;
    HojaSeguimiento1: TMenuItem;
    Label19: TLabel;
    tsAnexo: TEdit;
    estimacionessNumeroAnexo: TStringField;
    R2: TMenuItem;
    ResumenMN: TMenuItem;
    ResumenDLL: TMenuItem;
    txtPzas: TRxCalcEdit;
    fechasEstima: TStringField;
    memoriasEstimacion: TStringField;
    memoriadInstalado1: TFloatField;
    memoriadInstalado2: TFloatField;
    memoriadInstalado3: TFloatField;
    memoriadInstalado4: TFloatField;
    Datos: TRxMemoryData;
    IntegerField1: TIntegerField;
    StringField1: TStringField;
    StringField2: TStringField;
    StringField3: TStringField;
    StringField4: TStringField;
    FloatField1: TFloatField;
    FloatField2: TFloatField;
    StringField5: TStringField;
    FloatField3: TFloatField;
    StringField6: TStringField;
    DateField3: TDateField;
    DateField4: TDateField;
    StringField7: TStringField;
    memoriaTotal: TFloatField;
    DatosiItemOrden: TStringField;
    memoriaiItemOrden: TStringField;
    txtTag: TEdit;
    Label20: TLabel;
    dfecha: TDateTimePicker;
    Label21: TLabel;
    EstimacionxPartidasTag: TStringField;
    EstimacionxPartidadFecha: TDateField;
    txtTipoMoneda: TEdit;
    Label22: TLabel;
    fechassLeyenda1: TStringField;
    fechassLeyenda2: TStringField;
    fechassLeyenda3: TStringField;
    fechasiFirmas: TIntegerField;
    fechassIdUsuarioValida: TStringField;
    fechassIdUsuarioAutoriza: TStringField;
    EstimacionxPartidasWbsContrato: TStringField;
    actualizadatos1: TMenuItem;
    Actual1: TMenuItem;
    mObra: TMemo;
    Label23: TLabel;
    ActividadesIgualessWbsContrato: TStringField;
    ResumenMensualGeneracionObraGeneral1: TMenuItem;
    ResumenMN1: TMenuItem;
    ResumenDLL1: TMenuItem;
    N5: TMenuItem;
    N7: TMenuItem;
    Insertar1: TMenuItem;
    N4: TMenuItem;
    GenDespiezado: TMenuItem;
    N8: TMenuItem;
    GenTuberia: TMenuItem;
    GenIPR: TMenuItem;
    GenAngulo: TMenuItem;
    ActividadesIgualessWbs: TStringField;
    btnProveedores: TBitBtn;
    btnPlanos: TBitBtn;
    tsIsometrico: TEdit;
    CaratulaWbs: TMenuItem;
    GroupBox3: TGroupBox;
    Label24: TLabel;
    boxNiveles: TComboBox;

    function lExisteIsometrico(sParamGenerador, sParamIsometrico, sParamPrefijo: string): Boolean;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaInicialKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaFinalKeyPress(Sender: TObject; var Key: Char);
    procedure tiNumeroEstimacionKeyPress(Sender: TObject; var Key: Char);
    procedure BtnExitClick(Sender: TObject);
    procedure tsNumeroGeneradorKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroGeneradorExit(Sender: TObject);
    procedure tsNumeroGeneradorEnter(Sender: TObject);
    procedure tsNumeroActividadExit(Sender: TObject);
    procedure tsNumeroActividadKeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure tdCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure tsIsometricoKeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure ImprimirCarClick(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tsIsometricoExit(Sender: TObject);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tiConsecutivoKeyPress(Sender: TObject; var Key: Char);
    procedure tiNumeroEstimacionExit(Sender: TObject);
    procedure tdFechaInicialEnter(Sender: TObject);
    procedure tdFechaInicialExit(Sender: TObject);
    procedure tdFechaFinalEnter(Sender: TObject);
    procedure tdFechaFinalExit(Sender: TObject);
    procedure tsNumeroActividadEnter(Sender: TObject);
    procedure tdCantidadEnter(Sender: TObject);
    procedure tsIsometricoEnter(Sender: TObject);
    procedure tmComentariosEnter(Sender: TObject);
    procedure tmComentariosExit(Sender: TObject);
    procedure frGeneradorGetValue(const VarName: string;
      var Value: Variant);
    procedure tiNumeroEstimacionEnter(Sender: TObject);
    procedure Grid_GeneradoresEnter(Sender: TObject);
    procedure MenuItem9Click(Sender: TObject);
    procedure Salir2Click(Sender: TObject);
    procedure NumerosGeneradores1Click(Sender: TObject);
    procedure SemanalSImportes1Click(Sender: TObject);
    procedure EstimacionesAfterScroll(DataSet: TDataSet);
    procedure EstimacionxPartidaAfterScroll(DataSet: TDataSet);
    procedure tsIsometricoReferenciaKeyPress(Sender: TObject;
      var Key: Char);
    procedure tsIsometricoReferenciaEnter(Sender: TObject);
    procedure tsIsometricoReferenciaExit(Sender: TObject);
    procedure GridPartidasCellClick(Column: TColumn);
    procedure GridPartidasTitleBtnClick(Sender: TObject; ACol: Integer;
      Field: TField);
    procedure FormActivate(Sender: TObject);
    procedure GridPartidasEnter(Sender: TObject);
    procedure frmBarra2btnAddClick(Sender: TObject);
    procedure frmBarra2btnEditClick(Sender: TObject);
    procedure frmBarra2btnPostClick(Sender: TObject);
    procedure frmBarra2btnPrinterClick(Sender: TObject);
    procedure frmBarra2btnDeleteClick(Sender: TObject);
    procedure frmBarra2btnRefreshClick(Sender: TObject);
    procedure frmBarra2btnCancelClick(Sender: TObject);
    procedure frmBarra2btnExitClick(Sender: TObject);
    procedure ComentariosAdicionalesClick(Sender: TObject);
    procedure tsNumeroActividadChange(Sender: TObject);
    procedure imgNotasDblClick(Sender: TObject);
    procedure tiOrdenCambioEnter(Sender: TObject);
    procedure tiOrdenCambioExit(Sender: TObject);
    procedure tiOrdenCambioKeyPress(Sender: TObject; var Key: Char);
    procedure ListadeVerificacin1Click(Sender: TObject);
    procedure NumerosGeneradoresCIA1Click(Sender: TObject);
    procedure Historialdelapartidaanexo1Click(Sender: TObject);
    procedure tsPrefijoEnter(Sender: TObject);
    procedure tsPrefijoExit(Sender: TObject);
    procedure tsPrefijoKeyPress(Sender: TObject; var Key: Char);
    procedure mnHistorialClick(Sender: TObject);
    procedure tsInstalacionEnter(Sender: TObject);
    procedure tsInstalacionExit(Sender: TObject);
    procedure tsInstalacionKeyPress(Sender: TObject; var Key: Char);
    procedure Grid_GeneradoresDblClick(Sender: TObject);
    procedure EstimacionxPartidaCalcFields(DataSet: TDataSet);
    procedure grid_igualesGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure QryActividadesxOrdenCalcFields(DataSet: TDataSet);
    procedure mnResumenClick(Sender: TObject);
    procedure mnResumenGeneralClick(Sender: TObject);
    procedure mnConcentradoIsometricosClick(Sender: TObject);
    procedure mnConcentradoIsometricosGralClick(Sender: TObject);
    procedure SemanalCImportes1Click(Sender: TObject);
    procedure SemanalCImportesDLL1Click(Sender: TObject);
    procedure btnImprimirIsoRefClick(Sender: TObject);
    procedure ConcentradodeIsometricosdeReferenciaxGenerador1Click(
      Sender: TObject);
    procedure cmdSemanalClick(Sender: TObject);
    procedure txtMarcaEnter(Sender: TObject);
    procedure txtMarcaExit(Sender: TObject);
    procedure txtMarcaKeyPress(Sender: TObject; var Key: Char);
    procedure txtSubMarcaEnter(Sender: TObject);
    procedure txtSubMarcaExit(Sender: TObject);
    procedure txtLongAreaEnter(Sender: TObject);
    procedure txtLongAreaExit(Sender: TObject);
    procedure txtLongTotalEnter(Sender: TObject);
    procedure txtLongTotalExit(Sender: TObject);
    procedure txtPesoxUnidadEnter(Sender: TObject);
    procedure txtPesoxUnidadExit(Sender: TObject);
    procedure txtPesoTotalEnter(Sender: TObject);
    procedure txtPesoTotalExit(Sender: TObject);
    procedure txtSubMarcaKeyPress(Sender: TObject; var Key: Char);
    procedure txtLongAreaKeyPress(Sender: TObject; var Key: Char);
    procedure txtLongTotalKeyPress(Sender: TObject; var Key: Char);
    procedure NumGenDespiezadosClick(Sender: TObject);
    procedure txtPzasKeyPress(Sender: TObject; var Key: Char);
    procedure txtPzasExit(Sender: TObject);
    procedure txtPzasEnter(Sender: TObject);
    procedure txtPesoxUnidadKeyPress(Sender: TObject; var Key: Char);
    procedure GeneradorBarco1Click(Sender: TObject);
    procedure GeneradorEquipo1Click(Sender: TObject);
    procedure GeneradorPersonal1Click(Sender: TObject);
    procedure panelSemanalClick(Sender: TObject);
    procedure EquipoXOptativa1Click(Sender: TObject);
    procedure PersonalXOptativa1Click(Sender: TObject);
    procedure btSalirClick(Sender: TObject);
    procedure PernoctaXPlataforma1Click(Sender: TObject);
    procedure BarcoPorPlataformas1Click(Sender: TObject);
    procedure BarcoPorTotalOptativas1Click(Sender: TObject);
    procedure BarcoPorTotalProgramadas1Click(Sender: TObject);
    procedure Barco1Click(Sender: TObject);
    procedure HojaSeguimiento1Click(Sender: TObject);
    procedure tmComentariosGeneradorEnter(Sender: TObject);
    procedure R2Click(Sender: TObject);
    procedure ResumenMNClick(Sender: TObject);
    procedure ResumenDLLClick(Sender: TObject);
    procedure PopGeneradorPopup(Sender: TObject);
    procedure txtTagEnter(Sender: TObject);
    procedure txtTagExit(Sender: TObject);
    procedure ActividadesIgualesAfterScroll(DataSet: TDataSet);
    procedure Actual1Click(Sender: TObject);
    procedure tsAnexoEnter(Sender: TObject);
    procedure tsAnexoExit(Sender: TObject);
    procedure tsAnexoKeyPress(Sender: TObject;
      var Key: Char);

    procedure txtTipoMonedaEnter(Sender: TObject);
    procedure txtTipoMonedaExit(Sender: TObject);
    procedure txtTipoMonedaKeyPress(Sender: TObject;
      var Key: Char);
    procedure tdCantidadExit(Sender: TObject);
    procedure ResumenMN1Click(Sender: TObject);
    procedure ResumenDLL1Click(Sender: TObject);
    procedure GeneradorSemana(sParamTipo, sParamFrente: string);
    procedure actualizadatos1Click(Sender: TObject);
    procedure GenDespiezadoClick(Sender: TObject);
    procedure GenTuberiaClick(Sender: TObject);
    procedure GenIPRClick(Sender: TObject);
    procedure GenAnguloClick(Sender: TObject);
    procedure GridPartidasDblClick(Sender: TObject);
    procedure btnProveedoresClick(Sender: TObject);
    procedure btnPlanosClick(Sender: TObject);
    procedure CaratulaWbsClick(Sender: TObject);
  private
    sMenuP: string;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEstimaInstalado: TfrmEstimaInstalado;
  sIsometrico: string;
  sPrefijo: string;
  sIsometricoReferencia: string;
  sInstalacion: string;
  mComentarios: WideString;
  OpcButton1: string;
  Opcion: string;

  sPartidas: string;
  lExtraordinario: Boolean;
  sOpcion,
    lIniciado: Boolean;
  sTipoReporte: string;
  contador: Byte;

  lPerimetros: boolean;
  isometricoOld: string;

  BotonPermiso: TBotonesPermisos;
  BotonPermiso1: TBotonesPermisos;

implementation


uses  frm_isometricos, frm_trinomios, frm_connection, frm_comentariosxanexo, frm_EstimacionAlbum,
  frm_bitacoraxalcance, frm_bitacoradepartamental_2, frm_DespieceDX,
  frm_DespieceImagen;

{$R *.dfm}

procedure TfrmEstimaInstalado.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  frmBarra2.btnExitClick(Sender);
  action := cafree;
  BotonPermiso.free;
end;

procedure TfrmEstimaInstalado.FormShow(Sender: TObject);
var
  y: string;
  QryTrinomio: tzReadOnlyQuery;
  QryOrdenCambio: tzReadOnlyQuery;
  i: Integer;
begin
  sMenuP := stMenu;
  lPerimetros := False;

  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'opGeneradores', PopGenerador);

  if Connection.Configuracion.FieldValues['sGenDesp'] <> 'Despiezado' then
  begin
    tmComentarios.Top    := 146;
    tmComentarios.Height := 70;
    label17.Top := 146;
    label5.Visible := False;
    label8.Visible := False;
    label10.Visible := False;
    label8.Visible := False;
    label9.Visible := False;
    label12.Visible := False;
    label13.Visible := False;
    label21.Visible := False;
    label20.Visible := False;
    txtMarca.Visible := False;
    txtSubMarca.Visible := False;
    txtLongArea.Visible := False;
    txtPzas.Visible := False;
    txtLongTotal.Visible := False;
    txtPesoxUnidad.Visible := False;
    txtPesoTotal.Visible := False;
    txtTag.Visible := False;
    dFecha.Visible := False;
  end
  else
  begin
    txtMarca.Enabled := False;
    txtSubMarca.Enabled := False;
    txtLongArea.Enabled := False;
    txtPzas.Enabled := False;
    txtLongTotal.Enabled := False;
    txtPesoxUnidad.Enabled := False;
    txtPesoTotal.Enabled := False;
    label5.Visible := True;
    label8.Visible := True;
    label10.Visible := True;
    label8.Visible := True;
    label9.Visible := True;
    label12.Visible := True;
    label13.Visible := True;
    label21.Visible := True;
    label20.Visible := True;
    txtMarca.Visible := True;
    txtSubMarca.Visible := True;
    txtLongArea.Visible := True;
    txtPzas.Visible := True;
    txtLongTotal.Visible := True;
    txtPesoxUnidad.Visible := True;
    txtPesoTotal.Visible := True;
    tdCantidad.Enabled := False;
    tsIsometrico.Enabled := False;
    tsPrefijo.Enabled := False;
    txtTag.Visible := True;
    dFecha.Visible := True;

  end;

  fi.DateTime := date() - 7;
  ff.DateTime := date();

  QryTrinomio := tzReadOnlyQuery.Create(Self);
  QryTrinomio.Connection := connection.zConnection;

  Connection.EstimacionPeriodo.Active := False;
  Connection.EstimacionPeriodo.Open;

  tiFases.Clear;
  //codigo comentado, codigo eliminado..
  Connection.qryBusca.Active := False;
  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('Select sDescripcion From fasesxproyecto Where sDescripcion not like ("%C-%") Order By sDescripcion');
  Connection.qryBusca.Open;
  while not Connection.qryBusca.Eof do
  begin
    tiFases.Items.Add(Connection.qryBusca.FieldValues['sDescripcion']);
    Connection.qryBusca.Next
  end;

  EstimacionxPartida.Active := False;
  EstimacionxPartida.SQL.Clear;
  if Pos('WBS', global_checkGenerador) > 0 then
  begin
    pdPaquete.Visible := True;
    if ((Global_Optativa = 'PROGRAMADA') or (Global_Optativa = 'MIXTA')) and (Connection.Configuracion.FieldValues['sAnexos'] = 'No') then
      EstimacionxPartida.SQL.Add('Select a.iItemOrden, e1.sWbs, e1.sWbsContrato, e1.sNumeroActividad, e1.sIsometrico, e1.sPrefijo, e1.dCantidad as dGenerado, e1.sIsometricoReferencia, ' +
        'e1.nPiezas, e1.sMarcaRev, e1.sSubMca, e1.sLongArea, e1.sLongAreaTotal, e1.sPesoxUnidad, e1.dPesoTotal, ' +
        'e1.sInstalacion, e1.mComentarios, e1.lEstima, e1.iOrdenCambio, a.sMedida, a.mDescripcion, a.dVentaMN, e1.sTag, e1.dFecha from estimacionxpartida e1 ' +
        'inner join actividadesxorden a on (a.sContrato = e1.sContrato and a.sIdConvenio = :Convenio and a.sNumeroOrden = e1.sNumeroOrden and ' +
        'replace(a.sWbs," ","") = replace(e1.sWbs," ","") and replace(a.sNumeroActividad ," ","") = replace(e1.sNumeroActividad ," ","") And a.sTipoActividad = "Actividad") ' +
        'Where e1.sContrato = :Contrato And e1.sNumeroOrden = :Orden And e1.sNumeroGenerador = :Generador Order By :Ordenado ASC')
    else
      EstimacionxPartida.SQL.Add('Select a.iItemOrden, e1.sWbs, e1.sWbsContrato, e1.sNumeroActividad, e1.sIsometrico, e1.sPrefijo, e1.dCantidad as dGenerado, e1.sIsometricoReferencia, ' +
        'e1.nPiezas, e1.sMarcaRev, e1.sSubMca, e1.sLongArea, e1.sLongAreaTotal, e1.sPesoxUnidad, e1.dPesoTotal, ' +
        'e1.sInstalacion, e1.mComentarios, e1.lEstima, e1.iOrdenCambio, a.sMedida, a.mDescripcion, a.dVentaMN, e1.sTag, e1.dFecha from estimacionxpartida e1 ' +
        'inner join actividadesxanexo a on (a.sContrato = e1.sContrato And a.sIdConvenio =:Convenio And ' +
        'replace(a.sWbs," ","") = replace(e1.sWbsContrato ," ","") and replace(a.sNumeroActividad ," ","") = replace(e1.sNumeroActividad ," ","")  And a.sTipoActividad = "Actividad" ) ' +
        'Where e1.sContrato = :Contrato And e1.sNumeroOrden = :Orden And e1.sNumeroGenerador = :Generador Order By :Ordenado ASC');
  end
  else
  begin
    pdPaquete.Visible := False;
    EstimacionxPartida.SQL.Add('Select a.iItemOrden, e1.sWbs, e1.sWbsContrato, e1.sNumeroActividad, e1.sIsometrico, e1.sPrefijo, e1.dCantidad as dGenerado, e1.sIsometricoReferencia, ' +
      'e1.nPiezas, e1.sMarcaRev, e1.sSubMca, e1.sLongArea, e1.sLongAreaTotal, e1.sPesoxUnidad, e1.dPesoTotal, ' +
      'e1.sInstalacion, e1.mComentarios, e1.lEstima, e1.iOrdenCambio, a.sMedida, a.mDescripcion, a.dVentaMN, e1.sTag, e1.dFecha from estimacionxpartida e1 ' +
      'inner join actividadesxanexo a on (a.sContrato = e1.sContrato and a.sIdConvenio = :Convenio and replace(a.sNumeroActividad ," ","") = replace(e1.sNumeroActividad ," ","")  And a.sTipoActividad = "Actividad") ' +
      'Where e1.sContrato = :Contrato And e1.sNumeroOrden = :Orden And e1.sNumeroGenerador = :Generador Order By :Ordenado ASC');
  end;

  tsInstalacion.Items.Clear;
  QryTrinomio.Active := False;
  QryTrinomio.SQL.Clear;
  QryTrinomio.SQL.Add('Select sInstalacion from contrato_trinomio Where sContrato = :contrato and lVigente = "Si" Order By sInstalacion');
  QryTrinomio.Params.ParamByName('Contrato').DataType := ftString;
  QryTrinomio.Params.ParamByName('Contrato').Value := global_contrato;
  QryTrinomio.Open;
  if QryTrinomio.RecordCount > 0 then
    while not QryTrinomio.Eof do
    begin
      tsInstalacion.Items.Add(QryTrinomio.FieldValues['sInstalacion']);
      QryTrinomio.Next
    end
  else
    tsInstalacion.Items.Add(global_contrato);
  QryTrinomio.Destroy;
  tsInstalacion.ItemIndex := 0;

  QryOrdenCambio := tzReadOnlyQuery.Create(Self);
  QryOrdenCambio.Connection := connection.zConnection;
  tiOrdenCambio.Items.Clear;
  tiOrdenCambio.Items.Add('SIN ORDEN DE CAMBIO');
  QryOrdenCambio.Active := False;
  QryOrdenCambio.SQL.Clear;
  QryOrdenCambio.SQL.Add('Select iCedulaProcedencia, sNotificacionOficio from ordendecambio Where sContrato = :contrato order by iCedulaProcedencia');
  QryOrdenCambio.Params.ParamByName('Contrato').DataType := ftString;
  QryOrdenCambio.Params.ParamByName('Contrato').Value := global_contrato;
  QryOrdenCambio.Open;
  while not QryOrdenCambio.Eof do
  begin
    tiOrdenCambio.Items.Add('O.C. No. [' + IntToStr(QryOrdenCambio.FieldValues['iCedulaProcedencia']) + ']');
    QryOrdenCambio.Next
  end;
  QryOrdenCambio.Destroy;
  tiOrdenCambio.ItemIndex := 0;

  if global_orden_general <> '' then
  begin
    OrdenesdeTrabajo.Active := False;
    OrdenesdeTrabajo.SQL.Clear;
    OrdenesdeTrabajo.SQL.Add('Select sNumeroOrden, sDescripcionCorta, mDescripcion from ordenesdetrabajo where sContrato = :Contrato and ' +
      'sNumeroOrden = :orden');
    OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString;
    OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := Global_Contrato;
    ordenesdetrabajo.Params.ParamByName('orden').DataType := ftString;
    ordenesdetrabajo.Params.ParamByName('orden').Value := global_orden_general;
    OrdenesdeTrabajo.Open;
  end
  else
  begin
    OrdenesdeTrabajo.Active := False;
    OrdenesdeTrabajo.SQL.clear;
     //Agrega todoas las ordenes a Usuarios de Intel-Code
    if global_grupo = 'INTEL-CODE' then
      OrdenesdeTrabajo.SQL.Add('Select sNumeroOrden, sDescripcionCorta, mDescripcion from ordenesdetrabajo where sContrato = :Contrato and ' +
        'cIdStatus =:Status order by sIdFolio, sNumeroOrden')
    else
      OrdenesdeTrabajo.SQL.Add('Select  ot.sNumeroOrden, ot.sIdPlataforma, ot.sDescripcionCorta, ot.sIdPernocta ' +
        'from ordenesdetrabajo ot ' +
        'INNER JOIN ordenesxusuario ou On (ou.sContrato=ot.sContrato ' +
        'And ou.sNumeroOrden=ot.sNumeroOrden) ' +
        'where ot.sContrato =:Contrato And ou.sDerechos<>"BLOQUEADO" ' +
        'And ou.sIdUsuario =:Usuario And ot.cIdStatus =:Status order by ot.sNumeroOrden');
    OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString;
    OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := Global_Contrato;
    Ordenesdetrabajo.Params.ParamByName('status').DataType := ftString;
    Ordenesdetrabajo.Params.ParamByName('status').Value := connection.configuracion.FieldValues['cStatusProceso'];
    if global_grupo <> 'INTEL-CODE' then
    begin
      OrdenesdeTrabajo.Params.ParamByName('Usuario').DataType := ftString;
      OrdenesdeTrabajo.Params.ParamByName('Usuario').Value := Global_Usuario;
    end;
    OrdenesdeTrabajo.Open;
    if OrdenesdeTrabajo.RecordCount > 0 then
      tsNumeroOrden.KeyValue := OrdenesdeTrabajo.FieldValues['sNumeroOrden'];
     //termina termina codigo..
  end;

  QryPartidasEfectivas.Active := False;
  qryPartidasEfectivas.Sql.Clear;

  if (Connection.Contrato.FieldValues['sTipoObra'] = 'PROGRAMADA') or (Connection.Contrato.FieldValues['sTipoObra'] = 'MIXTA') then
  begin
    if (Connection.Configuracion.FieldValues['sAnexos'] = 'Si') then
      qryPartidasEfectivas.Sql.Add('(SELECT DISTINCT a.sNumeroActividad, a.mDescripcion FROM actividadesxanexo a ' +
        'WHERE a.sContrato = :Contrato And (a.sMedida <>"ACTIV" Or a.sMedida <>"ACTIVIDAD") ' +
        'And a.sIdConvenio = :Convenio and a.sTipoActividad = "Actividad"  And a.sAnexo <> "" Order By a.iItemOrden ) ' +
        'UNION ' +
        '(SELECT DISTINCT a.sNumeroActividad, a.mDescripcion FROM actividadesxorden a ' +
        'WHERE a.sContrato = :Contrato And a.sNumeroOrden = :Orden ' +
        'And (a.sMedida <>"ACTIV" Or a.sMedida <>"ACTIVIDAD" or a.sMedida<>"Actividad" ) ' +
        'And a.sIdConvenio = :Convenio and a.sTipoActividad = "Actividad" Order By a.iItemOrden )')
    else
      qryPartidasEfectivas.Sql.Add('SELECT DISTINCT a.sNumeroActividad, a.mDescripcion FROM actividadesxorden a ' +
        'INNER JOIN actividadesxanexo an on (an.sContrato = a.sContrato and an.sIdConvenio = a.sIdConvenio and ' +
        'a.sWbsContrato = an.sWbs and a.sNumeroActividad = an.sNumeroActividad) ' +
        'WHERE a.sContrato = :Contrato And a.sNumeroOrden = :Orden ' +
        'And (a.sMedida <>"ACTIV" Or a.sMedida <>"ACTIVIDAD" or a.sMedida<>"Actividad" ) ' +
        'And a.sIdConvenio = :Convenio and a.sTipoActividad = "Actividad" Order By an.iItemOrden')
  end
  else
    qryPartidasEfectivas.Sql.Add('SELECT DISTINCT a.sNumeroActividad, a.mDescripcion FROM actividadesxanexo a ' +
      'WHERE a.sContrato = :Contrato And (a.sMedida <>"ACTIV" Or a.sMedida <>"ACTIVIDAD") ' +
      'And a.sIdConvenio = :Convenio and a.sTipoActividad = "Actividad" And a.sAnexo <> "" Order By a.iItemOrden');
  QryPartidasEfectivas.Params.ParamByName('contrato').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('contrato').Value := global_contrato;
  QryPartidasEfectivas.Params.ParamByName('convenio').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('convenio').Value := global_convenio;
  if (Connection.Contrato.FieldValues['sTipoObra'] = 'PROGRAMADA') or (Connection.Contrato.FieldValues['sTipoObra'] = 'MIXTA') then
  begin
    QryPartidasEfectivas.Params.ParamByName('Orden').DataType := ftString;
    QryPartidasEfectivas.Params.ParamByName('Orden').Value := tsNumeroOrden.KeyValue;

  end;
  QryPartidasEfectivas.Open;

  tdFechaInicial.Date := Date;
  tdFechaFinal.Date := Date;
  pgControl.ActivePageIndex := 0;
  tsBaseGeneracion.Caption := 'BASE DE GENERACIÓN: ' + connection.configuracion.FieldValues['sBaseGeneracion'];
  lIniciado := False;

  Isometricos.Active := False;
  Isometricos.Params.ParamByName('contrato').DataType := ftString;
  Isometricos.Params.ParamByName('contrato').Value := global_contrato;
  Isometricos.Open;

  Estimaciones.Active := False;
  Estimaciones.Params.ParamByName('contrato').DataType := ftString;
  Estimaciones.Params.ParamByName('contrato').Value := global_contrato;
  Estimaciones.Params.ParamByName('Orden').DataType := ftString;
  Estimaciones.Params.ParamByName('Orden').Value := tsNumeroOrden.KeyValue;
  Estimaciones.Open;

  tdFechaInicial.Enabled := False;
  tdFechaFinal.Enabled := False;
  tsFaseObra.ReadOnly := True;
  tiFases.Enabled := False;
  tsNumeroGenerador.ReadOnly := True;
  tiNumeroEstimacion.ReadOnly := True;
  tmComentariosGenerador.ReadOnly := True;

  tsNumeroActividad.ReadOnly := True;
  tdCantidad.ReadOnly := True;
  tsIsometrico.ReadOnly := True;
  tsPrefijo.ReadOnly := True;
  tmComentarios.ReadOnly := True;
  tsInstalacion.Enabled := False;
  tsIsometricoReferencia.Enabled := False;
  tiOrdenCambio.Enabled := False;

  tsNumeroOrden.Color := global_color_salida;
  OpcButton1 := '';
  BotonPermiso.permisosBotones(nil);
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
  {Grid_Generadores.SetFocus; }
end;

procedure TfrmEstimaInstalado.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
  {if Key = #13 then
    Grid_Generadores.SetFocus}
end;

procedure TfrmEstimaInstalado.tdFechaInicialKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tdFechaFinal.SetFocus

end;

procedure TfrmEstimaInstalado.tdFechaFinalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tiFases.SetFocus
end;

procedure TfrmEstimaInstalado.tiNumeroEstimacionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    txtTipoMoneda.SetFocus;
end;

procedure TfrmEstimaInstalado.BtnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmEstimaInstalado.tsNumeroGeneradorKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tsanexo.SetFocus
end;

procedure TfrmEstimaInstalado.tsNumeroGeneradorExit(Sender: TObject);
begin
  tsNumeroGenerador.Color := global_color_salida
end;

procedure TfrmEstimaInstalado.tsNumeroGeneradorEnter(Sender: TObject);
begin
  tsNumeroGenerador.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.tsNumeroActividadExit(Sender: TObject);
begin
  tsNumeroActividad.Color := global_color_salida;
  if tsNumeroActividad.ReadOnly = False then
    if tsNumeroActividad.Text <> '' then
    begin
      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('Select sNumeroActividad From comentariosxanexo Where sContrato = :Contrato And sNumeroActividad = :Actividad');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca.Params.ParamByName('actividad').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
      Connection.QryBusca.Open;
      if Connection.QryBusca.RecordCount > 0 then
        imgNotas.Visible := True;

      AnexoConvenio.Active := False;
      AnexoConvenio.Params.ParamByName('Contrato').DataType := ftString;
      AnexoConvenio.Params.ParamByName('Contrato').Value := global_contrato;
      AnexoConvenio.Params.ParamByName('Actividad').DataType := ftString;
      AnexoConvenio.Params.ParamByName('Actividad').Value := tsNumeroActividad.Text;
      AnexoConvenio.Open;

      pdPaquete.Caption := '< < Seleccione un Paquete > >';
      pdPaquete.Hint := '< < Seleccione un Paquete > >';

      if Pos('WBS', global_checkGenerador) > 0 then
      begin

        ActividadesIguales.Active := False;
        ActividadesIguales.SQL.Clear;
        if ((Global_Optativa = 'PROGRAMADA') or (Global_Optativa = 'MIXTA') or (Global_Optativa = 'OPTATIVA')) and (Connection.Configuracion.FieldValues['sAnexos'] = 'No') then
          ActividadesIguales.SQL.Add('SELECT a.sWbsAnterior, a.sWbs, a.sWbsContrato, a.sNumeroActividad, a.mDescripcion, a.dCantidad, (a.dInstalado + a.dExcedente) as dInstalado, a.dPonderado, ' +
            'a.sMedida, a.dExcedente  FROM actividadesxorden a ' +
            'WHERE a.sContrato = :Contrato And a.sNumeroOrden = :orden And a.sIdConvenio = :convenio ' +
            'And a.sNumeroActividad = :Actividad And a.sTipoActividad = "Actividad" Order By a.iItemOrden');

  { ActividadesIguales.SQL.Add('SELECT a.sWbsAnterior, a.sWbs, a.sWbsContrato, a.sNumeroActividad, a.mDescripcion, a.dCantidadAnexo As dCantidad, (a.dInstalado + a.dExcedente) as dInstalado, a.dPonderado, ' +
                                   'a.sMedida, a.dExcedente  FROM actividadesxanexo a ' +
                                   'WHERE a.sContrato = :Contrato And a.sIdConvenio = :convenio ' +
                                   'And a.sNumeroActividad = :Actividad And a.sTipoActividad = "Actividad" and a.sIdFase <> "" Order By a.iItemOrden') ;}

  { ActividadesIguales.SQL.Add('SELECT a.sWbsAnterior, a.sWbs, a.sWbs as sWbsContrato, a.sNumeroActividad, a.mDescripcion, a.dCantidadAnexo As dCantidad, (a.dInstalado + a.dExcedente) as dInstalado, a.dPonderado, ' +
                                   'a.sMedida, a.dExcedente  FROM actividadesxanexo a ' +
                                   'WHERE a.sContrato = :Contrato And a.sIdConvenio = :convenio ' +                 //este es el bueno
                                   'And a.sNumeroActividad = :Actividad And a.sTipoActividad = "Actividad" and a.sIdFase <> "" Order By a.iItemOrden') ;}



        if Connection.Configuracion.FieldValues['sAnexos'] = 'Si' then
          ActividadesIguales.SQL.Add('SELECT a.sWbsAnterior, a.sWbs, a.sWbs as sWbsContrato, a.sNumeroActividad, a.mDescripcion, a.dCantidadAnexo As dCantidad, (a.dInstalado + a.dExcedente) as dInstalado, a.dPonderado, ' +
            'a.sMedida, a.dExcedente  FROM actividadesxanexo a ' +
            'WHERE a.sContrato = :Contrato And a.sIdConvenio = :convenio ' +
            'And a.sNumeroActividad = :Actividad And a.sTipoActividad = "Actividad" and a.sAnexo <> "" Order By a.iItemOrden');
        ActividadesIguales.Params.ParamByName('Contrato').DataType := ftString;
        ActividadesIguales.Params.ParamByName('Contrato').Value := global_contrato;
        ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
        ActividadesIguales.Params.ParamByName('convenio').Value := global_convenio;
        if ((Global_Optativa = 'PROGRAMADA') or (Global_Optativa = 'MIXTA') or (Global_Optativa = 'OPTATIVA')) and (Connection.Configuracion.FieldValues['sAnexos'] = 'No') then
        begin
          ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
          ActividadesIguales.Params.ParamByName('orden').Value := tsNumeroOrden.KeyValue;
        end;
        ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
        ActividadesIguales.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
        ActividadesIguales.Open;

        if OpcButton1 = 'Edit' then
          ActividadesIguales.Locate('sWbs', EstimacionxPartida.FieldValues['sWbs'], [loPartialKey]);
      end;

      Paquete.Active := False;
      Paquete.Params.ParamByName('contrato').DataType := ftString;
      Paquete.Params.ParamByName('contrato').Value := global_contrato;
      Paquete.Params.ParamByName('orden').DataType := ftString;
      Paquete.Params.ParamByName('orden').Value := tsNumeroOrden.KeyValue;
      Paquete.Params.ParamByName('Wbs').DataType := ftString;
      Paquete.Params.ParamByName('Wbs').Value := ActividadesIguales.FieldValues['sWbsAnterior'];
      Paquete.Open;
      if Paquete.RecordCount > 0 then
      begin
        pdPaquete.Caption := Paquete.FieldValues['mDescripcion'];
        pdPaquete.Hint := Paquete.FieldValues['mDescripcion']
      end
      else
      begin
        pdPaquete.Caption := '< < Seleccione un Paquete > >';
        pdPaquete.Hint := '< < Seleccione un Paquete > >';
      end
    end;
  txtMarca.Enabled := False;
  txtSubMarca.Enabled := False;
  txtPzas.Enabled := False;
  txtPesoTotal.Enabled := False;
  txtLongArea.Enabled := False;
  txtLongTotal.Enabled := False;
  txtPesoxUnidad.Enabled := False;
  if (ActividadesIguales.FieldValues['sMedida'] = 'TON') or (ActividadesIguales.FieldValues['sMedida'] = 'KG') or (ActividadesIguales.FieldValues['sMedida'] = 'M2') then
  begin
    txtLongArea.Enabled := True;
    txtPesoxUnidad.Enabled := True;
  end;
  if (ActividadesIguales.FieldValues['sMedida'] = 'M') then
    txtLongArea.Enabled := True;
  txtMarca.Enabled := True;
  txtSubMarca.Enabled := True;
  txtPzas.Enabled := True;
  txtPesoTotal.Enabled := True;
end;


procedure TfrmEstimaInstalado.tsNumeroActividadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
  begin
    if Connection.Configuracion.FieldValues['sGenDesp'] <> 'Despiezado' then
      tdCantidad.SetFocus
    else
      tsInstalacion.SetFocus;
  end;

end;

procedure TfrmEstimaInstalado.frmBarra1btnAddClick(Sender: TObject);
begin
  if (Estimaciones.RecordCount > 0) then
    if Estimaciones.FieldValues['lStatus'] = 'Pendiente' then
    begin
      OpcButton1 := 'New';

      txtMarca.Text := '';
      txtSubMarca.Text := '';
      txtTag.Text := '';
      txtLongArea.Value := 0;
      txtPzas.Text := '0';
      txtLongTotal.Value := 0;
      txtPesoxUnidad.Value := 0;
      txtPesoTotal.Value := 0;

      Insertar1.Enabled := False;
      Editar1.Enabled := False;
      Registrar1.Enabled := True;
      Can1.Enabled := True;
      Eliminar1.Enabled := False;
      Refresh1.Enabled := False;
      frmBarra1.btnAddClick(Sender);

      tsNumeroActividad.ReadOnly := False;
      tdCantidad.ReadOnly := False;
      tsIsometrico.ReadOnly := False;
      if Connection.Configuracion.FieldValues['sGenDesp'] <> 'Despiezado' then
        tsPrefijo.ReadOnly := True;
      tsPrefijo.ReadOnly := False;
      tdCantidad.ReadOnly := False;
      tmComentarios.ReadOnly := False;
      // Validaciones para ocultar o mostrar los siguientes valores ...
      if Pos('INSTALACION', global_checkGenerador) > 0 then
        tsInstalacion.Enabled := True
      else
        tsInstalacion.Enabled := False;

      if Pos('ORDENDECAMBIO', global_checkGenerador) > 0 then
        tiOrdenCambio.Enabled := True
      else
        tiOrdenCambio.Enabled := False;

      if Pos('REFERENCIA', global_checkGenerador) > 0 then
        tsIsometricoReferencia.Enabled := True
      else
        tsIsometricoReferencia.Enabled := False;
      // Termina Validaciones

      tdCantidad.Value := 0;
      tsPrefijo.Text := sPrefijo;
      tsIsometricoReferencia.KeyValue := sIsometricoReferencia;
      tsInstalacion.Text := sInstalacion;
      tiOrdenCambio.ItemIndex := 0;
      tmComentarios.Text := mComentarios;
      tsNumeroActividad.SetFocus;
    end
    else
      MessageDlg('Generador Aplicado, no pueden realizarse cambios', mtWarning, [mbOk], 0);

  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmEstimaInstalado.frmBarra1btnEditClick(Sender: TObject);
begin
  if Estimaciones.RecordCount > 0 then
    if Estimaciones.FieldValues['lStatus'] = 'Pendiente' then
    begin
      OpcButton1 := 'Edit';
      Insertar1.Enabled := False;
      Editar1.Enabled := False;
      Registrar1.Enabled := True;
      Can1.Enabled := True;
      Eliminar1.Enabled := False;
      Refresh1.Enabled := False;
      frmBarra1.btnEditClick(Sender);

      tsNumeroActividad.ReadOnly := False;
      tdCantidad.ReadOnly := False;
      tsIsometrico.ReadOnly := False;
      tsPrefijo.ReadOnly := False;
      tmComentarios.ReadOnly := False;
      // Validaciones para ocultar o mostrar los siguientes valores ...
      if Pos('INSTALACION', global_checkGenerador) > 0 then
        tsInstalacion.Enabled := True
      else
        lblInstalacion.Enabled := False;

      if Pos('ORDENDECAMBIO', global_checkGenerador) > 0 then
        tiOrdenCambio.Enabled := True
      else
        tiOrdenCambio.Enabled := False;

      if Pos('REFERENCIA', global_checkGenerador) > 0 then
        tsIsometricoReferencia.Enabled := True
      else
        tsIsometricoReferencia.Enabled := False;
      // Termina Validaciones


      if tsIsometricoReferencia.Text = '' then
        tsIsometricoReferencia.KeyValue := EstimacionxPartida.FieldValues['sIsometricoReferencia']; ;
      if tsInstalacion.Text = '' then
        tsInstalacion.Text := sInstalacion;

      tsNumeroActividad.SetFocus;
    end
    else
      MessageDlg('Generador Aplicado, no pueden realizarse cambios', mtWarning, [mbOk], 0);
end;

function TfrmEstimaInstalado.lExisteIsometrico(sParamGenerador, sParamIsometrico, sParamPrefijo: string): Boolean;
begin
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select sIsometrico from estimacionxpartida where sContrato = :Contrato and ' +
    'sNumeroGenerador <> :Generador And sNumeroOrden = :Orden And sIsometrico = :Isometrico And sPrefijo = :Prefijo');
  Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
  Connection.QryBusca.Params.ParamByName('Generador').DataType := ftString;
  Connection.QryBusca.Params.ParamByName('Generador').Value := sParamGenerador;
  Connection.QryBusca.Params.ParamByName('Isometrico').DataType := ftString;
  Connection.QryBusca.Params.ParamByName('Isometrico').Value := sParamIsometrico;
  Connection.QryBusca.Params.ParamByName('Prefijo').DataType := ftString;
  Connection.QryBusca.Params.ParamByName('Prefijo').Value := sParamPrefijo;
  Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
  Connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.KeyValue;
  connection.QryBusca.Open;
  if connection.QryBusca.RecordCount > 0 then
    lExisteIsometrico := True
  else
    lExisteIsometrico := False;
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmEstimaInstalado.tdCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
    tsIsometrico.SetFocus;
end;

procedure TfrmEstimaInstalado.frmBarra1btnPostClick(Sender: TObject);
var
  dCantidadInicial: Real;
  lContinua: Boolean;
  SavePlace: TBookmark;
  sWbs: string;
  sWbsContrato: string;
  dCantidad: Double;
  iResp: Byte;
  iOrdenCambio: Word;
  nombres, cadenas: TStringList;
  RepoCantidad, EstaCantidad, MaxCantidad : double;
begin
    {Perimetros..}
  if Connection.Configuracion.FieldValues['sGenDesp'] = 'Detallado' then
  begin
    lPerimetros := False;
        {Actualizamos el isometricos de los despieces...}
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('Update estimaciondespiece set sIsometrico =:Isometrico where sContrato = :Contrato ' +
      'And sNumeroOrden = :Orden And sNumeroGenerador = :Generador ' +
      'And sWbs = :Wbs And sNumeroActividad = :Actividad And sIsometrico = :IsometricoOld ');
    connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
    connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
    connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
    connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
    connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
    connection.zCommand.Params.ParamByName('Generador').value := Estimaciones.FieldValues['sNumeroGenerador'];
    connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
    connection.zCommand.Params.ParamByName('Wbs').value := ActividadesIguales.FieldValues['sWbs'];
    connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
    connection.zCommand.Params.ParamByName('Actividad').value := ActividadesIguales.FieldValues['sNumeroActividad'];
    connection.zCommand.Params.ParamByName('IsometricoOld').DataType := ftString;
    connection.zCommand.Params.ParamByName('IsometricoOld').value := IsometricoOld;
    connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
    connection.zCommand.Params.ParamByName('Isometrico').value := tsIsometrico.Text;
    connection.zCommand.ExecSQL;

        {Actualizamos el isometricos de los despieces imagen...}
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('Update estimaciondespiece_imagen set sIsometrico =:Isometrico where sContrato = :Contrato ' +
      'And sNumeroOrden = :Orden And sNumeroGenerador = :Generador ' +
      'And sWbs = :Wbs And sNumeroActividad = :Actividad And sIsometrico = :IsometricoOld ');
    connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
    connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
    connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
    connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
    connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
    connection.zCommand.Params.ParamByName('Generador').value := Estimaciones.FieldValues['sNumeroGenerador'];
    connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
    connection.zCommand.Params.ParamByName('Wbs').value := ActividadesIguales.FieldValues['sWbs'];
    connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
    connection.zCommand.Params.ParamByName('Actividad').value := ActividadesIguales.FieldValues['sNumeroActividad'];
    connection.zCommand.Params.ParamByName('IsometricoOld').DataType := ftString;
    connection.zCommand.Params.ParamByName('IsometricoOld').value := IsometricoOld;
    connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
    connection.zCommand.Params.ParamByName('Isometrico').value := tsIsometrico.Text;
    connection.zCommand.ExecSQL;
  end;

    {Validacion de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('Isometrico'); nombres.Add('Instalacion'); nombres.Add('Orden de Cambio'); nombres.Add('Isom. de Referencia');
  cadenas.Add(tsIsometrico.Text); cadenas.Add(tsInstalacion.Text); cadenas.Add(tiOrdenCambio.Text); cadenas.Add(tsIsometricoReferencia.Text);
  if not validaTexto(nombres, cadenas, 'Concepto/Partida', tsNumeroActividad.Text) then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
    {Continua insercion de datos}
  if txtLongArea.Enabled = false then
  begin
    if txtPzas.Value <> txtPesoTotal.Value then
      txtPesoTotal.Value := txtPzas.Value;
  end;
  if tsNumeroActividad.Text <> '' then
    lContinua := True
  else
    lContinua := False;

  if lContinua then
    if tsInstalacion.Enabled then
      if tsInstalacion.Text = '' then
        lContinua := False
      else
        lContinua := True
    else
      lContinua := True;

  if lContinua then
    if tsIsometricoReferencia.Enabled then
      if tsIsometricoReferencia.Text = '' then
        lContinua := False
      else
        lContinua := True
    else
      lContinua := True;

  dCantidad := tdCantidad.Value;

  if OpcButton1 = 'Edit' then
    if Pos('WBS', global_checkGenerador) > 0 then
    begin
      sWbs := ActividadesIguales.FieldValues['sWbs'];
      sWbsContrato := ActividadesIguales.FieldValues['sWbsContrato'];
      if ActividadesIguales.FieldValues['sWbs'] = EstimacionxPartida.FieldValues['sWbs'] then
        if dCantidad > EstimacionxPartida.FieldValues['dGenerado'] then
          dCantidad := dCantidad - EstimacionxPartida.FieldValues['dGenerado']
        else
          dCantidad := (0 - EstimacionxPartida.FieldValues['dGenerado']) + dCantidad
    end
    else
    begin
      sWbs := '';
      sWbsContrato := '';
      if dCantidad > EstimacionxPartida.FieldValues['dGenerado'] then
        dCantidad := dCantidad - EstimacionxPartida.FieldValues['dGenerado']
      else
        dCantidad := (0 - EstimacionxPartida.FieldValues['dGenerado']) + dCantidad
    end
  else
    if Pos('WBS', global_checkGenerador) > 0 then
    begin
      sWbs := ActividadesIguales.FieldValues['sWbs'];
      sWbsContrato := ActividadesIguales.FieldValues['sWbsContrato'];
    end
    else
    begin
      sWbs := '';
      sWbsContrato := '';
    end;

  MaxCantidad  := 0;
  RepoCantidad := 0;

  if (connection.configuracion.FieldValues['lAplicaAvisosGen'] = 'Si') then
  begin
      // Calcular la cantidad en base al avance
      EstaCantidad := tdCantidad.value ;

      // Validar si la cantidad captura es valida de acuerdo a sus recepciones
      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Text := 'select a.snumeroactividad, sum(a.dCantidad) as dCantidad ' +
        'from anexo_psuministro a ' +
        'inner join anexo_suministro b on (b.scontrato = a.scontrato and b.ifolio = a.iFolio) ' +
        'inner join actividadesxanexo c on (c.scontrato = a.scontrato and c.sidconvenio = :convenio and c.sNumeroActividad = a.sNumeroActividad) ' +
        'where b.scontrato = :contrato and b.snumeroorden = :orden and c.sTipoActividad = "Actividad" and a.sNumeroActividad = :Actividad and sTipoAnexo ="PU" ' +
        'group by a.sNumeroActividad';
      Connection.QryBusca.ParamByName('contrato').AsString  := global_contrato;
      Connection.QryBusca.ParamByName('convenio').AsString  := global_convenio;
      Connection.QryBusca.ParamByName('orden').AsString     := ordenesdetrabajo.FieldByName('snumeroorden').AsString;
      Connection.QryBusca.ParamByName('actividad').AsString := QryPartidasEfectivas.FieldValues['sNumeroActividad'];
      Connection.QryBusca.Open;

      // Cantidad reportada total de esta partida
      if Connection.QryBusca.RecordCount > 0 then
         MaxCantidad := Connection.QryBusca.FieldByName('dCantidad').AsFloat;

        // Calcular ahora el total de las cantidades capturadas en la bitácora correspondientes a esta partida
      Connection.QryBusca.SQL.Text := 'select	a.sNumeroActividad,	sum(a.dCantidad) as dCantidad ' +
        'from estimacionxpartida a ' +
        'where a.sContrato = :contrato and a.sNumeroOrden = :orden and a.sNumeroActividad = :actividad ' +
        'group by a.sNumeroActividad';
      Connection.QryBusca.ParamByName('contrato').AsString  := param_global_contrato;
      Connection.QryBusca.ParamByName('orden').AsString     := ordenesdetrabajo.FieldByName('snumeroorden').AsString;
      Connection.QryBusca.ParamByName('actividad').AsString := QryPartidasEfectivas.FieldValues['sNumeroActividad'];
      Connection.QryBusca.Open;

      if Connection.QryBusca.RecordCount > 0 then
        RepoCantidad := Connection.QryBusca.FieldByname('dCantidad').AsFloat;

      // Comparar ahora los datos           //HABILITAT CUANDO ESTEN LISTOS LOS AVISOSS.
      if RepoCantidad + EstaCantidad > MaxCantidad then
      begin
          messagedlg('Las cantidad generada para esta partida más la cantidad acumulada suman un volúmen mayor a las cantidades registradas en los avisos de embarque.' + #10 + #10 +
                     'No es posible generar mas volumenes de esta partida, verifique esto e intente de nuevo.', mtInformation, [mbOk], 0);
          abort;
      end;
  end;

  if lContinua then
  begin
      // iResp = 21  Que continue en el ciclo hasta que sea diferente de 21
      // iResp = 0   El usuario no desea generar la partida ...
      // iResp = 1   El usuario puede generar la partida
      // iResp = 13  El usuario no puede generar la partida pero quiere reportarla para poder generarla ...
    iResp := 21;
    lContinua := False;
    while iResp = 21 do
    begin
      try
        iResp := lVerificaGenerador(global_contrato, global_convenio, tsNumeroOrden.Text, sWbs, tsNumeroActividad.Text, Estimaciones.FieldValues['dFechaFinal'], Estimaciones.FieldValues['iConsecutivo'], dCantidad, frmEstimaInstalado);
      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al verificar generador al salvar nuevo registro', 0);
        end;
      end;
      if iResp = 1 then
        lContinua := True
      else
        if iResp = 0 then
          lContinua := False
        else
          if iResp = 13 then
          begin
                  // 1. Detectar que exista un reporte diario sin validar en una fecha inferior al dia final del generador y que este dentro del convenio vigente ...
                  // 2. Detectar si la partida tiene alcances, si es con alcances bitacora de alcances, si no bitacora de actividades ...
            connection.QryBusca.Active := False;
            connection.QryBusca.SQL.Clear;
            connection.QryBusca.SQL.Add('Select r.dIdFecha, r.sNumeroOrden, r.sIdTurno, r.sIdConvenio from reportediario r ' +
              'inner join turnos t on (r.sContrato = t.sContrato and r.sIdTurno = t.sIdTurno and t.sOrigenTierra = "No") ' +
              'Where r.sContrato = :contrato and r.sNumeroOrden = :Orden and ' +
              'r.lStatus = "Pendiente" and r.dIdFecha <= :Fecha and r.sIdConvenio = :Convenio order by r.dIdFecha DESC');
            connection.QryBusca.Params.ParamByName('contrato').DataType := ftString;
            connection.QryBusca.Params.ParamByName('contrato').Value := global_contrato;
            connection.QryBusca.Params.ParamByName('orden').DataType := ftString;
            connection.QryBusca.Params.ParamByName('orden').Value := tsNumeroOrden.Text;
            connection.QryBusca.Params.ParamByName('convenio').DataType := ftString;
            connection.QryBusca.Params.ParamByName('convenio').Value := global_convenio;
            connection.QryBusca.Params.ParamByName('fecha').DataType := ftDate;
            connection.QryBusca.Params.ParamByName('fecha').Value := Estimaciones.FieldValues['dFechaFinal'];
            connection.QryBusca.Open;

            if connection.QryBusca.RecordCount = 0 then
            begin
                      // No existe ningun reporte diario en status pendiente, se cancela la operacion ...
              MessageDlg('No se puede realizar la captura del volumen pendiente debido a que no existe ningun reporte diario ' +
                'en status de PENDIENTE perteneciente al convenio/acta vigente con fecha menor o igual a la fecha de ' +
                'termino de generacion.', mtWarning, [mbOk], 0);
              iResp := 0;
            end
            else
            begin
              global_fecha := connection.QryBusca.FieldValues['dIdFecha'];
              global_orden := connection.QryBusca.FieldValues['sNumeroOrden'];
              global_turno_reporte := connection.QryBusca.FieldValues['sIdTurno'];
              convenio_reporte := connection.QryBusca.FieldValues['sIdConvenio'];
              connection.QryBusca.Active := False;
              connection.QryBusca.SQL.Clear;
              connection.QryBusca.SQL.Add('Select sContrato from alcancesxactividad ' +
                'Where sContrato = :contrato and sNumeroActividad = :Actividad');
              connection.QryBusca.Params.ParamByName('contrato').DataType := ftString;
              connection.QryBusca.Params.ParamByName('contrato').Value := global_contrato;
              connection.QryBusca.Params.ParamByName('actividad').DataType := ftString;
              connection.QryBusca.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
              connection.QryBusca.Open;
              if connection.QryBusca.RecordCount > 0 then
                frmBitacoraxAlcance.showModal
              else
                frmBitacoraDepartamental_2.showmodal;

              iResp := 21;
            end;
          end;
    end;
  end;

  if lContinua then
  begin
    sIsometrico := tsIsometrico.Text;
    sPrefijo := tsPrefijo.Text;
    if tsIsometricoReferencia.Enabled then
      sIsometricoReferencia := tsIsometricoReferencia.KeyValue
    else
      sIsometricoReferencia := 'SIN REFERENCIA';
    if tsInstalacion.Enabled then
      sInstalacion := tsInstalacion.Text
    else
      sInstalacion := global_contrato;

    if tiOrdenCambio.Enabled then
      if tiOrdenCambio.ItemIndex > 0 then
        iOrdenCambio := StrToInt(MidStr(tiOrdenCambio.Text, 11, pos(']', tiOrdenCambio.Text) - 11))
      else
        iOrdenCambio := 0
    else
      iOrdenCambio := 0;

          {Detalle integridad estimacionxpartida_contratotrinomio 12 Junio 2011...}
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select sInstalacion from contrato_trinomio where sContrato =:Contrato and sInstalacion =:Instalacion ');
    connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
    connection.zCommand.ParamByName('Instalacion').AsString := sInstalacion;
    connection.zCommand.Open;

    if connection.zCommand.RecordCount = 0 then
      sInstalacion := '';
          {Termina implementacion de integridad..}

    mComentarios := tmComentarios.Text;
    if OpcButton1 = 'New' then
    begin
              {Insercion de partidas del generador..}
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('INSERT INTO estimacionxpartida ( sContrato , sNumeroOrden, sNumeroGenerador, ' +
        'sWbs, sWbsContrato, sNumeroActividad, sIsometrico, sPrefijo, dCantidad, dAcumulado, iOrdenCambio, sIsometricoReferencia, ' +
        'sInstalacion, mComentarios, lEstima, sMarcaRev, sSubMca, sLongArea, nPiezas, sLongAreaTotal, sPesoxUnidad, dPesoTotal, dFecha, sTag ) ' +
        'VALUES (:Contrato, :Orden, :Generador, :wbs, :WbsContrato, :Actividad, :Isometrico, :Prefijo, :Cantidad, :Acumulado, :OrdenCambio, ' +
        ':Referencia, :Instalacion, :Comentarios, :Genera, :MarcaRev, :SubMarca, :LongArea, :Piezas, :LongAreaTotal, :PesoxUnidad, :PesoTotal, :Fecha, :Tag)');
              // Corrige el error en progrmadas de insertar registros y la creacion de isometricos registros..
      connection.zCommand.Params.ParamByName('Contrato').DataType     := ftString;
      connection.zCommand.Params.ParamByName('Contrato').value        := Global_Contrato;
      connection.zCommand.Params.ParamByName('Orden').DataType        := ftString;
      connection.zCommand.Params.ParamByName('Orden').value           := Estimaciones.FieldValues['sNumeroOrden'];
      connection.zCommand.Params.ParamByName('Generador').DataType    := ftString;
      connection.zCommand.Params.ParamByName('Generador').value       := Estimaciones.FieldValues['sNumeroGenerador'];
      connection.zCommand.Params.ParamByName('wbs').DataType          := ftString;
      connection.zCommand.Params.ParamByName('wbs').value             := sWbs;
      connection.zCommand.Params.ParamByName('wbsContrato').DataType  := ftString;
      connection.zCommand.Params.ParamByName('wbsContrato').value     := sWbsContrato;
      connection.zCommand.Params.ParamByName('Actividad').DataType    := ftString;
      connection.zCommand.Params.ParamByName('Actividad').value       := tsNumeroActividad.Text;

      if Connection.Configuracion.FieldValues['sGenDesp'] = 'Despiezado' then
      begin
        Contador := Random(499) + 1 + 12345;
        connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
        connection.zCommand.Params.ParamByName('Isometrico').value := txtMarca.Text + '/' + txtSubMarca.Text + '-' + IntToStr(Contador);
      end;
      if Connection.Configuracion.FieldValues['sGenDesp'] <> 'Despiezado' then
      begin
        connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
        connection.zCommand.Params.ParamByName('Isometrico').value := tsIsometrico.Text;
      end;
      connection.zCommand.Params.ParamByName('Prefijo').DataType := ftString;
      connection.zCommand.Params.ParamByName('Prefijo').value := tsPrefijo.Text;
      if Connection.Configuracion.FieldValues['sGenDesp'] <> 'Despiezado' then
      begin
        connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Cantidad').value := tdCantidad.Value;
      end
      else
      begin
        connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Cantidad').value := txtPesoTotal.Text;
      end;
      connection.zCommand.Params.ParamByName('Acumulado').DataType := ftFloat;
      connection.zCommand.Params.ParamByName('Acumulado').value := 0;
      connection.zCommand.Params.ParamByName('OrdenCambio').DataType := ftInteger;
      if not iOrdenCambio = null then
        connection.zCommand.Params.ParamByName('OrdenCambio').value := iOrdenCambio
      else
        connection.zCommand.Params.ParamByName('OrdenCambio').value := null;
      connection.zCommand.Params.ParamByName('Referencia').DataType := ftString;
      connection.zCommand.Params.ParamByName('Referencia').value := sIsometricoReferencia;
      connection.zCommand.Params.ParamByName('Instalacion').DataType := ftString;
      if sInstalacion <> '' then
        connection.zCommand.Params.ParamByName('Instalacion').value := sInstalacion
      else
        connection.zCommand.Params.ParamByName('Instalacion').value := null;
      connection.zCommand.Params.ParamByName('Comentarios').DataType := ftMemo;
      connection.zCommand.Params.ParamByName('Comentarios').value := tmComentarios.Text;
      connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
      connection.zCommand.Params.ParamByName('Fecha').value := dFecha.Date;
      connection.zCommand.Params.ParamByName('Tag').DataType := ftString;
      connection.zCommand.Params.ParamByName('Tag').value := txtTag.Text;

      connection.zCommand.Params.ParamByName('Genera').DataType := ftString;
      if tlEstima.State = cbChecked then
        connection.zCommand.Params.ParamByName('Genera').value := 'Si'
      else
        connection.zCommand.Params.ParamByName('Genera').value := 'No';

      if Connection.Configuracion.FieldValues['sGenDesp'] <> 'Despiezado' then
      begin
        txtLongArea.Text := '0.01';
        txtPzas.text := '0.01';
        txtLongTotal.Text := '0.01';
        txtPesoxUnidad.text := '0.01';
        txtPesoTotal.text := '0.01';
      end;
      connection.zCommand.Params.ParamByName('MarcaRev').DataType := ftString;
      connection.zCommand.Params.ParamByName('MarcaRev').value := txtMarca.Text;
      connection.zCommand.Params.ParamByName('SubMarca').DataType := ftString;
      connection.zCommand.Params.ParamByName('SubMarca').value := txtSubMarca.Text;
      connection.zCommand.Params.ParamByName('LongArea').DataType := ftString;
      connection.zCommand.Params.ParamByName('LongArea').value := txtLongArea.Text;
      connection.zCommand.Params.ParamByName('Piezas').DataType := ftInteger;
      connection.zCommand.Params.ParamByName('Piezas').value := txtPzas.Text;
      connection.zCommand.Params.ParamByName('LongAreaTotal').DataType := ftString;
      connection.zCommand.Params.ParamByName('LongAreaTotal').value := txtLongTotal.Text;

      connection.zCommand.Params.ParamByName('PesoxUnidad').DataType := ftString;
      if txtPesoxUnidad.text = '' then
        connection.zCommand.Params.ParamByName('PesoxUnidad').value := '0'
      else
        connection.zCommand.Params.ParamByName('PesoxUnidad').value := txtPesoxUnidad.Text;

      connection.zCommand.Params.ParamByName('PesoTotal').DataType := ftString;
      if txtPesoTotal.Text = '' then
        connection.zCommand.Params.ParamByName('PesoTotal').value := txtLongTotal.Text
      else
        connection.zCommand.Params.ParamByName('PesoTotal').value := txtPesoTotal.Text;
      connection.zCommand.ExecSQL;
    end
    else
    begin
          {Edicion de los datos..}
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('UPDATE estimacionxpartida Set sWbs = :wbs, sWbsContrato = :wbsContrato, sNumeroActividad = :Actividad, dCantidad = :Cantidad, ' +
        'dAcumulado = :Acumulado, iOrdenCambio = :OrdenCambio, sIsometrico = :Isometrico, ' +
        'sMarcaRev = :MarcaRev, sSubMca= :SubMarca, sLongArea= :LongArea, nPiezas = :Piezas, sLongAreaTotal= :LongAreaTotal, sPesoxUnidad= :PesoxUnidad, dPesoTotal= :PesoTotal, ' +
        'sPrefijo = :Prefijo, sIsometricoReferencia = :Referencia, sInstalacion = :Instalacion,  mComentarios = :Comentarios, lEstima = :Genera Where ' +
        'sContrato = :Contrato And sNumeroOrden = :Orden And sNumeroGenerador = :Generador ' +
        'And sWbs = :OldWbs And sNumeroActividad = :OldActividad And sIsometrico = :OldIsometrico And sPrefijo = :OldPrefijo');
          // Corrige el error en programadas de acualizar registros..
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
      connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
      connection.zCommand.Params.ParamByName('Orden').value := Estimaciones.FieldValues['sNumeroOrden'];
      connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
      connection.zCommand.Params.ParamByName('Generador').value := Estimaciones.FieldValues['sNumeroGenerador'];
      connection.zCommand.Params.ParamByName('OldWbs').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldWbs').value := EstimacionxPartida.FieldValues['sWbs'];
      connection.zCommand.Params.ParamByName('OldActividad').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldActividad').value := EstimacionxPartida.FieldValues['sNumeroActividad'];
      connection.zCommand.Params.ParamByName('OldIsometrico').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldIsometrico').value := EstimacionxPartida.FieldValues['sIsometrico'];
      connection.zCommand.Params.ParamByName('OldPrefijo').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldPrefijo').value := EstimacionxPartida.FieldValues['sPrefijo'];
      connection.zCommand.Params.ParamByName('wbs').DataType := ftString;
      connection.zCommand.Params.ParamByName('wbs').value := sWbs;
      connection.zCommand.Params.ParamByName('wbsContrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('wbsContrato').value := sWbsContrato;
      connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
      connection.zCommand.Params.ParamByName('Actividad').value := tsNumeroActividad.Text;
      connection.zCommand.Params.ParamByName('Prefijo').DataType := ftString;
      connection.zCommand.Params.ParamByName('Prefijo').value := tsPrefijo.Text;
      if Connection.Configuracion.FieldValues['sGenDesp'] <> 'Despiezado' then
      begin
        connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Cantidad').value := tdCantidad.Value;
        connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
        connection.zCommand.Params.ParamByName('Isometrico').value := tsIsometrico.Text;
      end
      else
      begin
        connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
        connection.zCommand.Params.ParamByName('Cantidad').value := txtPesoTotal.Text;
        Contador := Random(499) + 1 + 12345;
        connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
        connection.zCommand.Params.ParamByName('Isometrico').value := txtMarca.Text + '/' + txtSubMarca.Text + '-' + IntToStr(Contador);
      end;
      connection.zCommand.Params.ParamByName('Acumulado').DataType := ftFloat;
      connection.zCommand.Params.ParamByName('Acumulado').value := 0;
      connection.zCommand.Params.ParamByName('OrdenCambio').DataType := ftInteger;
      if not iOrdenCambio = null then
        connection.zCommand.Params.ParamByName('OrdenCambio').value := iOrdenCambio
      else
        connection.zCommand.Params.ParamByName('OrdenCambio').value := null;
      connection.zCommand.Params.ParamByName('Referencia').DataType := ftString;
      connection.zCommand.Params.ParamByName('Referencia').value := sIsometricoReferencia;
      connection.zCommand.Params.ParamByName('Instalacion').DataType := ftString;
      if sInstalacion <> '' then
        connection.zCommand.Params.ParamByName('Instalacion').value := sInstalacion
      else
        connection.zCommand.Params.ParamByName('Instalacion').value := null;
      connection.zCommand.Params.ParamByName('Comentarios').DataType := ftMemo;
      connection.zCommand.Params.ParamByName('Comentarios').value := tmComentarios.Text;
      connection.zCommand.Params.ParamByName('Genera').DataType := ftString;
      if tlEstima.State = cbChecked then
        connection.zCommand.Params.ParamByName('Genera').value := 'Si'
      else
        connection.zCommand.Params.ParamByName('Genera').value := 'No';

      if Connection.Configuracion.FieldValues['sGenDesp'] <> 'Despiezado' then
      begin
        txtLongArea.Text := '0.01';
        txtPzas.text := '0.01';
        txtLongTotal.Text := '0.01';
        txtPesoxUnidad.text := '0.01';
        txtPesoTotal.text := '0.01';
        txtpzas.text := '0.01';
      end;
      connection.zCommand.Params.ParamByName('MarcaRev').DataType := ftString;
      connection.zCommand.Params.ParamByName('MarcaRev').value := txtMarca.Text;
      connection.zCommand.Params.ParamByName('SubMarca').DataType := ftString;
      connection.zCommand.Params.ParamByName('SubMarca').value := txtSubMarca.Text;
      connection.zCommand.Params.ParamByName('LongArea').DataType := ftString;
      connection.zCommand.Params.ParamByName('LongArea').value := txtLongArea.Text;
      connection.zCommand.Params.ParamByName('Piezas').DataType := ftInteger;
      connection.zCommand.Params.ParamByName('Piezas').value := txtPzas.Text;
      connection.zCommand.Params.ParamByName('LongAreaTotal').DataType := ftString;
      connection.zCommand.Params.ParamByName('LongAreaTotal').value := txtLongTotal.Text;
      connection.zCommand.Params.ParamByName('PesoxUnidad').DataType := ftString;
      if txtPesoxUnidad.text = '' then
        connection.zCommand.Params.ParamByName('PesoxUnidad').value := '0'
      else
        connection.zCommand.Params.ParamByName('PesoxUnidad').value := txtPesoxUnidad.Text;
      connection.zCommand.Params.ParamByName('PesoTotal').DataType := ftString;
      if txtPesoTotal.Text = '' then
        connection.zCommand.Params.ParamByName('PesoTotal').value := txtLongTotal.Text
      else
        connection.zCommand.Params.ParamByName('PesoTotal').value := txtPesoTotal.Text;
      connection.zCommand.ExecSQL;

    end;
      //termino de correccion..
    SavePlace := EstimacionxPartida.GetBookmark;
    EstimacionxPartida.Active := False;
    EstimacionxPartida.Open;
    try
      EstimacionxPartida.GotoBookmark(SavePlace);
    except
    else
      EstimacionxPartida.FreeBookmark(SavePlace);
    end;

    mObra.Text := '';
    tsNumeroActividad.ReadOnly := True;
    tdCantidad.ReadOnly := True;
    tsIsometrico.ReadOnly := True;
    tsPrefijo.ReadOnly := True;
    tmComentarios.ReadOnly := True;
    tsInstalacion.Enabled := False;
    tsIsometricoReferencia.Enabled := False;
    tiOrdenCambio.Enabled := False;

    Insertar1.Enabled := True;
    Editar1.Enabled := True;
    Registrar1.Enabled := False;
    Can1.Enabled := False;
    Eliminar1.Enabled := True;
    Refresh1.Enabled := True;
    frmBarra1.btnPostClick(Sender);
  end
  else
    MessageDlg('Debera seleccionar un isometrico de ingenieria que ampare la instalacion de la partida anexo asi como la plataforma de instalacion para el direccionamiento de costos.', mtWarning, [mbOk], 0);
  BotonPermiso.permisosBotones(frmBarra1);
end;




procedure TfrmEstimaInstalado.frmBarra1btnCancelClick(Sender: TObject);
begin
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  frmBarra1.btnCancelClick(Sender);
  lPerimetros := False;

  txtMarca.Enabled := False;
  txtSubMarca.Enabled := False;
  txtLongArea.Enabled := False;
  txtPzas.Enabled := False;
  txtLongTotal.Enabled := False;
  txtPesoxUnidad.Enabled := False;
  txtPesoTotal.Enabled := False;

  tsNumeroActividad.ReadOnly := True;
  tdCantidad.ReadOnly := True;
  tsIsometrico.ReadOnly := True;
  tsPrefijo.ReadOnly := True;
  tmComentarios.ReadOnly := True;

  tsInstalacion.Enabled := False;
  tsIsometricoReferencia.Enabled := False;
  tiOrdenCambio.Enabled := False;

  mObra.Text := '';
  tdCantidad.Value := 0;
  tsIsometrico.text := '';
  tsPrefijo.Text := '';
  tsIsometricoReferencia.KeyValue := '';

  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmEstimaInstalado.tsIsometricoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
    tsPrefijo.SetFocus
end;

procedure TfrmEstimaInstalado.frmBarra1btnExitClick(Sender: TObject);
begin
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  frmBarra1.btnExitClick(Sender);
end;

procedure TfrmEstimaInstalado.frmBarra1btnDeleteClick(Sender: TObject);
var
  dCantidadInicial: Real;
  lContinua: Boolean;
  SavePlace: TBookmark;
begin
  if EstimacionxPartida.RecordCount > 0 then
    if Estimaciones.FieldValues['lStatus'] = 'Pendiente' then
    begin
      if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        try
          SavePlace := EstimacionxPartida.GetBookmark;
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('Delete from estimacionxpartida where sContrato = :Contrato ' +
            'And sNumeroOrden = :Orden And sNumeroGenerador = :Generador ' +
            'And sWbs = :Wbs And sNumeroActividad = :Actividad And sIsometrico = :Isometrico And sPrefijo = :Prefijo');
          connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
          connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
          connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
          connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
          connection.zCommand.Params.ParamByName('Generador').value := Estimaciones.FieldValues['sNumeroGenerador'];
          connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
          connection.zCommand.Params.ParamByName('Wbs').value := EstimacionxPartida.FieldValues['sWbs'];
          connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
          connection.zCommand.Params.ParamByName('Actividad').value := EstimacionxPartida.FieldValues['sNumeroActividad'];
          connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
          connection.zCommand.Params.ParamByName('Isometrico').value := EstimacionxPartida.FieldValues['sIsometrico'];
          connection.zCommand.Params.ParamByName('Prefijo').DataType := ftString;
          connection.zCommand.Params.ParamByName('Prefijo').value := EstimacionxPartida.FieldValues['sPrefijo'];
          connection.zCommand.ExecSQL;

          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('Delete from estimaciondespiece where sContrato = :Contrato ' +
            'And sNumeroOrden = :Orden And sNumeroGenerador = :Generador ' +
            'And sWbs = :Wbs And sNumeroActividad = :Actividad And sIsometrico = :Isometrico ');
          connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
          connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
          connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
          connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
          connection.zCommand.Params.ParamByName('Generador').value := Estimaciones.FieldValues['sNumeroGenerador'];
          connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
          connection.zCommand.Params.ParamByName('Wbs').value := EstimacionxPartida.FieldValues['sWbs'];
          connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
          connection.zCommand.Params.ParamByName('Actividad').value := EstimacionxPartida.FieldValues['sNumeroActividad'];
          connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
          connection.zCommand.Params.ParamByName('Isometrico').value := EstimacionxPartida.FieldValues['sIsometrico'];
          connection.zCommand.ExecSQL;

          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('Delete from estimaciondespiece_imagen where sContrato = :Contrato ' +
            'And sNumeroOrden = :Orden And sNumeroGenerador = :Generador ' +
            'And sWbs = :Wbs And sNumeroActividad = :Actividad And sIsometrico = :Isometrico ');
          connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
          connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
          connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
          connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
          connection.zCommand.Params.ParamByName('Generador').value := Estimaciones.FieldValues['sNumeroGenerador'];
          connection.zCommand.Params.ParamByName('Wbs').DataType := ftString;
          connection.zCommand.Params.ParamByName('Wbs').value := EstimacionxPartida.FieldValues['sWbs'];
          connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
          connection.zCommand.Params.ParamByName('Actividad').value := EstimacionxPartida.FieldValues['sNumeroActividad'];
          connection.zCommand.Params.ParamByName('Isometrico').DataType := ftString;
          connection.zCommand.Params.ParamByName('Isometrico').value := EstimacionxPartida.FieldValues['sIsometrico'];
          connection.zCommand.ExecSQL;

          EstimacionxPartida.Active := False;
          EstimacionxPartida.Open;
          try
            EstimacionxPartida.FreeBookmark(SavePlace);
          except
          else
            EstimacionxPartida.GotoBookmark(SavePlace);
          end;
          {GridPartidas.SetFocus  }
        except
        end
      end
    end
    else
      MessageDlg('Generador Aplicado, no pueden realizarse cambios', mtWarning, [mbOk], 0);
end;

procedure TfrmEstimaInstalado.Insertar1Click(Sender: TObject);
begin
  frmBarra2.btnAdd.Click
end;

procedure TfrmEstimaInstalado.Editar1Click(Sender: TObject);
begin
  frmBarra2.btnEdit.Click
end;

procedure TfrmEstimaInstalado.Registrar1Click(Sender: TObject);
begin
  frmBarra2.btnPost.Click
end;


procedure TfrmEstimaInstalado.ResumenDLL1Click(Sender: TObject);
begin
  if grid_generadores.DataSource.DataSet.IsEmpty = false then
    GeneradorSemana('DLL', '');
end;

procedure TfrmEstimaInstalado.ResumenDLLClick(Sender: TObject);
begin
  if grid_generadores.DataSource.DataSet.IsEmpty = false then
    GeneradorSemana('DLL', tsNumeroOrden.Text);
end;

procedure TfrmEstimaInstalado.ResumenMN1Click(Sender: TObject);
begin
  if grid_generadores.DataSource.DataSet.IsEmpty = false then
    GeneradorSemana('MN', '');
end;

procedure TfrmEstimaInstalado.ResumenMNClick(Sender: TObject);
begin
  if grid_generadores.DataSource.DataSet.IsEmpty = false then
    GeneradorSemana('MN', tsNumeroOrden.Text);
end;


procedure TfrmEstimaInstalado.Can1Click(Sender: TObject);
begin
   frmBarra2.btnCancel.Click
end;

procedure TfrmEstimaInstalado.CaratulaWbsClick(Sender: TObject);
begin
   try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
      begin
        if Connection.configuracion.FieldValues['sCampPerf'] = 'No' then
                   //SE COMENTA ESTO DEBIDO A QUE IMPRIME SE REQUIRIO LA CARATULA DE UNA SOLA ORDEN
                   //procCaratulaGenerador( global_contrato, Estimaciones.FieldValues[ 'iNumeroEstimacion' ], Estimaciones.FieldValues[ 'sNumeroOrden' ], Estimaciones.FieldValues[ 'sNumeroGenerador' ], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, True )
          procCaratulaGenerador(StrToInt(boxNiveles.Text), global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, False)
        else
          MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
      end;

      if Connection.configuracion.FieldValues['sCampPerf'] = 'Si' then
        procCaratulaGeneradorPerf(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, True)
    end
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.Eliminar1Click(Sender: TObject);
begin
  frmBarra2.btnDelete.Click
end;

procedure TfrmEstimaInstalado.Refresh1Click(Sender: TObject);
begin
  frmBarra2.btnRefresh.Click
end;

procedure TfrmEstimaInstalado.ImprimirCarClick(Sender: TObject);
begin
  frmBarra2.btnPrinter.Click
end;

procedure TfrmEstimaInstalado.Salir1Click(Sender: TObject);
begin
  frmBarra1.btnExit.Click
end;

procedure TfrmEstimaInstalado.tsIsometricoExit(Sender: TObject);
begin
  tsIsometrico.Color := global_color_salida;
  if tsNumeroActividad.ReadOnly = False then
    if (tsIsometrico.Text <> '') or (tsPrefijo.Text <> '') then
      if lExisteIsometrico(Estimaciones.FieldValues['sNumeroGenerador'], tsIsometrico.Text, tsPrefijo.Text) then
      begin
        if tsPrefijo.Text <> '' then
        begin
          mnHistorial.Click;
          MessageDlg('El Isometrico: ' + tsIsometrico.Text + '-' + tsPrefijo.Text + ' se encuentra registrado en otro generador de la misma orden de trabajo.', mtInformation, [mbOk], 0);
        end;
        tsPrefijo.SetFocus
      end
      else
        gbIsometricos.Visible := False
end;

procedure TfrmEstimaInstalado.tsNumeroOrdenEnter(Sender: TObject);
begin
  tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.tsNumeroOrdenExit(Sender: TObject);
begin

  Estimaciones.Active := False;
  Estimaciones.Params.ParamByName('contrato').DataType := ftString;
  Estimaciones.Params.ParamByName('contrato').Value := global_contrato;
  Estimaciones.Params.ParamByName('Orden').DataType := ftString;
  Estimaciones.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
  Estimaciones.Open;

  QryPartidasEfectivas.Active := False;
  QryPartidasEfectivas.Params.ParamByName('contrato').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('contrato').Value := global_contrato;
  QryPartidasEfectivas.Params.ParamByName('convenio').DataType := ftString;
  QryPartidasEfectivas.Params.ParamByName('convenio').Value := global_convenio;
  if (Global_Optativa = 'PROGRAMADA') or (Global_optativa = 'MIXTA') then
  begin
    if Connection.Configuracion.FieldValues['sAnexos'] = 'No' then
    begin
      QryPartidasEfectivas.Params.ParamByName('Orden').DataType := ftString;
      QryPartidasEfectivas.Params.ParamByName('Orden').Value := tsNumeroOrden.KeyValue;
    end;
  end;
  QryPartidasEfectivas.Open;

  tdFechaInicial.Enabled := False;
  tdFechaFinal.Enabled := False;
  tsFaseObra.ReadOnly := True;
  tiFases.Enabled := False;
  tsNumeroGenerador.ReadOnly := True;
  tiNumeroEstimacion.ReadOnly := True;
  tmComentariosGenerador.ReadOnly := True;

  tsNumeroActividad.ReadOnly := True;
  tdCantidad.ReadOnly := True;
  tsIsometrico.ReadOnly := True;
  tsPrefijo.ReadOnly := True;
  tmComentarios.ReadOnly := True;

  tsInstalacion.Enabled := False;
  tsIsometricoReferencia.Enabled := False;
  tiOrdenCambio.Enabled := False;

  tsNumeroOrden.Color := global_color_salida;
  OpcButton1 := '';

end;

procedure TfrmEstimaInstalado.tiConsecutivoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    tdFechaInicial.SetFocus
end;

procedure TfrmEstimaInstalado.tiNumeroEstimacionExit(Sender: TObject);
begin
  tiNumeroEstimacion.Color := global_color_salida;
  if tiNumeroEstimacion.ReadOnly = False then
    if Connection.EstimacionPeriodo.FieldValues['lEstimado'] = 'Si' then
    begin
      MessageDlg('Estimacion Validada, no pueden adicionarse generadores.', mtWarning, [mbOk], 0);
      {tiNumeroEstimacion.SetFocus;}
    end;
end;



procedure TfrmEstimaInstalado.tdFechaInicialEnter(Sender: TObject);
begin
  tdFechaInicial.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.tdFechaInicialExit(Sender: TObject);
var
  dFechaFinal: tDate;
begin
  tdFechaInicial.Color := global_color_salida;
  if frmBarra2.btnCancel.Enabled = True then
  begin
    if connection.configuracion.FieldValues['sRangoEstimacion'] = 'Diario' then
      tdFechaFinal.Date := tdFechaInicial.Date + 1
    else if connection.configuracion.FieldValues['sRangoEstimacion'] = 'Semanal' then
      tdFechaFinal.Date := tdFechaInicial.Date + 7
    else if connection.configuracion.FieldValues['sRangoEstimacion'] = 'Mensual' then
      tdFechaFinal.Date := tdFechaInicial.Date + 30;

    if (MonthOf(tdFechaFinal.Date) <> MonthOf(tdFechaInicial.Date)) then
    begin
      if MonthOf(tdFechaInicial.Date) <= 11 then
        dFechaFinal := strToDate('01/' + Trim(IntToStr(MonthOf(tdFechaInicial.Date) + 1)) + '/' + Trim(IntToStr(YearOf(tdFechaInicial.Date))))
      else
        dFechaFinal := strToDate('01/01/' + Trim(IntToStr(YearOf(tdFechaInicial.Date) + 1)));
      dFechaFinal := dFechaFinal - 1;
      MessageDlg('El generador no puede abarcar un periodo de 2 meses. Periodo Propuesto [' + DateToStr(tdFechaInicial.Date) + ' al ' + DateToStr(dFechaFinal) + ']', mtWarning, [mbOk], 0);
      tdFechaFinal.Date := dFechaFinal;
    end;
    dFechaFinal := strToDate('01/' + Trim(IntToStr(MonthOf(tdFechaInicial.Date))) + '/' + Trim(IntToStr(YearOf(tdFechaInicial.Date))));
    tiSemana.Value := WeeksBetween(tdFechaFinal.date, dFechaFinal) + 1;
  end
end;



procedure TfrmEstimaInstalado.tdFechaFinalEnter(Sender: TObject);
begin
  tdFechaFinal.Color := global_Color_entrada
end;



procedure TfrmEstimaInstalado.tdFechaFinalExit(Sender: TObject);
var
  dFechaFinal: tDate;
begin
  tdFechaFinal.Color := global_Color_salida;
  if frmBarra2.btnCancel.Enabled = True then
  begin
    if (MonthOf(tdFechaFinal.Date) <> MonthOf(tdFechaInicial.Date)) then
    begin
      if MonthOf(tdFechaInicial.Date) <= 11 then
        dFechaFinal := strToDate('01/' + Trim(IntToStr(MonthOf(tdFechaInicial.Date) + 1)) + '/' + Trim(IntToStr(YearOf(tdFechaInicial.Date))))
      else
        dFechaFinal := strToDate('01/01/' + Trim(IntToStr(YearOf(tdFechaInicial.Date) + 1)));
      dFechaFinal := dFechaFinal - 1;
      MessageDlg('El generador no puede abarcar un periodo de 2 meses. Periodo Propuesto [' + DateToStr(tdFechaInicial.Date) + ' al ' + DateToStr(dFechaFinal) + ']', mtWarning, [mbOk], 0);
      tdFechaFinal.Date := dFechaFinal;
    end;
    dFechaFinal := strToDate('01/' + Trim(IntToStr(MonthOf(tdFechaInicial.Date))) + '/' + Trim(IntToStr(YearOf(tdFechaInicial.Date))));
    tiSemana.Value := WeeksBetween(tdFechaFinal.date, dFechaFinal) + 1;
  end
end;

procedure TfrmEstimaInstalado.tsNumeroActividadEnter(Sender: TObject);
begin
  tsNumeroActividad.Color := global_color_entrada;
end;

procedure TfrmEstimaInstalado.tdCantidadEnter(Sender: TObject);
begin
  if frmBarra1.btnPost.Enabled = True then
    if (tsNumeroActividad.Text <> '') then
      if OpcButton1 <> 'New' then
        tdCantidad.Color := global_color_entrada;

  if Connection.Configuracion.FieldValues['sGenDesp'] = 'Detallado' then
  begin
    if not tsNumeroActividad.ReadOnly then
    begin
          // Llamar la ventana de captura de cálculos
      try
        Application.CreateForm(TfrmDespieceDX, frmDespieceDX);
        frmDespieceDX.txtInicio.Text := OpcButton1;
        frmDespieceDX.ShowModal;
        lPerimetros := True;
      finally
        frmDespieceDX.Free;
      end;
      isometricoOld := tsIsometrico.Text;
    end;
  end;

end;

procedure TfrmEstimaInstalado.tsIsometricoEnter(Sender: TObject);
begin
  tsIsometrico.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.tmComentariosEnter(Sender: TObject);
begin
  tmComentarios.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.tmComentariosExit(Sender: TObject);
begin
  tmComentarios.Color := global_color_salida
end;

procedure TfrmEstimaInstalado.tmComentariosGeneradorEnter(Sender: TObject);
begin
  tmComentariosGenerador.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.frGeneradorGetValue(const VarName: string;
  var Value: Variant);
var
  sIsometricos: string;
  iValorNumerico: Longint;
  sCadena: string;
  Resultado: Real;
begin

  if CompareText(VarName, 'ISOMETRICOS') = 0 then
  begin
    sIsometricos := '';
    Connection.QryBusca.Active := False;
    Connection.QryBusca.SQL.Clear;
    Connection.QryBusca.SQL.Add('Select distinct sIsometrico, sPrefijo From estimacionxpartida Where sContrato = :Contrato And sNumeroOrden = :Orden And ' +
      'sNumeroGenerador = :Generador And sNumeroActividad = :Actividad And sIsometricoReferencia = :Referencia');
    Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
    Connection.QryBusca.Params.ParamByName('Orden').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Orden').Value := QryGenerador.FieldValues['sNumeroOrden'];
    Connection.QryBusca.Params.ParamByName('Generador').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Generador').Value := QryGenerador.FieldValues['sNumeroGenerador'];
    Connection.QryBusca.Params.ParamByName('Actividad').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Actividad').Value := QryGenerador.FieldValues['sNumeroActividad'];
    Connection.QryBusca.Params.ParamByName('Referencia').DataType := ftString;
    Connection.QryBusca.Params.ParamByName('Referencia').Value := QryGenerador.FieldValues['sIsometricoReferencia'];
    Connection.QryBusca.Open;
    while not Connection.QryBusca.Eof do
    begin
      if sIsometricos <> '' then
        sIsometricos := sIsometricos + ', ';
      sIsometricos := sIsometricos + Connection.QryBusca.FieldValues['sIsometrico'] + ' ' + Connection.QryBusca.FieldValues['sPrefijo'];
      Connection.QryBusca.Next
    end;
    Value := sIsometricos;
  end;


  if CompareText(VarName, 'SUPERINTENDENTE') = 0 then
    Value := sSuperIntendente;
  if CompareText(VarName, 'SUPERVISOR') = 0 then
    Value := sSupervisorGenerador;
  if CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
    Value := sSupervisorTierra;
  if CompareText(VarName, 'SUPERVISOR_RESIDENTE') = 0 then
    Value := sResidente;
  if CompareText(VarName, 'SUPERVISOR_SUBCONTRATISTA') = 0 then
    Value := sSupervisorSubContratista;

  if CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
    Value := sPuestoSuperIntendente;
  if CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
    Value := sPuestoSupervisorGenerador;
  if CompareText(VarName, 'PUESTO_SUPERVISOR_SUBCONTRATISTA') = 0 then
    Value := sPuestoSupervisorSubContratista;
  if CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
    Value := sPuestoSupervisorTierra;
  if CompareText(VarName, 'PUESTO_SUPERVISOR_RESIDENTE') = 0 then
    Value := sPuestoResidente;

  if CompareText(VarName, 'ESTIMACION') = 0 then
    Value := Estimaciones.FieldValues['iNumeroEstimacion'];
  if CompareText(VarName, 'GENERADOR') = 0 then
    Value := Estimaciones.FieldValues['sNumeroGenerador'];
  if CompareText(VarName, 'PERIODO') = 0 then
    Value := DateToStr(connection.EstimacionPeriodo.FieldValues['dFechaInicio']) + ' AL ' + DateToStr(connection.EstimacionPeriodo.FieldValues['dFechaFinal']);

  if CompareText(VarName, 'ORDEN') = 0 then
    Value := Estimaciones.FieldValues['sNumeroOrden'];
  if CompareText(VarName, 'DESCRIPCION_CORTA') = 0 then
    Value := OrdenesdeTrabajo.FieldValues['sDescripcionCorta'];
  if CompareText(VarName, 'DESCRIPCION') = 0 then
    Value := OrdenesdeTrabajo.FieldValues['mDescripcion'];

  if CompareText(VarName, 'PERGENOPT') = 0 then
    Value := sDiarioPeriodo;

  if CompareText(VarName, 'TIPOACUM') = 0 then
    Value := sTipoAcumulado;

  if CompareText(VarName, 'DESCRIPCION_PDA') = 0 then
    Value := sDescripcionPda;


end;

procedure TfrmEstimaInstalado.tiNumeroEstimacionEnter(Sender: TObject);
begin
  tiNumeroEstimacion.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.Grid_GeneradoresEnter(Sender: TObject);
begin
  if frmBarra2.btnCancel.Enabled = True then
    frmBarra2.btnCancel.Click;
  if Estimaciones.RecordCount = 0 then
  begin
    EstimacionxPartida.Active := False;
    tdFechaInicial.Date := Date;
    tdFechaFinal.Date := Date;
    tsNumeroGenerador.Text := '';
    tsFaseObra.Text := '';
    tmComentariosGenerador.Text := '';
    tiNumeroEstimacion.KeyValue := '';
    tiConsecutivo.Text := '0';
    tdCantidad.Value := 0;
    tsIsometrico.text:= '';
    tsPrefijo.Text := '';
    tsIsometricoReferencia.KeyValue := '';
    tmComentarios.Text := '';
  end
end;

procedure TfrmEstimaInstalado.MenuItem9Click(Sender: TObject);
begin
  frmBarra2.btnPrinter.Click
end;

procedure TfrmEstimaInstalado.Salir2Click(Sender: TObject);
begin
  frmBarra2.btnExit.Click
end;

procedure TfrmEstimaInstalado.NumerosGeneradores1Click(Sender: TObject);
begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
        procNumeroGenerador(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, 'Cliente', frmEstimaInstalado, frGenerador.OnGetValue, True)
      else
        MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al salvar nuevo registro', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.SemanalSImportes1Click(Sender: TObject);
begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
        procSemanalSinConImportes(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, 'Sin Importes', frmEstimaInstalado, frGenerador.OnGetValue, True)
      else
        MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
    end
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al click en Resumen Semanal Numeros Generadores en el menu', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.EstimacionesAfterScroll(DataSet: TDataSet);
var
  iCheck: Byte;
  sDescompone: string;
  sFase: string;
begin
  if frmBarra2.btnCancel.Enabled = True then
    frmBarra2.btnCancel.Click;

  if frmBarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;

  for iCheck := 0 to tiFases.Items.Count - 1 do
    tiFases.Checked[iCheck] := False;

  tmComentariosGenerador.Color := global_color_salida;
  if Estimaciones.RecordCount > 0 then
  begin
    mComentarios := '';
    tdFechaInicial.Date := Estimaciones.FieldValues['dFechaInicio'];
    tdFechaFinal.Date := Estimaciones.FieldValues['dFechaFinal'];
    tsNumeroGenerador.Text := Estimaciones.FieldValues['sNumeroGenerador'];
    tsAnexo.Text := Estimaciones.FieldValues['sNumeroAnexo'];
    tmComentariosGenerador.Text := Estimaciones.FieldValues['mComentarios'];
    tiNumeroEstimacion.KeyValue := Estimaciones.FieldValues['iNumeroEstimacion'];
    tiConsecutivo.Text := Estimaciones.FieldValues['iConsecutivo'];
    tsFaseObra.Text := Estimaciones.FieldValues['sFaseObra'];
    sDescompone := tsFaseObra.Text;
    while sDescompone <> '' do
    begin
      if Pos('-', sDescompone) > 0 then
      begin
        sFase := MidStr(sDescompone, 1, Pos('-', sDescompone) - 1);
        sDescompone := MidStr(sDescompone, Pos('-', sDescompone) + 1, Length(sDescompone));
      end
      else
      begin
        sFase := sDescompone;
        sDescompone := ''
      end;
      iCheck := 0;
      while (iCheck < tiFases.Items.Count) and (sFase <> '') do
      begin
        if tiFases.Items.Strings[iCheck] = sFase then
        begin
          tiFases.Checked[iCheck] := True;
          sFase := '';
        end;
        iCheck := iCheck + 1;
      end;
    end;

    tdCantidad.Value  := 0;
    tsIsometrico.text := '';
    tsPrefijo.Text := '';
    tsIsometricoReferencia.KeyValue := '';
    tmComentarios.Text := '';
    txtTag.Text := '';

    EstimacionxPartida.Active := False;
    EstimacionxPartida.Params.ParamByName('Contrato').DataType := ftString;
    EstimacionxPartida.Params.ParamByName('Contrato').Value := Global_Contrato;
    EstimacionxPartida.Params.ParamByName('Convenio').DataType := ftString;
    EstimacionxPartida.Params.ParamByName('Convenio').Value := Global_Convenio;
    EstimacionxPartida.Params.ParamByName('Orden').DataType := ftString;
    EstimacionxPartida.Params.ParamByName('Orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
    EstimacionxPartida.Params.ParamByName('Generador').DataType := ftString;
    EstimacionxPartida.Params.ParamByName('Generador').Value := Estimaciones.FieldValues['sNumeroGenerador'];
    EstimacionxPartida.Params.ParamByName('Ordenado').DataType := ftString;
    EstimacionxPartida.Params.ParamByName('Ordenado').Value := 'sWbsContrato, iItemOrden';
    EstimacionxPartida.Open;
    //tdCantidad.Value := EstimacionxPartida.FieldValues['dGenerado'] ;
    if EstimacionxPartida.RecordCount = 0 then
    begin
      AnexoConvenio.Active := False;
      AnexoConvenio.Params.ParamByName('Contrato').DataType := ftString;
      AnexoConvenio.Params.ParamByName('Contrato').Value := global_contrato;
      AnexoConvenio.Params.ParamByName('Actividad').DataType := ftString;
      AnexoConvenio.Params.ParamByName('Actividad').Value := '';
      AnexoConvenio.Open;
    end;
    QryPrefijos.Active := False;
    QryPrefijos.Params.ParamByName('Contrato').DataType := ftString;
    QryPrefijos.Params.ParamByName('Contrato').Value := global_contrato;
    QryPrefijos.Params.ParamByName('Isometrico').DataType := ftString;
    QryPrefijos.Params.ParamByName('Isometrico').Value := tsIsometrico.Text;
    QryPrefijos.Params.ParamByName('Ordenado').DataType := ftString;
    QryPrefijos.Params.ParamByName('Ordenado').Value := 'sPrefijo';
    QryPrefijos.Open;
  end
  else
  begin
    EstimacionxPartida.Active := False;
    tdFechaInicial.Date := Date;
    tdFechaFinal.Date := Date;
    tsNumeroGenerador.Text := '';
    tmComentariosGenerador.Text := '';
    tsFaseObra.Text := '';
    tiNumeroEstimacion.KeyValue := '';
    tiConsecutivo.Text := '0';
    tdCantidad.Value  := 0;
    tsIsometrico.text := '';
    tsPrefijo.Text := '';
    tsIsometricoReferencia.KeyValue := '';
    tmComentarios.Text := '';
    txtTag.Text := '';
    dFecha.Date := Date;
  end
end;

procedure TfrmEstimaInstalado.EstimacionxPartidaAfterScroll(
  DataSet: TDataSet);
begin
  ImgNotas.Visible := False;
  try
    if EstimacionxPartida.RecordCount > 0 then
    begin
      if EstimacionxPartida.FieldValues['iOrdenCambio'] > 0 then
        tiOrdenCambio.ItemIndex := tiOrdenCambio.Items.IndexOf('O.C. No. [' + IntToStr(EstimacionxPartida.FieldValues['iOrdenCambio']) + ']')
      else
        tiOrdenCambio.ItemIndex := 0;
      tsNumeroActividad.KeyValue := EstimacionxPartida.fieldByName('sNumeroActividad').AsString;
      tdCantidad.Value  := EstimacionxPartida.FieldValues['dGenerado'];
      tsIsometrico.text := EstimacionxPartida.FieldValues['sIsometrico'];
      tsPrefijo.Text := EstimacionxPartida.FieldValues['sPrefijo'];
      tsIsometricoReferencia.KeyValue := EstimacionxPartida.FieldValues['sIsometricoReferencia'];
      if EstimacionxPartida.FieldValues['sInstalacion'] = null then
        tsInstalacion.Text := ''
      else
        tsInstalacion.Text := EstimacionxPartida.FieldValues['sInstalacion'];
      txtTag.Text := EstimacionxPartida.FieldValues['sTag'];
      dFecha.Date := EstimacionxPartida.FieldValues['dFecha'];

      txtMarca.Text := EstimacionxPartida.FieldValues['sMarcaRev'];
      txtSubMarca.Text := EstimacionxPartida.FieldValues['sSubMca']; ;
      txtLongArea.Text := EstimacionxPartida.FieldValues['sLongArea'];
      txtPzas.text := EstimacionxPartida.FieldValues['nPiezas']; ;
      txtLongTotal.Text := EstimacionxPartida.FieldValues['sLongAreaTotal'];
      txtPesoxUnidad.Text := EstimacionxPartida.FieldValues['sPesoxUnidad'];
      txtPesoTotal.Text := EstimacionxPartida.FieldValues['dPesoTotal'];
      tmComentarios.Text := EstimacionxPartida.FieldValues['mComentarios'];
      if EstimacionxPartida.FieldValues['lEstima'] = 'Si' then
        tlEstima.State := cbChecked
      else
        tlEstima.State := cbUnChecked;

      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('Select sNumeroActividad From comentariosxanexo Where sContrato = :Contrato And sNumeroActividad = :Actividad');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca.Params.ParamByName('actividad').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
      Connection.QryBusca.Open;
      if Connection.QryBusca.RecordCount > 0 then
        imgNotas.Visible := True;

      AnexoConvenio.Active := False;
      AnexoConvenio.Params.ParamByName('Contrato').DataType := ftString;
      AnexoConvenio.Params.ParamByName('Contrato').Value := global_contrato;
      AnexoConvenio.Params.ParamByName('Actividad').DataType := ftString;
      AnexoConvenio.Params.ParamByName('Actividad').Value := tsNumeroActividad.Text;
      AnexoConvenio.Open;

      if Pos('WBS', global_checkGenerador) > 0 then
      begin
        ActividadesIguales.Active := False;
        ActividadesIguales.SQL.Clear;
        if (Global_Optativa = 'PROGRAMADA') or (Global_Optativa = 'MIXTA') then
          ActividadesIguales.SQL.Add('SELECT a.swbscontrato,a.sWbsAnterior, a.sWbs, a.sNumeroActividad, a.mDescripcion, a.dCantidad, (a.dInstalado + a.dExcedente) as dInstalado, a.dPonderado, ' +
            'a.sMedida, a.dExcedente  FROM actividadesxorden a ' +
            'WHERE a.sContrato = :Contrato And a.sNumeroOrden = :orden And a.sIdConvenio = :convenio ' +
            'And a.sNumeroActividad = :Actividad And a.sTipoActividad = "Actividad" Order By a.iItemOrden')
        else
          ActividadesIguales.SQL.Add('SELECT a.swbs as swbscontrato,a.sWbsAnterior, a.sWbs, a.sNumeroActividad, a.mDescripcion, a.dCantidadAnexo As dCantidad, (a.dInstalado + a.dExcedente) as dInstalado, a.dPonderado, ' +
            'a.sMedida, a.dExcedente  FROM actividadesxanexo a ' +
            'WHERE a.sContrato = :Contrato And a.sIdConvenio = :convenio And a.sMedida <> "ACTIVIDAD" ' +
            'And a.sNumeroActividad = :Actividad And a.sTipoActividad = "Actividad" Order By a.iItemOrden');
        ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
        ActividadesIguales.Params.ParamByName('contrato').Value := global_contrato;
        ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
        ActividadesIguales.Params.ParamByName('Convenio').Value := global_convenio;
        if (Global_Optativa = 'PROGRAMADA') or (Global_Optativa = 'MIXTA') then
        begin
          ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
          ActividadesIguales.Params.ParamByName('orden').Value := tsNumeroOrden.KeyValue;
        end;
        ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
        ActividadesIguales.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
        ActividadesIguales.Open;
        ActividadesIguales.Locate('sWbs', EstimacionxPartida.FieldValues['sWbs'], [loPartialKey]);

        Paquete.Active := False;
        Paquete.Params.ParamByName('contrato').DataType := ftString;
        Paquete.Params.ParamByName('contrato').Value := global_contrato;
        Paquete.Params.ParamByName('orden').DataType := ftString;
        Paquete.Params.ParamByName('orden').Value := tsNumeroOrden.KeyValue;
        Paquete.Params.ParamByName('wbs').DataType := ftString;
        Paquete.Params.ParamByName('wbs').Value := ActividadesIguales.FieldValues['sWbsAnterior'];
        Paquete.Open;
        if Paquete.RecordCount > 0 then
        begin
          pdPaquete.Caption := Paquete.FieldValues['mDescripcion'];
          pdPaquete.Hint := Paquete.FieldValues['mDescripcion'];
        end
        else
        begin
          pdPaquete.Caption := '< < Seleccione un Paquete > >';
          pdPaquete.Hint := '< < Seleccione un Paquete > >';
        end
      end
        // Calculo de Instalado, Estimado y Diferencia
    end
    else
    begin
      mObra.Text := '';
      tdCantidad.Value := 0;
      tsIsometrico.text := '';
      tsPrefijo.Text := '';
      tsIsometricoReferencia.KeyValue := '';
      tmComentarios.Text := '';

      AnexoConvenio.Active := False;
      AnexoConvenio.Params.ParamByName('Contrato').DataType := ftString;
      AnexoConvenio.Params.ParamByName('Contrato').Value := global_contrato;
      AnexoConvenio.Params.ParamByName('Actividad').DataType := ftString;
      AnexoConvenio.Params.ParamByName('Actividad').Value := '';
      AnexoConvenio.Open;
    end;
  except

  end
end;

procedure TfrmEstimaInstalado.tsIsometricoReferenciaKeyPress(
  Sender: TObject; var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
    if Connection.Configuracion.FieldValues['sGenDesp'] <> 'Despiezado' then
      tmComentarios.SetFocus
    else
      txtMarca.SetFocus;
end;

procedure TfrmEstimaInstalado.tsIsometricoReferenciaEnter(Sender: TObject);
begin
  tsIsometricoReferencia.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.tsIsometricoReferenciaExit(Sender: TObject);
begin
  tsIsometricoReferencia.Color := global_color_salida;
end;

procedure TfrmEstimaInstalado.GridPartidasCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;
end;

procedure TfrmEstimaInstalado.GridPartidasDblClick(Sender: TObject);
begin
  actualizadatos1.Click
end;

procedure TfrmEstimaInstalado.GridPartidasTitleBtnClick(Sender: TObject;
  ACol: Integer; Field: TField);
var
  sCampo: string;
begin
 //SE COMENTAN ESTA LINEAS YA QUE AL HACER DOBE CLICK DESAPARECEN DATOS...
 { sCampo := Field.FieldName;
  EstimacionxPartida.Active := False;
  EstimacionxPartida.SQL.Clear;
  if sTipoOrden = 'DESC' then
  begin
    if Pos( 'WBS', global_checkGenerador ) > 0 then
      EstimacionxPartida.SQL.Add( 'Select a.iItemOrden, e1.sWbs, e1.sNumeroActividad, e1.sIsometrico, e1.sPrefijo, e1.dCantidad as dGenerado, e1.sIsometricoReferencia, ' +
        'e1.nPiezas, e1.sMarcaRev, e1.sSubMca, e1.sLongArea, e1.sLongAreaTotal, e1.sPesoxUnidad, e1.dPesoTotal, ' +
        'e1.sInstalacion, e1.mComentarios, e1.lEstima, e1.iOrdenCambio, a.sMedida, a.mDescripcion, a.dVentaMN from estimacionxpartida e1 ' +
        'inner join actividadesxorden a on (a.sContrato = e1.sContrato and a.sIdConvenio = :Convenio and a.sNumeroOrden = e1.sNumeroOrden and ' +
        'a.sWbs = e1.sWbs and a.sNumeroActividad = e1.sNumeroActividad And a.sTipoActividad = "Actividad") ' +
        'Where e1.sContrato = :Contrato And e1.sNumeroOrden = :Orden And e1.sNumeroGenerador = :Generador Order By :Ordenado ASC' )
    else
      EstimacionxPartida.SQL.Add( 'Select a.iItemOrden, e1.sWbs, e1.sNumeroActividad, e1.sIsometrico, e1.sPrefijo, e1.dCantidad as dGenerado, e1.sIsometricoReferencia, ' +
        'e1.nPiezas, e1.sMarcaRev, e1.sSubMca, e1.sLongArea, e1.sLongAreaTotal, e1.sPesoxUnidad, e1.dPesoTotal, ' +
        'e1.sInstalacion, e1.mComentarios, e1.lEstima, e1.iOrdenCambio, a.sMedida, a.mDescripcion, a.dVentaMN from estimacionxpartida e1 ' +
        'inner join actividadesxanexo a on (a.sContrato = e1.sContrato and a.sIdConvenio = :Convenio and a.sNumeroActividad = e1.sNumeroActividad And a.sTipoActividad = "Actividad") ' +
        'Where e1.sContrato = :Contrato And e1.sNumeroOrden = :Orden And e1.sNumeroGenerador = :Generador Order By :Ordenado ASC' );
    sTipoOrden := 'ASC'
  end
  else
  begin
    if Pos( 'WBS', global_checkGenerador ) > 0 then
      EstimacionxPartida.SQL.Add( 'Select a.iItemOrden, e1.sWbs, e1.sNumeroActividad, e1.sIsometrico, e1.sPrefijo, e1.dCantidad as dGenerado, e1.sIsometricoReferencia, ' +
        'e1.nPiezas, e1.sMarcaRev, e1.sSubMca, e1.sLongArea, e1.sLongAreaTotal, e1.sPesoxUnidad, e1.dPesoTotal, ' +
        'e1.sInstalacion, e1.mComentarios, e1.lEstima, e1.iOrdenCambio, a.sMedida, a.mDescripcion, a.dVentaMN from estimacionxpartida e1 ' +
        'inner join actividadesxorden a on (a.sContrato = e1.sContrato and a.sIdConvenio = :Convenio and a.sNumeroOrden = e1.sNumeroOrden and ' +
        'a.sWbs = e1.sWbs and a.sNumeroActividad = e1.sNumeroActividad And a.sTipoActividad = "Actividad") ' +
        'Where e1.sContrato = :Contrato And e1.sNumeroOrden = :Orden And e1.sNumeroGenerador = :Generador Order By :Ordenado DESC' )
    else
      EstimacionxPartida.SQL.Add( 'Select a.iItemOrden, e1.sWbs, e1.sNumeroActividad, e1.sIsometrico, e1.sPrefijo, e1.dCantidad as dGenerado, e1.sIsometricoReferencia, ' +
        'e1.nPiezas, e1.sMarcaRev, e1.sSubMca, e1.sLongArea, e1.sLongAreaTotal, e1.sPesoxUnidad, e1.dPesoTotal, ' +
        'e1.sInstalacion, e1.mComentarios, e1.lEstima, e1.iOrdenCambio, a.sMedida, a.mDescripcion, a.dVentaMN from estimacionxpartida e1 ' +
        'inner join actividadesxanexo a on (a.sContrato = e1.sContrato and a.sIdConvenio = :Convenio and a.sNumeroActividad = e1.sNumeroActividad And a.sTipoActividad = "Actividad") ' +
        'Where e1.sContrato = :Contrato And e1.sNumeroOrden = :Orden And e1.sNumeroGenerador = :Generador Order By :Ordenado DESC' );
    sTipoOrden := 'DESC'
  end;
  EstimacionxPartida.Params.ParamByName( 'Contrato' ).DataType := ftString;
  EstimacionxPartida.Params.ParamByName( 'Contrato' ).Value := Global_Contrato;
  EstimacionxPartida.Params.ParamByName( 'Convenio' ).DataType := ftString;
  EstimacionxPartida.Params.ParamByName( 'Convenio' ).Value := Global_Convenio;
  EstimacionxPartida.Params.ParamByName( 'Orden' ).DataType := ftString;
  EstimacionxPartida.Params.ParamByName( 'Orden' ).Value := Estimaciones.FieldValues[ 'sNumeroOrden' ];
  EstimacionxPartida.Params.ParamByName( 'Generador' ).DataType := ftString;
  EstimacionxPartida.Params.ParamByName( 'Generador' ).Value := Estimaciones.FieldValues[ 'sNumeroGenerador' ];
  EstimacionxPartida.Params.ParamByName( 'Ordenado' ).DataType := ftString;
  EstimacionxPartida.Params.ParamByName( 'Ordenado' ).Value := sCampo;
  EstimacionxPartida.Open;     }
end;

procedure TfrmEstimaInstalado.FormActivate(Sender: TObject);
var
  iPosicion: Integer;
begin
  if lIniciado then
    tsBaseGeneracion.Caption := 'BASE DE GENERACIÓN: ' + connection.configuracion.FieldValues['sBaseGeneracion'];
  lIniciado := True;
end;

procedure TfrmEstimaInstalado.GridPartidasEnter(Sender: TObject);
begin
  if gridpartidas.DataSource.DataSet.IsEmpty = false then
  begin
    if frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click;

    ImgNotas.Visible := False;
    if tsNumeroOrden.Text <> '' then
      if EstimacionxPartida.RecordCount > 0 then
      begin
        if EstimacionxPartida.FieldValues['iOrdenCambio'] > 0 then
          tiOrdenCambio.ItemIndex := tiOrdenCambio.Items.IndexOf('O.C. No. [' + IntToStr(EstimacionxPartida.FieldValues['iOrdenCambio']) + ']')
        else
          tiOrdenCambio.ItemIndex := 0;

        tsNumeroActividad.KeyValue := EstimacionxPartida.fieldByName('sNumeroActividad').AsString;
        tdCantidad.Value := EstimacionxPartida.FieldValues['dGenerado'];
        tsIsometrico.text := EstimacionxPartida.FieldValues['sIsometrico'];
        tsPrefijo.Text := EstimacionxPartida.FieldValues['sPrefijo'];
        tsIsometricoReferencia.KeyValue := EstimacionxPartida.FieldValues['sIsometricoReferencia'];
        tsInstalacion.Text := EstimacionxPartida.FieldByName('sInstalacion').AsString;
        tmComentarios.Text := EstimacionxPartida.FieldValues['mComentarios'];
        txtTag.Text := EstimacionxPartida.FieldValues['sTag'];

        txtPzas.Text := EstimacionxPartida.FieldValues['nPiezas'];
        txtMarca.Text := EstimacionxPartida.FieldValues['sMarcaRev'];
        txtSubMarca.Text := EstimacionxPartida.FieldValues['sSubMca']; ;
        txtLongArea.Text := EstimacionxPartida.FieldValues['sLongArea']; ;
        txtLongTotal.Text := EstimacionxPartida.FieldValues['sLongAreaTotal'];
        txtPesoxUnidad.Text := EstimacionxPartida.FieldValues['sPesoxUnidad'];
        txtPesoTotal.Text := EstimacionxPartida.FieldValues['dPesoTotal'];
        if EstimacionxPartida.FieldValues['lEstima'] = 'Si' then
          tlEstima.State := cbChecked
        else
          tlEstima.State := cbUnChecked;

        AnexoConvenio.Active := False;
        AnexoConvenio.Params.ParamByName('Contrato').DataType := ftString;
        AnexoConvenio.Params.ParamByName('Contrato').Value := global_contrato;
        AnexoConvenio.Params.ParamByName('Actividad').DataType := ftString;
        AnexoConvenio.Params.ParamByName('Actividad').Value := '';
        AnexoConvenio.Open;

        if Pos('WBS', global_checkGenerador) > 0 then
        begin
          ActividadesIguales.Active := False;
          ActividadesIguales.SQL.Clear;
          if (Global_Optativa = 'PROGRAMADA') or (Global_Optativa = 'MIXTA') then
            ActividadesIguales.SQL.Add('SELECT a.sWbsContrato, a.sWbsAnterior, a.sWbs, a.sNumeroActividad, a.mDescripcion, a.dCantidad, (a.dInstalado + a.dExcedente) as dInstalado, a.dPonderado, ' +
              'a.sMedida, a.dExcedente  FROM actividadesxorden a ' +
              'WHERE a.sContrato = :Contrato And a.sNumeroOrden = :orden And a.sIdConvenio = :convenio ' +
              'And a.sNumeroActividad = :Actividad And a.sTipoActividad = "Actividad" Order By a.iItemOrden')
          else
            ActividadesIguales.SQL.Add('SELECT a.sWbs as sWbsContrato, a.sWbsAnterior, a.sWbs, a.sNumeroActividad, a.mDescripcion, a.dCantidadAnexo As dCantidad, (a.dInstalado + a.dExcedente) as dInstalado, a.dPonderado, ' +
              'a.sMedida, a.dExcedente  FROM actividadesxanexo a ' +
              'WHERE a.sContrato = :Contrato And a.sIdConvenio = :convenio ' +
              'And a.sNumeroActividad = :Actividad And a.sTipoActividad = "Actividad" Order By a.iItemOrden');

          ActividadesIguales.Params.ParamByName('contrato').DataType := ftString;
          ActividadesIguales.Params.ParamByName('contrato').Value := global_contrato;
          ActividadesIguales.Params.ParamByName('convenio').DataType := ftString;
          if convenio_reporte = '' then
            ActividadesIguales.Params.ParamByName('Convenio').Value := global_convenio
          else
            ActividadesIguales.Params.ParamByName('Convenio').Value := convenio_reporte;
          if (Global_Optativa = 'PROGRAMADA') or (Global_Optativa = 'MIXTA') then
          begin
            ActividadesIguales.Params.ParamByName('orden').DataType := ftString;
            ActividadesIguales.Params.ParamByName('orden').Value := tsNumeroOrden.KeyValue;
          end;
          ActividadesIguales.Params.ParamByName('actividad').DataType := ftString;
          ActividadesIguales.Params.ParamByName('actividad').Value := tsNumeroActividad.Text;
          ActividadesIguales.Open;
          ActividadesIguales.Locate('sWbs', EstimacionxPartida.FieldValues['sWbs'], [loPartialKey]);

          Paquete.Active := False;
          Paquete.Params.ParamByName('contrato').DataType := ftString;
          Paquete.Params.ParamByName('contrato').Value := global_contrato;
          Paquete.Params.ParamByName('orden').DataType := ftString;
          Paquete.Params.ParamByName('orden').Value := tsNumeroOrden.KeyValue;
          Paquete.Params.ParamByName('wbs').DataType := ftString;
          if EstimacionxPartida.RecordCount > 0 then
            Paquete.Params.ParamByName('wbs').Value := ActividadesIguales.FieldValues['sWbsAnterior']
          else
            Paquete.Params.ParamByName('wbs').Value := '';
          Paquete.Open;
          if Paquete.RecordCount > 0 then
          begin
            pdPaquete.Caption := Paquete.FieldValues['mDescripcion'];
            pdPaquete.Hint := Paquete.FieldValues['mDescripcion'];
          end
          else
          begin
            pdPaquete.Caption := '< < Seleccione un Paquete > >';
            pdPaquete.Hint := '< < Seleccione un Paquete > >';
          end
        end;

        QryPrefijos.Active := False;
        QryPrefijos.Params.ParamByName('Contrato').DataType := ftString;
        QryPrefijos.Params.ParamByName('Contrato').Value := global_contrato;
        QryPrefijos.Params.ParamByName('Isometrico').DataType := ftString;
        QryPrefijos.Params.ParamByName('Isometrico').Value := tsIsometrico.Text;
        QryPrefijos.Params.ParamByName('Ordenado').DataType := ftString;
        QryPrefijos.Params.ParamByName('Ordenado').Value := 'sPrefijo';
        QryPrefijos.Open;
      end;

  end;

end;

procedure TfrmEstimaInstalado.frmBarra2btnAddClick(Sender: TObject);
var
  dFechaFinal: tDate;
  iCheck: Integer;
begin
  if tsNumeroOrden.Text <> '' then
  begin
    frmBarra2.btnCancel.Click;
    mComentarios := '';
    tdFechaInicial.Enabled := True;
    tdFechaFinal.Enabled := True;
    tsFaseObra.ReadOnly := False;
    tiFases.Enabled := True;

    tsNumeroGenerador.ReadOnly := False;
    tmComentariosGenerador.ReadOnly := False;
    tiNumeroEstimacion.ReadOnly := False;
    if Estimaciones.RecordCount > 0 then
    begin
      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('Select Max(iConsecutivo) as iConsecutivo From estimaciones Where sContrato = :Contrato Group By sContrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca.Open;
      if Connection.QryBusca.RecordCount > 0 then
        tiConsecutivo.Text := Connection.QryBusca.FieldValues['iConsecutivo'] + 1
      else
        tiConsecutivo.Text := '1';

      tiNumeroEstimacion.KeyValue := Estimaciones.FieldValues['iNumeroEstimacion'];
      tdFechaInicial.Date := Estimaciones.FieldValues['dFechaFinal'] + 1;
      if connection.configuracion.FieldValues['sRangoEstimacion'] = 'Diario' then
        tdFechaFinal.Date := Estimaciones.FieldValues['dFechaFinal'] + 1
      else if connection.configuracion.FieldValues['sRangoEstimacion'] = 'Semanal' then
        tdFechaFinal.Date := Estimaciones.FieldValues['dFechaFinal'] + 7
      else if connection.configuracion.FieldValues['sRangoEstimacion'] = 'Mensual' then
        tdFechaFinal.Date := Estimaciones.FieldValues['dFechaFinal'] + 30;

    end
    else
    begin
      tdFechaInicial.Date := Date;
      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Clear;
      Connection.QryBusca.SQL.Add('Select Max(iConsecutivo) as iConsecutivo From estimaciones Where sContrato = :Contrato Group By sContrato');
      Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato;
      Connection.QryBusca.Open;
      if Connection.QryBusca.RecordCount > 0 then
        tiConsecutivo.Text := Connection.QryBusca.FieldValues['iConsecutivo'] + 1
      else
        tiConsecutivo.Text := '1';
      if connection.configuracion.FieldValues['sRangoEstimacion'] = 'Diario' then
        tdFechaFinal.Date := Date + 1
      else if connection.configuracion.FieldValues['sRangoEstimacion'] = 'Semanal' then
        tdFechaFinal.Date := Date + 7
      else if connection.configuracion.FieldValues['sRangoEstimacion'] = 'Mensual' then
        tdFechaFinal.Date := Date + 30
    end;

    if (MonthOf(tdFechaFinal.Date) <> MonthOf(tdFechaInicial.Date)) then
    begin
      if MonthOf(tdFechaInicial.Date) <= 11 then
        dFechaFinal := strToDate('01/' + Trim(IntToStr(MonthOf(tdFechaInicial.Date) + 1)) + '/' + Trim(IntToStr(YearOf(tdFechaInicial.Date))))
      else
        dFechaFinal := strToDate('01/01/' + Trim(IntToStr(YearOf(tdFechaInicial.Date) + 1)));
      dFechaFinal := dFechaFinal - 1;
      MessageDlg('El generador no puede abarcar un periodo de 2 meses. Periodo Propuesto [' + DateToStr(tdFechaInicial.Date) + ' al ' + DateToStr(dFechaFinal) + ']', mtWarning, [mbOk], 0);
      tdFechaFinal.Date := dFechaFinal;
    end;
    dFechaFinal := strToDate('01/' + Trim(IntToStr(MonthOf(tdFechaInicial.Date))) + '/' + Trim(IntToStr(YearOf(tdFechaInicial.Date))));
    tiSemana.Value := WeeksBetween(tdFechaFinal.date, dFechaFinal) + 1;

    for iCheck := 0 to tiFases.Items.Count - 1 do
      tiFases.Checked[iCheck] := False;

    OpcButton1 := 'New';

    frmBarra2.btnAddClick(Sender);
    EstimacionxPartida.Active := False;
    tsNumeroGenerador.Text := '';
    tmComentariosGenerador.Text := '';
    tsFaseObra.Text := '';
    tdCantidad.Value := 0;
    tsIsometrico.text := '';
    tsPrefijo.Text := '';
    tsIsometricoReferencia.KeyValue := '';
    tmComentarios.Text := '';
    pgControl.ActivePageIndex := 0;
    tsNumeroGenerador.SetFocus
  end;
  activapop(frmEstimaInstalado, popupprincipal);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmEstimaInstalado.frmBarra2btnEditClick(Sender: TObject);
begin
  if tsNumeroOrden.Text <> '' then
    if Estimaciones.RecordCount > 0 then
      if Estimaciones.FieldValues['lStatus'] = 'Pendiente' then
      begin
        frmBarra2.btnCancel.Click;
        OpcButton1 := 'False';
        frmBarra2.btnEditClick(Sender);

        if global_grupo <> 'INTEL-CODE' then
          if EstimacionxPartida.RecordCount = 0 then
          begin
            tdFechaInicial.Enabled := True;
            tdFechaFinal.Enabled := True;
          end
          else
            MessageDlg('Existen partidas registradas en el generador, no podra modificar el periodo de generación.', mtWarning, [mbOk], 0)
        else
        begin
          tdFechaInicial.Enabled := True;
          tdFechaFinal.Enabled := True;
        end;
        tsFaseObra.ReadOnly := False;
        tiFases.Enabled := True;
        tsNumeroGenerador.ReadOnly := False;
        tmComentariosGenerador.ReadOnly := False;
        tiNumeroEstimacion.ReadOnly := False;
        pgControl.ActivePageIndex := 0;
        tsNumeroGenerador.SetFocus;
        activapop(frmEstimaInstalado, popupprincipal)
      end
      else
        MessageDlg('Generador Aplicado no se pueden realizar cambios', mtWarning, [mbOk], 0);

  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmEstimaInstalado.frmBarra2btnPostClick(Sender: TObject);
var
  Posicion: Integer;
  iCheck: Byte;
  dFechaFinal: tDate;
  nombres, cadenas: TStringList;
begin
  {Validacion de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('Estimacion'); nombres.Add('Numero Anexo'); nombres.Add('Tipo Moneda');
  cadenas.Add(tiNumeroEstimacion.Text); cadenas.Add(tsAnexo.Text); cadenas.Add(txtTipoMoneda.Text);
  if not validaTexto(nombres, cadenas, 'Generador', tsNumeroGenerador.Text) then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
  {Continua insercion de datos}
  // Fases Igual a ""
  tsFaseObra.Text := '';
  for iCheck := 0 to tiFases.Items.Count - 1 do
    if tiFases.Checked[iCheck] = True then
    begin
      if tsFaseObra.Text <> '' then
        tsFaseObra.Text := tsFaseObra.Text + '-';
      tsFaseObra.Text := tsFaseObra.Text + tiFases.Items.Strings[iCheck];
      desactivapop(popupprincipal);
    end;

  if Connection.configuracion.FieldValues['sCampPerf'] = 'No' then
  begin
    // se determina la fecha final del generador y la semana segun la fecha final...
    if (MonthOf(tdFechaFinal.Date) <> MonthOf(tdFechaInicial.Date)) then
    begin
      if MonthOf(tdFechaInicial.Date) <= 11 then
        dFechaFinal := strToDate('01/' + Trim(IntToStr(MonthOf(tdFechaInicial.Date) + 1)) + '/' + Trim(IntToStr(YearOf(tdFechaInicial.Date))))
      else
        dFechaFinal := strToDate('01/01/' + Trim(IntToStr(YearOf(tdFechaInicial.Date) + 1)));
      dFechaFinal := dFechaFinal - 1;
      MessageDlg('El generador no puede abarcar un periodo de 2 meses. Periodo Propuesto [' + DateToStr(tdFechaInicial.Date) + ' al ' + DateToStr(dFechaFinal) + ']', mtWarning, [mbOk], 0);
      tdFechaFinal.Date := dFechaFinal;
    end;
  end;
  dFechaFinal := strToDate('01/' + Trim(IntToStr(MonthOf(tdFechaInicial.Date))) + '/' + Trim(IntToStr(YearOf(tdFechaInicial.Date))));
  tiSemana.Value := WeeksBetween(tdFechaFinal.date, dFechaFinal) + 1;
  if tmComentariosGenerador.Text = '' then
    tmComentariosGenerador.Text := ' ';

  if OpcButton1 = 'New' then
  begin
    try
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('INSERT INTO estimaciones ( sContrato , sNumeroOrden, iNumeroEstimacion, sNumeroGenerador, ' +
        'iSemana, iConsecutivo, dFechaInicio, dFechaFinal, dBitacoraInicio, dBitacoraFinal, sFaseObra, lStatus, mComentarios, ' +
        'sIdUsuario, sNumeroAnexo, sTipoMoneda ) ' +
        'VALUES (:Contrato, :Orden, :EStimacion, :Generador, :Semana, :Consecutivo, :FechaI, :FechaF, ' +
        ':BitacoraI, :BitacoraF, :Fase, :Status, :Comentarios, :Usuario, :Anexo, :TipoMoneda)');
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
      connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
      connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
      connection.zCommand.Params.ParamByName('Estimacion').DataType := ftString;
      connection.zCommand.Params.ParamByName('Estimacion').value := tiNumeroEstimacion.KeyValue;
      connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
      connection.zCommand.Params.ParamByName('Generador').value := tsNumeroGenerador.Text;
      connection.zCommand.Params.ParamByName('Semana').DataType := ftInteger;
      connection.zCommand.Params.ParamByName('Semana').value := tiSemana.Value;
      connection.zCommand.Params.ParamByName('Consecutivo').DataType := ftInteger;
      connection.zCommand.Params.ParamByName('Consecutivo').value := tiConsecutivo.Text;
      connection.zCommand.Params.ParamByName('FechaI').DataType := ftDate;
      connection.zCommand.Params.ParamByName('FechaI').value := tdFechaInicial.Date;
      connection.zCommand.Params.ParamByName('FechaF').DataType := ftDate;
      connection.zCommand.Params.ParamByName('FechaF').value := tdFechaFinal.Date;
      connection.zCommand.Params.ParamByName('BitacoraI').DataType := ftDate;
      connection.zCommand.Params.ParamByName('BitacoraI').value := Date;
      connection.zCommand.Params.ParamByName('BitacoraF').DataType := ftDate;
      connection.zCommand.Params.ParamByName('BitacoraF').value := Date;
      connection.zCommand.Params.ParamByName('Fase').DataType := ftString;
      connection.zCommand.Params.ParamByName('Fase').value := tsFaseObra.Text;
      connection.zCommand.Params.ParamByName('Comentarios').DataType := ftMemo;
      connection.zCommand.Params.ParamByName('Comentarios').value := tmComentariosGenerador.Text;
      connection.zCommand.Params.ParamByName('Status').DataType := ftString;
      connection.zCommand.Params.ParamByName('Status').value := 'Pendiente';
      connection.zCommand.Params.ParamByName('Usuario').DataType := ftString;
      connection.zCommand.Params.ParamByName('Usuario').value := global_usuario;
      connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
      connection.zCommand.Params.ParamByName('Anexo').value := tsAnexo.Text;
      connection.zCommand.Params.ParamByName('TipoMoneda').DataType := ftString;
      connection.zCommand.Params.ParamByName('TipoMoneda').value := txtTipoMoneda.Text;
      connection.zCommand.ExecSQL;

      // Actualizo Kardex del Sistema ....
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen)' +
        'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)');
      connection.zcommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zcommand.Params.ParamByName('Contrato').Value := Global_Contrato;
      connection.zcommand.Params.ParamByName('Usuario').DataType := ftString;
      connection.zcommand.Params.ParamByName('Usuario').Value := Global_Usuario;
      connection.zcommand.Params.ParamByName('Fecha').DataType := ftDate;
      connection.zcommand.Params.ParamByName('Fecha').Value := Date;
      connection.zcommand.Params.ParamByName('Hora').DataType := ftString;
      connection.zcommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now);
      connection.zcommand.Params.ParamByName('Descripcion').DataType := ftString;
      connection.zcommand.Params.ParamByName('Descripcion').Value := 'Creación del Generador No. [' + tsNumeroGenerador.Text + '] de la Orden [' + tsNumeroOrden.Text + ']';
      connection.zcommand.Params.ParamByName('Origen').DataType := ftString;
      connection.zcommand.Params.ParamByName('Origen').Value := 'Generadores';
      connection.zcommand.ExecSQL();

      Estimaciones.Active := False;
      Estimaciones.Open;
    except
      on e: exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al salvar nuevo registro', 0);
      end;
    end
  end
  else
  begin
    try
      Posicion := Estimaciones.RecNo;
      // Actualizo los isometricos del generador

      // Actualizo todas las partidas del generador, enviandola al nuevo generador ...
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('UPDATE estimacionxpartida Set sNumeroGenerador = :Generador Where ' +
        'sContrato = :Contrato And sNumeroOrden = :Orden And sNumeroGenerador = :OldGenerador');
      connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
      connection.zCommand.Params.ParamByName('Generador').value := tsNumeroGenerador.Text;
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
      connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
      connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
      connection.zCommand.Params.ParamByName('OldGenerador').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldGenerador').value := Estimaciones.FieldValues['sNumeroGenerador'];
      connection.zCommand.ExecSQL;

      // Actualizo todas las partidas prov del generador, enviandola al nuevo generador ...
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('UPDATE estimacionxpartidaprov Set sNumeroGenerador = :Generador Where ' +
        'sContrato = :Contrato And sNumeroOrden = :Orden And sNumeroGenerador = :OldGenerador');
      connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
      connection.zCommand.Params.ParamByName('Generador').value := tsNumeroGenerador.Text;
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
      connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
      connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
      connection.zCommand.Params.ParamByName('OldGenerador').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldGenerador').value := Estimaciones.FieldValues['sNumeroGenerador'];
      connection.zCommand.ExecSQL;

      // Actualizo los equipos del generador
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('UPDATE estimacionxequipo Set sNumeroGenerador = :Generador Where ' +
        'sContrato = :Contrato And sNumeroOrden = :Orden And sNumeroGenerador = :OldGenerador');
      connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
      connection.zCommand.Params.ParamByName('Generador').value := tsNumeroGenerador.Text;
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
      connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
      connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
      connection.zCommand.Params.ParamByName('OldGenerador').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldGenerador').value := Estimaciones.FieldValues['sNumeroGenerador'];
      connection.zCommand.ExecSQL;

      // Actualizo Kardex del Sistema ....
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen)' +
        'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)');
      connection.zcommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zcommand.Params.ParamByName('Contrato').Value := Global_Contrato;
      connection.zcommand.Params.ParamByName('Usuario').DataType := ftString;
      connection.zcommand.Params.ParamByName('Usuario').Value := Global_Usuario;
      connection.zcommand.Params.ParamByName('Fecha').DataType := ftDate;
      connection.zcommand.Params.ParamByName('Fecha').Value := Date;
      connection.zcommand.Params.ParamByName('Hora').DataType := ftString;
      connection.zcommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now);
      connection.zcommand.Params.ParamByName('Descripcion').DataType := ftString;
      connection.zcommand.Params.ParamByName('Descripcion').Value := 'Modificación del Generador Original No. [' + Estimaciones.FieldValues['sNumeroGenerador'] + '] de la Orden [' + tsNumeroOrden.Text + '], No. de Generador Final [ ' + tsNumeroGenerador.Text + ']';
      connection.zcommand.Params.ParamByName('Origen').DataType := ftString;
      connection.zcommand.Params.ParamByName('Origen').Value := 'Generadores';
      connection.zcommand.ExecSQL();

      // Ahora si actualizo el encabezado ...

      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('UPDATE estimaciones Set iNumeroEstimacion = :Estimacion, sNumeroGenerador = :Generador, iSemana = :Semana, ' +
        'dFechaInicio = :FechaI, dFechaFinal = :FechaF, sFaseObra = :Fase, sNumeroAnexo = :Anexo, mComentarios = :Comentarios Where ' +
        'sContrato = :Contrato And sNumeroOrden = :Orden And iNumeroEstimacion = :OldEstimacion And sNumeroGenerador = :OldGenerador');
      connection.zCommand.Params.ParamByName('Estimacion').DataType := ftString;
      connection.zCommand.Params.ParamByName('Estimacion').value := tiNumeroEstimacion.KeyValue;
      connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
      connection.zCommand.Params.ParamByName('Generador').value := tsNumeroGenerador.Text;
      connection.zCommand.Params.ParamByName('Semana').DataType := ftInteger;
      connection.zCommand.Params.ParamByName('Semana').value := tiSemana.Value;
      connection.zCommand.Params.ParamByName('FechaI').DataType := ftDate;
      connection.zCommand.Params.ParamByName('FechaI').value := tdFechaInicial.Date;
      connection.zCommand.Params.ParamByName('FechaF').DataType := ftDate;
      connection.zCommand.Params.ParamByName('FechaF').value := tdFechaFinal.Date;
      connection.zCommand.Params.ParamByName('Fase').DataType := ftString;
      connection.zCommand.Params.ParamByName('Fase').value := tsFaseObra.Text;
      connection.zCommand.Params.ParamByName('Comentarios').DataType := ftMemo;
      connection.zCommand.Params.ParamByName('Comentarios').value := tmComentariosGenerador.Text;
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
      connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
      connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
      connection.zCommand.Params.ParamByName('OldEstimacion').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldEstimacion').value := Estimaciones.FieldValues['iNumeroEstimacion'];
      connection.zCommand.Params.ParamByName('OldGenerador').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldGenerador').value := Estimaciones.FieldValues['sNumeroGenerador'];
      connection.zCommand.Params.ParamByName('Anexo').DataType := ftString;
      connection.zCommand.Params.ParamByName('Anexo').value := tsAnexo.Text;
      connection.zCommand.ExecSQL;

      //Isometicos,,
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('UPDATE estimacionxisometrico Set sNumeroGenerador = :Generador Where ' +
        'sContrato = :Contrato And sNumeroOrden = :Orden And sNumeroGenerador = :OldGenerador');
      connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
      connection.zCommand.Params.ParamByName('Generador').value := tsNumeroGenerador.Text;
      connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
      connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
      connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
      connection.zCommand.Params.ParamByName('OldGenerador').DataType := ftString;
      connection.zCommand.Params.ParamByName('OldGenerador').value := Estimaciones.FieldValues['sNumeroGenerador'];
      connection.zCommand.ExecSQL;

      Estimaciones.Active := False;
      Estimaciones.Open;
      Estimaciones.RecNo := Posicion;
    except
      on e: exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al salvar cambios en registro', 0);
      end;
    end
  end;
  OpcButton1 := '';
  frmBarra2.btnPostClick(Sender);
  {Grid_Generadores.SetFocus }
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmEstimaInstalado.frmBarra2btnPrinterClick(Sender: TObject);
begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
      begin
        if Connection.configuracion.FieldValues['sCampPerf'] = 'No' then
                   //SE COMENTA ESTO DEBIDO A QUE IMPRIME SE REQUIRIO LA CARATULA DE UNA SOLA ORDEN
                   //procCaratulaGenerador( global_contrato, Estimaciones.FieldValues[ 'iNumeroEstimacion' ], Estimaciones.FieldValues[ 'sNumeroOrden' ], Estimaciones.FieldValues[ 'sNumeroGenerador' ], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, True )
          procCaratulaGenerador(0, global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, False)
        else
          MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
      end;

      if Connection.configuracion.FieldValues['sCampPerf'] = 'Si' then
        procCaratulaGeneradorPerf(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, True)
    end
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.frmBarra2btnDeleteClick(Sender: TObject);
var
  lContinua: Boolean;
begin
  lContinua := False;
  if Estimaciones.RecordCount > 0 then
  begin
    if (Estimaciones.FieldValues['lStatus'] = 'Pendiente') and (Estimaciones.FieldValues['sIdUsuario'] = global_usuario) then
      lContinua := True
    else if (Estimaciones.FieldValues['lStatus'] = 'Pendiente') and (global_grupo = 'INTEL-CODE') then
      lContinua := True;

    if lContinua then
      if EstimacionxPartida.RecordCount = 0 then
      begin
        if MessageDlg('Desea eliminar el generador seleccionado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          try
            // Actualizo Kardex del Sistema ....
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen)' +
              'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)');
            connection.zcommand.Params.ParamByName('Contrato').DataType := ftString;
            connection.zcommand.Params.ParamByName('Contrato').Value := Global_Contrato;
            connection.zcommand.Params.ParamByName('Usuario').DataType := ftString;
            connection.zcommand.Params.ParamByName('Usuario').Value := Global_Usuario;
            connection.zcommand.Params.ParamByName('Fecha').DataType := ftDate;
            connection.zcommand.Params.ParamByName('Fecha').Value := Date;
            connection.zcommand.Params.ParamByName('Hora').DataType := ftString;
            connection.zcommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now);
            connection.zcommand.Params.ParamByName('Descripcion').DataType := ftString;
            connection.zcommand.Params.ParamByName('Descripcion').Value := 'Eliminación del Generador No. [' + Estimaciones.FieldValues['sNumeroGenerador'] + '] de la Orden [' + tsNumeroOrden.Text + ']';
            connection.zcommand.Params.ParamByName('Origen').DataType := ftString;
            connection.zcommand.Params.ParamByName('Origen').Value := 'Generadores';
            connection.zcommand.ExecSQL();

            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Delete from estimacionxisometrico where sContrato = :Contrato And sNumeroOrden = :Orden And ' +
              'sNumeroGenerador = :Generador');
            connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
            connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
            connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
            connection.zCommand.Params.ParamByName('Orden').value := Estimaciones.FieldValues['sNumeroOrden'];
            connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
            connection.zCommand.Params.ParamByName('Generador').value := Estimaciones.FieldValues['sNumeroGenerador'];
            connection.zCommand.ExecSQL;

            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Delete from estimacionxpartida where sContrato = :Contrato And sNumeroOrden = :Orden And ' +
              'sNumeroGenerador = :Generador');
            connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
            connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
            connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
            connection.zCommand.Params.ParamByName('Orden').value := Estimaciones.FieldValues['sNumeroOrden'];
            connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
            connection.zCommand.Params.ParamByName('Generador').value := Estimaciones.FieldValues['sNumeroGenerador'];
            connection.zCommand.ExecSQL;

            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Delete from estimaciones where sContrato = :Contrato And sNumeroOrden = :Orden And ' +
              'iNumeroEstimacion = :Estimacion And sNumeroGenerador = :Generador');
            connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
            connection.zCommand.Params.ParamByName('Contrato').value := Global_Contrato;
            connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
            connection.zCommand.Params.ParamByName('Orden').value := Estimaciones.FieldValues['sNumeroOrden'];
            connection.zCommand.Params.ParamByName('Estimacion').DataType := ftString;
            connection.zCommand.Params.ParamByName('Estimacion').value := Estimaciones.FieldValues['iNumeroEstimacion'];
            connection.zCommand.Params.ParamByName('Generador').DataType := ftString;
            connection.zCommand.Params.ParamByName('Generador').value := Estimaciones.FieldValues['sNumeroGenerador'];
            connection.zCommand.ExecSQL;
            Estimaciones.Active := False;
            Estimaciones.Open;

            {Grid_Generadores.SetFocus }
          except
            on e: exception do begin
              UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al eliminar registro', 0);
            end;
          end
        end
      end
      else
        MessageDlg('Existen partidas registradas en el generador seleccionado, elimine las partidas para poder eliminar el generador.', mtInformation, [mbOk], 0)
    else
      MessageDlg('El Generador ha sido aplicado o fue creado por un usuario distinto por lo tanto no puede eliminarse.', mtInformation, [mbOk], 0);
  end
end;

procedure TfrmEstimaInstalado.frmBarra2btnRefreshClick(Sender: TObject);
var
  QryTrinomio: tzReadOnlyQuery;
  QryOrdenCambio: tzReadOnlyQuery;
begin
  QryTrinomio := tzReadOnlyQuery.Create(Self);
  QryTrinomio.Connection := connection.zConnection;
  tsInstalacion.Items.Clear;
  QryTrinomio.Active := False;
  QryTrinomio.SQL.Clear;
  QryTrinomio.SQL.Add('Select sInstalacion from contrato_trinomio Where sContrato = :contrato and lVigente = "Si" Order By sInstalacion');
  QryTrinomio.Params.ParamByName('Contrato').DataType := ftString;
  QryTrinomio.Params.ParamByName('Contrato').Value := global_contrato;
  QryTrinomio.Open;
  if QryTrinomio.RecordCount > 0 then
    while not QryTrinomio.Eof do
    begin
      tsInstalacion.Items.Add(QryTrinomio.FieldValues['sInstalacion']);
      QryTrinomio.Next
    end
  else
    tsInstalacion.Items.Add(global_contrato);
  QryTrinomio.Destroy;

  QryOrdenCambio := tzReadOnlyQuery.Create(Self);
  QryOrdenCambio.Connection := connection.zConnection;
  tiOrdenCambio.Items.Clear;
  tiOrdenCambio.Items.Add('SIN ORDEN DE CAMBIO');
  QryOrdenCambio.Active := False;
  QryOrdenCambio.SQL.Clear;
  QryOrdenCambio.SQL.Add('Select iCedulaProcedencia, sNotificacionOficio from ordendecambio Where sContrato = :contrato order by iCedulaProcedencia');
  QryOrdenCambio.Params.ParamByName('Contrato').DataType := ftString;
  QryOrdenCambio.Params.ParamByName('Contrato').Value := global_contrato;
  QryOrdenCambio.Open;
  while not QryOrdenCambio.Eof do
  begin
    tiOrdenCambio.Items.Add('O.C. No. [' + IntToStr(QryOrdenCambio.FieldValues['iCedulaProcedencia']) + ']');
    QryOrdenCambio.Next
  end;
  QryOrdenCambio.Destroy;

  tsInstalacion.ItemIndex := 0;
  tiOrdenCambio.ItemIndex := 0;

  OrdenesdeTrabajo.Active := False;
  OrdenesdeTrabajo.Open;

  Estimaciones.Active := False;
  Estimaciones.Open;

  Connection.EstimacionPeriodo.Active := False;
  connection.EstimacionPeriodo.Open;

  Isometricos.Active := False;
  Isometricos.Open;

  QryPartidasEfectivas.Active := False;
  QryPartidasEfectivas.Open;
end;

procedure TfrmEstimaInstalado.frmBarra2btnCancelClick(Sender: TObject);
begin
  tdFechaInicial.Enabled := False;
  tdFechaFinal.Enabled := False;
  tsFaseObra.ReadOnly := True;
  tiFases.Enabled := False;
  tsNumeroGenerador.ReadOnly := True;
  tiNumeroEstimacion.ReadOnly := True;
  tmComentariosGenerador.ReadOnly := True;
  OpcButton1 := '';
  desactivapop(popupprincipal);
  frmBarra2.btnCancelClick(Sender);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmEstimaInstalado.frmBarra2btnExitClick(Sender: TObject);
begin
  frmBarra2.btnExitClick(Sender);
  close
end;

procedure TfrmEstimaInstalado.ComentariosAdicionalesClick(
  Sender: TObject);
begin
  global_partida := tsNumeroActividad.Text;
  Application.CreateForm(TfrmComentariosxAnexo, frmComentariosxAnexo);
  frmComentariosxAnexo.show;
end;

procedure TfrmEstimaInstalado.tsNumeroActividadChange(Sender: TObject);
begin
  if (OpcButton1 = 'New') or (OpcButton1 = 'Edit') then
  begin
    if lPerimetros then
    begin
      messagedlg('No puedes seleccionar otra partida mientras ingresas Datos!', mtInformation, [mbOk], 0);
      exit;
    end;
  end;

  global_partida := tsNumeroActividad.Text;
  tsNumeroActividad.Hint := QryPartidasEfectivas.FieldValues['mDescripcion'];
  mObra.Text := '';
  mObra.Lines.Add(QryPartidasEfectivas.FieldValues['mDescripcion']);
  imgNotas.Visible := False;

  AnexoConvenio.Active := False;
  AnexoConvenio.Params.ParamByName('Contrato').DataType := ftString;
  AnexoConvenio.Params.ParamByName('Contrato').Value := global_contrato;
  AnexoConvenio.Params.ParamByName('Actividad').DataType := ftString;
  AnexoConvenio.Params.ParamByName('Actividad').Value := '';
  AnexoConvenio.Open;

  txtMarca.Enabled := False;
  txtSubMarca.Enabled := False;
  txtPzas.Enabled := False;
  txtPesoTotal.Enabled := False;
  txtLongArea.Enabled := False;
  txtLongTotal.Enabled := False;
  txtPesoxUnidad.Enabled := False;
end;

procedure TfrmEstimaInstalado.imgNotasDblClick(Sender: TObject);
begin
  ComentariosAdicionales.Click
end;

procedure TfrmEstimaInstalado.tiOrdenCambioEnter(Sender: TObject);
begin
  tiOrdenCambio.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.tiOrdenCambioExit(Sender: TObject);
begin
  tiOrdenCambio.Color := global_color_salida
end;

procedure TfrmEstimaInstalado.tiOrdenCambioKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
    if tsIsometricoReferencia.Enabled then
      tsIsometricoReferencia.SetFocus
    else
      tmComentarios.SetFocus
end;

procedure TfrmEstimaInstalado.ListadeVerificacin1Click(Sender: TObject);
begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
        procListadeVerificacion(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, False)
      else
        MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
    end
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Administración de Contratos', 'Al ejecutar Lista de Verificación', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.NumerosGeneradoresCIA1Click(Sender: TObject);
begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
        procNumeroGenerador(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, 'Contratista', frmEstimaInstalado, frGenerador.OnGetValue, False)
      else
        MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
    end
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al click en Resumen Generadores (Isometrico CIA)', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.Historialdelapartidaanexo1Click(
  Sender: TObject);
begin
  Panel.Visible := not Panel.Visible;
end;

procedure TfrmEstimaInstalado.HojaSeguimiento1Click(Sender: TObject);
var
  estimacion: string;
  qry, qry2: TZReadOnlyQuery;
begin
  try

    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
      begin
         //SE OBTIENE EL NUEMRO DE ESTIMACION JUTO CON SUS FECHAS..
        qry2 := TZReadOnlyQuery.Create(Self);
        qry2.Connection := connection.zConnection;
        qry2.Active := False;
        qry2.SQL.Clear;
        qry2.SQL.Add('select iNumeroEstimacion, dFechaInicio, dFechaFinal, sIdTipoEstimacion from estimacionperiodo ' +
          'where sContrato =:Contrato and iNumeroEstimacion =:Estimacion');
        qry2.Params.ParamByName('Contrato').DataType := ftString;
        qry2.Params.ParamByName('Contrato').Value := global_contrato;
        qry2.Params.ParamByName('Estimacion').DataType := ftString;
        qry2.Params.ParamByName('Estimacion').Value := tiNumeroEstimacion.KeyValue;
        qry2.Open;
        if qry2.RecordCount > 0 then
        begin
          Fi.Date := qry2.FieldValues['dFechaInicio'];
          Ff.Date := qry2.FieldValues['dFechaFinal'];
          estimacion := qry2.FieldValues['iNumeroEstimacion'];
        end;
        fechas.Active := true;
        fechas.Open;
        fechas.EmptyTable;
        fechas.Append;
        fechas.FieldValues['fi'] := Fi.Date;
        fechas.FieldValues['ff'] := Ff.Date;
        fechas.FieldValues['Estima'] := estimacion;
        fechas.Post;
        procHojasegGeneradores(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado, frGenerador.OnGetValue, True)
      end
      else
        MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
    end

  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al Imprimir Hoja Seguimiento de Entrega', 0);
    end;
  end;
end;



procedure TfrmEstimaInstalado.tsPrefijoEnter(Sender: TObject);
begin
  tsPrefijo.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.tsPrefijoExit(Sender: TObject);
begin
  tsPrefijo.Color := global_color_salida;
  if tsPrefijo.ReadOnly = False then
    if (tsIsometrico.Text <> '') or (tsPrefijo.Text <> '') then
    begin
      if lExisteIsometrico(Estimaciones.FieldValues['sNumeroGenerador'], tsIsometrico.Text, tsPrefijo.Text) then
      begin
        mnHistorial.Click;
        MessageDlg('El Isometrico: ' + tsIsometrico.Text + ' ' + tsPrefijo.Text + ' se encuentra registrado en otro generador de la misma orden de trabajo.', mtInformation, [mbOk], 0);
        tsIsometrico.SetFocus
      end
      else
        gbIsometricos.Visible := False
    end

end;

procedure TfrmEstimaInstalado.tsPrefijoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
    if tsInstalacion.Enabled then
      tsInstalacion.SetFocus
    else if tiOrdenCambio.Enabled then
      tiOrdenCambio.SetFocus
    else if tsIsometricoReferencia.Enabled then
      tsIsometricoReferencia.SetFocus
    else
      tmComentarios.SetFocus
end;

procedure TfrmEstimaInstalado.mnHistorialClick(Sender: TObject);
begin
  if (tsIsometrico.Text <> '') or (tsPrefijo.Text <> '') then
  begin
    gbIsometricos.Visible := True;
    QryPrefijos.Active := False;
    QryPrefijos.Params.ParamByName('Contrato').DataType := ftString;
    QryPrefijos.Params.ParamByName('Contrato').Value := global_contrato;
    QryPrefijos.Params.ParamByName('Isometrico').DataType := ftString;
    QryPrefijos.Params.ParamByName('Isometrico').Value := tsIsometrico.Text;
    QryPrefijos.Params.ParamByName('Ordenado').DataType := ftString;
    QryPrefijos.Params.ParamByName('Ordenado').Value := 'sPrefijo';
    QryPrefijos.Open;
    gbIsometricos.Caption := 'Historial de Isometrico [' + tsIsometrico.Text + ']'
  end;
end;

procedure TfrmEstimaInstalado.tsInstalacionEnter(Sender: TObject);
begin
  tsInstalacion.Color := global_color_entrada;
end;

procedure TfrmEstimaInstalado.tsInstalacionExit(Sender: TObject);
begin
  tsInstalacion.Color := global_color_salida;
end;

procedure TfrmEstimaInstalado.tsInstalacionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
    if tiOrdenCambio.Enabled then
      tiOrdenCambio.SetFocus
    else if tsIsometricoReferencia.Enabled then
      tsIsometricoReferencia.SetFocus
    else
      tmComentarios.SetFocus
end;

procedure TfrmEstimaInstalado.Grid_GeneradoresDblClick(Sender: TObject);
begin
  if Estimaciones.RecordCount > 0 then
  begin
    sGeneradorContrato := Estimaciones.FieldValues['sContrato'];
    sGeneradorOrden := Estimaciones.FieldValues['sNumeroOrden'];
    sGeneradorNumero := Estimaciones.FieldValues['sNumeroGenerador'];
    sGeneradorStatus := Estimaciones.FieldValues['lStatus'];

    Application.CreateForm(TfrmEstimacionAlbum, frmEstimacionAlbum);
    frmEstimacionAlbum.show;
  end
end;

procedure TfrmEstimaInstalado.EstimacionxPartidaCalcFields(
  DataSet: TDataSet);
begin
  EstimacionxPartidadMontoMN.Value := EstimacionxPartidadVentaMN.Value * EstimacionxPartidadGenerado.Value;
end;



procedure TfrmEstimaInstalado.grid_igualesGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
  if Pos('WBS', global_checkGenerador) > 0 then
    if (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sWbs').AsString = ActividadesIguales.FieldValues['sWbs'] then
      Background := clGradientInactiveCaption

end;

procedure TfrmEstimaInstalado.QryActividadesxOrdenCalcFields(
  DataSet: TDataSet);
var
  dGenerado,
    dAcumulado,
    dMonto,
    dMontoAcumulado: Currency;
begin
  if sTipoReporte <> '' then
  begin
    if QryActividadesxOrden.FieldValues['sTipoActividad'] = 'Actividad' then
    begin
      Connection.qryBusca2.Active := False;
      Connection.qryBusca2.SQL.Clear;
      Connection.qryBusca2.SQL.Add('Select dCantidadAnexo From actividadesxanexo ' +
        'Where sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroActividad = :Actividad and sTipoActividad = "Actividad"');
      Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato;
      Connection.QryBusca2.Params.ParamByName('convenio').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('convenio').Value := global_convenio;
      Connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Actividad').Value := QryActividadesxOrden.FieldValues['sNumeroActividad'];
      Connection.qryBusca2.Open;
      if Connection.qryBusca2.RecordCount > 0 then
        QryActividadesxOrdendCantidadAnexo.Value := Connection.qryBusca2.FieldValues['dCantidadAnexo']
      else
        QryActividadesxOrdendCantidadAnexo.Value := 0;

      Connection.qryBusca2.Active := False;
      Connection.qryBusca2.SQL.Clear;
      if sTipoReporte = 'individual' then
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as Generado From estimacionxpartida e ' +
          'inner join estimaciones e2 on (e.sContrato = e2.sContrato and e.sNumeroOrden = e2.sNumeroOrden and e.sNumeroGenerador = e2.sNumeroGenerador ) ' +
          'Where e.sContrato = :Contrato And e.sNumeroOrden = :Orden And e.sNumeroGenerador = :Generador And e.sNumeroActividad = :Actividad ' +
          'Group By e.sNumeroActividad')
      else
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as Generado From estimacionxpartida e ' +
          'inner join estimaciones e2 on (e.sContrato = e2.sContrato and e.sNumeroOrden = e2.sNumeroOrden and e.sNumeroGenerador = e2.sNumeroGenerador ) ' +
          'Where e.sContrato = :Contrato And e.sNumeroOrden = :Orden And e.sNumeroActividad = :Actividad ' +
          'Group By e.sNumeroActividad');
      Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato;
      Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
      if sTipoReporte = 'individual' then
      begin
        Connection.QryBusca2.Params.ParamByName('Generador').DataType := ftString;
        Connection.QryBusca2.Params.ParamByName('Generador').Value := Estimaciones.FieldValues['sNumeroGenerador'];
      end;
      Connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Actividad').Value := QryActividadesxOrden.FieldValues['sNumeroActividad'];
      Connection.qryBusca2.Open;
      if Connection.qryBusca2.RecordCount > 0 then
        dGenerado := Connection.qryBusca2.FieldValues['Generado']
      else
        dGenerado := 0;

      Connection.qryBusca2.Active := False;
      Connection.qryBusca2.SQL.Clear;
      if sTipoReporte = 'individual' then
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as Generado From estimacionxpartida e ' +
          'inner join estimaciones e2 on (e.sContrato = e2.sContrato and e.sNumeroOrden = e2.sNumeroOrden and e.sNumeroGenerador = e2.sNumeroGenerador And e2.iConsecutivo <= :Consecutivo ) ' +
          'Where e.sContrato = :Contrato And e.sNumeroOrden = :Orden And e.sWbs = :Wbs And e.sNumeroActividad = :Actividad ' +
          'Group By e.sWbs, e.sNumeroActividad')
      else
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as Generado From estimacionxpartida e ' +
          'inner join estimaciones e2 on (e.sContrato = e2.sContrato and e.sNumeroOrden = e2.sNumeroOrden and e.sNumeroGenerador = e2.sNumeroGenerador ) ' +
          'Where e.sContrato = :Contrato And e.sNumeroOrden = :Orden And e.sWbs = :Wbs And e.sNumeroActividad = :Actividad ' +
          'Group By e.sWbs, e.sNumeroActividad');

      Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato;
      Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
      if sTipoReporte = 'individual' then
      begin
        Connection.QryBusca2.Params.ParamByName('Consecutivo').DataType := ftInteger;
        Connection.QryBusca2.Params.ParamByName('Consecutivo').Value := Estimaciones.FieldValues['iConsecutivo'];
      end;
      Connection.QryBusca2.Params.ParamByName('Wbs').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Wbs').Value := QryActividadesxOrden.FieldValues['sWbs'];
      Connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Actividad').Value := QryActividadesxOrden.FieldValues['sNumeroActividad'];
      Connection.qryBusca2.Open;
      if Connection.qryBusca2.RecordCount > 0 then
        dAcumulado := Connection.qryBusca2.FieldValues['Generado']
      else
        dAcumulado := 0;

      QryActividadesxOrdendCantidad.Value := dGenerado;
      QryActividadesxOrdendTotal.Value := dGenerado * QryActividadesxOrdendVentaMN.Value;
      QryActividadesxOrdendCantidadAcumulado.Value := dAcumulado;
      QryActividadesxOrdendTotalAcumulado.Value := dAcumulado * QryActividadesxOrdendVentaMN.Value;

    end
    else if QryActividadesxOrdeniNivel.Value <= 1 then
    begin
      // Es Paquete ...
      Connection.qryBusca2.Active := False;
      Connection.qryBusca2.SQL.Clear;
      if sTipoReporte = 'individual' then
        Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad * a.dVentaMN) as dTotal From estimacionxpartida b ' +
          'inner join estimaciones e2 on (b.sContrato = e2.sContrato and b.sNumeroOrden = e2.sNumeroOrden and b.sNumeroGenerador = e2.sNumeroGenerador ) ' +
          'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sNumeroOrden = b.sNumeroOrden and a.sWbs = b.sWbs And a.sNumeroActividad = b.sNumeroActividad and a.sIdConvenio = :Convenio) ' +
          'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.sNumeroGenerador = :Generador And b.sWbs Like :Wbs ' +
          'Group By b.sContrato')
      else
        Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad * a.dVentaMN) as dTotal From estimacionxpartida b ' +
          'inner join estimaciones e2 on (b.sContrato = e2.sContrato and b.sNumeroOrden = e2.sNumeroOrden and b.sNumeroGenerador = e2.sNumeroGenerador ) ' +
          'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sNumeroOrden = b.sNumeroOrden and a.sWbs = b.sWbs And a.sNumeroActividad = b.sNumeroActividad and a.sIdConvenio = :Convenio) ' +
          'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.sWbs Like :Wbs ' +
          'Group By b.sContrato');

      Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato;
      Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
      Connection.QryBusca2.Params.ParamByName('convenio').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('convenio').Value := global_convenio;
      if sTipoReporte = 'individual' then
      begin
        Connection.QryBusca2.Params.ParamByName('generador').DataType := ftString;
        Connection.QryBusca2.Params.ParamByName('generador').Value := Estimaciones.FieldValues['sNumeroGenerador'];
      end;
      Connection.QryBusca2.Params.ParamByName('Wbs').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Wbs').Value := Trim(QryActividadesxOrden.FieldValues['sWbs']) + '.%';
      Connection.qryBusca2.Open;
      if Connection.qryBusca2.RecordCount > 0 then
        dMonto := connection.QryBusca2.FieldValues['dTotal']
      else
        dMonto := 0;

      Connection.qryBusca2.Active := False;
      Connection.qryBusca2.SQL.Clear;
      if sTipoReporte = 'individual' then
        Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad * a.dVentaMN) as dTotal From estimacionxpartida b ' +
          'inner join estimaciones e2 on (b.sContrato = e2.sContrato and b.sNumeroOrden = e2.sNumeroOrden and b.sNumeroGenerador = e2.sNumeroGenerador And e2.iConsecutivo <= :Consecutivo ) ' +
          'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sNumeroOrden = b.sNumeroOrden and a.sWbs = b.sWbs And a.sNumeroActividad = b.sNumeroActividad and a.sIdConvenio = :Convenio) ' +
          'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.sWbs Like :Wbs ' +
          'Group By b.sContrato')
      else
        Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad * a.dVentaMN) as dTotal From estimacionxpartida b ' +
          'inner join estimaciones e2 on (b.sContrato = e2.sContrato and b.sNumeroOrden = e2.sNumeroOrden and b.sNumeroGenerador = e2.sNumeroGenerador ) ' +
          'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sNumeroOrden = b.sNumeroOrden and a.sWbs = b.sWbs And a.sNumeroActividad = b.sNumeroActividad and a.sIdConvenio = :Convenio) ' +
          'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.sWbs Like :Wbs ' +
          'Group By b.sContrato');
      Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato;
      Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
      Connection.QryBusca2.Params.ParamByName('convenio').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('convenio').Value := global_convenio;
      if sTipoReporte = 'individual' then
      begin
        Connection.QryBusca2.Params.ParamByName('Consecutivo').DataType := ftInteger;
        Connection.QryBusca2.Params.ParamByName('Consecutivo').Value := Estimaciones.FieldValues['iConsecutivo'];
      end;
      Connection.QryBusca2.Params.ParamByName('Wbs').DataType := ftString;
      Connection.QryBusca2.Params.ParamByName('Wbs').Value := Trim(QryActividadesxOrden.FieldValues['sWbs']) + '.%';
      Connection.qryBusca2.Open;
      if Connection.qryBusca2.RecordCount > 0 then
        dMontoAcumulado := connection.QryBusca2.FieldValues['dTotal']
      else
        dMontoAcumulado := 0;
      QryActividadesxOrdendCantidad.Value := 0;
      QryActividadesxOrdendCantidadAcumulado.Value := 0;
      QryActividadesxOrdendTotal.Value := dMonto;
      QryActividadesxOrdendTotalAcumulado.Value := dMontoAcumulado;
    end
  end
end;

procedure TfrmEstimaInstalado.mnResumenClick(Sender: TObject);
var
  iRecord: Byte;
  QryConfiguracion: tzReadOnlyQuery;

begin
  try

    if grid_generadores.DataSource.DataSet.IsEmpty = false then
    begin
      QryConfiguracion := tzReadOnlyQuery.Create(Self);
      QryConfiguracion.Connection := connection.zConnection;

      QryConfiguracion.Sql.Add('Select iFirmas from configuracion Where sContrato= :Contrato');
      QryConfiguracion.Params.ParamByName('contrato').DataType := ftString;
      QryConfiguracion.Params.ParamByName('contrato').Value := global_Contrato;
      QryConfiguracion.Open;

      sTipoReporte := 'individual';
      if rxActividadesxOrden.RecordCount > 0 then
      begin //********************************************por esto no habre el reporte
        rxActividadesxOrden.EmptyTable;

        QryActividadesxOrden.Active := False;
        QryActividadesxOrden.Params.ParamByName('Contrato').DataType := ftString;
        QryActividadesxOrden.Params.ParamByName('Contrato').Value := global_contrato;
        QryActividadesxOrden.Params.ParamByName('Convenio').DataType := ftString;
        QryActividadesxOrden.Params.ParamByName('Convenio').Value := global_convenio;
        QryActividadesxOrden.Params.ParamByName('Orden').DataType := ftString;
        QryActividadesxOrden.Params.ParamByName('Orden').Value := estimaciones.FieldValues['sNumeroOrden'];
        QryActividadesxOrden.Open;
        QryActividadesxOrden.First;
        while not QryActividadesxOrden.Eof do
        begin
          if QryActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
          begin
            if QryActividadesxOrden.FieldValues['iNivel'] <= 1 then
            begin
              rxActividadesxOrden.Insert;
              for iRecord := 0 to QryActividadesxOrden.Fields.Count - 1 do
                rxActividadesxOrden.FieldValues[QryActividadesxOrden.Fields.Fields[iRecord].DisplayName] := QryActividadesxOrden.Fields.Fields[iRecord].Value;
              rxActividadesxOrden.Post;
            end
          end
          else if QryActividadesxOrden.FieldValues['dTotalAcumulado'] > 0 then
          begin
            rxActividadesxOrden.Insert;
            for iRecord := 0 to QryActividadesxOrden.Fields.Count - 1 do
              rxActividadesxOrden.FieldValues[QryActividadesxOrden.Fields.Fields[iRecord].DisplayName] := QryActividadesxOrden.Fields.Fields[iRecord].Value;
            rxActividadesxOrden.Post;
          end;
          QryActividadesxOrden.Next
        end;
        rDiarioFirmas(global_contrato, estimaciones.FieldValues['sNumeroOrden'], 'A', estimaciones.FieldValues['dFechaFinal'], frmEstimaInstalado);
        frGenerador.PreviewOptions.MDIChild := True;
        frGenerador.PreviewOptions.Modal := False;
        frGenerador.PreviewOptions.Maximized := lCheckMaximized();
        frGenerador.PreviewOptions.ShowCaptions := False;
        frGenerador.Previewoptions.ZoomMode := zmPageWidth;

        if QryConfiguracion.FieldValues['iFirmas'] = '2' then
          frGenerador.LoadFromFile(global_files + 'ResumenGenerador2.fr3');
        if QryConfiguracion.FieldValues['iFirmas'] = '3' then
          frGenerador.LoadFromFile(global_files + 'ResumenGenerador.fr3');

        frGenerador.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
        sTipoReporte := '';
        QryConfiguracion.Destroy;
      end; //**********************************************porloqestaarriba
    end;

  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir resumen de generacion de obra', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.mnResumenGeneralClick(Sender: TObject);
var
  iRecord: Byte;
  QryConfiguracion: tzReadOnlyQuery;

begin
  try

    if grid_generadores.DataSource.DataSet.IsEmpty = false then
    begin
      QryConfiguracion := tzReadOnlyQuery.Create(Self);
      QryConfiguracion.Connection := connection.zConnection;

      QryConfiguracion.Sql.Add('Select iFirmas from configuracion Where sContrato= :Contrato');
      QryConfiguracion.Params.ParamByName('contrato').DataType := ftString;
      QryConfiguracion.Params.ParamByName('contrato').Value := global_Contrato;
      QryConfiguracion.Open;

      sTipoReporte := 'general';
      if rxActividadesxOrden.RecordCount > 0 then
        rxActividadesxOrden.EmptyTable;

      QryActividadesxOrden.Active := False;
      QryActividadesxOrden.Params.ParamByName('Contrato').DataType := ftString;
      QryActividadesxOrden.Params.ParamByName('Contrato').Value := global_contrato;
      QryActividadesxOrden.Params.ParamByName('Convenio').DataType := ftString;
      QryActividadesxOrden.Params.ParamByName('Convenio').Value := global_convenio;
      QryActividadesxOrden.Params.ParamByName('Orden').DataType := ftString;
      QryActividadesxOrden.Params.ParamByName('Orden').Value := estimaciones.FieldValues['sNumeroOrden'];
      QryActividadesxOrden.Open;

      QryActividadesxOrden.First;
      while not QryActividadesxOrden.Eof do
      begin
        if QryActividadesxOrden.FieldValues['sTipoActividad'] = 'Paquete' then
        begin
          if QryActividadesxOrden.FieldValues['iNivel'] <= 1 then
          begin
            rxActividadesxOrden.Insert;
            for iRecord := 0 to QryActividadesxOrden.Fields.Count - 1 do
              rxActividadesxOrden.FieldValues[QryActividadesxOrden.Fields.Fields[iRecord].DisplayName] := QryActividadesxOrden.Fields.Fields[iRecord].Value;
            rxActividadesxOrden.Post;
          end
        end
        else if QryActividadesxOrden.FieldValues['dTotalAcumulado'] > 0 then
        begin
         //Buscamos si la partida esta dentro de la estiamcion..
//         connection.zCommand.Active := False;
//         connection.zCommand.SQL.Clear;
//         connection.zCommand.SQL.Add('select ep.sNumeroActividad from estimaciones e '+
//                                     'inner join estimacionxpartida ep on (e.sContrato = ep.sContrato and e.sNumeroOrden = ep.sNumeroOrden and e.sNumeroGenerador = ep.sNumeroGenerador and ep.sWbs =:Wbs and ep.sNumeroActividad =:Actividad ) '+
//                                     'where e.sContrato =:Contrato and e.iNumeroEstimacion =:Estimacion and e.sNumeroOrden =:Orden ');
//         connection.zCommand.ParamByName('Contrato').AsString   := global_contrato;
//         connection.zCommand.ParamByName('Wbs').AsString        := QryActividadesxOrden.FieldValues['sWbs'];
//         connection.zCommand.ParamByName('Actividad').AsString  := QryActividadesxOrden.FieldValues['sNumeroActividad'];
//         connection.zCommand.ParamByName('Estimacion').AsString := estimaciones.FieldValues['iNumeroEstimacion'];
//         connection.zCommand.ParamByName('Orden').AsString      := estimaciones.FieldValues['sNumeroOrden' ];
//         connection.zCommand.Open;
//
//         if connection.zCommand.RecordCount > 0 then
//         begin
          rxActividadesxOrden.Insert;
          for iRecord := 0 to QryActividadesxOrden.Fields.Count - 1 do
            rxActividadesxOrden.FieldValues[QryActividadesxOrden.Fields.Fields[iRecord].DisplayName] := QryActividadesxOrden.Fields.Fields[iRecord].Value;
          rxActividadesxOrden.Post;
//         end;
        end;
        QryActividadesxOrden.Next
      end;
      rDiarioFirmas(global_contrato, estimaciones.FieldValues['sNumeroOrden'], 'A', estimaciones.FieldValues['dFechaFinal'], frmEstimaInstalado);
      frGenerador.PreviewOptions.MDIChild := True;
      frGenerador.PreviewOptions.Modal := False;
      frGenerador.PreviewOptions.Maximized := lCheckMaximized();
      frGenerador.PreviewOptions.ShowCaptions := False;
      frGenerador.Previewoptions.ZoomMode := zmPageWidth;

      if QryConfiguracion.FieldValues['iFirmas'] = '2' then
        frGenerador.LoadFromFile(global_files + 'ResumendeOrdendeTrabajo2.fr3');

      if QryConfiguracion.FieldValues['iFirmas'] = '3' then
        frGenerador.LoadFromFile(global_files + 'ResumendeOrdendeTrabajo.fr3');

      frGenerador.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
      sTipoReporte := '';
    end;

  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir Resumen General por Orden de Trabajo', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.mnConcentradoIsometricosClick(
  Sender: TObject);
var
  QryConcentradoIsometricos: tzReadOnlyQuery;
  dsConcentradoIsometricos: tfrxDBDataSet;
  frIsometricos: TfrxReport;
  QryConfiguracion: tzReadOnlyQuery;
begin
  try
    if grid_generadores.DataSource.DataSet.IsEmpty = false then
    begin
      QryConfiguracion := tzReadOnlyQuery.Create(Self);
      QryConfiguracion.Connection := connection.zConnection;

      QryConfiguracion.Sql.Add('Select iFirmas from configuracion Where sContrato= :Contrato');
      QryConfiguracion.Params.ParamByName('contrato').DataType := ftString;
      QryConfiguracion.Params.ParamByName('contrato').Value := global_Contrato;
      QryConfiguracion.Open;

      QryConcentradoIsometricos := tzReadOnlyQuery.Create(Self);
      QryConcentradoIsometricos.Connection := connection.zConnection;
      QryConcentradoIsometricos.Active := False;
      QryConcentradoIsometricos.SQL.Clear;

      QryConcentradoIsometricos.SQL.Add('select e.sWbs, e.sNumeroActividad, a.mDescripcion, a.sMedida, a.dVentaMN, a.dCantidad as dCantidadInstalar, ' +
        'e.sIsometrico, e.sPrefijo, e.dCantidad from estimacionxpartida e ' +
        'inner join actividadesxorden a on (e.sContrato = a.sContrato and e.sNumeroOrden = a.sNumeroOrden and ' +
        'a.sIdConvenio = :Convenio and e.sWbs = a.sWbs and e.sNumeroActividad = a.sNumeroActividad and a.sTipoActividad = "Actividad") ' +
        'Where e.sContrato = :Contrato and e.sNumeroOrden = :Orden and e.sNumeroGenerador = :Generador order by a.iItemOrden, e.sIsometrico');
      QryConcentradoIsometricos.Params.ParamByName('contrato').DataType := ftString;
      QryConcentradoIsometricos.Params.ParamByName('contrato').Value := global_contrato;
      QryConcentradoIsometricos.Params.ParamByName('convenio').DataType := ftString;
      QryConcentradoIsometricos.Params.ParamByName('convenio').Value := global_convenio;
      QryConcentradoIsometricos.Params.ParamByName('orden').DataType := ftString;
      QryConcentradoIsometricos.Params.ParamByName('orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
      QryConcentradoIsometricos.Params.ParamByName('generador').DataType := ftString;
      QryConcentradoIsometricos.Params.ParamByName('generador').Value := Estimaciones.FieldValues['sNumeroGenerador'];
      QryConcentradoIsometricos.Open;

      dsConcentradoIsometricos := tfrxDBDataSet.Create(Self);
      dsConcentradoIsometricos.DataSet := QryConcentradoIsometricos;
      dsConcentradoIsometricos.UserName := 'dsConcentradoIsometricos';

      rDiarioFirmas(global_contrato, estimaciones.FieldValues['sNumeroOrden'], 'A', estimaciones.FieldValues['dFechaFinal'], frmEstimaInstalado);

      frIsometricos := tFrxReport.Create(Self);
      frIsometricos.DataSets.Add(connection.rpt_contrato);
      frIsometricos.DataSets.Add(connection.rpt_setup);
      frIsometricos.DataSets.Add(dsConcentradoIsometricos);
      frIsometricos.OnGetValue := frGenerador.OnGetValue;
      frIsometricos.PreviewOptions.MDIChild := False;
      frIsometricos.PreviewOptions.Modal := True;
      frIsometricos.PreviewOptions.Maximized := lCheckMaximized();
      frIsometricos.PreviewOptions.ShowCaptions := False;
      frIsometricos.Previewoptions.ZoomMode := zmPageWidth;
      if QryConfiguracion.FieldValues['iFirmas'] = '2' then
        frIsometricos.LoadFromFile(global_files + 'IsometricosxGenerador.fr3');
      if QryConfiguracion.FieldValues['iFirmas'] = '3' then
        frIsometricos.LoadFromFile(global_files + 'IsometricosxGenerador3.fr3');
      frIsometricos.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));

      frIsometricos.Destroy;
      dsConcentradoIsometricos.Destroy;
      QryConcentradoIsometricos.Destroy;

    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir Concentrado de Isometricos x Generador', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.mnConcentradoIsometricosGralClick(
  Sender: TObject);
var
  QryConcentradoIsometricos: tzReadOnlyQuery;
  dsConcentradoIsometricos: tfrxDBDataSet;
  frIsometricos: TfrxReport;
  QryConfiguracion: tzReadOnlyQuery;

begin
  try
    if grid_generadores.DataSource.DataSet.IsEmpty = false then
    begin
      QryConfiguracion := tzReadOnlyQuery.Create(Self);
      QryConfiguracion.Connection := connection.zConnection;

      QryConfiguracion.Sql.Add('Select iFirmas from configuracion Where sContrato= :Contrato');
      QryConfiguracion.Params.ParamByName('contrato').DataType := ftString;
      QryConfiguracion.Params.ParamByName('contrato').Value := global_Contrato;
      QryConfiguracion.Open;

      QryConcentradoIsometricos := tzReadOnlyQuery.Create(Self);
      QryConcentradoIsometricos.Connection := connection.zConnection;
      QryConcentradoIsometricos.Active := False;
      QryConcentradoIsometricos.SQL.Clear;
      QryConcentradoIsometricos.SQL.Add('select e.sWbs, e.sNumeroActividad, a.mDescripcion, a.sMedida, a.dVentaMN, a.dCantidad as dCantidadInstalar, ' +
        'e.sNumeroGenerador, e.sIsometrico, e.sPrefijo, e.dCantidad from estimacionxpartida e ' +
        'inner join actividadesxorden a on (e.sContrato = a.sContrato and e.sNumeroOrden = a.sNumeroOrden and ' +
        'a.sIdConvenio = :Convenio and e.sWbsContrato = a.sWbsContrato and e.sNumeroActividad = a.sNumeroActividad and a.sTipoActividad = "Actividad") ' +
        'Where e.sContrato = :Contrato and e.sNumeroOrden = :Orden order by a.iItemOrden, e.sNumeroGenerador, e.sIsometrico');
      QryConcentradoIsometricos.Params.ParamByName('contrato').DataType := ftString;
      QryConcentradoIsometricos.Params.ParamByName('contrato').Value := global_contrato;
      QryConcentradoIsometricos.Params.ParamByName('convenio').DataType := ftString;
      QryConcentradoIsometricos.Params.ParamByName('convenio').Value := global_convenio;
      QryConcentradoIsometricos.Params.ParamByName('orden').DataType := ftString;
      QryConcentradoIsometricos.Params.ParamByName('orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
      QryConcentradoIsometricos.Open;

      dsConcentradoIsometricos := tfrxDBDataSet.Create(Self);
      dsConcentradoIsometricos.DataSet := QryConcentradoIsometricos;
      dsConcentradoIsometricos.UserName := 'dsConcentradoIsometricos';

      rDiarioFirmas(global_contrato, estimaciones.FieldValues['sNumeroOrden'], 'A', estimaciones.FieldValues['dFechaFinal'], frmEstimaInstalado);

      frIsometricos := tFrxReport.Create(Self);
      frIsometricos.DataSets.Add(connection.rpt_contrato);
      frIsometricos.DataSets.Add(connection.rpt_setup);
      frIsometricos.DataSets.Add(dsConcentradoIsometricos);
      frIsometricos.OnGetValue := frGenerador.OnGetValue;
      frIsometricos.PreviewOptions.MDIChild := False;
      frIsometricos.PreviewOptions.Modal := True;
      frIsometricos.PreviewOptions.Maximized := lCheckMaximized();
      frIsometricos.PreviewOptions.ShowCaptions := False;
      frIsometricos.Previewoptions.ZoomMode := zmPageWidth;
      if QryConfiguracion.FieldValues['iFirmas'] = '2' then
        frIsometricos.LoadFromFile(global_files + 'IsometricosxOrden.fr3');
      if QryConfiguracion.FieldValues['iFirmas'] = '3' then
        frIsometricos.LoadFromFile(global_files + 'IsometricosxOrden3.fr3');
      frIsometricos.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));

      frIsometricos.Destroy;
      dsConcentradoIsometricos.Destroy;
      QryConcentradoIsometricos.Destroy;
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir Concentrado de Isometricos x Orden de Trabajo', 0);
    end;
  end;
end;
//*********************aqui estuvo carmen

procedure TfrmEstimaInstalado.tsAnexoEnter(Sender: TObject);
begin
  tsAnexo.Color := global_Color_entrada
end;

procedure TfrmEstimaInstalado.tsAnexoExit(Sender: TObject);

begin
  tsAnexo.Color := global_color_salida
end;

procedure TfrmEstimaInstalado.tsAnexoKeyPress(Sender: TObject;
  var Key: Char);

begin
  if Key = #13 then
  begin
    if tiNumeroEstimacion.Enabled then
      tiNumeroEstimacion.SetFocus
    else
      txtTipoMoneda.SetFocus
  end;
end;


procedure TfrmEstimaInstalado.txtTipoMonedaEnter(Sender: TObject);
begin
  txtTipomoneda.Color := global_Color_entrada
end;

procedure TfrmEstimaInstalado.txtTipoMonedaExit(Sender: TObject);

begin
  txtTipomoneda.Color := global_color_salida
end;

procedure TfrmEstimaInstalado.txtTipoMonedaKeyPress(Sender: TObject;
  var Key: Char);

begin
  if Key = #13 then
  begin
    if tdfechafinal.Enabled then
      tdFechaInicial.SetFocus
    else
      tmComentariosGenerador.SetFocus;
  end;
end;

procedure TfrmEstimaInstalado.tdCantidadExit(Sender: TObject);
begin
  tdCantidad.Color := global_color_salida
end;
//********************************************************

procedure TfrmEstimaInstalado.SemanalCImportes1Click(Sender: TObject);
begin
  try
    if not grid_generadores.DataSource.DataSet.IsEmpty then
    begin
      global_Caratula := 'MN';
      if GridPartidas.Fields[0].Value = NULL then
        MessageDlg(' No Existen partidas Cargadas al Generador ', mtWarning, [mbOk], 0)
      else
      begin
        if Estimaciones.RecordCount > 0 then
        begin
          frmBarra1.btnCancel.Click;
          if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
            procSemanalSinConImportes(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, 'Con Importes', frmEstimaInstalado, frGenerador.OnGetValue, True)
          else
            MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
        end;
      end
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al click en Semanal C/Importes M.N.', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.SemanalCImportesDLL1Click(Sender: TObject);
begin
  try
    if grid_generadores.DataSource.DataSet.IsEmpty = false then
    begin
      global_Caratula := 'DLL';
      if GridPartidas.Fields[0].Value = NULL then
        MessageDlg(' No Existen partidas Cargadas al Generador ', mtWarning, [mbOk], 0)
      else
      begin
        if Estimaciones.RecordCount > 0 then
        begin
          frmBarra1.btnCancel.Click;
          if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
            procSemanalSinConImportes(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, 'Con Importes', frmEstimaInstalado, frGenerador.OnGetValue, True)
          else
            MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
        end
      end;
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al click en Semanal C/Importes DLL', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.btnImprimirIsoRefClick(Sender: TObject);
var
  frIsometricos: TfrxReport;
begin
  try
    panelIsometrico.Visible := False;
    if grid_generadores.DataSource.DataSet.IsEmpty = false then
    begin
      qryIsometricoReferencia.Active := False;
      qryIsometricoReferencia.Params.ParamByName('Contrato').DataType := ftString;
      qryIsometricoReferencia.Params.ParamByName('Contrato').Value := global_contrato;
      qryIsometricoReferencia.Params.ParamByName('Convenio').DataType := ftString;
      qryIsometricoReferencia.Params.ParamByName('Convenio').Value := global_convenio;
      qryIsometricoReferencia.Params.ParamByName('Orden').DataType := ftString;
      qryIsometricoReferencia.Params.ParamByName('Orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
      qryIsometricoReferencia.Params.ParamByName('IsoReferencia').DataType := ftString;
      qryIsometricoReferencia.Params.ParamByName('IsoReferencia').Value := IsoReferencia.KeyValue;
      qryIsometricoReferencia.Open;

      rDiarioFirmas(global_contrato, estimaciones.FieldValues['sNumeroOrden'], 'A', estimaciones.FieldValues['dFechaFinal'], frmEstimaInstalado);

      frIsometricos := tFrxReport.Create(Self);
      frIsometricos.DataSets.Add(connection.rpt_contrato);
      frIsometricos.DataSets.Add(connection.rpt_setup);
      frIsometricos.DataSets.Add(dsIsometricoReferencia);
      frIsometricos.OnGetValue := frGenerador.OnGetValue;
      frIsometricos.PreviewOptions.MDIChild := False;
      frIsometricos.PreviewOptions.Modal := True;
      frIsometricos.PreviewOptions.Maximized := lCheckMaximized();
      frIsometricos.PreviewOptions.ShowCaptions := False;
      frIsometricos.Previewoptions.ZoomMode := zmPageWidth;
      frIsometricos.LoadFromFile(global_files + 'IsometricosReferenciaxGenerador.fr3');
      frIsometricos.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));

      frIsometricos.Destroy;
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al Imprimir Isometrico de Referencia', 0);
    end;
  end;

end;

procedure TfrmEstimaInstalado.btnPlanosClick(Sender: TObject);
begin
    if tsIsometricoReferencia.Enabled = False then
    begin
        messageDLG('Debe Agregar una Partida al Generador', mtInformation, [mbOk], 0);
        exit;
    end;
    global_frmActivo := 'frm_estimainstalado';
    Application.CreateForm(TfrmIsometricos, frmIsometricos);
    frmIsometricos.show;
end;

procedure TfrmEstimaInstalado.btnProveedoresClick(Sender: TObject);
begin
    if tsInstalacion.Enabled = False then
    begin
        messageDLG('Debe Agregar una Partida al Generador', mtInformation, [mbOk], 0);
        exit;
    end;         
    global_frmActivo := 'frm_estimainstalado';
    Application.CreateForm(Tfrmtrinomios, frmtrinomios);
    frmtrinomios.show;
end;

procedure TfrmEstimaInstalado.ConcentradodeIsometricosdeReferenciaxGenerador1Click(
  Sender: TObject);
begin
  panelIsometrico.Visible := True;
end;


procedure TfrmEstimaInstalado.cmdSemanalClick(Sender: TObject);
var
  iSemanaInicio, iSemanaFinal, w, j: Integer;
  sFechaInicio, sFechaFinal, tmpsFechaFinal: string;
  lContinuar: Boolean;
  qry, qry2: TZReadOnlyQuery;
begin
  try

    if Opcion = 'Personal' then
      procCaratulaPersonalEquipo(global_contrato, tsNumeroOrden.KeyValue, Global_Convenio, Opcion, Fi.DateTime, Ff.DateTime, frmEstimaInstalado, frGenerador.OnGetValue);

    if Opcion = 'Equipo' then
      procCaratulaPersonalEquipo(global_contrato, tsNumeroOrden.KeyValue, Global_Convenio, Opcion, Fi.DateTime, Ff.DateTime, frmEstimaInstalado, frGenerador.OnGetValue);

    if Opcion = 'Personalxoptativa' then
      procCaratulaPersonalEquipo(global_contrato, tsNumeroOrden.KeyValue, Global_Convenio, Opcion, Fi.DateTime, Ff.DateTime, frmEstimaInstalado, frGenerador.OnGetValue);

    if Opcion = 'Equipoxoptativa' then
      procCaratulaPersonalEquipo(global_contrato, tsNumeroOrden.KeyValue, Global_Convenio, Opcion, Fi.DateTime, Ff.DateTime, frmEstimaInstalado, frGenerador.OnGetValue);

    if Opcion = 'Pernoctas' then
      procCaratulaPersonalEquipo(global_contrato, tsNumeroOrden.KeyValue, Global_Convenio, Opcion, Fi.DateTime, Ff.DateTime, frmEstimaInstalado, frGenerador.OnGetValue);

    if Opcion = 'barcoxoptativas' then
      procCaratulaPersonalEquipo(global_contrato, tsNumeroOrden.KeyValue, Global_Convenio, Opcion, Fi.DateTime, Ff.DateTime, frmEstimaInstalado, frGenerador.OnGetValue);

    if Opcion = 'barcoxtotaloptativas' then
      procCaratulaPersonalEquipo(global_contrato, tsNumeroOrden.KeyValue, Global_Convenio, Opcion, Fi.DateTime, Ff.DateTime, frmEstimaInstalado, frGenerador.OnGetValue);

    if Opcion = 'barcoxtotalprogramadas' then
      procCaratulaPersonalEquipo(global_contrato, tsNumeroOrden.KeyValue, Global_Convenio, Opcion, Fi.DateTime, Ff.DateTime, frmEstimaInstalado, frGenerador.OnGetValue);

    if (Opcion = 'Semanal') then
    begin
      qry := TZReadOnlyQuery.Create(Self);
      qry2 := TZReadOnlyQuery.Create(Self);
      qry.Connection := connection.zConnection;
      qry2.Connection := connection.zConnection;

      memoria.Active := true;
      memoria.Open;
      memoria.EmptyTable;

      qry.Active := False;
      qry.SQL.Clear;
      qry.SQL.Add('select :FechaI as FI, :FechaF as FF');
      qry.Params.ParamByName('FechaI').DataType := ftDate;
      qry.Params.ParamByName('FechaI').Value := Fi.Date;
      qry.Params.ParamByName('FechaF').DataType := ftDate;
      qry.Params.ParamByName('FechaF').Value := Ff.Date;
      qry.Open;
      if qry.RecordCount > 0 then
      begin
        sFechaInicio := qry.FieldValues['FI'];
        sFechaFinal := qry.FieldValues['FF'];
        tmpsFechaFinal := sFechaFinal;
      end;

      j := 1;

      qry.Active := False;
      qry.SQL.Clear;
      qry.SQL.Add('select WEEK(:FechaI) as iSemanaInicio,WEEK(:FechaF) as iSemanaFinal');
      qry.Params.ParamByName('FechaI').DataType := ftDate;
      qry.Params.ParamByName('FechaI').Value := sFechaInicio;
      qry.Params.ParamByName('FechaF').DataType := ftDate;
      qry.Params.ParamByName('FechaF').Value := sFechaFinal;
      qry.Open;

      sFechaFinal := sFechaInicio;
      if qry.RecordCount > 0 then
      begin
        iSemanaInicio := qry.FieldValues['iSemanaInicio'];
        iSemanaFinal := qry.FieldValues['iSemanaFinal'];
      end;
  //INICIA FOR
      for w := iSemanaInicio to iSemanaFinal do
      begin
        sFechaInicio := sFechaFinal;
        lContinuar := True;
        while lContinuar and (sFechaFinal <= tmpsFechaFinal) do
        begin
          qry.Active := false;
          qry.SQL.Clear;
          qry.SQL.Add('select WEEK(:Fecha) as iSemana');
          qry.Params.ParamByName('Fecha').DataType := ftDate;
          qry.Params.ParamByName('Fecha').Value := sFechaFinal;
          qry.Open;
          if qry.RecordCount > 0 then
          begin
            if qry.FieldValues['iSemana'] = w then
            begin
              qry2.Active := False;
              qry2.SQL.Clear;
              qry2.SQL.Add('select adddate(:Fecha,1) as dNuevaFechaFinal');
              qry2.Params.ParamByName('Fecha').DataType := ftDate;
              qry2.Params.ParamByName('Fecha').Value := sFechaFinal;
              qry2.Open;
              if qry2.RecordCount > 0 then
                sFechaFinal := qry2.FieldValues['dNuevaFechaFinal'];
              lContinuar := true;
            end
            else
              lContinuar := false;

          end
          else
            lContinuar := false;

        end; {end while}
        if w <> iSemanaInicio then
        begin
          qry2.Active := False;
          qry2.SQL.Clear;
          qry2.SQL.Add('select adddate(:Fecha,1) as sFechaInicio');
          qry2.Params.ParamByName('Fecha').DataType := ftDate;
          qry2.Params.ParamByName('Fecha').Value := sFechaInicio;
          qry2.Open;
          if qry2.RecordCount > 0 then
            sFechaInicio := qry2.FieldValues['sFechaInicio'];
        end;
        if w = iSemanaFinal then
        begin
          sFechaFinal := tmpsFechaFinal;
        end;

    //Memo1.Text := Memo1.Text + '  ,  ' + sFechaInicio + ' < al >' + sFechaFinal;
    {llenar el memory data}
        qry2.Active := false;
        qry2.SQL.Clear;
        if (Global_Optativa = 'PROGRAMADA') or (Global_Optativa = 'MIXTA') then
          qry2.SQL.Add('select   ' +
            ' a.sContrato, ' +
            ' a.sNumeroOrden, ' +
            ' a.sNumeroActividad, ' +
            ' a.mDescripcion, ' +
            ' a.sMedida, a.iItemOrden, ' +
            ' a.dVentaMN, ' +
            ' a.dVentaDLL, ' +
            ' e.sNumeroGenerador, ' +
            ' es.iNumeroEstimacion,' +
            ' sum( e.dCantidad ) as dInstalado ' +
            ' from actividadesxorden a ' +
            ' inner join estimacionxpartida e on( ' +
            '   a.sContrato = e.sContrato and a.sNumeroActividad = e.sNumeroActividad ' +
            '   and a.sWbs = e.sWbsContrato and a.sTipoActividad = "Actividad" ) ' +
            ' inner join estimaciones es on( es.sContrato = e.sContrato and es.sNumeroGenerador = e.sNumeroGenerador ) ' +
            '   where ' +
            '   a.sContrato =:Contrato ' +
            '   and a.sIdConvenio =:Convenio ' +
            '   and a.sNumeroOrden <> "PATIO-AKC" ' +
            '   and es.dFechaInicio >=:FechaInicio ' +
            '   and es.dFechaInicio  <=:FechaFinal ' +
            '   group by a.iItemOrden, a.sNumeroActividad ' +
            '   Order By a.iItemOrden, a.sNumeroActividad')
        else
          qry2.SQL.Add('select   ' +
            ' a.sContrato, ' +
            ' a.sNumeroActividad, ' +
            ' a.mDescripcion, ' +
            ' a.sMedida, a.iItemOrden, ' +
            ' a.dVentaMN, ' +
            ' a.dVentaDLL, ' +
            ' e.sNumeroGenerador, ' +
            ' es.iNumeroEstimacion,' +
            ' sum( e.dCantidad ) as dInstalado ' +
            ' from actividadesxanexo a ' +
            ' inner join estimacionxpartida e on( ' +
            '   a.sContrato = e.sContrato and a.sNumeroActividad = e.sNumeroActividad ' +
            '   and a.sWbs = e.sWbsContrato and a.sTipoActividad = "Actividad" ) ' +
            ' inner join estimaciones es on( es.sContrato = e.sContrato and es.sNumeroGenerador = e.sNumeroGenerador ) ' +
            '   where ' +
            '   a.sContrato =:Contrato ' +
            '   and a.sIdConvenio =:Convenio ' +
            '   and es.dFechaInicio >=:FechaInicio ' +
            '   and es.dFechaInicio  <=:FechaFinal ' +
            '   group by a.iItemOrden, a.sNumeroActividad ' +
            '   Order By a.iItemOrden, a.sNumeroActividad');


        qry2.Params.ParamByName('Contrato').DataType := ftString;
        qry2.Params.ParamByName('Contrato').Value := global_contrato;

        qry2.Params.ParamByName('Convenio').DataType := ftString;
        qry2.Params.ParamByName('Convenio').Value := global_convenio;

        qry2.Params.ParamByName('FechaInicio').DataType := ftDate;
        qry2.Params.ParamByName('FechaInicio').Value := sFechaInicio;

        qry2.Params.ParamByName('FechaFinal').DataType := ftDate;
        qry2.Params.ParamByName('FechaFinal').Value := sFechaFinal;

        qry2.Open;
        if qry2.RecordCount <= 0 then
        begin
          memoria.Append;
          memoria.FieldValues['iSemana'] := j;
          memoria.FieldValues['sContrato'] := global_contrato;
          memoria.FieldValues['sNumeroOrden'] := Estimaciones.FieldValues['sNumeroOrden'];
          memoria.FieldValues['sNumeroActividad'] := '-';
          memoria.FieldValues['mDescripcion'] := '-';
          memoria.FieldValues['sMedida'] := '-';
          memoria.FieldValues['dVentaMN'] := 0;
          memoria.FieldValues['dVentaDLL'] := 0;
          memoria.FieldValues['sNumeroGenerador'] := '-';
          memoria.FieldValues['dInstalado'] := 0;
          memoria.FieldValues['Fi'] := Fi.Date;
          memoria.FieldValues['Ff'] := Ff.Date;
          memoria.Post;
        end;
        while not qry2.Eof do
        begin

          memoria.Append;
          memoria.FieldValues['iSemana'] := j;
          memoria.FieldValues['sContrato'] := global_contrato;
          memoria.FieldValues['sNumeroOrden'] := Estimaciones.FieldValues['sNumeroOrden'];
          memoria.FieldValues['sNumeroActividad'] := qry2.FieldValues['sNumeroActividad'];
          memoria.FieldValues['mDescripcion'] := qry2.FieldValues['mDescripcion'];
          memoria.FieldValues['sMedida'] := qry2.FieldValues['sMedida'];
          memoria.FieldValues['dVentaMN'] := qry2.FieldValues['dVentaMN'];
          memoria.FieldValues['dVentaDLL'] := qry2.FieldValues['dVentaDLL'];
          memoria.FieldValues['sNumeroGenerador'] := qry2.FieldValues['sNumeroGenerador'];
          memoria.FieldValues['dInstalado'] := qry2.FieldValues['dInstalado'];
          memoria.FieldValues['Fi'] := Fi.Date;
          memoria.FieldValues['Ff'] := Ff.Date;
          memoria.Post;

          qry2.Next;
        end;

        j := j + 1;
      end;
    end; //TERMINA FOR

    panelSemanal.Visible := False;

    fechas.Active := true;
    fechas.Open;
    fechas.EmptyTable;
    fechas.Append;
    fechas.FieldValues['fi'] := Fi.Date;
    fechas.FieldValues['ff'] := Ff.Date;
    fechas.Post;
    try
      rDiarioFirmas(global_contrato, estimaciones.FieldValues['sNumeroOrden'], 'A', estimaciones.FieldValues['dFechaFinal'], frmEstimaInstalado);
    except
      on e: exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al generar firmas al imprimir resumen mensual', 0);
      end;
    end;
    frxReport1.DataSets.Add(connection.rpt_contrato);
    frxReport1.DataSets.Add(connection.rpt_setup);
    frxReport1.DataSets.Add(frxDBDataset1);
    frxReport1.DataSets.Add(dsFechas);
    frxReport1.OnGetValue := frGenerador.OnGetValue;
    frxReport1.PreviewOptions.MDIChild := False;
    frxReport1.PreviewOptions.Modal := True;
    frxReport1.PreviewOptions.Maximized := lCheckMaximized();
    frxReport1.PreviewOptions.ShowCaptions := False;
    frxReport1.Previewoptions.ZoomMode := zmPageWidth;
    frxReport1.LoadFromFile(global_files + 'reporteActividadesxSemanaxGenerador.fr3');
    frxReport1.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir resumen mensual', 0);
    end;
  end;
end;


procedure TfrmEstimaInstalado.R2Click(Sender: TObject);
begin
  if grid_generadores.DataSource.DataSet.IsEmpty = true then
  begin
    Opcion := 'Semanal';
    panelSemanal.Visible := True;
  end;
end;

procedure TfrmEstimaInstalado.txtMarcaEnter(Sender: TObject);
begin
  txtMarca.Color := Global_Color_Entrada;
end;

procedure TfrmEstimaInstalado.txtMarcaExit(Sender: TObject);
begin
  txtMarca.Color := Global_Color_Salida;
end;

procedure TfrmEstimaInstalado.txtMarcaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
    txtSubMarca.SetFocus;
end;

procedure TfrmEstimaInstalado.txtSubMarcaEnter(Sender: TObject);
begin
  txtSubMarca.Color := Global_color_entrada;
end;

procedure TfrmEstimaInstalado.txtSubMarcaExit(Sender: TObject);
begin
  txtSubMarca.Color := Global_Color_Salida;
end;

procedure TfrmEstimaInstalado.txtLongAreaEnter(Sender: TObject);
begin
  txtLongArea.Color := Global_Color_Entrada;
end;

procedure TfrmEstimaInstalado.txtLongAreaExit(Sender: TObject);
begin
  txtLongArea.Color := Global_Color_Salida;
  if txtLongArea.Text = '' then
    txtLongArea.Text := '0';
end;

procedure TfrmEstimaInstalado.txtLongTotalEnter(Sender: TObject);
begin
  txtLongTotal.Color := Global_Color_Entrada;
end;

procedure TfrmEstimaInstalado.txtLongTotalExit(Sender: TObject);
begin
  txtLongTotal.Color := Global_Color_salida;
end;

procedure TfrmEstimaInstalado.txtPesoxUnidadEnter(Sender: TObject);
begin
  txtPesoxUnidad.Color := Global_Color_Entrada;
  txtLongTotal.Text := CurrToStr(txtLongArea.Value * txtPzas.Value);
  if txtPesoxUnidad.Value > 0 then
    txtPesoTotal.Text := CurrToStr(txtLongTotal.Value * txtPesoxUnidad.Value);
end;

procedure TfrmEstimaInstalado.txtPesoxUnidadExit(Sender: TObject);
begin
  txtPesoxUnidad.Color := Global_Color_Salida;
  if txtPesoxUnidad.Value > 0 then
  begin
    if (ActividadesIguales.FieldValues['sMedida'] = 'TON') or (ActividadesIguales.FieldValues['sMedida'] = 'ton') then
      txtPesoTotal.Text := CurrToStr((txtLongTotal.Value * txtPesoxUnidad.Value) / 1000);
  end;
end;

procedure TfrmEstimaInstalado.txtPesoTotalEnter(Sender: TObject);
begin
  txtPesoTotal.Color := Global_Color_entrada;
  txtLongTotal.Text := CurrToStr(txtLongArea.Value * txtPzas.Value);
  if txtPesoxUnidad.text = '0' then
    txtPesoxUnidad.Value := 0;

  if txtPesoxUnidad.Enabled = True then
  begin
    if (ActividadesIguales.FieldValues['sMedida'] = 'TON') or (ActividadesIguales.FieldValues['sMedida'] = 'ton') then
      txtPesoTotal.Text := CurrToStr((txtLongTotal.value * txtPesoxUnidad.value) / 1000)
    else
      txtPesoTotal.Text := CurrToStr(txtLongTotal.Value * txtPesoxUnidad.Value);
  end
  else
    txtPesoTotal.Text := txtLongTotal.Text;

  if txtLongArea.Enabled = False then
  begin
    txtPesoTotal.Text := txtPzas.Text;
  end;


end;

procedure TfrmEstimaInstalado.txtPesoTotalExit(Sender: TObject);
begin
  txtPesoTotal.Color := Global_color_salida;
end;



procedure TfrmEstimaInstalado.txtSubMarcaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
  begin
    if txtLongArea.Enabled = true then
      txtLongArea.SetFocus
    else
      txtpzas.SetFocus;
  end;
end;

procedure TfrmEstimaInstalado.txtTagEnter(Sender: TObject);
begin
  txtTag.Color := global_color_entrada
end;

procedure TfrmEstimaInstalado.txtTagExit(Sender: TObject);
begin
  txtTag.Color := global_color_salida
end;

procedure TfrmEstimaInstalado.txtLongAreaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
  begin
    if txtLongArea.Value = 0 then
    begin
      MessageDlg('Esrciba un valor para Longitud', mtWarning, [mbOk], 0);
    end
    else
    begin
      txtPzas.SetFocus;
    end;
  end;
end;

procedure TfrmEstimaInstalado.txtLongTotalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
  begin
    if txtPesoxUnidad.Enabled = False then
    begin
      txtPesoTotal.SetFocus;
      txtPesoTotal.Text := txtLongTotal.Text;
    end
    else
      txtPesoxUnidad.SetFocus
  end;
end;

procedure TfrmEstimaInstalado.NumGenDespiezadosClick(Sender: TObject);
begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if Connection.Configuracion.FieldValues['sGenDesp'] = 'Despiezado' then
      begin
        if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
          procNumeroGeneradorDespiezado(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, 'Cliente', frmEstimaInstalado, frGenerador.OnGetValue, True)
        else
          MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
      end;
      if (Connection.Configuracion.FieldValues['sGenDesp'] = 'Normales') or (Connection.Configuracion.FieldValues['sGenDesp'] = 'Detallado') then
      begin
        if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
          procNumeroGenerador(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, 'Cliente', frmEstimaInstalado, frGenerador.OnGetValue, True)
        else
          MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' + chr(13) + ' Partidas Adicionales ' + sPartidas, mtWarning, [mbOk], 0);
      end;
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Estimaciones', 'Al click en Numeros Generadores', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.txtPzasKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
  begin
    txtLongTotal.Text := CurrToStr(txtLongArea.Value * txtPzas.Value);
    if txtPesoxUnidad.Enabled = False then
    begin
      txtPesoTotal.Text := txtLongTotal.Text;
      txtPesoTotal.SetFocus;
    end
    else
      txtPesoxUnidad.SetFocus;
  end;

end;

procedure TfrmEstimaInstalado.txtPzasExit(Sender: TObject);
begin
  txtPzas.Color := Global_Color_Salida;
end;

procedure TfrmEstimaInstalado.txtPzasEnter(Sender: TObject);
begin
  txtPzas.Color := Global_Color_Entrada;
end;

procedure TfrmEstimaInstalado.txtPesoxUnidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (Key = #13) or (Key = #9) then
  begin
    if txtPesoxUnidad.Value > 0 then
    begin
      txtPesoTotal.Text := CurrToStr(txtLongTotal.value * txtPesoxUnidad.value);
      txtPesoTotal.SetFocus;
    end
    else
    begin
      txtPesoTotal.Text := txtLongTotal.Text;
      txtPesoTotal.SetFocus;
    end;

  end;
end;

procedure TfrmEstimaInstalado.GeneradorBarco1Click(Sender: TObject);
begin
  Opcion := 'Barco';
  gpTitulo.Caption := 'Rango de Fechas General de Barco';
  panelSemanal.Visible := True;
end;

procedure TfrmEstimaInstalado.GenIPRClick(Sender: TObject);
begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
        procGeneradorConversiones(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, 'ReportePerimetro', False);
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.GenDespiezadoClick(Sender: TObject);

begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
        procGeneradorConversiones(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, 'DespiezadoGeneral', False);
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.GeneradorEquipo1Click(Sender: TObject);
begin
  panelSemanal.Visible := True;
  gpTitulo.Caption := 'Rango de Fechas de Equipos Por Plataformas';
  Opcion := 'Equipo';
end;

procedure TfrmEstimaInstalado.GenAnguloClick(Sender: TObject);
begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
        procGeneradorConversiones(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, 'ReporteAngulo', False);
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir', 0);
    end;
  end;
end;

procedure TfrmEstimaInstalado.GeneradorPersonal1Click(Sender: TObject);
begin
  Opcion := 'Personal';
  gpTitulo.Caption := 'Rango de Fechas de Personal Por Plataformas';
  PanelSemanal.Visible := True;
end;

procedure TfrmEstimaInstalado.panelSemanalClick(Sender: TObject);
begin
  Panel.Visible := False;
end;

procedure TfrmEstimaInstalado.EquipoXOptativa1Click(Sender: TObject);
begin
  panelSemanal.Visible := True;
  gpTitulo.Caption := 'Rango de Fechas de Equipos Por Optativas';
  Opcion := 'Equipoxoptativa';
end;

procedure TfrmEstimaInstalado.PersonalXOptativa1Click(Sender: TObject);
begin
  panelSemanal.Visible := True;
  gpTitulo.Caption := 'Rango de Fechas de Personal Por Optativas';
  Opcion := 'Personalxoptativa';
end;

procedure TfrmEstimaInstalado.PopGeneradorPopup(Sender: TObject);
begin
  if Connection.Configuracion.FieldValues['sGenDesp'] = 'Despiezado' then
  begin
    PopGenerador.Items[0].Visible := False;
    PopGenerador.Items[7].Visible := False;
    PopGenerador.Items[10].Visible := False;
    PopGenerador.Items[11].Visible := False;
    PopGenerador.Items[12].Visible := False;
    PopGenerador.Items[13].Visible := False;
    PopGenerador.Items[14].Visible := False;
  end;

  if Connection.Configuracion.FieldValues['sGenDesp'] = 'Detallado' then
  begin
      GenDespiezado.Visible := True;
      GenTuberia.Visible    := True;
      GenIPR.Visible        := True;
      GenAngulo.Visible     := True;
  end;
end;

procedure TfrmEstimaInstalado.btSalirClick(Sender: TObject);
begin
  PanelSemanal.Visible := False;
end;

procedure TfrmEstimaInstalado.PernoctaXPlataforma1Click(Sender: TObject);
begin
  panelSemanal.Visible := True;
  gpTitulo.Caption := 'Fechas de Pernoctas Por Optativas';
  Opcion := 'Pernoctas'
end;

procedure TfrmEstimaInstalado.ActividadesIgualesAfterScroll(DataSet: TDataSet);
begin
  if ActividadesIguales.State <> dsInactive then
     // if tsNumeroActividad.ReadOnly = False then
    if ActividadesIguales.RecordCount > 0 then
    begin
      Grid_Iguales.Hint := ActividadesIguales.FieldValues['mDescripcion'];
      Paquete.Active := False;
      Paquete.Params.ParamByName('contrato').DataType := ftString;
      Paquete.Params.ParamByName('contrato').Value := global_contrato;
      Paquete.Params.ParamByName('orden').DataType := ftString;
      Paquete.Params.ParamByName('orden').Value := tsNumeroOrden.Text;
      Paquete.Params.ParamByName('wbs').DataType := ftString;
      Paquete.Params.ParamByName('wbs').Value := ActividadesIguales.FieldValues['sWbsAnterior'];
      Paquete.Open;

      if Paquete.RecordCount > 0 then
      begin
        pdPaquete.Caption := '  ' + Paquete.FieldValues['sNumeroActividad'] + ' .- ' + Paquete.FieldValues['mDescripcion'];
        pdPaquete.Hint := '  ' + Paquete.FieldValues['sNumeroActividad'] + ' .- ' + Paquete.FieldValues['mDescripcion'];
      end
      else
      begin
        pdPaquete.Caption := '< < Seleccione un Paquete > >';
        pdPaquete.Hint := '< < Seleccione un Paquete > >'
      end;
    end;
end;

procedure TfrmEstimaInstalado.Actual1Click(Sender: TObject);
begin
  FrmBarra2.btnRefresh.Click;
end;

procedure TfrmEstimaInstalado.actualizadatos1Click(Sender: TObject);
begin
  try
    Application.CreateForm(TfrmDespieceImagen, frmDespieceImagen);
    frmDespieceImagen.ShowModal;
  finally
    frmDespieceImagen.Free;
  end;
end;

procedure TfrmEstimaInstalado.Barco1Click(Sender: TObject);
begin
  panelSemanal.Visible := True;
  gpTitulo.Caption := 'Fechas de Barco Por Optativas X Plataformas';
  Opcion := 'barcoxoptativas'
end;

procedure TfrmEstimaInstalado.BarcoPorPlataformas1Click(Sender: TObject);
begin
  panelSemanal.Visible := True;
  gpTitulo.Caption := 'Fechas de Barco Por Optativas';
  Opcion := 'barcoxoptativas'
end;

procedure TfrmEstimaInstalado.BarcoPorTotalOptativas1Click(
  Sender: TObject);
begin
  panelSemanal.Visible := True;
  gpTitulo.Caption := 'Fechas de Barco Por Total Optativas';
  Opcion := 'barcoxtotaloptativas';
end;

procedure TfrmEstimaInstalado.BarcoPorTotalProgramadas1Click(
  Sender: TObject);
begin
  panelSemanal.Visible := True;
  gpTitulo.Caption := 'Fechas de Barco Por Total Programadas';
  Opcion := 'barcoxtotalprogramadas';
end;


procedure TfrmEstimaInstalado.GeneradorSemana(sParamTipo: string; sParamFrente: string);
var
  iSemanaInicio, iSemanaFinal, w, j, x, total: Integer;
  sFechaInicio,
    sFechaFinal,
    tmpsFechaFinal,
    TipoE,
    Actividad: string;
  lContinuar: Boolean;
  qry,
    qry2,
    QryConfiguracion: TZReadOnlyQuery;
  c1, c2, c3, c4, c5, i: Double;
  EsMixto: string;
  cbander: boolean;
  cadena, cadena2, cadena3: string;
begin
  cbander := false;
    //Revisado por <ivan> ... 17 Sept 2010..

  QryConfiguracion := TZReadOnlyQuery.Create(self);
  QryConfiguracion.Connection := connection.zconnection;

  cadena := '';
  if sParamFrente <> '' then
    cadena := ' and e.sNumeroOrden =:Orden ';

  QryConfiguracion.Active := False;
  QryConfiguracion.SQL.Clear;
  QryConfiguracion.SQL.Add('select c.sLeyenda1, c.sLeyenda2, c.sLeyenda3, c.iFirmas, e.sIdUsuarioValida, e.sIdUsuarioAutoriza  ' +
    'From contratos c2 INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
    'inner join estimaciones e on (e.sContrato = c2.sContrato ' + cadena + 'and e.iNumeroEstimacion =:Estimacion and e.sNumeroGenerador =:Generador ) ' +
    'Where c2.sContrato = :Contrato');
  QryConfiguracion.Params.ParamByName('contrato').DataType := ftString;
  QryConfiguracion.Params.ParamByName('contrato').Value := global_contrato;
  QryConfiguracion.Params.ParamByName('Estimacion').DataType := ftString;
  QryConfiguracion.Params.ParamByName('Estimacion').Value := Estimaciones.FieldValues['iNumeroEstimacion'];
  if sParamFrente <> '' then
  begin
    QryConfiguracion.Params.ParamByName('Orden').DataType := ftString;
    QryConfiguracion.Params.ParamByName('Orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
  end;
  QryConfiguracion.Params.ParamByName('Generador').DataType := ftString;
  QryConfiguracion.Params.ParamByName('Generador').Value := Estimaciones.FieldValues['sNumeroGenerador'];
  QryConfiguracion.Open;


  if (sParamTipo = 'MN') or (sParamTipo = 'DLL') then
  begin
    qry := TZReadOnlyQuery.Create(Self);
    qry.Connection := connection.zConnection;
    qry2 := TZReadOnlyQuery.Create(Self);
    qry2.Connection := connection.zConnection;

    qry2.Active := False;
    qry2.SQL.Clear;
    qry2.SQL.Add('select dFechaInicio, dFechaFinal, sIdTipoEstimacion from estimacionperiodo where sContrato =:Contrato and iNumeroEstimacion =:Estimacion');
    qry2.Params.ParamByName('Contrato').DataType := ftString;
    qry2.Params.ParamByName('Contrato').Value := global_contrato;
    qry2.Params.ParamByName('Estimacion').DataType := ftString;
    qry2.Params.ParamByName('Estimacion').Value := tiNumeroEstimacion.KeyValue;
    qry2.Open;

    if qry2.RecordCount > 0 then
    begin
      Fi.Date := qry2.FieldValues['dFechaInicio'];
      Ff.Date := qry2.FieldValues['dFechaFinal'];
      TipoE := qry2.FieldValues['sIdTipoEstimacion'];
    end;

    Datos.Active := true;
    Datos.Open;
    Datos.EmptyTable;

    qry.Active := False;
    qry.SQL.Clear;
    qry.SQL.Add('select :FechaI as FI, :FechaF as FF');
    qry.Params.ParamByName('FechaI').DataType := ftDate;
    qry.Params.ParamByName('FechaI').Value := Fi.Date;
    qry.Params.ParamByName('FechaF').DataType := ftDate;
    qry.Params.ParamByName('FechaF').Value := Ff.Date;
    qry.Open;
    if qry.RecordCount > 0 then
    begin
      sFechaInicio := qry.FieldValues['FI'];
      sFechaFinal := qry.FieldValues['FF'];
    end;
       //INICIA FOR
    for w := 1 to 5 do
    begin
      qry2.Active := false;
      qry2.SQL.Clear;

      if ((Global_Optativa = 'PROGRAMADA') or (Global_Optativa = 'MIXTA') or (Global_Optativa = 'OPTATIVA')) and (Connection.Configuracion.FieldValues['sAnexos'] = 'No') then
      begin
                {Si es General el Resumen Mensual..}
        cadena := '';
        cadena2 := ' from actividadesxanexo a ' +
          ' inner join estimacionxpartida e on(a.sContrato = e.sContrato and a.sNumeroActividad = e.sNumeroActividad ' +
          ' and a.sWbs = e.sWbsContrato and a.sTipoActividad = "Actividad" ) ';
        cadena3 := '';
                {Si es por Frente el Resumen Mensual..}
        if sParamFrente <> '' then
        begin
          cadena := ' a.sNumeroOrden, ';
          cadena2 := ' from actividadesxorden a ' +
            ' inner join estimacionxpartida e on(a.sContrato = e.sContrato and e.sNumeroOrden =:Orden and a.sNumeroActividad = e.sNumeroActividad ' +
            ' and a.sWbs = e.sWbs and a.sTipoActividad = "Actividad" ) ';
          cadena3 := ' and a.sNumeroOrden =:Orden ';
        end;
        qry2.SQL.Add('select   ' +
          ' a.sContrato, ' +
          cadena +
          ' a.sNumeroActividad, ' +
          ' a.mDescripcion, ' +
          ' a.sMedida, a.iItemOrden, ' +
          ' a.dVentaMN, ' +
          ' a.dVentaDLL, ' +
          ' e.sNumeroGenerador, ' +
          ' es.iNumeroEstimacion,' +
          ' es.iSemana,' +
          ' sum( e.dCantidad ) as dInstalado ' +
          cadena2 +
          ' inner join estimaciones es on( es.sContrato = e.sContrato and es.sNumeroGenerador = e.sNumeroGenerador  ' +
          '                                and es.iSemana  =:W and es.iNumeroEstimacion =:Estimacion)' +
          ' inner join estimacionperiodo ep on( ep.sContrato = e.sContrato and ep.iNumeroEstimacion = es.iNumeroEstimacion ' +
          '                                and ep.sIdTipoEstimacion =:Tipo)' +
          ' where ' +
          ' a.sContrato =:Contrato ' +
          ' and a.sIdConvenio =:Convenio ' +
          cadena3 +
          ' and es.dFechaInicio >=:FechaInicio ' +
          ' and es.dFechaInicio <=:FechaFinal ' +
          ' group by a.sWbs, a.sNumeroActividad ' +
          ' Order By a.iItemOrden, a.sNumeroActividad');
        cbander := true;
      end;
      if Connection.Configuracion.FieldValues['sAnexos'] = 'Si' then
      begin
                {Si es General el Resumen Mensual..}
        cadena := '';
        cadena2 := '';
                {Si es por Frente el Resumen Mensual..}
        if sParamFrente <> '' then
        begin
          cadena := ' and es.sNumeroOrden =:Orden ';
          cadena2 := ' and e.sNumeroOrden  =:Orden ';
        end;
        qry2.SQL.Add('select   ' +
          ' a.sContrato, ' +
          ' a.sNumeroActividad, ' +
          ' a.mDescripcion, ' +
          ' a.sMedida, a.iItemOrden, ' +
          ' a.dVentaMN, ' +
          ' a.dVentaDLL, ' +
          ' e.sNumeroGenerador, ' +
          ' es.iNumeroEstimacion, ' +
          ' es.iSemana,' +
          ' sum( e.dCantidad ) as dInstalado ' +
          ' from estimacionxpartida e ' +
          ' inner join actividadesxanexo a on( ' +
          '   a.sContrato = e.sContrato and sIdConvenio=:Convenio and a.sNumeroActividad = e.sNumeroActividad ' +
          '   and a.sWbs = e.sWbsContrato and a.sTipoActividad = "Actividad" ) ' +
          ' inner join estimaciones es on( es.sContrato = e.sContrato and es.sNumeroGenerador = e.sNumeroGenerador ' +
          cadena + ' and es.iSemana  =:W  and es.iNumeroEstimacion =:Estimacion)' +
          ' inner join estimacionperiodo ep on( ep.sContrato = e.sContrato and ep.iNumeroEstimacion = es.iNumeroEstimacion ' +
          '   and ep.sIdTipoEstimacion =:Tipo)' +
          ' where ' +
          ' e.sContrato         =:Contrato ' +
          cadena2 +
          ' and es.dFechaInicio >=:FechaInicio ' +
          ' and es.dFechaInicio <=:FechaFinal ' +
          ' group by e.sWbsContrato, e.sNumeroActividad ' +
          ' Order By a.iItemOrden, a.sNumeroActividad');
        cbander := True;
      end;

      if cbander then
      begin
        qry2.Params.ParamByName('Contrato').DataType := ftString;
        qry2.Params.ParamByName('Contrato').Value := global_contrato;
        qry2.Params.ParamByName('Convenio').DataType := ftString;
        qry2.Params.ParamByName('Convenio').Value := global_convenio;
        if sParamFrente <> '' then
        begin
          qry2.Params.ParamByName('Orden').DataType := ftString;
          qry2.Params.ParamByName('Orden').Value := Estimaciones.FieldValues['sNumeroOrden'];
        end;
        qry2.Params.ParamByName('FechaInicio').DataType := ftDate;
        qry2.Params.ParamByName('FechaInicio').Value := sFechaInicio;
        qry2.Params.ParamByName('FechaFinal').DataType := ftDate;
        qry2.Params.ParamByName('FechaFinal').Value := sFechaFinal;
        qry2.Params.ParamByName('W').DataType := ftInteger;
        qry2.Params.ParamByName('W').Value := w;
        qry2.Params.ParamByName('Estimacion').DataType := ftString;
        qry2.Params.ParamByName('Estimacion').Value := Estimaciones.FieldValues['iNumeroEstimacion'];
        qry2.Params.ParamByName('Tipo').DataType := ftString;
        qry2.Params.ParamByName('Tipo').Value := TipoE;

        qry2.Open;

        while not qry2.Eof do
        begin
          Datos.Append;
          Datos.FieldValues['iSemana'] := qry2.FieldValues['iSemana'];
          Datos.FieldValues['sContrato'] := global_contrato;
          if sParamFrente <> '' then
            Datos.FieldValues['sNumeroOrden'] := Estimaciones.FieldValues['sNumeroOrden']
          else
            Datos.FieldValues['sNumeroOrden'] := '';
          Datos.FieldValues['sEstimacion'] := qry2.FieldValues['iNumeroEstimacion'];
          Datos.FieldValues['sNumeroActividad'] := qry2.FieldValues['sNumeroActividad'];
          Datos.FieldValues['mDescripcion'] := qry2.FieldValues['mDescripcion'];
          Datos.FieldValues['sMedida'] := qry2.FieldValues['sMedida'];
          Datos.FieldValues['dVentaMN'] := qry2.FieldValues['dVentaMN'];
          Datos.FieldValues['dVentaDLL'] := qry2.FieldValues['dVentaDLL'];
          Datos.FieldValues['sNumeroGenerador'] := qry2.FieldValues['sNumeroGenerador'];
          Datos.FieldValues['Fi'] := Fi.Date;
          Datos.FieldValues['Ff'] := Ff.Date;
          Datos.FieldValues['dInstalado'] := qry2.FieldValues['dInstalado'];
          Datos.Post;
          qry2.Next;
        end;
      end;
    end;

        // Datos Agreagado al final para completar funcion..
    Datos.Append;
    Datos.FieldValues['iSemana'] := 1;
    Datos.FieldValues['sNumeroActividad'] := 'www';
    Datos.FieldValues['dInstalado'] := 0;
    Datos.Post;

    memoria.Active := true;
    memoria.Open;
    memoria.EmptyTable;

    if Datos.RecordCount > 0 then
    begin
      Datos.First;
      Datos.SortOnFields('sNumeroActividad');
      i := 0;
      Datos.First;
      actividad := Datos.FieldValues['sNumeroActividad'];
      while not Datos.Eof do
      begin
        if actividad = Datos.FieldValues['sNumeroActividad'] then
        begin
          if Datos.FieldValues['iSemana'] = 1 then
            c1 := c1 + Datos.FieldValues['dInstalado'];
          if Datos.FieldValues['iSemana'] = 2 then
            c2 := c2 + Datos.FieldValues['dInstalado'];
          if Datos.FieldValues['iSemana'] = 3 then
            c3 := c3 + Datos.FieldValues['dInstalado'];
          if Datos.FieldValues['iSemana'] = 4 then
            c4 := c4 + Datos.FieldValues['dInstalado'];
          if Datos.FieldValues['iSemana'] = 5 then
            c5 := c5 + Datos.FieldValues['dInstalado'];
          i := i + Datos.FieldValues['dInstalado'];
        end
        else
        begin
          actividad := Datos.FieldValues['sNumeroActividad'];
          Datos.Prior;
          memoria.Append;
          memoria.FieldValues['iSemana'] := Datos.FieldValues['iSemana'];
          memoria.FieldValues['sContrato'] := global_contrato;
          memoria.FieldValues['sNumeroOrden'] := Datos.FieldValues['sNumeroOrden'];
          memoria.FieldValues['sEstimacion'] := Datos.FieldValues['sEstimacion'];
          memoria.FieldValues['sNumeroActividad'] := Datos.FieldValues['sNumeroActividad'];
          memoria.FieldValues['mDescripcion'] := Datos.FieldValues['mDescripcion'];
          memoria.FieldValues['sMedida'] := Datos.FieldValues['sMedida'];
          memoria.FieldValues['dVentaMN'] := Datos.FieldValues['dVentaMN'];
          memoria.FieldValues['dVentaDLL'] := Datos.FieldValues['dVentaDLL'];
          memoria.FieldValues['sNumeroGenerador'] := Datos.FieldValues['sNumeroGenerador'];
          memoria.FieldValues['Fi'] := Datos.FieldValues['Fi'];
          memoria.FieldValues['Ff'] := Datos.FieldValues['Ff'];
          memoria.FieldValues['dInstalado'] := c1;
          memoria.FieldValues['dInstalado1'] := c2;
          memoria.FieldValues['dInstalado2'] := c3;
          memoria.FieldValues['dInstalado3'] := c4;
          memoria.FieldValues['dInstalado4'] := c5;
          memoria.FieldValues['Total'] := i;
          memoria.FieldValues['iItemOrden'] := Datos.FieldValues['sNumeroGenerador'];
          memoria.Post;
          i := 0;
          c1 := 0;
          c2 := 0;
          c3 := 0;
          c4 := 0;
          c5 := 0;
        end;
        Datos.Next;
      end;
    end;
    if cbander then
      Qry2.First;
    panelSemanal.Visible := False;
    fechas.Active := true;
    fechas.Open;
    fechas.EmptyTable;
    fechas.Append;
    fechas.FieldValues['fi'] := Fi.Date;
    fechas.FieldValues['ff'] := Ff.Date;
    fechas.FieldValues['sLeyenda1'] := QryConfiguracion.FieldValues['sLeyenda1'];
    fechas.FieldValues['sLeyenda2'] := QryConfiguracion.FieldValues['sLeyenda2'];
    fechas.FieldValues['sLeyenda3'] := QryConfiguracion.FieldValues['sLeyenda3'];
    fechas.FieldValues['iFirmas'] := QryConfiguracion.FieldValues['iFirmas'];
    fechas.FieldValues['sIdUsuarioValida'] := QryConfiguracion.FieldValues['sIdUsuarioValida'];
    fechas.FieldValues['sIdUsuarioAutoriza'] := QryConfiguracion.FieldValues['sIdUsuarioAutoriza'];
    fechas.Post;

    try
      rDiarioFirmas(global_contrato, estimaciones.FieldValues['sNumeroOrden'], 'A', estimaciones.FieldValues['dFechaFinal'], frmEstimaInstalado);
    except
      on e: exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al generar firmas al salir de resumen mensual', 0);
      end;
    end;

    frxReport1.DataSets.Add(connection.rpt_contrato);
    frxReport1.DataSets.Add(connection.rpt_setup);
    frxReport1.DataSets.Add(frxDBDataset1);
    frxReport1.DataSets.Add(dsFechas);

    frxReport1.OnGetValue := frGenerador.OnGetValue;
    frxReport1.PreviewOptions.MDIChild := False;
    frxReport1.PreviewOptions.Modal := True;
    frxReport1.PreviewOptions.Maximized := lCheckMaximized();
    frxReport1.PreviewOptions.ShowCaptions := False;
    frxReport1.Previewoptions.ZoomMode := zmPageWidth;

    if sParamTipo = 'MN' then
      frxReport1.LoadFromFile(global_files + 'ResumenMensualGeneradorMN.fr3');

    if sParamTipo = 'DLL' then
      frxReport1.LoadFromFile(global_files + 'ResumenMensualGeneradorDLL.fr3');

    frxReport1.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));

  end;
  QryConfiguracion.Destroy;

end;


procedure TfrmEstimaInstalado.GenTuberiaClick(Sender: TObject);
begin
  try
    if Estimaciones.RecordCount > 0 then
    begin
      frmBarra1.btnCancel.Click;
      if lfnValidaGenerador(global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], frmEstimaInstalado) then
        procGeneradorConversiones(global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'], Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'], global_convenio, frmEstimaInstalado, frGenerador.OnGetValue, 'ReporteTuberia', False);
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Generadores de Obra', 'Al imprimir', 0);
    end;
  end;
end;

end.

