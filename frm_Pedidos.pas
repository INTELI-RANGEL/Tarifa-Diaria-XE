unit frm_Pedidos;

interface

uses
  Windows, Messages, Math, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, DB, ADODB, DBCtrls, global, strUtils,
  Mask, OleCtrls, Grids, DBGrids, frm_barra, ExtCtrls, Utilerias,
  Menus, frxClass, frxDBSet, RXDBCtrl, RxLookup, unitactivapop,
  RXCtrls, CheckLst, RxMemDS, ExcelXP, OleServer, 
  Excel2000, ZAbstractRODataset, ZDataset, ZAbstractDataset, rxCurrEdit,
  rxToolEdit, JvExMask, JvToolEdit, JvCombobox, NxCollection,
  AdvGlowButton,udbgrid, unitexcepciones, unittbotonespermisos,UnitValidaTexto,
  UFunctionsGHH, UnitValidacion, AdvGrid, AsgLinks;

type
  TfrmPedidos = class(TForm)
    ds_anexo_ocompras: TDataSource;
    ds_anexo_prequisicion: TDataSource;
    PgControl: TPageControl;
    TabSheet2: TTabSheet;
    frxDBReporte: TfrxDBDataset;
    GroupBox3: TGroupBox;
    frmBarra2: TfrmBarra;
    TabSheet1: TTabSheet;
    Grid_Entradas: TRxDBGrid;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    N3: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    N4: TMenuItem;
    tsPlataforma: TLabel;
    imgNotas: TImage;
    frmBarra1: TfrmBarra;
    ComentariosAdicionales: TMenuItem;
    ds_proveedores: TDataSource;
    OpenXLS: TOpenDialog;
    ExcelApplication1: TExcelApplication;
    ExcelWorkbook1: TExcelWorkbook;
    ExcelWorksheet1: TExcelWorksheet;
    reporte: TZQuery;
    anexo_pocompras: TZQuery;
    Proveedores: TZReadOnlyQuery;
    anexo_pocomprasdMontoMN: TFloatField;
    anexo_pocomprassContrato: TStringField;
    anexo_pocomprasdCantidad: TFloatField;
    anexo_pocomprasiFolioPedido: TIntegerField;
    anexo_pocomprassNumeroActividad: TStringField;
    anexo_pocomprasiItem: TIntegerField;
    anexo_pocomprasmDescripcion: TMemoField;
    anexo_pocomprassMedida: TStringField;
    frxEntrada: TfrxReport;
    anexo_pocomprassIdInsumo: TStringField;
    anexo_pocomprasdCosto: TFloatField;
    anexo_pocomprassNumeroOrden: TStringField;
    anexo_pocomprassDescripcion: TStringField;
    requisiciones: TZReadOnlyQuery;
    ds_requisiciones: TDataSource;
    NxHeaderPanel1: TNxHeaderPanel;
    Label7: TLabel;
    Label3: TLabel;
    tdIdFecha: TDateTimePicker;
    Label15: TLabel;
    dbRequisicion: TJvCheckedComboBox;
    Label4: TLabel;
    tsIdProveedor: TRxDBLookupCombo;
    Label17: TLabel;
    tsNumeroOrden: TComboBox;
    Label6: TLabel;
    tmComentarios: TMemo;
    NxHeaderPanel2: TNxHeaderPanel;
    lblPartida: TLabel;
    GridPartidas: TRxDBGrid;
    GroupBox1: TGroupBox;
    Label5: TLabel;
    Label9: TLabel;
    ds_FormaPago: TDataSource;
    FormaPago: TZReadOnlyQuery;
    tsFormaPago: TRxDBLookupCombo;
    Grid_Insumos: TDBGrid;
    dbPartidas: TDBLookupComboBox;
    dtsPartidas: TDataSource;
    txtMaterial: TEdit;
    ds_insumo: TDataSource;
    Insumos: TZReadOnlyQuery;
    rxSeguimiento_Mat: TRxMemoryData;
    rxSeguimiento_MatsContrato: TStringField;
    rxSeguimiento_MatPartida: TStringField;
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
    rxSeguimiento_MatFolioReq: TIntegerField;
    rxSeguimiento_MatItemReq: TIntegerField;
    rxSeguimiento_MatdCantidadReq: TFloatField;
    rxSeguimiento_MatdRestanteReq: TFloatField;
    rxSeguimiento_MatdExcedenteReq: TFloatField;
    rxSeguimiento_MatdPorcentajeReq: TFloatField;
    rxSeguimiento_MatdPorcentajeReq_T: TFloatField;
    rxSeguimiento_MatFolioOC: TIntegerField;
    rxSeguimiento_MatItemOC: TIntegerField;
    rxSeguimiento_MatdCantidadOC: TFloatField;
    rxSeguimiento_MatdRestanteOC: TFloatField;
    rxSeguimiento_MatdExcedenteOC: TFloatField;
    rxSeguimiento_MatdPorcentajeOC: TFloatField;
    rxSeguimiento_MatdPorcentajeOC_T: TFloatField;
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
    rxSeguimiento_MatNumeroReporte: TStringField;
    rxSeguimiento_MatFechaRD: TDateField;
    rxSeguimiento_MatFrenteRD: TStringField;
    rxSeguimiento_MatdCantidadRD: TFloatField;
    rxSeguimiento_MatdRestanteRD: TFloatField;
    rxSeguimiento_MatdExcedenteRD: TFloatField;
    rxSeguimiento_MatdPorcentajeRD: TFloatField;
    rxSeguimiento_MatdPorcentajeRD_T: TFloatField;
    rxSeguimiento_MatiNumeroEstimacion: TIntegerField;
    rxSeguimiento_MatsNumeroOrden: TStringField;
    rxSeguimiento_MatsNumeroGenerador: TStringField;
    rxSeguimiento_MatdCantidadGen: TFloatField;
    rxSeguimiento_MatdExcedenteGen: TFloatField;
    rxSeguimiento_MatdRestanteGen: TFloatField;
    rxSeguimiento_MatdPorcentajeGen: TFloatField;
    frxSeguimiento_Mat: TfrxDBDataset;
    frxSeguimiento_Mat1: TfrxDBDataset;
    rxSeguimiento_Mat1: TRxMemoryData;
    StringField9: TStringField;
    FloatField6: TFloatField;
    rxSeguimiento_Mat1Unidad: TStringField;
    IntegerField1: TIntegerField;
    IntegerField2: TIntegerField;
    FloatField9: TFloatField;
    FloatField10: TFloatField;
    FloatField11: TFloatField;
    FloatField12: TFloatField;
    rxSeguimiento_Mat1dCantidadReq_T: TFloatField;
    rxSeguimiento_Mat1dRestanteReq_T: TFloatField;
    rxSeguimiento_Mat1dExcedenteReq_T: TFloatField;
    FloatField13: TFloatField;
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
    rxSeguimiento_Mat2dCantidadOC_T: TFloatField;
    rxSeguimiento_Mat2dRestanteOC_T: TFloatField;
    rxSeguimiento_Mat2dExcedenteOC_T: TFloatField;
    FloatField49: TFloatField;
    frxSeguimiento_Mat3: TfrxDBDataset;
    rxSeguimiento_Mat3: TRxMemoryData;
    StringField7: TStringField;
    FloatField4: TFloatField;
    IntegerField5: TIntegerField;
    IntegerField6: TIntegerField;
    FloatField17: TFloatField;
    FloatField18: TFloatField;
    FloatField19: TFloatField;
    FloatField20: TFloatField;
    rxSeguimiento_Mat3dCantidadIn_T: TFloatField;
    rxSeguimiento_Mat3dExcedenteIn_T: TFloatField;
    frxSeguimiento_Mat4: TfrxDBDataset;
    rxSeguimiento_Mat4: TRxMemoryData;
    StringField8: TStringField;
    FloatField5: TFloatField;
    IntegerField9: TIntegerField;
    IntegerField10: TIntegerField;
    FloatField26: TFloatField;
    FloatField27: TFloatField;
    FloatField28: TFloatField;
    FloatField29: TFloatField;
    rxSeguimiento_Mat4dCantidadOut_T: TFloatField;
    rxSeguimiento_Mat4dExcedenteOut_T: TFloatField;
    frxSeguimiento_Mat5: TfrxDBDataset;
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
    rxSeguimiento_Mat5dCantidadRD_T: TFloatField;
    rxSeguimiento_Mat5dExcedenteRD_T: TFloatField;
    rxSeguimiento_Mat5dRestanteRD_T: TFloatField;
    frxSeguimiento_Mat6: TfrxDBDataset;
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
    rxSeguimiento_Mat6CantidadAnexo: TFloatField;
    SeguimientoMaterialGeneral1: TMenuItem;
    SeguimientoMaterialxPartida1: TMenuItem;
    SeguimientoMaterialxPartidaDetalle1: TMenuItem;
    chkRequisicion: TCheckBox;
    tsMoneda: TRxDBLookupCombo;
    ds_moneda: TDataSource;
    Moneda: TZReadOnlyQuery;
    tiFolio: TEdit;
    Label11: TLabel;
    tsEmbarca: TEdit;
    Label12: TLabel;
    tsCondiciones: TEdit;
    Label16: TLabel;
    tsEntrega: TEdit;
    Label19: TLabel;
    tsPrecios: TEdit;
    Label20: TLabel;
    tsVigencia: TEdit;
    Label8: TLabel;
    chkTipoCambio: TCheckBox;
    tsTipoCambio: TRxCalcEdit;
    Label13: TLabel;
    tsVendedor: TEdit;
    Label18: TLabel;
    tsMail: TEdit;
    reportesContrato: TStringField;
    reporteiFolioPedido: TIntegerField;
    reportesOrdenCompra: TStringField;
    reportesFolioRequisicion: TStringField;
    reportesIdProveedor: TStringField;
    reportesNumeroOrden: TStringField;
    reportedIdFecha: TDateField;
    reportedFechaEntrega: TDateField;
    reportesReferencia: TStringField;
    reportesElaboro: TStringField;
    reportesReviso1: TStringField;
    reportesAutorizo: TStringField;
    reportesFormaPago: TStringField;
    reportesMoneda: TStringField;
    reportedCambio: TFloatField;
    reportemComentarios: TMemoField;
    reportesStatus: TStringField;
    reportesLugarEntrega: TStringField;
    reportesCondiciones: TStringField;
    reportesEntrega: TStringField;
    reportesPrecios: TStringField;
    reportesVigencia: TStringField;
    reportesVendedor: TStringField;
    reportesMail: TStringField;
    reportesRazon: TStringField;
    reportesDomicilio: TStringField;
    reportesEstado: TStringField;
    reportesRFC: TStringField;
    reportesTelefono: TStringField;
    reporteiItem: TIntegerField;
    reportedCantidad: TFloatField;
    reportemDescripcion: TMemoField;
    reportesMedida: TStringField;
    reportedCosto: TFloatField;
    reporteiItemOrden: TStringField;
    reportemoneda: TStringField;
    reportedAcumulado: TFloatField;
    reportesElabora: TStringField;
    reportedIVA: TFloatField;
    btnProveedores: TBitBtn;
    btnReferencia: TBitBtn;
    Label14: TLabel;
    tdCantidad: TRxCalcEdit;
    Label10: TLabel;
    tDescuentoMat: TRxCalcEdit;
    Label1: TLabel;
    tDescuentoGral: TRxCalcEdit;
    anexo_ocompras: TZQuery;
    reportedDescuento: TFloatField;
    anexo_pocomprassStatus: TStringField;
    N5: TMenuItem;
    N6: TMenuItem;
    Imprimir1: TMenuItem;
    ImprimirComprasxProveedor1: TMenuItem;
    ImprimirComprasxReferencia1: TMenuItem;
    PanelImprime: TNxPanel;
    cmdTitle: TButton;
    cmdCancelar: TButton;
    cmdAcepta: TButton;
    zqDatos: TZReadOnlyQuery;
    anexo_ocomprasdMontoMN: TFloatField;
    anexo_ocomprassContrato: TStringField;
    anexo_ocomprassOrdenCompra: TStringField;
    anexo_ocomprassFolioRequisicion: TStringField;
    anexo_ocomprassIdProveedor: TStringField;
    anexo_ocomprassNumeroOrden: TStringField;
    anexo_ocomprasdIdFecha: TDateField;
    anexo_ocomprasdFechaEntrega: TDateField;
    anexo_ocomprassReferencia: TStringField;
    anexo_ocomprassElaboro: TStringField;
    anexo_ocomprassReviso1: TStringField;
    anexo_ocomprassReviso2: TStringField;
    anexo_ocomprassAutorizo: TStringField;
    anexo_ocomprassFormaPago: TStringField;
    anexo_ocomprassMoneda: TStringField;
    anexo_ocomprasdCambio: TFloatField;
    anexo_ocomprasdIVA: TFloatField;
    anexo_ocomprasdDescuento: TFloatField;
    anexo_ocomprasmComentarios: TMemoField;
    anexo_ocomprassStatus: TStringField;
    anexo_ocomprassLugarEntrega: TStringField;
    anexo_ocomprassCondiciones: TStringField;
    anexo_ocomprassEntrega: TStringField;
    anexo_ocomprassPrecios: TStringField;
    anexo_ocomprassVigencia: TStringField;
    anexo_ocomprassVendedor: TStringField;
    anexo_ocomprassMail: TStringField;
    anexo_ocomprasiFolioPedido: TIntegerField;
    Label2: TLabel;
    tdCostoNuevo: TRxCalcEdit;
    anexo_pocomprasdDescuento: TFloatField;
    ImprimirVariaciondePrecios1: TMenuItem;
    reporteDescuentoMat: TFloatField;
    ImprimirResumendeOC1: TMenuItem;
    FiltroUno: TGroupBox;
    lblDatos: TLabel;
    cbxDatos: TComboBox;
    FiltroDos: TGroupBox;
    Label21: TLabel;
    FechaInicio: TDateTimePicker;
    FechaFinal: TDateTimePicker;
    Label22: TLabel;
    Label23: TLabel;
    rxPrecios: TRxMemoryData;
    rxPreciossContrato: TStringField;
    rxPreciossIdInsumo: TStringField;
    rxPreciosmDescripcion: TMemoField;
    rxPreciossIdProveedor: TStringField;
    rxPreciosdFecha1: TDateField;
    rxPreciosdFecha2: TDateField;
    rxPreciosdFecha3: TDateField;
    rxPreciosdFecha4: TDateField;
    rxPreciosdFecha5: TDateField;
    rxPreciosdFecha6: TDateField;
    rxPreciosdFecha7: TDateField;
    rxPreciosdFecha8: TDateField;
    rxPreciosdCosto1: TFloatField;
    rxPreciosdCosto2: TFloatField;
    rxPreciosdCosto3: TFloatField;
    rxPreciosdCosto4: TFloatField;
    rxPreciosdCosto5: TFloatField;
    rxPreciosdCosto6: TFloatField;
    rxPreciosdCosto7: TFloatField;
    rxPreciosdCosto8: TFloatField;
    btnAdd: TAdvGlowButton;
    frxPrecios: TfrxDBDataset;
    rxPreciosItem: TIntegerField;
    rxPreciossOrdenCompra1: TStringField;
    rxPreciossOrdenCompra2: TStringField;
    rxPreciossOrdenCompra3: TStringField;
    rxPreciossOrdenCompra4: TStringField;
    rxPreciossOrdenCompra5: TStringField;
    rxPreciossOrdenCompra6: TStringField;
    rxPreciossOrdenCompra7: TStringField;
    rxPreciossOrdenCompra8: TStringField;
    CheckEditLink1: TCheckEditLink;
    rxPartidas: TRxMemoryData;
    rxPartidassNumeroActividad: TStringField;
    rxPartidassWbs: TStringField;
    rxPartidasiItemOrden: TStringField;
    rxPartidasiFolioRequisicion: TIntegerField;
    rxPartidassPartidaSolicitud: TStringField;
    anexo_pocomprassNumeroSolicitud: TStringField;
    rxPartidassNumeroSolicitud: TStringField;
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
    procedure Salir1Click(Sender: TObject);
    procedure anexo_ocomprasAfterScroll(DataSet: TDataSet);
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

    procedure GridPartidasEnter(Sender: TObject);
    procedure frmBarra2btnPrinterClick(Sender: TObject);
    procedure tsNumeroActividadExit(Sender: TObject);
    procedure tsNumeroActividadEnter(Sender: TObject);
    procedure tsNumeroActividadKeyPress(Sender: TObject; var Key: Char);
    procedure frxEntradaGetValue(const VarName: String;
      var Value: Variant);
    procedure imgNotasDblClick(Sender: TObject);
    procedure ComentariosAdicionalesClick(Sender: TObject);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure anexo_ocomprasCalcFields(DataSet: TDataSet);
    procedure anexo_pocomprasCalcFields(DataSet: TDataSet);
    procedure ReporteCalcFields(DataSet: TDataSet);
    procedure tsIdProveedorEnter(Sender: TObject);
    procedure tsIdProveedorExit(Sender: TObject);
    procedure tsIdProveedorKeyPress(Sender: TObject; var Key: Char);
    procedure FormActivate(Sender: TObject);
    procedure Grid_EntradasCellClick(Column: TColumn);
    procedure tsFormaPagoKeyPress(Sender: TObject; var Key: Char);
    procedure tsFormaPagoEnter(Sender: TObject);
    procedure tsentregaKeyPress(Sender: TObject; var Key: Char);
    procedure tsentregaEnter(Sender: TObject);
    procedure tsentregaExit(Sender: TObject);
    procedure tsMonedaEnter(Sender: TObject);
    procedure tsMonedaExit(Sender: TObject);
    procedure validaciones();
    procedure dbRequisicionKeyPress(Sender: TObject; var Key: Char);
    procedure dbPartidasClick(Sender: TObject);
    procedure txtMaterialKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure InsumosAfterScroll(DataSet: TDataSet);
    procedure Grid_InsumosCellClick(Column: TColumn);
    procedure txtMaterialKeyPress(Sender: TObject; var Key: Char);
    procedure Grid_InsumosKeyPress(Sender: TObject; var Key: Char);
    procedure validaPedido();
    procedure dbRequisicionEnter(Sender: TObject);
    procedure SeguimientoMaterialxPartida1Click(Sender: TObject);
    procedure Seguimiento_Material(dParamActividad : string);
    procedure SeguimientoMaterialGeneral1Click(Sender: TObject);
    procedure SeguimientoMaterialxPartidaDetalle1Click(Sender: TObject);
    procedure dbRequisicionExit(Sender: TObject);
    procedure dbPartidasKeyPress(Sender: TObject; var Key: Char);
    procedure Grid_EntradasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_EntradasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_EntradasTitleClick(Column: TColumn);
    procedure Grid_InsumosTitleClick(Column: TColumn);
    procedure GridPartidasTitleClick(Column: TColumn);
    procedure Grid_InsumosMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GridPartidasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GridPartidasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_InsumosMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure tdCantidadChange(Sender: TObject);
    procedure chkRequisicionClick(Sender: TObject);
    procedure tiFolioEnter(Sender: TObject);
    procedure tiFolioExit(Sender: TObject);
    procedure tiFolioKeyPress(Sender: TObject; var Key: Char);
    procedure tsEmbarcaKeyPress(Sender: TObject; var Key: Char);
    procedure tsCondicionesKeyPress(Sender: TObject; var Key: Char);
    procedure tsPreciosKeyPress(Sender: TObject; var Key: Char);
    procedure tsVigenciaKeyPress(Sender: TObject; var Key: Char);
    procedure tsEmbarcaEnter(Sender: TObject);
    procedure tsEmbarcaExit(Sender: TObject);
    procedure tsCondicionesEnter(Sender: TObject);
    procedure tsCondicionesExit(Sender: TObject);
    procedure tsPreciosEnter(Sender: TObject);
    procedure tsPreciosExit(Sender: TObject);
    procedure tsVigenciaEnter(Sender: TObject);
    procedure tsVigenciaExit(Sender: TObject);
    procedure txtMaterialEnter(Sender: TObject);
    procedure txtMaterialExit(Sender: TObject);
    procedure tsVendedorKeyPress(Sender: TObject; var Key: Char);
    procedure tsVendedorEnter(Sender: TObject);
    procedure tsVendedorExit(Sender: TObject);
    procedure tsMailKeyPress(Sender: TObject; var Key: Char);
    procedure tsMailEnter(Sender: TObject);
    procedure tsMailExit(Sender: TObject);
    procedure tipocambio();
    procedure PgControlChange(Sender: TObject);
    procedure btnProveedoresClick(Sender: TObject);
    procedure btnReferenciaClick(Sender: TObject);
    procedure GridPartidasGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure lblPartidaClick(Sender: TObject);
    procedure ImprimirComprasxProveedor1Click(Sender: TObject);
    procedure cmdCancelarClick(Sender: TObject);
    procedure ImprimirComprasxReferencia1Click(Sender: TObject);
    procedure cmdAceptaClick(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure ImprimirVariaciondePrecios1Click(Sender: TObject);
    procedure ImprimirResumendeOC1Click(Sender: TObject);

  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmPedidos: TfrmPedidos;
  SavePlace     : TBookmark;
  sDescripcion    : String ;
  txtAux          : String ;
  lNuevo          : Boolean ;
  OpcButton1      : String ;
  Valida          : boolean;
  filtro, buscar  : string;
  iFolio          : integer;
  CadenaOrden, CadenaOrden2, cadenaReq  : string;
  lValidaReq      : boolean;
  CadenaReporte, CadenaReporte2 : string;

  sSuperintendente, sSupervisor : String ;
  sPuestoSuperintendente, sPuestoSupervisor : String ;
  sSupervisorTierra, sPuestoSupervisorTierra : String ;
  sBackup       : String ;
  MontoTotal    : Currency  ;
  TipoExplosion : string;       
  utgrid:ticdbgrid;
  utgrid2:ticdbgrid;
  utgrid3:ticdbgrid;
  botonpermiso:tbotonespermisos;
implementation

uses frm_connection , frm_comentariosxanexo, frm_ordenes, frm_proveedores;

{$R *.dfm}

procedure filtra;
begin
  filtro := '';
  if length(trim(frmPedidos.txtMaterial.Text)) > 0  then
  begin
      buscar := frmPedidos.txtMaterial.Text;
      buscar := buscar + '*';

      if frmPedidos.lblPartida.Caption = 'ID Msterial' then
         filtro := ' sIdInsumo like ' + QuotedStr(buscar)
      else
         filtro := ' mDescripcion like ' + QuotedStr(buscar);
  end;

  if filtro <> '' then
  begin
      frmPedidos.insumos.Filtered := False;
      frmPedidos.insumos.Filter := filtro;
      frmPedidos.insumos.Filtered := True;
  end
  else
     frmPedidos.insumos.Filtered := False;
end;

procedure TfrmPedidos.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  action := cafree ;
  utgrid.Destroy;
  utgrid2.Destroy;
  utgrid3.Destroy;
  botonpermiso.Free;
end;


procedure TfrmPedidos.FormShow(Sender: TObject);
var x : integer;
begin
  try
    sMenuP:=stMenu;
    BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'opPedidos', PopupPrincipal);
    UtGrid:=TicdbGrid.create(grid_entradas);
    UtGrid2:=TicdbGrid.create(grid_insumos);
    UtGrid3:=TicdbGrid.create(gridpartidas);
    tdIdFecha.Enabled      := True ;
    tsNumeroOrden.Enabled  := False ;
    tsIdProveedor.ReadOnly := False ;
    tmComentarios.ReadOnly := True ;
    txtMaterial.ReadOnly   := True;
    pgControl.ActivePageIndex := 0 ;

    CadenaOrden := 'Select a1.*, a.sMedida, a.dCantidad, a3.iItemOrden from anexo_ppedido a1 '+
                   'left join actividadesxanexo a3 on (a3.sContrato = a1.sContrato and a3.sIdConvenio =:Convenio '+
                   'and a3.sNumeroActividad = a1.sNumeroActividad and a3.sTipoActividad ="Actividad") '+
                   'inner join anexo_pedidos a2  on (a1.sContrato = a2.sContrato And a1.iFolioPedido = a2.iFolioPedido) '+
                   'inner join insumos a on (a1.sContrato = a.sContrato And a1.sIdInsumo = a.sIdInsumo ) '+
                   'Where a1.sContrato = :Contrato And a1.iFolioPedido = :Folio order by a3.iItemOrden ';

    CadenaOrden2 := 'Select a1.*, "" as sMedida, 0.0 as dCantidad  from anexo_ppedido a1 '+
                   'inner join anexo_pedidos a2  on (a1.sContrato = a2.sContrato And a1.iFolioPedido = a2.iFolioPedido) '+
                   'inner join insumos a on (a1.sContrato = a.sContrato And a1.sIdInsumo = a.sIdInsumo ) '+
                   'Where a1.sContrato =:Contrato And a1.iFolioPedido = :Folio order by a.sIdInsumo ';

    CadenaReporte  := 'Select a2.*, p.*, a1.iItem, a1.dCantidad, a1.mDescripcion, a1.sMedida, a1.dCosto, a3.iItemOrden, u.sNombre as sElabora, m.sDescripcion as moneda, (a1.dCosto - a1.dDescuento) as DescuentoMat '+
                      'from anexo_ppedido a1 '+
                      'inner join anexo_pedidos a2 on (a1.sContrato = a2.sContrato And a1.iFolioPedido = a2.iFolioPedido) '+
                      'left join actividadesxanexo a3 on (a3.sContrato = a1.sContrato and a3.sIdConvenio =:Convenio '+
                      'and a3.sNumeroActividad = a1.sNumeroActividad and a3.sTipoActividad ="Actividad") '+
                      'inner join proveedores p on (a2.sIdProveedor = p.sIdProveedor) '+
                      'left join usuarios u on (u.sIdUsuario = a2.sElaboro) '+
                      'inner join tiposdemoneda m on (a2.sMoneda = m.sIdMoneda) '+
                      'Where a1.sContrato = :Contrato And a1.iFolioPedido = :Folio order by a3.iItemOrden ';

    CadenaReporte2 := 'Select a2.*, p.*, a1.iItem, a1.dCantidad, a1.mDescripcion, a1.sMedida, a1.dCosto, "" as iItemOrden, u.sNombre as sElabora, m.sDescripcion as moneda, (a1.dCosto - a1.dDescuento) as DescuentoMat '+
                      'from anexo_ppedido a1 '+
                      'inner join anexo_pedidos a2 on (a1.sContrato = a2.sContrato And a1.iFolioPedido = a2.iFolioPedido) '+
                      'inner join proveedores p on (a2.sIdProveedor = p.sIdProveedor) '+
                      'left join usuarios u on (u.sIdUsuario = a2.sElaboro) '+
                      'inner join tiposdemoneda m on (a2.sMoneda = m.sIdMoneda) '+
                      'Where a1.sContrato = :Contrato And a1.iFolioPedido = :Folio order by a1.sIdInsumo ';

    tsNumeroOrden.Items.Clear ;
    //tsNumeroOrden.Items.Add('CONTRATO No. ' + global_contrato)  ;
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select sNumeroOrden from ordenesdetrabajo where sContrato = :Contrato and ' +
                                'cIdStatus = :status order by sNumeroOrden') ;
    connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    connection.qryBusca.Params.ParamByName('Contrato').Value    := Global_Contrato ;
    connection.qryBusca.Params.ParamByName('status').DataType   := ftString ;
    connection.qryBusca.Params.ParamByName('status').Value      := connection.configuracion.FieldValues [ 'cStatusProceso' ];
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 Then
        While NOT Connection.qryBusca.Eof Do
        Begin
            tsNumeroOrden.Items.Add(Connection.qryBusca.FieldValues['sNumeroOrden']) ;
            Connection.qryBusca.Next
        End ;
    tsNumeroOrden.ItemIndex := 0 ;

    Proveedores.Active := False ;
    Proveedores.Open ;

    FormaPago.Active := False ;
    FormaPago.Open ;

    Moneda.Active := False ;
    Moneda.Open ;

    anexo_ocompras.Active := False ;
    anexo_ocompras.Params.ParamByName('Contrato').DataType := ftString ;
    anexo_ocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
    anexo_ocompras.Open ;

    anexo_pocompras.Active := False ;
    anexo_pocompras.SQL.Clear;
    anexo_pocompras.SQL.Add(CadenaOrden);
    anexo_pocompras.Params.ParamByName('Contrato').DataType := ftString ;
    anexo_pocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
    anexo_pocompras.Params.ParamByName('Convenio').DataType := ftString ;
    anexo_pocompras.Params.ParamByName('Convenio').Value    := global_convenio ;
    anexo_pocompras.Params.ParamByName('Folio').DataType    := ftInteger ;
    If anexo_ocompras.RecordCount > 0 Then
        anexo_pocompras.Params.ParamByName('Folio').Value := anexo_ocompras.FieldValues['iFolioPedido']
    Else
        anexo_pocompras.Params.ParamByName('Folio').Value := 0 ;
    anexo_pocompras.Open ;
    grid_entradas.enabled:=true;
    Grid_Entradas.SetFocus   ;
    requisiciones.Params.ParamByName('Contrato').DataType := ftString ;
    requisiciones.Params.ParamByName('Contrato').Value    := global_contrato ;
    requisiciones.Open ;

    if requisiciones.RecordCount > 0 then
    begin
         dbRequisicion.Clear;
         while not requisiciones.Eof do
         begin
              dbRequisicion.Items.Add(IntToStr(requisiciones.FieldValues['iFolioRequisicion'])+ '.- '+requisiciones.FieldValues['sNumeroSolicitud']  );
              requisiciones.Next;
         end;
    end;

    if connection.configuracion.FieldValues['sExplosion'] = 'Recursos por Concepto/Partida' then
       TipoExplosion := 'recursosanexo'
    else
       TipoExplosion := 'recursosanexosnuevos';

    frmBarra2.btnAdd.Enabled := true;
    frmBarra2.btnAdd.Enabled := true;
    frmBarra2.btnEdit.Enabled := true;
    frmBarra2.btnDelete.Enabled := true;
    frmBarra2.btnPrinter.Enabled := true;
    BotonPermiso.permisosBotones(frmBarra1);
    BotonPermiso.permisosBotones(frmBarra2);
    dbRequisicion.Enabled := False;
  except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al iniciar formulario', 0);
  end;
  end;
end;

procedure TfrmPedidos.btnReferenciaClick(Sender: TObject);
begin
    global_frmActivo := 'frm_pedidos';
    Application.CreateForm(TfrmOrdenes, frmOrdenes);
    frmOrdenes.show;
end;

procedure TfrmPedidos.BtnExitClick(Sender: TObject);
begin
    Close ;
end;

procedure TfrmPedidos.frmBarra1btnAddClick(Sender: TObject);
var
   sRequisiciones, sReq : string;
   inicio, indice : integer;
begin
    validaciones;

    if valida then
    begin
          frmBarra1.btnCancel.Click ;
          exit;
    end;
    tdCantidad.ReadOnly := False;

    if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
    begin
        lblPartida.Caption  := 'No. Partida';
        txtMaterial.Visible := False;
        dbPartidas.Visible  := True;
     end
     else
     begin
        lblPartida.Caption  := 'Material';
        txtMaterial.SetFocus;
        txtMaterial.Visible := True;
        dbPartidas.Visible  := False;
     end;

    txtMaterial.ReadOnly := False;
    if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
    begin
        rxPartidas.EmptyTable;
        rxPartidas.Active := True;
        sRequisiciones := anexo_ocompras.FieldValues['sFolioRequisicion'];
        sReq           := sRequisiciones;
        inicio         := 1;
        while sReq <> '' do
        begin
            indice := pos('.-', sRequisiciones);
            if indice > 0 then
            begin
                sReq   := copy(sRequisiciones, inicio, (indice - inicio));
                sRequisiciones := copy(sRequisiciones, indice + 3, length(sRequisiciones));
                inicio := pos(',', sRequisiciones) + 1;
                connection.QryBusca.Active := False;
                connection.QryBusca.SQL.Clear;
                connection.QryBusca.SQL.Add('select distinct r.sNumeroSolicitud, p.sNumeroActividad, a.iItemOrden, a.sWbs from anexo_prequisicion  p '+
                                   'inner join anexo_requisicion r on (r.sContrato = p.sContrato and r.iFolioRequisicion = p.iFolioRequisicion ) '+
                                   'left join actividadesxanexo a '+
                                   'on (a.sContrato = p.sContrato and a.sIdConvenio =:Convenio and a.sNumeroActividad = p.sNumeroActividad and a.sTipoActividad = "Actividad") '+
                                   'where p.sContrato =:Contrato and p.iFolioRequisicion =:Folio order by a.iItemOrden ');
                connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
                connection.QryBusca.ParamByName('Convenio').AsString := global_convenio;
                connection.QryBusca.ParamByName('Folio').AsInteger   := StrToInt(sReq);
                connection.QryBusca.Open;

                while not connection.QryBusca.Eof do
                begin
                    rxPartidas.Append;
                    rxPartidas.FieldValues['sPartidaSolicitud'] := connection.QryBusca.FieldValues['sNumeroActividad']+' - Solic. '+connection.QryBusca.FieldValues['sNumeroSolicitud'];
                    rxPartidas.FieldValues['sNumeroActividad']  := connection.QryBusca.FieldValues['sNumeroActividad'];
                    rxPartidas.FieldValues['iItemOrden']        := connection.QryBusca.FieldValues['iItemOrden'];
                    rxPartidas.FieldValues['sWbs']              := connection.QryBusca.FieldValues['sWbs'];
                    rxPartidas.FieldValues['iFolioRequisicion'] := StrToInt(sReq);
                    rxPartidas.FieldValues['sNumeroSolicitud']  := connection.QryBusca.FieldValues['sNumeroSolicitud'];
                    rxPartidas.Post;
                    connection.QryBusca.Next;
                end;
            end
            else
                sReq := '';
        end;
    end
    else
    begin
        insumos.Active := False;
        insumos.SQL.Clear;
        insumos.SQL.Add('Select *,SubStr(mDescripcion, 1, 255) as sDescripcion, "" as Requisicion from insumos where sContrato =:Contrato ');
        insumos.ParamByName('Contrato').AsString := global_contrato;
        insumos.Open;

        if insumos.RecordCount = 0 then
           messageDLG('No existen Materiales para el Almacen principal', mtInformation, [mbOk],0);
    end ;

    If (anexo_ocompras.RecordCount > 0) Then
    Begin
        Insertar1.Enabled := False ;
        Editar1.Enabled := False ;
        Registrar1.Enabled := True ;
        Can1.Enabled := True ;
        Eliminar1.Enabled := False ;
        Refresh1.Enabled := False ;
        frmBarra1.btnAddClick(Sender);
        tdCantidad.Value := 0 ;
    End;
    tDescuentoMat.Value := 0;
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmPedidos.frmBarra1btnEditClick(Sender: TObject);
begin
    if anexo_pocompras.FieldValues['sStatus'] = 'Entregado' then
    begin
        messageDLG('No se puede modificar el Material ya fue Entregado!', mtWarning, [mbOk], 0);
        exit;
    end;
    
    validaciones;
    if valida then
    begin
        frmBarra1.btnCancel.Click ;
        exit;
    end;
    txtMaterial.ReadOnly := False;
    tdCantidad.ReadOnly := False;
    If anexo_ocompras.RecordCount > 0 Then
    Begin
        Insertar1.Enabled := False ;
        Editar1.Enabled := False ;
        Registrar1.Enabled := True ;
        Can1.Enabled := True ;
        Eliminar1.Enabled := False ;
        Refresh1.Enabled := False;
        frmBarra1.btnEditClick(Sender);
        tdCantidad.ReadOnly := False ;
    End;
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;


procedure TfrmPedidos.frmBarra1btnPostClick(Sender: TObject);
Var
  SavePlace  : TBookmark ;
  dCantidad  : Currency ;
  conteo     : integer;
begin
    if tDescuentoMat.Value > (tdCantidad.Value * tdCostoNuevo.Value) then
    begin
        messageDLG('El Descuento Applicado es Mayor al Costo Total, Favor de verificar! ', mtInformation, [mbOk], 0);
        frmBarra1.btnCancel.Click;
        exit;
    end;
        If OpcButton = 'New' then
        Begin
             //Consultamos primero si ya existe el material dado de alta en la orden de compra..
             Connection.QryBusca.Active := False ;
             Connection.QryBusca.SQL.Clear ;
             Connection.QryBusca.SQL.Add('select sIdInsumo from anexo_ppedido where sContrato =:Contrato and iFolioPedido =:Folio and sIdInsumo =:Insumo ');
             Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
             Connection.QryBusca.Params.ParamByName('Contrato').value    := Global_Contrato ;
             Connection.QryBusca.Params.ParamByName('Folio').DataType    := ftInteger ;
             Connection.QryBusca.Params.ParamByName('Folio').value       := anexo_ocompras.FieldValues['iFolioPedido'] ;
             Connection.QryBusca.Params.ParamByName('Insumo').DataType   := ftString ;
             Connection.QryBusca.Params.ParamByName('Insumo').value      := insumos.FieldValues['sIdInsumo'] ;
             Connection.QryBusca.Open;

             if connection.QryBusca.RecordCount = 0 then
             begin
                  Connection.zCommand.Active := False ;
                  Connection.zcommand.SQL.Clear ;
                  Connection.zCommand.SQL.Add('INSERT INTO anexo_ppedido ( sContrato, iFolioPedido, sIdInsumo, iItem, mDescripcion, sMedida, dCantidad, dCosto, sNumeroActividad, sNumeroOrden, sStatus, dDescuento, sNumeroSolicitud ) '+
                                              'VALUES (:Contrato, :Folio, :Insumo, 1, :Descripcion, :Medida, :Cantidad, :Costo, :Actividad, :Orden, "Pendiente", :Descuento, :solicitud) ');
                  Connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
                  Connection.zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                  Connection.zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                  Connection.zcommand.Params.ParamByName('Folio').value          := anexo_ocompras.FieldValues['iFolioPedido'] ;
                  Connection.zcommand.Params.ParamByName('Insumo').DataType      := ftString ;
                  Connection.zcommand.Params.ParamByName('Insumo').value         := insumos.FieldValues['sIdInsumo'] ;
                  Connection.zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
                  Connection.zcommand.Params.ParamByName('Descripcion').value    := insumos.FieldValues['mDescripcion'] ;
                  Connection.zcommand.Params.ParamByName('Medida').DataType      := ftString ;
                  Connection.zcommand.Params.ParamByName('Medida').value         := insumos.FieldValues['sMedida'];
                  Connection.zcommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
                  Connection.zcommand.Params.ParamByName('Cantidad').value       := tdCantidad.Value ;
                  Connection.zcommand.Params.ParamByName('Costo').DataType       := ftFloat ;
                  if (tsMoneda.KeyValue = 'DLL') or (tsMoneda.KeyValue = 'dll') then
                  begin
                      if tdCostoNuevo.Value <> insumos.FieldValues['dVentaDLL'] then
                      begin
                          if anexo_ocompras.FieldValues['dCambio'] > 0 then
                             Connection.zcommand.Params.ParamByName('Costo').value   := (insumos.FieldValues['dVentaDLL'] * anexo_ocompras.FieldValues['dCambio'])
                          else
                             Connection.zcommand.Params.ParamByName('Costo').value   := insumos.FieldValues['dVentaDLL'];
                      end
                      else
                          Connection.zcommand.Params.ParamByName('Costo').value   := tdCostoNuevo.Value;
                  end
                  else
                     if tdCostoNuevo.Value <> insumos.FieldValues['dVentaMN'] then
                        Connection.zcommand.Params.ParamByName('Costo').value    := tdCostoNuevo.Value
                     else
                        Connection.zcommand.Params.ParamByName('Costo').value    := insumos.FieldValues['dVentaMN'];
                  Connection.zcommand.Params.ParamByName('Actividad').DataType   := ftString ;
                  if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
                     Connection.zcommand.Params.ParamByName('Actividad').value   := rxPartidas.FieldValues['sNumeroActividad']
                  else
                     Connection.zcommand.Params.ParamByName('Actividad').value   := 'S/N';
                  Connection.zcommand.Params.ParamByName('Orden').DataType       := ftString ;
                  Connection.zcommand.Params.ParamByName('Orden').value          := anexo_ocompras.FieldValues['sNumeroOrden'] ;
                  Connection.zcommand.Params.ParamByName('Descuento').DataType   := ftFloat ;
                  Connection.zcommand.Params.ParamByName('Descuento').value      := tDescuentoMat.Value ;
                  Connection.zcommand.Params.ParamByName('Solicitud').DataType   := ftString ;
                  Connection.zcommand.Params.ParamByName('Solicitud').value      := rxPartidas.FieldValues['sNumeroSolicitud'] ;
                  Connection.zcommand.ExecSQL  ;
            end
            else
            begin
                  If MessageDlg('Se encontro una coincidencia con el material en la orden de compra actual, desea modificar el registro encontrado?. Si asi no lo desea, se creara un registro nuevo.', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
                  Begin
                      Connection.qryBusca.Active := False ;
                      Connection.qryBusca.SQL.Clear ;
                      Connection.qryBusca.SQL.Add('Select Max(iItem) as iItem From anexo_ppedido Where sContrato = :Contrato And iFolioPedido = :Folio and sIdInsumo = :Insumo ') ;
                      connection.qryBusca.Params.ParamByName('Contrato').DataType  := ftString ;
                      connection.qryBusca.Params.ParamByName('Contrato').value     := Global_Contrato ;
                      connection.qryBusca.Params.ParamByName('Folio').DataType     := ftInteger ;
                      connection.qryBusca.Params.ParamByName('Folio').value        := anexo_ocompras.FieldValues['iFolioPedido'] ;
                      Connection.qryBusca.Params.ParamByName('Insumo').DataType    := ftString ;
                      Connection.qryBusca.Params.ParamByName('Insumo').value       := insumos.FieldValues['sIdInsumo'] ;
                      Connection.qryBusca.Open ;
                      If NOT Connection.qryBusca.FieldByName('iItem').IsNull Then
                      Begin
                          Connection.zCommand.Active := False ;
                          Connection.zCommand.SQL.Clear ;
                          Connection.zCommand.SQL.Add('INSERT INTO anexo_ppedido (sContrato, iFolioPedido, sIdInsumo, iItem, mDescripcion, sMedida, dCantidad, dCosto, sNumeroActividad, sNumeroOrden, sStatus, sNumeroSolicitud ) '  +
                                            'VALUES (:Contrato, :Folio, :Insumo, :Item, :Descripcion, :Medida, :Cantidad, :Costo, :Actividad, :Orden, "Pendiente", :solicitud) ');
                          Connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
                          Connection.zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                          Connection.zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                          Connection.zcommand.Params.ParamByName('Folio').value          := anexo_ocompras.FieldValues['iFolioPedido'] ;
                          Connection.zcommand.Params.ParamByName('Insumo').DataType      := ftString ;
                          Connection.zcommand.Params.ParamByName('Insumo').value         := insumos.FieldValues['sIdInsumo'] ;
                          Connection.zcommand.Params.ParamByName('Actividad').DataType   := ftString ;
                          if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
                              Connection.zcommand.Params.ParamByName('Actividad').value  := rxPartidas.FieldValues['sNumeroActividad']
                          else
                              Connection.zcommand.Params.ParamByName('Actividad').value  := 'S/N';
                          Connection.zcommand.Params.ParamByName('Item').DataType        := ftInteger ;
                          Connection.zcommand.Params.ParamByName('Item').value           := Connection.qryBusca.FieldValues['iItem'] + 1  ;
                          Connection.zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
                          Connection.zcommand.Params.ParamByName('Descripcion').value    := insumos.FieldValues['mDescripcion'] ;
                          Connection.zcommand.Params.ParamByName('Medida').DataType      := ftString ;
                          Connection.zcommand.Params.ParamByName('Medida').value         := insumos.FieldValues['sMedida'] ;
                          Connection.zcommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
                          Connection.zcommand.Params.ParamByName('Cantidad').value       := tdCantidad.Value ;
                          Connection.zcommand.Params.ParamByName('Costo').DataType       := ftFloat ;
                          if (tsMoneda.KeyValue = 'DLL') or (tsMoneda.KeyValue = 'dll') then
                          begin
                              if tdCostoNuevo.Value <> insumos.FieldValues['dVentaDLL'] then
                              begin
                                  if anexo_ocompras.FieldValues['dCambio'] > 0 then
                                     Connection.zcommand.Params.ParamByName('Costo').value   := (insumos.FieldValues['dVentaDLL'] * anexo_ocompras.FieldValues['dCambio'])
                                  else
                                     Connection.zcommand.Params.ParamByName('Costo').value   := insumos.FieldValues['dVentaDLL'];
                              end
                              else
                                  Connection.zcommand.Params.ParamByName('Costo').value   := tdCostoNuevo.Value;
                          end
                          else
                             if tdCostoNuevo.Value <> insumos.FieldValues['dVentaMN'] then
                                Connection.zcommand.Params.ParamByName('Costo').value    := tdCostoNuevo.Value
                             else
                                Connection.zcommand.Params.ParamByName('Costo').value    := insumos.FieldValues['dVentaMN'];
                          Connection.zcommand.Params.ParamByName('Orden').DataType       := ftString ;
                          Connection.zcommand.Params.ParamByName('Orden').value          := anexo_ocompras.FieldValues['sNumeroOrden'];
                          Connection.zcommand.Params.ParamByName('Solicitud').DataType   := ftString;
                          Connection.zcommand.Params.ParamByName('Solicitud').value      := rxPartidas.FieldValues['sNumeroSolicitud'];
                          Connection.zcommand.ExecSQL ;
                          MessageDlg('Se inserto correctemente el Material '+ insumos.FieldValues['sIdInsumo']+' en el CPT #' + IntToStr(Connection.qryBusca.FieldValues['iItem'] + 1) , mtInformation, [mbOk], 0);
                      End
                      Else
                      Begin
                          Connection.zCommand.Active := False ;
                          Connection.zCommand.SQL.Clear ;
                          Connection.zCommand.SQL.Add('INSERT INTO anexo_ppedido (sContrato, iFolioPedido, sIdInsumo, iItem, mDescripcion, sMedida, dCantidad, dCosto, sNumeroActividad, sNumeroOrden, sStatus, sNumeroSolicitud ) '  +
                                            'VALUES (:Contrato, :Folio, :Insumo, :Item, :Descripcion, :Medida, :Cantidad, :Costo, :Actividad, :Orden, "Pendiente", :solicitud) ');
                          Connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
                          Connection.zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                          Connection.zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                          Connection.zcommand.Params.ParamByName('Folio').value          := anexo_ocompras.FieldValues['iFolioPedido'] ;
                          Connection.zcommand.Params.ParamByName('Insumo').DataType      := ftString ;
                          Connection.zcommand.Params.ParamByName('Insumo').value         := insumos.FieldValues['sIdInsumo'] ;
                          Connection.zcommand.Params.ParamByName('Actividad').DataType   := ftString ;
                          if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
                              Connection.zcommand.Params.ParamByName('Actividad').value  := rxPartidas.FieldValues['sNumeroActividad']
                          else
                              Connection.zcommand.Params.ParamByName('Actividad').value  := 'S/N';
                          Connection.zcommand.Params.ParamByName('Item').DataType        := ftInteger ;
                          Connection.zcommand.Params.ParamByName('Item').value           := Connection.qryBusca.FieldValues['iItem'] + 1  ;
                          Connection.zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
                          Connection.zcommand.Params.ParamByName('Descripcion').value    := insumos.FieldValues['mDescripcion'] ;
                          Connection.zcommand.Params.ParamByName('Medida').DataType      := ftString ;
                          Connection.zcommand.Params.ParamByName('Medida').value         := insumos.FieldValues['sMedida'] ;
                          Connection.zcommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
                          Connection.zcommand.Params.ParamByName('Cantidad').value       := tdCantidad.Value ;
                          Connection.zcommand.Params.ParamByName('Costo').DataType       := ftFloat ;
                          if (tsMoneda.KeyValue = 'DLL') or (tsMoneda.KeyValue = 'dll') then
                          begin
                              if tdCostoNuevo.Value <> insumos.FieldValues['dVentaDLL'] then
                              begin
                                  if anexo_ocompras.FieldValues['dCambio'] > 0 then
                                     Connection.zcommand.Params.ParamByName('Costo').value   := (insumos.FieldValues['dVentaDLL'] * anexo_ocompras.FieldValues['dCambio'])
                                  else
                                     Connection.zcommand.Params.ParamByName('Costo').value   := insumos.FieldValues['dVentaDLL'];
                              end
                              else
                                  Connection.zcommand.Params.ParamByName('Costo').value   := tdCostoNuevo.Value;
                          end
                          else
                             if tdCostoNuevo.Value <> insumos.FieldValues['dVentaMN'] then
                                Connection.zcommand.Params.ParamByName('Costo').value    := tdCostoNuevo.Value
                             else
                                Connection.zcommand.Params.ParamByName('Costo').value    := insumos.FieldValues['dVentaMN'];
                          Connection.zcommand.Params.ParamByName('Orden').DataType       := ftString ;
                          Connection.zcommand.Params.ParamByName('Orden').value          := anexo_ocompras.FieldValues['sNumeroOrden'] ;
                          Connection.zcommand.Params.ParamByName('Solicitud').DataType   := ftString ;
                          Connection.zcommand.Params.ParamByName('Solicitud').value      := rxPartidas.FieldValues['sNumeroSolicitud'] ;
                          Connection.zcommand.ExecSQL ;
                          MessageDlg('Se inserto correctemente el Material '+ insumos.FieldValues['sIdInsumo']+' en el CPT #' + IntToStr(Connection.qryBusca.FieldValues['iItem'] + 1) , mtInformation, [mbOk], 0);
                      End
                  End
                  Else
                  Begin
                      dCantidad := 0 ;
                      Connection.qryBusca.Active := False ;
                      Connection.qryBusca.SQL.Clear ;
                      Connection.qryBusca.SQL.Add('Select dCantidad From anexo_ppedido Where sContrato = :Contrato And ' +
                                                  'iFolioPedido = :Folio and sIdInsumo = :Insumo And iItem = 1 ') ;
                      connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
                      connection.qryBusca.Params.ParamByName('Contrato').value    := Global_Contrato ;
                      connection.qryBusca.Params.ParamByName('Folio').DataType    := ftInteger ;
                      connection.qryBusca.Params.ParamByName('Folio').value       := anexo_ocompras.FieldValues['iFolioPedido'] ;
                      connection.qryBusca.Params.ParamByName('Insumo').DataType   := ftString ;
                      connection.qryBusca.Params.ParamByName('Insumo').value      := insumos.FieldValues['sIdInsumo'] ;
                      Connection.qryBusca.Open ;

                      If Connection.qryBusca.RecordCount > 0 Then
                         dCantidad := Connection.qryBusca.FieldValues['dCantidad'] ;

                      Connection.zCommand.Active := False ;
                      Connection.zCommand.SQL.Clear ;
                      Connection.zCommand.SQL.Add('UPDATE anexo_ppedido SET mDescripcion =:Descripcion, dCantidad =:Cantidad, sMedida =:Medida, sNumeroActividad =:Actividad, dCosto =:Costo ' +
                                                  'WHERE sContrato =:Contrato And iFolioPedido =:Folio and sIdInsumo =:Insumo And iItem = 1 ');
                      Connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
                      Connection.zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                      Connection.zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                      Connection.zcommand.Params.ParamByName('Folio').value          := anexo_ocompras.FieldValues['iFolioPedido'] ;
                      Connection.zcommand.Params.ParamByName('Insumo').DataType      := ftString ;
                      Connection.zcommand.Params.ParamByName('Insumo').value         := insumos.FieldValues['sIdInsumo'] ;
                      Connection.zcommand.Params.ParamByName('Actividad').DataType   := ftString ;
                      if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
                         Connection.zcommand.Params.ParamByName('Actividad').value   := rxPartidas.FieldValues['sNumeroActividad']
                      else
                         Connection.zcommand.Params.ParamByName('Actividad').value   := 'S/N';
                      Connection.zcommand.Params.ParamByName('Descripcion').DataType := ftMemo ;
                      Connection.zcommand.Params.ParamByName('Descripcion').value    := insumos.FieldValues['mDescripcion'] ;
                      Connection.zcommand.Params.ParamByName('Medida').DataType      := ftString ;
                      Connection.zcommand.Params.ParamByName('Medida').value         := insumos.FieldValues['sMedida'] ;
                      Connection.zcommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
                      Connection.zcommand.Params.ParamByName('Cantidad').value       := dCantidad + tdCantidad.Value ;
                      Connection.zcommand.Params.ParamByName('Costo').DataType       := ftFloat ;
                      if (tsMoneda.KeyValue = 'DLL') or (tsMoneda.KeyValue = 'dll') then
                      begin
                          if tdCostoNuevo.Value <> insumos.FieldValues['dVentaDLL'] then
                          begin
                              if anexo_ocompras.FieldValues['dCambio'] > 0 then
                                 Connection.zcommand.Params.ParamByName('Costo').value   := (insumos.FieldValues['dVentaDLL'] * anexo_ocompras.FieldValues['dCambio'])
                              else
                                 Connection.zcommand.Params.ParamByName('Costo').value   := insumos.FieldValues['dVentaDLL'];
                          end
                          else
                              Connection.zcommand.Params.ParamByName('Costo').value   := tdCostoNuevo.Value;
                      end
                      else
                         if tdCostoNuevo.Value <> insumos.FieldValues['dVentaMN'] then
                            Connection.zcommand.Params.ParamByName('Costo').value    := tdCostoNuevo.Value
                         else
                            Connection.zcommand.Params.ParamByName('Costo').value    := insumos.FieldValues['dVentaMN'];
                      Connection.zcommand.ExecSQL ;
                  End;
            end;
            SavePlace := anexo_ocompras.GetBookmark ;
            anexo_pocompras.Refresh ;

            Try
               anexo_pocompras.GotoBookmark(SavePlace);
            Except
            Else
              anexo_pocompras.FreeBookmark(SavePlace);
            End ;

            //Ahora actualizamos los precios del almacen..
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Update insumos set dNuevoPrecio =:NvoCosto where sIdInsumo =:Insumo ');
            connection.zCommand.ParamByName('Insumo').AsString  := insumos.FieldValues['sIdInsumo'];
            connection.zCommand.ParamByName('NvoCosto').AsFloat := tdCostoNuevo.Value;
            connection.zCommand.ExecSQL;
        End
        Else
        Try
            SavePlace := anexo_pocompras.GetBookmark ;

            //Primero buscamos los insumos..
            Connection.QryBusca.Active := False ;
            Connection.QryBusca.SQL.Clear ;
            Connection.QryBusca.SQL.Add('select p.*, i.dVentaMN, i.dVentaDLL from anexo_ppedido p '+
                                        'inner join insumos i on (i.sContrato = p.sContrato and i.sIdInsumo = p.sIdInsumo) '+
                                        'WHERE p.sContrato =:Contrato And p.iFolioPedido =:Folio and p.sIdInsumo =:Insumo and p.iItem =:Item ');
            Connection.QryBusca.Params.ParamByName('Contrato').DataType  := ftString ;
            Connection.QryBusca.Params.ParamByName('Contrato').value     := Global_Contrato ;
            Connection.QryBusca.Params.ParamByName('Folio').DataType     := ftInteger ;
            Connection.QryBusca.Params.ParamByName('Folio').value        := anexo_ocompras.FieldValues['iFolioPedido'];
            Connection.QryBusca.Params.ParamByName('Insumo').DataType    := ftString ;
            Connection.QryBusca.Params.ParamByName('Insumo').value       := anexo_pocompras.FieldValues['sIdInsumo'];
            Connection.QryBusca.Params.ParamByName('Item').DataType      := ftString ;
            Connection.QryBusca.Params.ParamByName('Item').value         := anexo_pocompras.FieldValues['iItem'];
            Connection.QryBusca.Open ;

            if connection.QryBusca.RecordCount > 0 then
            begin
                Connection.zCommand.Active := False ;
                Connection.zCommand.SQL.Clear ;
                Connection.zCommand.SQL.Add('UPDATE anexo_ppedido SET dCantidad = :Cantidad, dCosto =:Costo, dDescuento =:Descuento ' +
                                            'WHERE sContrato = :Contrato And iFolioPedido = :Folio And sIdInsumo =:Insumo And iItem = :Item');
                Connection.zcommand.Params.ParamByName('Contrato').DataType     := ftString ;
                Connection.zcommand.Params.ParamByName('Contrato').value        := Global_Contrato ;
                Connection.zcommand.Params.ParamByName('Folio').DataType        := ftInteger ;
                Connection.zcommand.Params.ParamByName('Folio').value           := anexo_ocompras.FieldValues['iFolioPedido'] ;
                Connection.zcommand.Params.ParamByName('Insumo').DataType       := ftString ;
                Connection.zcommand.Params.ParamByName('Insumo').value          := anexo_pocompras.FieldValues['sIdInsumo'] ;
                Connection.zcommand.Params.ParamByName('Item').DataType         := ftString ;
                Connection.zcommand.Params.ParamByName('Item').value            := anexo_pocompras.FieldValues['iItem'] ;
                Connection.zcommand.Params.ParamByName('Cantidad').DataType     := ftFloat ;
                Connection.zcommand.Params.ParamByName('Cantidad').value        := tdCantidad.Value ;
                Connection.zcommand.Params.ParamByName('Costo').DataType        := ftFloat ;
                if (tsMoneda.KeyValue = 'DLL') or (tsMoneda.KeyValue = 'dll') then
                begin
                    if tdCostoNuevo.Value <> insumos.FieldValues['dVentaDLL'] then
                    begin
                        if anexo_ocompras.FieldValues['dCambio'] > 0 then
                           Connection.zcommand.Params.ParamByName('Costo').value   := (insumos.FieldValues['dVentaDLL'] * anexo_ocompras.FieldValues['dCambio'])
                        else
                           Connection.zcommand.Params.ParamByName('Costo').value   := insumos.FieldValues['dVentaDLL'];
                    end
                    else
                        Connection.zcommand.Params.ParamByName('Costo').value   := tdCostoNuevo.Value;
                end
                else
                   if tdCostoNuevo.Value <> insumos.FieldValues['dVentaMN'] then
                      Connection.zcommand.Params.ParamByName('Costo').value    := tdCostoNuevo.Value
                   else
                      Connection.zcommand.Params.ParamByName('Costo').value    := insumos.FieldValues['dVentaMN'];
                Connection.zcommand.Params.ParamByName('Descuento').DataType   := ftFloat ;
                Connection.zcommand.Params.ParamByName('Descuento').value      := tDescuentoMat.Value ;
                Connection.zcommand.ExecSQL ;

                 //Ahora actualizamos los precios del almacen..
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('Update insumos set dNuevoPrecio =:NvoCosto where sIdInsumo =:Insumo ');
                connection.zCommand.ParamByName('Insumo').AsString  := anexo_pocompras.FieldValues['sIdInsumo'];
                connection.zCommand.ParamByName('NvoCosto').AsFloat := tdCostoNuevo.Value;
                connection.zCommand.ExecSQL;

                anexo_pocompras.Refresh ;
                anexo_pocompras.GotoBookmark(SavePlace);
                anexo_pocompras.FreeBookmark(SavePlace);
            end
            else
               messageDLG('No se encontró el Insumo en el Catalogo de Materiales!', mtWarning, [mbOk], 0);
        except
          //  MessageDlg('Ocurrio un error al actualizar el registro', mtWarning, [mbOk], 0);
          on e : exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Ordenes de Compra', 'Al actualizar registro', 0);
          end;
        end ;

        If dbRequisicion.Text <> '' then
        begin
            //Actualizamos el estatus del material si es una requisicion;
            Connection.zCommand.Active := False ;
            Connection.zCommand.SQL.Clear ;
            Connection.zCommand.SQL.Add('UPDATE anexo_prequisicion SET sEstado ="COMPRADO", iFolioPedido =:Pedido ' +
                                        'WHERE sContrato =:Contrato And iFolioRequisicion =:Folio And sIdInsumo =:Insumo And iItem =:Item and sNumeroActividad =:Actividad ');
            Connection.zcommand.Params.ParamByName('Contrato').DataType  := ftString ;
            Connection.zcommand.Params.ParamByName('Contrato').value     := Global_Contrato ;
            Connection.zcommand.Params.ParamByName('Folio').DataType     := ftInteger ;
            Connection.zcommand.Params.ParamByName('Folio').value        := rxPartidas.FieldValues['iFolioRequisicion'];
            Connection.zcommand.Params.ParamByName('Insumo').DataType    := ftString ;
            Connection.zcommand.Params.ParamByName('Insumo').value       := insumos.FieldValues['sIdInsumo'];
            Connection.zcommand.Params.ParamByName('Item').DataType      := ftInteger;
            Connection.zcommand.Params.ParamByName('Item').value         := insumos.FieldValues['iItem'];
            Connection.zcommand.Params.ParamByName('Actividad').DataType := ftString ;
            Connection.zcommand.Params.ParamByName('Actividad').value    := insumos.FieldValues['sNumeroActividad'];
            Connection.zcommand.Params.ParamByName('Pedido').DataType    := ftInteger ;
            Connection.zcommand.Params.ParamByName('Pedido').value       := anexo_ocompras.FieldValues['iFolioPedido'];
            Connection.zcommand.ExecSQL;

            //Ahora consultamos cuales partidas estan pendientes para seguir mostrando o no la requisicion..
            Connection.zCommand.Active := False ;
            Connection.zCommand.SQL.Clear ;
            Connection.zCommand.SQL.Add('select * from anexo_prequisicion ' +
                                        'WHERE sContrato =:Contrato And iFolioRequisicion =:Folio ');
            Connection.zcommand.Params.ParamByName('Contrato').DataType  := ftString ;
            Connection.zcommand.Params.ParamByName('Contrato').value     := Global_Contrato ;
            Connection.zcommand.Params.ParamByName('Folio').DataType     := ftInteger ;
            Connection.zcommand.Params.ParamByName('Folio').value        := insumos.FieldValues['Requisicion'];
            Connection.zcommand.Open;

            conteo := 0;
            while not connection.zCommand.Eof do
            begin
                if connection.zCommand.FieldValues['sEstado'] <> 'PENDIENTE' then
                   inc(conteo);
                connection.zCommand.Next;
            end;

            if conteo = connection.zCommand.RecordCount then
            begin
                Connection.zCommand.Active := False ;
                Connection.zCommand.SQL.Clear ;
                Connection.zCommand.SQL.Add('Update anexo_requisicion set sEstado = "ACEPTADA" ' +
                                            'WHERE sContrato =:Contrato And iFolioRequisicion =:Folio ');
                Connection.zcommand.Params.ParamByName('Contrato').DataType  := ftString ;
                Connection.zcommand.Params.ParamByName('Contrato').value     := Global_Contrato ;
                Connection.zcommand.Params.ParamByName('Folio').DataType     := ftInteger ;
                Connection.zcommand.Params.ParamByName('Folio').value        := StrToInt(anexo_ocompras.FieldValues['sFolioRequisicion']);
                Connection.zcommand.ExecSQL;
            end;

            Insumos.Refresh;
            Requisiciones.Refresh;
        end;

        txtMaterial.ReadOnly := True ;
        tdCantidad.ReadOnly  := True ;
        Insertar1.Enabled := True ;
        Editar1.Enabled := True ;
        Registrar1.Enabled := False ;
        Can1.Enabled := False ;
        Eliminar1.Enabled := True ;
        Refresh1.Enabled := True ;
        frmBarra1.btnPostClick(Sender);
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmPedidos.frmBarra1btnCancelClick(Sender: TObject);
begin
  Insertar1.Enabled    := True ;
  Editar1.Enabled      := True ;
  Registrar1.Enabled   := False ;
  Can1.Enabled         := False ;
  Eliminar1.Enabled    := True ;
  Refresh1.Enabled     := True ;
  txtMaterial.ReadOnly := True ;
  tdCantidad.ReadOnly  := True;
  if txtMaterial.Enabled then
  begin
      txtMaterial.Text := '';
      filtra;
  end;
  frmBarra1.btnCancelClick(Sender);
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmPedidos.frmBarra1btnExitClick(Sender: TObject);
begin
  Insertar1.Enabled := True ;
  Editar1.Enabled := True ;
  Registrar1.Enabled := False ;
  Can1.Enabled := False ;
  Eliminar1.Enabled := True ;
  Refresh1.Enabled := True ;
  frmBarra1.btnExitClick(Sender);
end;

procedure TfrmPedidos.frmBarra1btnDeleteClick(Sender: TObject);
Var
  SavePlace : TBookmark;
begin
    if anexo_pocompras.FieldValues['sStatus'] = 'Entregado' then
    begin
        messageDLG('No se puede modificar el Material ya fue Entregado!', mtWarning, [mbOk], 0);
        exit;
    end;

    validaciones;
    if valida then
    begin
          frmBarra1.btnCancel.Click ;
          exit;
    end;
    If anexo_pocompras.RecordCount > 0 Then
    Begin
        With connection do
        Begin
            try
                SavePlace := anexo_pocompras.GetBookmark ;
                zCommand.Active := False ;
                zCommand.SQL.Clear ;
                zCommand.SQL.Add('Delete from anexo_ppedido where sContrato = :Contrato And ' +
                                 'iFolioPedido = :Folio and sIdInsumo =:Insumo And sNumeroActividad = :Actividad And iItem = :Item' );
                zcommand.Params.ParamByName('Contrato').DataType  := ftString ;
                zcommand.Params.ParamByName('Contrato').value     := Global_Contrato ;
                zcommand.Params.ParamByName('Folio').DataType     := ftInteger ;
                zcommand.Params.ParamByName('Folio').value        := anexo_pocompras.FieldValues['iFolioPedido'] ;
                zcommand.Params.ParamByName('Insumo').DataType    := ftString ;
                zcommand.Params.ParamByName('Insumo').value       := anexo_pocompras.FieldValues['sIdInsumo'] ;
                zcommand.Params.ParamByName('Actividad').DataType := ftString ;
                zcommand.Params.ParamByName('Actividad').value    := anexo_pocompras.FieldValues['sNumeroActividad'] ;
                zcommand.Params.ParamByName('Item').DataType      := ftInteger ;
                zcommand.Params.ParamByName('Item').value         := anexo_pocompras.FieldValues['iItem'] ;
                zcommand.ExecSQL;

                if dbRequisicion.Text <> '' then
                begin
                    Connection.zCommand.Active := False ;
                    Connection.zCommand.SQL.Clear ;
                    Connection.zCommand.SQL.Add('Update anexo_prequisicion set sEstado = "PENDIENTE", iFolioPedido = 0 ' +
                                                'WHERE sContrato =:Contrato And iFolioRequisicion =:Folio and sIdInsumo =:Insumo and sNumeroActividad =:Actividad');
                    Connection.zcommand.Params.ParamByName('Contrato').DataType  := ftString ;
                    Connection.zcommand.Params.ParamByName('Contrato').value     := Global_Contrato ;
                    Connection.zcommand.Params.ParamByName('Folio').DataType     := ftInteger ;
                    Connection.zcommand.Params.ParamByName('Folio').value        := rxPartidas.FieldValues['iFolioRequisicion'];
                    Connection.zcommand.Params.ParamByName('Insumo').DataType    := ftString ;
                    Connection.zcommand.Params.ParamByName('Insumo').value       := anexo_pocompras.FieldValues['sIdInsumo'] ;
                    Connection.zcommand.Params.ParamByName('Actividad').DataType := ftString ;
                    Connection.zcommand.Params.ParamByName('Actividad').value    := anexo_pocompras.FieldValues['sNumeroActividad'] ;
                    Connection.zcommand.ExecSQL;

                    Connection.zCommand.Active := False ;
                    Connection.zCommand.SQL.Clear ;
                    Connection.zCommand.SQL.Add('Update anexo_requisicion set sEstado = "PENDIENTE" ' +
                                                'WHERE sContrato =:Contrato And iFolioRequisicion =:Folio ');
                    Connection.zcommand.Params.ParamByName('Contrato').DataType  := ftString ;
                    Connection.zcommand.Params.ParamByName('Contrato').value     := Global_Contrato ;
                    Connection.zcommand.Params.ParamByName('Folio').DataType     := ftInteger ;
                    Connection.zcommand.Params.ParamByName('Folio').value        := rxPartidas.FieldValues['iFolioRequisicion'];
                    Connection.zcommand.ExecSQL;

                    insumos.Refresh;
                    Requisiciones.Refresh;
                end;

                SavePlace := anexo_ocompras.GetBookmark ;
                anexo_pocompras.Refresh ;

                Try
                   anexo_pocompras.GotoBookmark(SavePlace);
                Except
                Else
                  anexo_pocompras.FreeBookmark(SavePlace);
                End ;
            Except
              on e : exception do begin
              UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Ordenes de Compra', 'Al eliminar registro', 0);
              end;
            End ;
            GridPartidas.SetFocus
        End
    End
end;

procedure TfrmPedidos.Insertar1Click(Sender: TObject);
begin
    frmBarra2.btnAdd.Click
end;

procedure TfrmPedidos.InsumosAfterScroll(DataSet: TDataSet);
begin
      if insumos.RecordCount > 0 then
      begin
          if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
             tdCantidad.Value := insumos.FieldValues['dCantidad']
          else
             tdCantidad.Value := 0;
          tdCostoNuevo.Value := insumos.FieldValues['dNuevoPrecio'];
      end;
end;

procedure TfrmPedidos.Editar1Click(Sender: TObject);
begin
    frmBarra2.btnEdit.Click
end;

procedure TfrmPedidos.Registrar1Click(Sender: TObject);
begin
    frmBarra2.btnPost.Click
end;

procedure TfrmPedidos.Can1Click(Sender: TObject);
begin
    frmBarra2.btnCancel.Click
end;

procedure TfrmPedidos.chkRequisicionClick(Sender: TObject);
begin
    if chkRequisicion.Checked then
       dbRequisicion.Enabled := True
    else
       dbRequisicion.Enabled := False;
end;

procedure TfrmPedidos.cmdAceptaClick(Sender: TObject);
var
   Item    : string;
   conteo, lineas  : integer;
   Insumo  : string;
begin
     try
        if filtroUno.Visible then
        begin
            //Reporte Para proveedores...
            if lblDatos.Caption = 'Proveedores:' then
            begin
                if cbxDatos.Text = 'Todos los provedores' then
                   Item := '%'
                else
                   Item := copy(cbxDatos.Text, 0, pos('.-', cbxDatos.Text) - 2);

                Reporte.Active := False ;
                Reporte.SQL.Clear;
                Reporte.SQL.Add('Select a2.*, p.*, a1.iItem, a1.dCantidad, a1.mDescripcion, a1.sMedida, sum(a1.dCantidad * (a1.dCosto - a1.dDescuento)) as dCosto, "" as iItemOrden, u.sNombre as sElabora, m.sDescripcion as moneda, a1.dDescuento as DescuentoMat '+
                                'from anexo_ppedido a1 '+
                                'inner join anexo_pedidos a2 on (a1.sContrato = a2.sContrato And a1.iFolioPedido = a2.iFolioPedido) '+
                                'inner join proveedores p on (a2.sIdProveedor = p.sIdProveedor) '+
                                'left join usuarios u on (u.sIdUsuario = a2.sElaboro) '+
                                'inner join tiposdemoneda m on (a2.sMoneda = m.sIdMoneda) '+
                                'Where a1.sContrato =:Contrato and a2.sIdProveedor like "'+Item+'" group by a2.iFolioPedido order by a2.dIdFecha, a2.sOrdenCompra');
                Reporte.Params.ParamByName('Contrato').DataType  := ftString ;
                Reporte.Params.ParamByName('Contrato').Value     := global_contrato ;
                Reporte.Open ;

                frxEntrada.PreviewOptions.MDIChild  := False ;
                frxEntrada.PreviewOptions.Modal     := True ;
                frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
                frxEntrada.PreviewOptions.ShowCaptions := False ;
                frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
                frxEntrada.LoadFromFile (global_files +global_miReporte+ '_oCompraxProveedor.fr3') ;
                frxentrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
            end;

            //Reporte Para Referencias...
            if lblDatos.Caption = 'Referencia:' then
            begin
                if cbxDatos.Text = 'Todos las Referencias' then
                   Item := '%'
                else
                   Item := cbxDatos.Text;

                Reporte.Active := False ;
                Reporte.SQL.Clear;
                Reporte.SQL.Add('Select a2.*, p.*, a1.iItem, a1.dCantidad, a1.mDescripcion, a1.sMedida, sum(a1.dCantidad * (a1.dCosto - a1.dDescuento)) as dCosto, "" as iItemOrden, u.sNombre as sElabora, m.sDescripcion as moneda, a1.dDescuento as DescuentoMat '+
                                'from anexo_ppedido a1 '+
                                'inner join anexo_pedidos a2 on (a1.sContrato = a2.sContrato And a1.iFolioPedido = a2.iFolioPedido) '+
                                'inner join proveedores p on (a2.sIdProveedor = p.sIdProveedor) '+
                                'left join usuarios u on (u.sIdUsuario = a2.sElaboro) '+
                                'inner join tiposdemoneda m on (a2.sMoneda = m.sIdMoneda) '+
                                'Where a1.sContrato =:Contrato and a2.sNumeroOrden like "'+Item+'" group by a2.iFolioPedido order by a2.dIdFecha, a2.sOrdenCompra');
                Reporte.Params.ParamByName('Contrato').DataType   := ftString ;
                Reporte.Params.ParamByName('Contrato').Value      := global_contrato ;
                Reporte.Open ;

                frxEntrada.PreviewOptions.MDIChild  := False ;
                frxEntrada.PreviewOptions.Modal     := True ;
                frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
                frxEntrada.PreviewOptions.ShowCaptions := False ;
                frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
                frxEntrada.LoadFromFile (global_files +global_miReporte+ '_oCompraxReferencia.fr3') ;
                frxentrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
            end;

            //Reporte Para Comparativo...
            if lblDatos.Caption = 'Precios Materiales:' then
            begin
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                if cbxDatos.Text = 'Todos los Materiales de la O.C.' then
                begin
                    connection.zCommand.SQL.Add('select distinct sIdInsumo, iFolioPedido, mDescripcion, sMedida from anexo_ppedido where sContrato =:Contrato and iFolioPedido =:Pedido ');
                    connection.zCommand.ParamByName('Contrato').AsString := global_contrato ;
                    connection.zCommand.ParamByName('Pedido').AsInteger  := anexo_ocompras.FieldValues['iFolioPedido'];
                end
                else
                begin
                    connection.zCommand.SQL.Add('select distinct sIdInsumo, iFolioPedido, mDescripcion, sMedida from anexo_ppedido where sContrato =:Contrato and iFolioPedido =:Pedido and sIdInsumo =:Insumo');
                    connection.zCommand.ParamByName('Contrato').AsString := global_contrato ;
                    connection.zCommand.ParamByName('Pedido').AsInteger  := anexo_ocompras.FieldValues['iFolioPedido'];
                    connection.zCommand.ParamByName('Insumo').AsString   := copy(cbxDatos.Text, 0, pos('.-', cbxDatos.Text) - 2);
                end;
                connection.zCommand.Open;

                rxPrecios.EmptyTable;
                conteo := 1;
                Insumo := '';
                lineas := 1;
                while not connection.zCommand.Eof do
                begin
                    connection.QryBusca.Active := False;
                    connection.QryBusca.SQL.Clear;
                    connection.QryBusca.SQL.Add('select pp.dCosto, p.dIdFecha, p.sOrdenCompra, p.sIdProveedor from anexo_ppedido pp '+
                                                'inner join anexo_pedidos p on (p.sContrato = pp.sContrato and p.iFolioPedido = pp.iFolioPedido ) '+
                                                'where pp.sContrato = :Contrato and pp.sIdInsumo =:Insumo group by p.dIdFecha, pp.dCosto order by p.dIdFecha');
                    connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
                    connection.QryBusca.ParamByName('Insumo').AsString   := connection.zCommand.FieldValues['sIdInsumo'];
                    connection.QryBusca.Open;

                    while not connection.QryBusca.Eof do
                    begin
                        if conteo > 8 then
                        begin
                           conteo := 1;
                           inc(lineas);
                        end;

                        if conteo = 1 then
                           rxPrecios.Append
                        else
                           rxPrecios.Edit;

                        rxPrecios.FieldValues['sContrato']    := global_contrato;
                        rxPrecios.FieldValues['Item']         := lineas;
                        rxPrecios.FieldValues['sIdInsumo']    := connection.zCommand.FieldValues['sIdInsumo'];
                        rxPrecios.FieldValues['mDescripcion'] := connection.zCommand.FieldValues['mDescripcion'];
                        rxPrecios.FieldValues['sIdProveedor'] := connection.QryBusca.FieldValues['sIdProveedor'];
                        rxPrecios.FieldValues['sOrdenCompra'+ IntToStr(conteo) ] := connection.QryBusca.FieldValues['sOrdenCompra'];
                        rxPrecios.FieldValues['dFecha'+ IntToStr(conteo)] := connection.QryBusca.FieldValues['dIdFecha'];
                        rxPrecios.FieldValues['dCosto'+ IntToStr(conteo)] := connection.QryBusca.FieldValues['dCosto'];
                        rxPrecios.Post;

                        inc(conteo);
                        connection.QryBusca.Next;

                    end;
                    connection.zCommand.Next;
                    Insumo := connection.zCommand.FieldValues['sIdInsumo'];
                    lineas := 1;
                    conteo := 1;
                end;

                frxEntrada.PreviewOptions.MDIChild  := False ;
                frxEntrada.PreviewOptions.Modal     := True ;
                frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
                frxEntrada.PreviewOptions.ShowCaptions := False ;
                frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
                frxEntrada.LoadFromFile (global_files +global_miReporte+ '_oCompraComparativo.fr3') ;
                frxentrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
            end;

        end
        else
        begin
            Reporte.Active := False ;
            Reporte.SQL.Clear;
            Reporte.SQL.Add('Select a2.*, p.*, a1.iItem, a1.dCantidad, a1.mDescripcion, a1.sMedida, sum(a1.dCantidad * (a1.dCosto - a1.dDescuento)) as dCosto, "" as iItemOrden, u.sNombre as sElabora, m.sDescripcion as moneda, a1.dDescuento as DescuentoMat '+
                            'from anexo_ppedido a1 '+
                            'inner join anexo_pedidos a2 on (a1.sContrato = a2.sContrato And a1.iFolioPedido = a2.iFolioPedido) '+
                            'inner join proveedores p on (a2.sIdProveedor = p.sIdProveedor) '+
                            'left join usuarios u on (u.sIdUsuario = a2.sElaboro) '+
                            'inner join tiposdemoneda m on (a2.sMoneda = m.sIdMoneda) '+
                            'Where a1.sContrato =:Contrato and a2.dIdFecha >=:FechaI and a2.dIdFecha <=:FechaF group by a2.iFolioPedido order by a2.dIdFecha, a2.sOrdenCompra');
            Reporte.Params.ParamByName('Contrato').DataType := ftString ;
            Reporte.Params.ParamByName('Contrato').Value    := global_contrato ;
            Reporte.Params.ParamByName('FechaI').DataType   := ftDate ;
            Reporte.Params.ParamByName('FechaI').Value      := FechaInicio.Date;
            Reporte.Params.ParamByName('FechaF').DataType   := ftDate ;
            Reporte.Params.ParamByName('FechaF').Value      := FechaFinal.Date;
            Reporte.Open ;

            frxEntrada.PreviewOptions.MDIChild  := False ;
            frxEntrada.PreviewOptions.Modal     := True ;
            frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
            frxEntrada.PreviewOptions.ShowCaptions := False ;
            frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
            frxEntrada.LoadFromFile (global_files +global_miReporte+ '_oCompraxPeriodo.fr3') ;
            frxentrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
        end;

     except
     on e : exception do
     begin
         UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al imprimir', 0);
     end;
  end;
  Grid_Entradas.Enabled := True;
  Grid_insumos.Enabled  := True;
  GridPartidas.Enabled  := True;
end;

procedure TfrmPedidos.cmdCancelarClick(Sender: TObject);
begin
     cbxDatos.Clear;
     Grid_Entradas.Enabled := True;
     Grid_insumos.Enabled  := True;
     GridPartidas.Enabled  := True;
     PanelImprime.Visible  := False;
end;

procedure TfrmPedidos.Eliminar1Click(Sender: TObject);
begin
    frmBarra2.btnDelete.Click
end;

procedure TfrmPedidos.Refresh1Click(Sender: TObject);
begin
    frmBarra2.btnRefresh.Click
end;

procedure TfrmPedidos.Imprimir1Click(Sender: TObject);
begin
try
    If (tsNumeroOrden.Text = 'CONTRATO No. ' + global_contrato) Then
        rDiarioFirmas (global_contrato, '', 'A',anexo_ocompras.FieldValues['dIdFecha'], frmPedidos )
    Else
        rDiarioFirmas (
        global_contrato, tsNumeroOrden.Text, 'A',anexo_ocompras.FieldValues['dIdFecha'], frmPedidos) ;


    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select Sum(r.dCantidad * r.dCosto) as dMontoMN From anexo_ppedido r ' +
                                'Where r.sContrato = :Contrato And r.iFolioPedido = :Folio Group By r.iFolioPedido');
    connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato ;
    connection.qryBusca.Params.ParamByName('Folio').DataType := ftInteger ;
    connection.qryBusca.Params.ParamByName('Folio').Value := anexo_ocompras.FieldValues['iFolioPedido'] ;
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 Then
            MontoTotal :=  Connection.qryBusca.fieldValues['dMontoMN'] ;

    Reporte.Active := False ;
    Reporte.SQL.Clear;
    if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
       Reporte.SQL.Add(CadenaReporte)
    else
       Reporte.SQL.Add(CadenaReporte2);
    Reporte.Params.ParamByName('Contrato').DataType := ftString ;
    Reporte.Params.ParamByName('Contrato').Value    := global_contrato ;
    if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
    begin
        Reporte.Params.ParamByName('Convenio').DataType := ftString ;
        Reporte.Params.ParamByName('Convenio').Value    := global_convenio ;
    end;
    Reporte.Params.ParamByName('Folio').DataType    := ftInteger ;
    Reporte.Params.ParamByName('Folio').Value       := anexo_ocompras.FieldValues['iFolioPedido'] ;
    Reporte.Open ;

    frxEntrada.PreviewOptions.MDIChild  := False ;
    frxEntrada.PreviewOptions.Modal     := True ;
    frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
    frxEntrada.PreviewOptions.ShowCaptions := False ;
    frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
    frxEntrada.LoadFromFile (global_files +global_miReporte+ '_oCompra.fr3') ;
    frxentrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
  except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al imprimir', 0);
  end;
  end;
end;

procedure TfrmPedidos.ImprimirComprasxProveedor1Click(Sender: TObject);
begin
    zqDatos.Active := False;
    zqDatos.SQL.Clear;
    zqDatos.SQL.Add('select sIdProveedor, sRazon from proveedores ');
    zqDatos.Open;

    cbxDatos.Clear;
    cbxDatos.Items.Add('Todos los provedores');
    while not zqDatos.Eof do
    begin
        cbxDatos.Items.Add(zqDatos.FieldValues['sIdProveedor']+ ' .- '+ zqDatos.FieldValues['sRazon']);
        zqDatos.Next;
    end;
    cbxDatos.ItemIndex := 0;
    Grid_Entradas.Enabled := False;
    Grid_insumos.Enabled  := False;
    GridPartidas.Enabled  := False;
    lblDatos.Caption := 'Proveedores:';
    PanelImprime.Visible := True;
    FiltroUno.Visible    := True;
    FiltroDos.Visible    := False;

end;

procedure TfrmPedidos.ImprimirComprasxReferencia1Click(Sender: TObject);
begin
    zqDatos.Active := False;
    zqDatos.SQL.Clear;
    zqDatos.SQL.Add('select sNumeroOrden from ordenesdetrabajo where sContrato =:Contrato');
    zqDatos.ParamByName('Contrato').AsString := global_contrato;
    zqDatos.Open;

    cbxDatos.Clear;
    cbxDatos.Items.Add('Todos las Referencias');
    while not zqDatos.Eof do
    begin
        cbxDatos.Items.Add(zqDatos.FieldValues['sNumeroOrden']);
        zqDatos.Next;
    end;
    cbxDatos.ItemIndex := 0;
    Grid_Entradas.Enabled := False;
    Grid_insumos.Enabled  := False;
    GridPartidas.Enabled  := False;
    lblDatos.Caption := 'Referencia:';
    PanelImprime.Visible := True;
    FiltroUno.Visible    := True;
    FiltroDos.Visible    := False;
end;

procedure TfrmPedidos.ImprimirResumendeOC1Click(Sender: TObject);
begin
   Grid_Entradas.Enabled := False;
   Grid_insumos.Enabled  := False;
   GridPartidas.Enabled  := False;
   PanelImprime.Visible  := True;
   FiltroUno.Visible     := False;
   FiltroDos.Visible     := True;
   FechaInicio.Date      := date;
   FechaFinal.Date       := date;
end;

procedure TfrmPedidos.ImprimirVariaciondePrecios1Click(Sender: TObject);
begin
    zqDatos.Active := False;
    zqDatos.SQL.Clear;
    zqDatos.SQL.Add('select sIdInsumo, substr(mDescripcion, 1,255) as Descripcion from anexo_ppedido where sContrato =:Contrato and iFolioPedido =:Folio ');
    zqDatos.ParamByName('Contrato').AsString := global_contrato;
    zqDatos.ParamByName('Folio').AsInteger   := anexo_ocompras.FieldValues['iFolioPedido'];
    zqDatos.Open;

    cbxDatos.Clear;
    cbxDatos.Items.Add('Todos los Materiales de la O.C.');
    while not zqDatos.Eof do
    begin
        cbxDatos.Items.Add(zqDatos.FieldValues['sIdInsumo'] + ' .- '+ zqDatos.FieldValues['Descripcion']);
        zqDatos.Next;
    end;
    cbxDatos.ItemIndex := 0;
    Grid_Entradas.Enabled := False;
    Grid_insumos.Enabled  := False;
    GridPartidas.Enabled  := False;
    lblDatos.Caption := 'Precios Materiales:';
    PanelImprime.Visible := True;
    FiltroUno.Visible    := True;
    FiltroDos.Visible    := False;
end;

procedure TfrmPedidos.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

procedure TfrmPedidos.SeguimientoMaterialGeneral1Click(Sender: TObject);
begin
  try
    Seguimiento_Material('');
    frxEntrada.PreviewOptions.MDIChild := False ;
    frxEntrada.PreviewOptions.Modal := True ;
    frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
    frxEntrada.PreviewOptions.ShowCaptions := False ;
    frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
    frxEntrada.LoadFromFile (global_files + 'seguimiento_materialxpartida.fr3') ;
    frxEntrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
  except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'en el proceso Seguimiento Material General', 0);
  end;
  end;
end;

procedure TfrmPedidos.SeguimientoMaterialxPartida1Click(Sender: TObject);
begin
  try
   if anexo_pocompras.RecordCount > 0 then
       Seguimiento_Material(anexo_pocompras.FieldValues['sNumeroActividad'])
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
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'en el proceso Seguimiento Material x Partida', 'Al cancelar registro', 0);
  end;
  end;

end;

procedure TfrmPedidos.SeguimientoMaterialxPartidaDetalle1Click(Sender: TObject);
var
   x, y, num, i, contador : integer;
   SumCantidad, SumTotal, SumExcedente, SumRestante : double;
   Cadena   : string;
begin
  try
    if anexo_pocompras.RecordCount = 0 then
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
    connection.zCommand.ParamByName('Actividad').AsString := anexo_pocompras.FieldValues['sNumeroActividad'];
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
                                          'on (i.sContrato = ra.sContrato and i.sIdInsumo = ra.sIdInsumo) '+
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

              SumTotal     := 0;
              SumExcedente := 0;

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

  except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'En el proceso Seguimiento Material x Partida Detalle', 0);
  end;
  end;
end;

procedure TfrmPedidos.anexo_ocomprasAfterScroll(DataSet: TDataSet);
begin
    if (OpcButton1 = '') then
    begin
        If frmbarra2.btnCancel.Enabled = False then
           frmBarra2.btnCancel.Click ;
    end;

    If anexo_ocompras.RecordCount > 0 Then
    Begin
        if Not anexo_ocompras.FieldByName('iFolioPedido').IsNull then
           iFolio         := anexo_ocompras.FieldValues['iFolioPedido'] ;
        if Not anexo_ocompras.FieldByName('sOrdenCompra').IsNull then
           tiFolio.Text   := anexo_ocompras.FieldValues['sOrdenCompra'] ;
        if Not anexo_ocompras.FieldByName('dIdFecha').IsNull then
           tdIdFecha.Date         := anexo_ocompras.FieldValues['dIdFecha'] ;
        if Not anexo_ocompras.FieldByName('sIdProveedor').IsNull then
           tsIdProveedor.KeyValue := anexo_ocompras.FieldValues['sIdProveedor'] ;
        if not anexo_ocompras.FieldByName('sNumeroOrden').IsNull then
           tsNumeroOrden.Text := anexo_ocompras.FieldValues['sNumeroOrden'] ;
        if not anexo_ocompras.FieldByName('mComentarios').IsNull then
           tmComentarios.Text     := anexo_ocompras.FieldValues['mComentarios'] ;
        if not anexo_ocompras.FieldByName('sFormaPago').IsNull then
           tsFormaPago.KeyValue   := anexo_ocompras.FieldValues['sFormaPago'];
        if not anexo_ocompras.FieldByName('sFolioRequisicion').IsNull then
          dbRequisicion.Text     := anexo_ocompras.FieldValues['sFolioRequisicion'];
        if not anexo_ocompras.FieldByName('sMoneda').IsNull then
          tsMoneda.KeyValue    := anexo_ocompras.FieldValues['sMoneda'];

        if not anexo_ocompras.FieldByName('sLugarEntrega').IsNull then
          tsEmbarca.Text     := anexo_ocompras.FieldValues['sLugarEntrega'];
        if not anexo_ocompras.FieldByName('sEntrega').IsNull then
          tsEntrega.Text     := anexo_ocompras.FieldValues['sEntrega'];
        if not anexo_ocompras.FieldByName('sCondiciones').IsNull then
          tsCondiciones.Text := anexo_ocompras.FieldValues['sCondiciones'];
        if not anexo_ocompras.FieldByName('sPrecios').IsNull then
          tsPrecios.Text     := anexo_ocompras.FieldValues['sPrecios'];
        if not anexo_ocompras.FieldByName('sVigencia').IsNull then
          tsVigencia.Text    := anexo_ocompras.FieldValues['sVigencia'];
        if not anexo_ocompras.FieldByName('sVendedor').IsNull then
          tsVendedor.Text    := anexo_ocompras.FieldValues['sVendedor'];
        if not anexo_ocompras.FieldByName('sMail').IsNull then
          tsMail.Text        := anexo_ocompras.FieldValues['sMail'];
        if not anexo_ocompras.FieldByName('dDescuento').IsNull then
          tDescuentoGral.Value  := anexo_ocompras.FieldValues['dDescuento'];

       if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
       begin
            lblPartida.Caption  := 'No. Partida';
            txtMaterial.Visible := False;
            dbPartidas.Visible  := True;
       end
       else
       begin
            lblPartida.Caption  := 'Material';
            txtMaterial.Visible := True;
            dbPartidas.Visible  := False;
       end;

       anexo_pocompras.Active := False ;
       anexo_pocompras.SQL.Clear;
       if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
          anexo_pocompras.SQL.Add(CadenaOrden)
       else
          anexo_pocompras.SQL.Add(CadenaOrden2);
       anexo_pocompras.Params.ParamByName('Contrato').DataType := ftString ;
       anexo_pocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
       if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
       begin
            anexo_pocompras.Params.ParamByName('Convenio').DataType := ftString ;
            anexo_pocompras.Params.ParamByName('Convenio').Value    := global_convenio;
       end;
       anexo_pocompras.Params.ParamByName('Folio').DataType := ftInteger ;
       if Not anexo_ocompras.FieldByName('iFolioPedido').IsNull then
          anexo_pocompras.Params.ParamByName('Folio').Value    := anexo_ocompras.FieldValues['iFolioPedido']
       else
          anexo_pocompras.Params.ParamByName('Folio').Value    := 0;
       anexo_pocompras.Open ;

       If anexo_pocompras.RecordCount > 0 then
       Begin
           tdCantidad.Value := anexo_pocompras.FieldValues['dCantidad'] ;

           Connection.qryBusca.Active := False ;
           Connection.qryBusca.SQL.Clear ;
           Connection.qryBusca.SQL.Add('Select sContrato From comentariosxanexo Where sContrato = :Contrato And sNumeroActividad = :Actividad') ;
           connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
           connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
           connection.qryBusca.Params.ParamByName('actividad').DataType := ftString ;
           connection.qryBusca.Params.ParamByName('actividad').Value := '' ;
           Connection.qryBusca.Open ;
           If Connection.qryBusca.RecordCount > 0 Then
               imgNotas.Visible := True ;
       End
       Else
       Begin
           tdCantidad.Value := 0 ;
       End;
       dbPartidas.KeyValue := null;
       txtMaterial.Text    := '';
       insumos.Active := False;
       insumos.SQL.Clear;
       insumos.SQL.Add('select a.iFolioRequisicion as Requisicion, a.dCantidad as dCantidad, i.*, SubStr(i.mDescripcion, 1, 255) as sDescripcion from anexo_prequisicion a  '+
                       'inner join insumos i on (i.sContrato = a.sContrato and a.sIdInsumo = i.sIdInsumo) '+
                       'where a.sContrato =:Contrato and a.iFolioPedido =:Folio and a.sNumeroActividad =:Partida ');
       insumos.ParamByName('Contrato').AsString := global_contrato;
       if Not anexo_ocompras.FieldByName('iFolioPedido').IsNull then
          insumos.ParamByName('Folio').AsInteger    := anexo_ocompras.FieldValues['iFolioPedido']
       else
          insumos.ParamByName('Folio').AsInteger    := 0;
       insumos.ParamByName('Partida').AsString  := '';
       insumos.Open;
   End
   Else
   Begin
       tiFolio.text := '' ;
       tdIdFecha.Date := Date ;

       tsIdProveedor.KeyValue := '' ;
       tmComentarios.Text := '' ;
       tdCantidad.Value := 0 ;
       tsNumeroOrden.Text := '' ;
       
       anexo_pocompras.Active := False ;
       anexo_pocompras.SQL.Clear;
       if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
          anexo_pocompras.SQL.Add(CadenaOrden)
       else
          anexo_pocompras.SQL.Add(CadenaOrden2);
       anexo_pocompras.Params.ParamByName('Contrato').DataType := ftString ;
       anexo_pocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
       if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
       begin
            anexo_pocompras.Params.ParamByName('Convenio').DataType := ftString ;
            anexo_pocompras.Params.ParamByName('Convenio').Value    := global_convenio;
       end;
       anexo_pocompras.Params.ParamByName('Folio').DataType    := ftInteger ;
       anexo_pocompras.Params.ParamByName('Folio').Value       := 0 ;
       anexo_pocompras.Open ;
   End
end;

procedure TfrmPedidos.tsIsometricoReferenciaKeyPress(
  Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tmComentarios.SetFocus
end;


procedure TfrmPedidos.tsMailEnter(Sender: TObject);
begin
     tsMail.Color := global_color_entrada
end;

procedure TfrmPedidos.tsMailExit(Sender: TObject);
begin
     tsMail.Color := global_color_salida
end;

procedure TfrmPedidos.tsMailKeyPress(Sender: TObject; var Key: Char);
begin
    if key = #13 then
       tmComentarios.SetFocus
end;

procedure TfrmPedidos.tsMonedaEnter(Sender: TObject);
begin
      tsMoneda.Color := global_color_entrada
end;

procedure TfrmPedidos.tsMonedaExit(Sender: TObject);
begin
    tsMoneda.Color := global_color_salida;
    tipocambio;
end;

procedure TfrmPedidos.GridPartidasCellClick(Column: TColumn);
begin
   If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;

   If anexo_pocompras.RecordCount > 0 then
   begin
      tdCantidad.Value    := anexo_pocompras.FieldValues['dCantidad'] ;
      tDescuentoMat.Value := anexo_pocompras.FieldValues['dDescuento'] ;
      tdCostoNuevo.Value  := anexo_pocompras.FieldValues['dCosto'] ;
   end;

end;

procedure TfrmPedidos.frxReport50GetValue(const VarName: String;
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


procedure TfrmPedidos.GridPartidasTitleBtnClick(Sender: TObject;
  ACol: Integer; Field: TField);
Var
  sCampo : String ;
begin
  sCampo := Field.FieldName ;
  anexo_pocompras.Active := False ;
  anexo_pocompras.SQL.Clear;
  if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
     anexo_pocompras.SQL.Add(CadenaOrden)
  else
     anexo_pocompras.SQL.Add(CadenaOrden2);
  anexo_pocompras.Params.ParamByName('Contrato').DataType := ftString ;
  anexo_pocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
  if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
  begin
       anexo_pocompras.Params.ParamByName('Convenio').DataType := ftString ;
       anexo_pocompras.Params.ParamByName('Convenio').Value    := global_convenio;
  end;
  anexo_pocompras.Params.ParamByName('Folio').DataType    := ftInteger ;
  anexo_pocompras.Params.ParamByName('Folio').Value       := anexo_ocompras.FieldValues['iFolioPedido'] ;
  anexo_pocompras.Open ;
end;

procedure TfrmPedidos.GridPartidasTitleClick(Column: TColumn);
begin
 UtGrid3.DbGridTitleClick(Column);
end;

procedure TfrmPedidos.frmBarra2btnAddClick(Sender: TObject);
Var
  dFechaFinal : tDate ;
begin
  try
    OpcButton1 := 'New' ;
    frmBarra2.btnAddClick(Sender);
    frmBarra1.btnCancel.Click ;
    pgControl.ActivePageIndex := 0 ;
    tdIdFecha.Enabled := True ;
    tsIdProveedor.ReadOnly := False ;
    tsNumeroOrden.Enabled  := True ;
    tmComentarios.ReadOnly := False ;
    tiFolio.text := '' ;
    tdIdFecha.Date := Date ;
    dbRequisicion.Text  := '';

    tmComentarios.Text := '';
    tsEmbarca.Text     := '';
    tsCondiciones.Text := '';
    tsEntrega.Text     := '';
    tsPrecios.Text     := '';
    tsVigencia.Text    := '';
    tsVendedor.Text    := '';
    tsMail.Text        := '';
    tdCantidad.Value := 0 ;
    tDescuentoGral.Value := 0;
    tiFolio.SetFocus ;
    BotonPermiso.permisosBotones(frmBarra1);
    Grid_Entradas.Enabled := False;
  Except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al añadir registro', 0);
  end;
  end;
end;

procedure TfrmPedidos.frmBarra2btnEditClick(Sender: TObject);
begin
if not grid_entradas.DataSource.DataSet.IsEmpty then
begin              
try
   validaciones;
   if valida then
   begin
          frmBarra1.btnCancel.Click ;
          exit;
   end;
   If anexo_ocompras.RecordCount > 0 then
   Begin
       OpcButton1 := 'Edit' ;
       activapop(frmPedidos,popupprincipal);
       cadenaReq  := dbRequisicion.Text;
       frmBarra2.btnEditClick(Sender);
       pgControl.ActivePageIndex := 0 ;
       tdIdFecha.Enabled := True ;
       tsIdProveedor.ReadOnly := False ;
       tsNumeroOrden.Enabled := True ;
       tmComentarios.ReadOnly := False ;
      tiFolio.SetFocus
   End;
except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al editar registro', 0);
  end;
end;
   BotonPermiso.permisosBotones(frmBarra1);
   BotonPermiso.permisosBotones(frmBarra2);

end;
end;

procedure TfrmPedidos.frmBarra2btnPostClick(Sender: TObject);
Var
    Posicion, x, i : Integer ;
    cadena, cad, texto : string;
    nombres, cadenas: TStringList;
begin
  //empieza validacion
  nombres:=TStringList.Create;cadenas:=TStringList.Create;
  nombres.Add('Orden de Compra');
  nombres.Add('Proveedor');nombres.Add('Frente de Trabajo');
  nombres.Add('Forma de Pago');
  nombres.Add('Moneda');

  cadenas.Add(tiFolio.Text);
  cadenas.Add(tsIdProveedor.Text);cadenas.Add(tsNumeroOrden.Text);
  cadenas.Add(tsFormaPago.Text);
  cadenas.Add(tsMoneda.Text);

  if not validaTexto(nombres, cadenas, '', '') then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
  //continuainserccion de datos
  try
     if (CadenaReq <> '') and (dbRequisicion.Text = '') then
     begin
          validaPedido;
          if lValidaReq then
          begin
              messageDLG('Existen materiales para la Orden de Compra '+anexo_ocompras.FieldValues['sOrdenCompra']+', No se pueden Cambiar la Orden de Compra por Requisicion. ', mtInformation, [mbOk],0);
              exit;
          end;
     end;

     if OpcButton1 = 'Edit' then
        if (CadenaReq = '') and (dbRequisicion.Text <> '') then
        begin
            validaPedido;
            if lValidaReq then
            begin
                messageDLG('Existen materiales para la Orden de Compra '+anexo_ocompras.FieldValues['sOrdenCompra']+', No se pueden Cambiar la Orden de Compra sin Requisicion. ', mtInformation, [mbOk],0);
                exit;
            end;
        end;

  cad := dbRequisicion.Text;

  if chkTipoCambio.Checked then
  begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('select * from tiposdecambio where sIdMonedaConv =:Moneda and dIdFecha =:Fecha order by dIdFecha DESC ');
      connection.zCommand.ParamByName('Moneda').AsString := moneda.FieldValues['sIdMoneda'];
      connection.zCommand.ParamByName('Fecha').AsDate    := tdIdFecha.Date;
      connection.zCommand.Open;

      If connection.zCommand.RecordCount = 0 then
      begin
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('select * from tiposdecambio where sIdMoneda =:Moneda and dIdFecha < :Fecha order by dIdFecha DESC ');
          connection.zCommand.ParamByName('Moneda').AsString := moneda.FieldValues['sIdMoneda'];
          connection.zCommand.ParamByName('Fecha').AsDate    := tdIdFecha.Date;
          connection.zCommand.Open;

          if connection.zCommand.RecordCount > 0 then
          begin
             If messageDLG('No se encontró un Tipo de Cambio para '+DateToStr(tdIdFecha.Date)+ ' el Sistema Cotizará a : $'+FloatToStr(tsTipoCambio.Value)+ #13+' -> ¿Desea Ccontinuar?', mtInformation, [mbYes, mbNo], 0) = mrNo then
                exit;
          end
          else
          begin
             messageDLG('No se encontró un Tipo de Cambio para la Fecha '+DateToStr(tdIdFecha.Date)+' en el Sistema!', mtInformation, [mbOk], 0);
             tsTipoCambio.Value := 0;
          end;
      end;
  end;

  If OpcButton1 = 'New' then
  Begin
      With connection do
      Begin
          try
              zCommand.Active := False ;
              zCommand.SQL.Clear ;
              zcommand.SQL.Add('INSERT INTO anexo_pedidos ( sContrato, sOrdenCompra, sIdProveedor, sNumeroOrden, dIdFecha, '+
                               'sReferencia,  mComentarios, sFolioRequisicion, sFormaPago, sMoneda, dCambio, dIVA, sElaboro, sLugarEntrega, sCondiciones, sEntrega, '+
                               'sPrecios, sVigencia, sVendedor, sMail, dDescuento ) ' +
                               'VALUES (:Contrato, :oCompra, :Proveedor, :Orden, :Fecha, :Orden, :Comentarios, :Requisicion, :sPago, :Moneda, :Cambio, :IVA, :Elaboro, '+
                               ':Lugar, :Condiciones, :Entrega, :Precios, :Vigencia, :Vendedor, :Mail, :Descuento )');
              zcommand.Params.ParamByName('Contrato').DataType     := ftString ;
              zcommand.Params.ParamByName('Contrato').value        := Global_Contrato ;
              zcommand.Params.ParamByName('oCompra').DataType      := ftString ;
              zcommand.Params.ParamByName('oCompra').value         := tiFolio.Text ;
              zcommand.Params.ParamByName('Proveedor').DataType    := ftString ;
              zcommand.Params.ParamByName('Proveedor').value       := tsIdProveedor.KeyValue ;
              zcommand.Params.ParamByName('Orden').DataType        := ftString ;
              zcommand.Params.ParamByName('Orden').value           := tsNumeroOrden.Text ;
              zcommand.Params.ParamByName('Fecha').DataType        := ftDate ;
              zcommand.Params.ParamByName('Fecha').value           := tdIdFecha.Date ;
              zcommand.Params.ParamByName('Comentarios').DataType  := ftMemo ;
              zcommand.Params.ParamByName('Comentarios').value     := tmCOmentarios.Text ;
              zcommand.Params.ParamByName('Requisicion').AsString  := dbRequisicion.Text ;
              zcommand.Params.ParamByName('sPago').DataType        := ftString ;
              zcommand.Params.ParamByName('sPago').value           := tsFormaPago.KeyValue ;
              zcommand.Params.ParamByName('Entrega').DataType      := ftString ;
              zcommand.Params.ParamByName('Entrega').value         := tsEntrega.Text ;
              zcommand.Params.ParamByName('Moneda').DataType       := ftString ;
              zcommand.Params.ParamByName('Moneda').value          := tsMoneda.KeyValue ;
              zcommand.Params.ParamByName('Cambio').DataType       := ftFloat ;
              if chkTipoCambio.Checked then
                 zcommand.Params.ParamByName('Cambio').value       := tsTipoCambio.Value
              else
                 zcommand.Params.ParamByName('Cambio').value       := 0;
              zcommand.Params.ParamByName('IVA').DataType          := ftFloat ;
              zcommand.Params.ParamByName('IVA').value             := moneda.FieldValues['dIVA'] ;
              zcommand.Params.ParamByName('Elaboro').DataType      := ftString ;
              zcommand.Params.ParamByName('Elaboro').value         := global_usuario;
              zcommand.Params.ParamByName('Lugar').DataType        := ftString ;
              zcommand.Params.ParamByName('Lugar').value           := tsEmbarca.Text;
              zcommand.Params.ParamByName('Condiciones').DataType  := ftString ;
              zcommand.Params.ParamByName('Condiciones').value     := tsCondiciones.Text;
              zcommand.Params.ParamByName('Precios').DataType      := ftString ;
              zcommand.Params.ParamByName('Precios').value         := tsPrecios.Text;
              zcommand.Params.ParamByName('Vigencia').DataType     := ftString ;
              zcommand.Params.ParamByName('Vigencia').value        := tsVigencia.Text;
              zcommand.Params.ParamByName('Vendedor').DataType     := ftString ;
              zcommand.Params.ParamByName('Vendedor').value        := tsVendedor.Text;
              zcommand.Params.ParamByName('Mail').DataType         := ftString ;
              zcommand.Params.ParamByName('Mail').value            := tsMail.Text;
              zcommand.Params.ParamByName('Descuento').DataType    := ftFloat ;
              zcommand.Params.ParamByName('Descuento').value       := tDescuentoGral.Value;
              zcommand.ExecSQL ;

              SavePlace := anexo_ocompras.GetBookmark ;
              anexo_ocompras.Refresh ;

              Try
                  anexo_ocompras.GotoBookmark(SavePlace);
              Except
              Else
                  anexo_ocompras.FreeBookmark(SavePlace);
              End ;

              // Actualizo Kardex del Sistema ....
              Connection.zCommand.Active := False;
              Connection.zCommand.SQL.Clear ;
              connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                          'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
              connection.zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
              connection.zcommand.Params.ParamByName('Contrato').Value       := Global_Contrato ;
              connection.zcommand.Params.ParamByName('Usuario').DataType     := ftString ;
              connection.zcommand.Params.ParamByName('Usuario').Value        := Global_Usuario ;
              connection.zcommand.Params.ParamByName('Fecha').DataType       := ftDate ;
              connection.zcommand.Params.ParamByName('Fecha').Value          := Date ;
              connection.zcommand.Params.ParamByName('Hora').DataType        := ftString ;
              connection.zcommand.Params.ParamByName('Hora').value           := FormatDateTime('hh:mm:ss', Now) ;
              connection.zcommand.Params.ParamByName('Descripcion').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Descripcion').Value    := 'Registro de Pedido ' + tiFolio.Text +  '] Registrado el [' + DateToStr(tdIdFecha.Date) +  '] Usuario [ ' + global_usuario + ']' ;
              connection.zcommand.Params.ParamByName('Origen').DataType      := ftString ;
              connection.zcommand.Params.ParamByName('Origen').Value         := 'Requisiciones y Compras' ;
              connection.zcommand.ExecSQL ;

              desactivapop(popupprincipal);
              tdIdFecha.Enabled := False ;
              tsIdProveedor.ReadOnly := True ;
              tsNumeroOrden.Enabled := False ;
              tmComentarios.ReadOnly := True ;
              frmBarra2.btnCancelClick(Sender);
          Except
            on e : exception do begin
            UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Ordenes de Compra', 'Al salvar registro', 0);
            end;
          End
      End
  End
  Else
      If OpcButton1 = 'Edit' then
      Begin
         With connection do
         Begin
            try
                zCommand.Active := False ;
                zCommand.SQL.Clear ;
                zcommand.SQL.Add('UPDATE anexo_pedidos SET sOrdenCompra = :oCompra, sIdProveedor = :Proveedor, dIdFecha = :Fecha, ' +
                                 'sFormaPago = :sPago, sEntrega = :Entrega, sMoneda = :Moneda, dCambio =:Cambio, dIVA =:IVA, sLugarEntrega =:Lugar, sCondiciones =:Condiciones, sPrecios =:Precio, sVigencia =:Vigencia, ' +
                                 'sNumeroOrden = :Orden, sReferencia = :Orden, mComentarios = :Comentarios, sFolioRequisicion = :Requisicion, sVendedor =:Vendedor, sMail =:Mail, dDescuento =:Descuento ' +
                                 'WHERE sContrato = :Contrato And iFolioPedido = :Folio');
                zcommand.Params.ParamByName('Contrato').DataType    := ftString ;
                zcommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                zcommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                zcommand.Params.ParamByName('Folio').value          := anexo_ocompras.FieldValues['iFoliopedido'] ;
                zcommand.Params.ParamByName('oCompra').DataType     := ftString ;
                zcommand.Params.ParamByName('oCompra').value        := tiFolio.Text ;
                zcommand.Params.ParamByName('Proveedor').DataType   := ftString ;
                zcommand.Params.ParamByName('Proveedor').value      := tsIdProveedor.KeyValue ;
                zcommand.Params.ParamByName('Orden').DataType       := ftString ;
                zcommand.Params.ParamByName('Orden').value          := tsNumeroOrden.Text ;
                zcommand.Params.ParamByName('Fecha').DataType       := ftDate ;
                zcommand.Params.ParamByName('Fecha').value          := tdIdFecha.Date ;
                zcommand.Params.ParamByName('sPago').DataType       := ftString ;
                zcommand.Params.ParamByName('sPago').value          := tsFormaPago.KeyValue ;
                zcommand.Params.ParamByName('Entrega').DataType     := ftString ;
                zcommand.Params.ParamByName('Entrega').value        := tsEntrega.Text ;
                zcommand.Params.ParamByName('Moneda').DataType      := ftString ;
                zcommand.Params.ParamByName('Moneda').value         := tsMoneda.KeyValue ;
                zcommand.Params.ParamByName('Cambio').DataType      := ftFloat ;
                if chkTipoCambio.Checked then
                   zcommand.Params.ParamByName('Cambio').value      := tsTipoCambio.Value
                else
                   zcommand.Params.ParamByName('Cambio').value      := 0;
                zcommand.Params.ParamByName('IVA').DataType         := ftFloat ;
                zcommand.Params.ParamByName('IVA').value            := moneda.FieldValues['dIVA'] ;
                zcommand.Params.ParamByName('Comentarios').DataType := ftMemo ;
                zcommand.Params.ParamByName('Comentarios').value    := tmCOmentarios.Text ;
                zcommand.Params.ParamByName('Requisicion').AsString := dbRequisicion.Text ;
                zcommand.Params.ParamByName('Lugar').DataType       := ftString ;
                zcommand.Params.ParamByName('Lugar').value          := tsEmbarca.Text;
                zcommand.Params.ParamByName('Condiciones').DataType := ftString ;
                zcommand.Params.ParamByName('Condiciones').value    := tsCondiciones.Text;
                zcommand.Params.ParamByName('Precio').DataType      := ftString ;
                zcommand.Params.ParamByName('Precio').value         := tsPrecios.Text;
                zcommand.Params.ParamByName('Vigencia').DataType    := ftString ;
                zcommand.Params.ParamByName('Vigencia').value       := tsVigencia.Text;
                zcommand.Params.ParamByName('Vendedor').DataType    := ftString ;
                zcommand.Params.ParamByName('Vendedor').value       := tsVendedor.Text;
                zcommand.Params.ParamByName('Mail').DataType        := ftString ;
                zcommand.Params.ParamByName('Mail').value           := tsMail.Text;
                zcommand.Params.ParamByName('Descuento').DataType   := ftFloat ;
                zcommand.Params.ParamByName('Descuento').value      := tDescuentoGral.Value;
                zcommand.ExecSQL ;

                SavePlace := anexo_ocompras.GetBookmark ;
                anexo_ocompras.Refresh ;

                Try
                   anexo_ocompras.GotoBookmark(SavePlace);
                Except
                Else
                  anexo_ocompras.FreeBookmark(SavePlace);
                End;

                //Ahora actualizamos el tipo de cambio si aplica.
                Connection.QryBusca.Active := False ;
                Connection.QryBusca.SQL.Clear ;
                Connection.QryBusca.SQL.Add('select p.iFolioPedido, p.sIdInsumo, p.iItem, i.dNuevoPrecio from anexo_ppedido p '+
                                            'inner join insumos i on (i.sContrato = p.sContrato and i.sIdInsumo = p.sIdInsumo) '+
                                            'WHERE p.sContrato =:Contrato And p.iFolioPedido =:Folio ');
                Connection.QryBusca.Params.ParamByName('Contrato').DataType  := ftString ;
                Connection.QryBusca.Params.ParamByName('Contrato').value     := Global_Contrato ;
                Connection.QryBusca.Params.ParamByName('Folio').DataType     := ftInteger ;
                Connection.QryBusca.Params.ParamByName('Folio').value        := anexo_ocompras.FieldValues['iFolioPedido'] ;
                Connection.QryBusca.Open ;

                if Connection.QryBusca.RecordCount > 0 then
                begin
                    if (anexo_ocompras.FieldValues['sMoneda'] = 'DLL') or (anexo_ocompras.FieldValues['sMoneda'] = 'dll') then
                    begin
                        if anexo_ocompras.FieldValues['dCambio'] > 0 then
                        begin
                            while not connection.QryBusca.Eof do
                            begin
                                Connection.zCommand.Active := False ;
                                Connection.zCommand.SQL.Clear ;
                                Connection.zCommand.SQL.Add('UPDATE anexo_ppedido SET dCosto =:Costo ' +
                                                            'WHERE sContrato = :Contrato And iFolioPedido = :Folio and sIdInsumo =:Insumo And iItem = :Item');
                                Connection.zCommand.Params.ParamByName('Contrato').DataType  := ftString ;
                                Connection.zCommand.Params.ParamByName('Contrato').value     := Global_Contrato ;
                                Connection.zCommand.Params.ParamByName('Folio').DataType     := ftInteger ;
                                Connection.zCommand.Params.ParamByName('Folio').value        := Connection.QryBusca.FieldValues['iFolioPedido'] ;
                                Connection.zCommand.Params.ParamByName('Insumo').DataType    := ftString ;
                                Connection.zCommand.Params.ParamByName('Insumo').value       := connection.QryBusca.FieldValues['sIdInsumo'];
                                Connection.zCommand.Params.ParamByName('Item').DataType      := ftInteger ;
                                Connection.zCommand.Params.ParamByName('Item').value         := Connection.QryBusca.FieldValues['iItem'] ;
                                Connection.zCommand.Params.ParamByName('Costo').DataType     := ftFloat ;
                                Connection.zCommand.Params.ParamByName('Costo').value        := (Connection.QryBusca.FieldValues['dNuevoPrecio'] * anexo_ocompras.FieldValues['dCambio']);
                                Connection.zCommand.ExecSQL ;
                                connection.QryBusca.Next;
                            end;
                        end ;
                    end;
                end;

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
                connection.zcommand.Params.ParamByName('Descripcion').Value := 'Modificación de Pedido ' + tiFolio.Text +  '] Registrado el [' + DateToStr(tdIdFecha.Date) +  '] Usuario [ ' + global_usuario + ']' ;
                connection.zcommand.Params.ParamByName('Origen').DataType := ftString ;
                connection.zcommand.Params.ParamByName('Origen').Value := 'Requisiciones y Compras' ;
                connection.zcommand.ExecSQL ;

                tdIdFecha.Enabled := False ;
                tsNumeroOrden.Enabled := False ;
                tsIdProveedor.ReadOnly := True ;
                tmComentarios.ReadOnly := True ;
                frmBarra2.btnCancelClick(Sender);
            except
              on e : exception do begin
              UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Ordenes de Compra', 'Al actualizar registro', 0);
              end;
            End
         End
      End ;   

      {connection.QryBusca.Active;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('Delete from anexo_prequisicion Where sContrato =:Contrato and iFolioPedido =:Pedido ');
      connection.QryBusca.ParamByName('Contrato').AsString  := global_contrato ;
      connection.QryBusca.ParamByName('Pedido').AsInteger   := iFolio ;
      connection.QryBusca.ExecSQL;

      x := 1;
      i := length(cad);
      while x = 1 do
      begin
           x     := pos(',',cad);
           texto := cad;
           if x > 0 then
           begin
                texto := copy(cad,1, x - 1);
                cad   := copy(cad,x + 1, i);
                x     := 1;
           end;

           if texto <> '' then
           begin
                connection.zCommand.Active;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('Select * from anexo_prequisicion Where sContrato =:Contrato and iFolioRequisicion =:Folio and iFolioPedido = 0');
                connection.zCommand.ParamByName('Contrato').AsString := global_contrato ;
                connection.zCommand.ParamByName('Folio').AsInteger   := StrToInt(texto) ;
                connection.zCommand.Open;

                if connection.zCommand.RecordCount > 0 then
                begin
                    while not connection.zCommand.Eof do
                    begin
                        Connection.QryBusca.Active := False ;
                        Connection.QryBusca.SQL.Clear ;
                        Connection.QryBusca.SQL.Add('INSERT INTO anexo_prequisicion ( sContrato, iFolioRequisicion, sIdInsumo, iFolioPedido, iItem, mDescripcion, sMedida, dFechaRequerimiento, dCantidad, dCosto, sNumeroActividad, sNumeroOrden, sWbs) '  +
                                            'VALUES (:Contrato, :Folio, :Insumo, :Pedido, :Item, :Descripcion, :Medida, :Requerido, :Cantidad, :Costo, :Actividad, :Orden, :Wbs) ' );
                        Connection.QryBusca.Params.ParamByName('Contrato').DataType     := ftString ;
                        Connection.QryBusca.Params.ParamByName('Contrato').value        := Global_Contrato ;
                        Connection.QryBusca.Params.ParamByName('Folio').DataType        := ftInteger ;
                        Connection.QryBusca.Params.ParamByName('Folio').value           := connection.zCommand.FieldValues['iFolioRequisicion'] ;
                        Connection.QryBusca.Params.ParamByName('Pedido').DataType       := ftInteger ;
                        Connection.QryBusca.Params.ParamByName('Pedido').value          := iFolio  ;
                        Connection.QryBusca.Params.ParamByName('Item').DataType         := ftInteger ;
                        Connection.QryBusca.Params.ParamByName('Item').value            := connection.zCommand.fieldValues['iItem'] ;
                        Connection.QryBusca.Params.ParamByName('Insumo').DataType       := ftString ;
                        Connection.QryBusca.Params.ParamByName('Insumo').value          := connection.zCommand.FieldValues['sIdInsumo'] ;
                        Connection.QryBusca.Params.ParamByName('Descripcion').DataType  := ftMemo ;
                        Connection.QryBusca.Params.ParamByName('Descripcion').value     := connection.zCommand.fieldValues['mDescripcion'] ;
                        Connection.QryBusca.Params.ParamByName('Medida').DataType       := ftString ;
                        Connection.QryBusca.Params.ParamByName('Medida').value          := connection.zCommand.fieldValues['sMedida'] ;
                        Connection.QryBusca.Params.ParamByName('Requerido').DataType    := ftDate ;
                        Connection.QryBusca.Params.ParamByName('Requerido').value       := connection.zCommand.fieldValues['dFechaRequerimiento'] ;
                        Connection.QryBusca.Params.ParamByName('Cantidad').DataType     := ftFloat ;
                        Connection.QryBusca.Params.ParamByName('Cantidad').value        := connection.zCommand.fieldValues['dCantidad'] ;
                        Connection.QryBusca.Params.ParamByName('Costo').DataType        := ftFloat ;
                        Connection.QryBusca.Params.ParamByName('Costo').value           := connection.zCommand.fieldValues['dCosto'] ;
                        Connection.QryBusca.Params.ParamByName('Actividad').DataType    := ftString ;
                        Connection.QryBusca.Params.ParamByName('Actividad').value       := connection.zCommand.fieldValues['sNumeroActividad'] ;
                        Connection.QryBusca.Params.ParamByName('Wbs').DataType          := ftString ;
                        Connection.QryBusca.Params.ParamByName('Wbs').value             := connection.zCommand.fieldValues['sWbs'] ;
                        Connection.QryBusca.Params.ParamByName('Orden').DataType        := ftString ;
                        Connection.QryBusca.Params.ParamByName('Orden').value           := connection.zCommand.FieldValues['sNumeroOrden'] ;
                        Connection.QryBusca.ExecSQL  ;
                        connection.zCommand.Next;
                    end;
                end
           end;
      end;      }

      SavePlace := anexo_ocompras.GetBookmark ;
      anexo_ocompras.Refresh ;

      Try
         anexo_ocompras.GotoBookmark(SavePlace);
      Except
      Else
        anexo_ocompras.FreeBookmark(SavePlace);
      End ;

      OpcButton1 := '' ;

      BotonPermiso.permisosBotones(frmBarra1);
      BotonPermiso.permisosBotones(frmBarra2);
      except
      on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al salvar registro', 0);
      end;
      end;
      Grid_Entradas.Enabled := True;
end;

procedure TfrmPedidos.frmBarra2btnDeleteClick(Sender: TObject);
begin
 try
       validaciones;
       if valida then
       begin
           frmBarra1.btnCancel.Click ;
           exit;
       end;
       If anexo_ocompras.RecordCount > 0 Then
          If MessageDlg('Desea eliminar el folio seleccionado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          Begin
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear ;
              connection.zCommand.SQL.Add('select iFolioPedido from anexo_ppedido where sContrato = :Contrato And iFolioPedido = :Folio') ;
              connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Contrato').value    := Global_Contrato ;
              connection.zcommand.Params.ParamByName('Folio').DataType    := ftInteger ;
              connection.zcommand.Params.ParamByName('Folio').value       := anexo_ocompras.FieldValues['iFolioPedido'] ;
              connection.zcommand.Open ;

              if connection.zCommand.RecordCount > 0 then
              begin
                  MessageDLG(' No se puede eliminar la Orden de Compra '+anexo_ocompras.FieldValues['sOrdenCompra']+' , Elimine materiales.', mtInformation, [mbOk],0);
                  exit;
              end;

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
              connection.zcommand.Params.ParamByName('Descripcion').Value := 'Eliminación de Pedido ' + tiFolio.Text +  '] Registrado el [' + DateToStr(tdIdFecha.Date) +  '] Usuario [ ' + global_usuario + ']' ;
              connection.zcommand.Params.ParamByName('Origen').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Origen').Value := 'Requisiciones y Compras' ;
              connection.zcommand.ExecSQL ;

              With connection do
              Begin
                 try
                     //Eliminamos las referecnias en Requisiciones de la Ordne de Copra hacia las requisiciones asignadas..
                     zCommand.Active := False;
                     zCommand.SQL.Clear ;
                     zCommand.SQL.Add('Delete from anexo_prequisicion where sContrato = :Contrato And iFolioPedido = :Folio') ;
                     zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                     zcommand.Params.ParamByName('Contrato').value := Global_Contrato ;
                     zcommand.Params.ParamByName('Folio').DataType := ftInteger ;
                     zcommand.Params.ParamByName('Folio').value := anexo_ocompras.FieldValues['iFolioPedido'] ;
                     zcommand.ExecSQL ;
                     //Eliminamos la orden de compraa..
                     zCommand.Active := False;
                     zCommand.SQL.Clear ;
                     zCommand.SQL.Add('Delete from anexo_pedidos where sContrato = :Contrato And iFolioPedido = :Folio') ;
                     zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                     zcommand.Params.ParamByName('Contrato').value := Global_Contrato ;
                     zcommand.Params.ParamByName('Folio').DataType := ftInteger ;
                     zcommand.Params.ParamByName('Folio').value := anexo_ocompras.FieldValues['iFolioPedido'] ;
                     zcommand.ExecSQL ;

                     SavePlace := anexo_ocompras.GetBookmark ;
                     anexo_ocompras.Refresh ;

                     Try
                        anexo_ocompras.GotoBookmark(SavePlace);
                     Except
                     Else
                        anexo_ocompras.FreeBookmark(SavePlace);
                     End;
                 Except
                  on e : exception do begin
                  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Ordenes de Compra', 'Al eliminar registro', 0);
                  end;
                     //MessageDlg('Ocurrio un error al eliminar el registro.', mtInformation, [mbOk], 0);
                 End
              End
          End
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al eliminar registro', 0);
  end;
 end;
end;


procedure TfrmPedidos.frmBarra2btnRefreshClick(Sender: TObject);
begin
 try
    anexo_ocompras.Refresh ;
    Proveedores.Refresh;
    FormaPago.Refresh;

    anexo_pocompras.Active := False ;
    anexo_pocompras.SQL.Clear;
    anexo_pocompras.SQL.Add(CadenaOrden);
    anexo_pocompras.Params.ParamByName('Contrato').DataType := ftString ;
    anexo_pocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
    anexo_pocompras.Params.ParamByName('Convenio').DataType := ftString ;
    anexo_pocompras.Params.ParamByName('Convenio').Value    := global_convenio ;
    anexo_pocompras.Params.ParamByName('Folio').DataType    := ftInteger ;
    If anexo_ocompras.RecordCount > 0 Then
        anexo_pocompras.Params.ParamByName('Folio').Value := anexo_ocompras.FieldValues['iFolioPedido']
    Else
        anexo_pocompras.Params.ParamByName('Folio').Value := 0 ;
    anexo_pocompras.Open ;

    requisiciones.Params.ParamByName('Contrato').DataType := ftString ;
    requisiciones.Params.ParamByName('Contrato').Value    := global_contrato ;
    requisiciones.Open ;

    if requisiciones.RecordCount > 0 then
    begin
         dbRequisicion.Clear;
         while not requisiciones.Eof do
         begin
              dbRequisicion.Items.Add(IntToStr(requisiciones.FieldValues['iFolioRequisicion'])+ '.- '+requisiciones.FieldValues['sNumeroSolicitud']);
              requisiciones.Next;
         end;
    end;
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al actualizar grid', 0);
  end;
 end;
end;

procedure TfrmPedidos.frmBarra2btnCancelClick(Sender: TObject);
begin
TRY
  tdIdFecha.Enabled := False ;
  tsIdProveedor.ReadOnly := True ;
  tsNumeroOrden.Enabled := False ;
  tmComentarios.ReadOnly := True ;
  OpcButton1 := '';
  frmBarra2.btnCancelClick(Sender);
  grid_entradas.enabled:=true;
  Grid_Entradas.SetFocus ;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
EXCEPT
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al cancelar registro', 0);
  end;
END;
end;

procedure TfrmPedidos.frmBarra2btnExitClick(Sender: TObject);
begin
  frmBarra2.btnExitClick(Sender);
  close
end;

procedure TfrmPedidos.tdIdFechaEnter(Sender: TObject);
begin
    tdIdFecha.Color := global_color_entrada
end;

procedure TfrmPedidos.tdIdFechaExit(Sender: TObject);
begin
    tdIdFecha.Color := global_color_salida;
    tipocambio;
end;

procedure TfrmPedidos.tdIdFechaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if chkRequisicion.Checked then
    begin
        If Key = #13 Then
          dbRequisicion.SetFocus
    end
    else
    begin
        If Key = #13 Then
          tsIdProveedor.SetFocus
    end;
end;

procedure TfrmPedidos.tiFolioEnter(Sender: TObject);
begin
    tiFolio.Color := global_color_entrada;
end;

procedure TfrmPedidos.tiFolioExit(Sender: TObject);
begin
    tiFolio.Color := clGradientActiveCaption;
    tipocambio;
end;

procedure TfrmPedidos.tiFolioKeyPress(Sender: TObject; var Key: Char);
begin
     if key = #13 then
        tdIdFecha.SetFocus;
end;

procedure TfrmPedidos.tsVendedorEnter(Sender: TObject);
begin
    tsVendedor.Color := global_color_entrada
end;

procedure TfrmPedidos.tsVendedorExit(Sender: TObject);
begin
     tsVendedor.Color := global_color_salida
end;

procedure TfrmPedidos.tsVendedorKeyPress(Sender: TObject; var Key: Char);
begin
    if key = #13 then
       tsMail.SetFocus
end;

procedure TfrmPedidos.tsVigenciaEnter(Sender: TObject);
begin
    tsVigencia.Color := global_color_entrada;
end;

procedure TfrmPedidos.tsVigenciaExit(Sender: TObject);
begin
    tsVigencia.Color := global_color_salida;
end;

procedure TfrmPedidos.tsVigenciaKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
       tsVendedor.SetFocus
end;

procedure TfrmPedidos.txtMaterialEnter(Sender: TObject);
begin
     txtMaterial.color := Global_color_entrada;
end;

procedure TfrmPedidos.txtMaterialExit(Sender: TObject);
begin
    txtMaterial.color := Global_color_salida;
end;

procedure TfrmPedidos.txtMaterialKeyPress(Sender: TObject; var Key: Char);
begin
      if Key = #13 then
         Grid_Insumos.SetFocus;

end;

procedure TfrmPedidos.txtMaterialKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     if frmBarra1.btnPost.Enabled = True then
        filtra;
end;

procedure TfrmPedidos.Grid_EntradasEnter(Sender: TObject);
begin
    frmBarra1.btnCancel.Click ;
    If frmbarra2.btnCancel.Enabled = True then
       frmBarra2.btnCancel.Click ;
end;

procedure TfrmPedidos.tmComentariosEnter(Sender: TObject);
begin
    tmComentarios.Color := global_color_entrada
end;

procedure TfrmPedidos.tmComentariosExit(Sender: TObject);
begin
    tmComentarios.Color := global_color_salida
end;
procedure TfrmPedidos.tdCantidadChange(Sender: TObject);
begin
TRxCalcEditChangef(tdCantidad,'Cantidad');
end;

procedure TfrmPedidos.tdCantidadEnter(Sender: TObject);
begin
    tdCantidad.Color := global_color_entrada
end;

procedure TfrmPedidos.tdCantidadExit(Sender: TObject);
begin
    tdCantidad.Color := global_color_salida
end;

procedure TfrmPedidos.tdCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
if not keyFiltroTRxCalcEdit(tdCantidad,key) then
   key:=#0;
//    If Key = #13 Then

end;


//procedure TfrmPedidos.anexo_pocomprasAfterScroll(DataSet: TDataSet);



procedure TfrmPedidos.GridPartidasEnter(Sender: TObject);
begin
    If frmBarra1.btnCancel.Enabled = True Then
        frmBarra1.btnCancel.Click ;

    If anexo_pocompras.RecordCount > 0 then
    begin
        tdCantidad.Value := anexo_pocompras.FieldValues['dCantidad'];
        tDescuentoMat.Value := anexo_pocompras.FieldValues['dDescuento'] ;
        tdCostoNuevo.Value  := anexo_pocompras.FieldValues['dCosto'] ;
    end
    Else
    begin
        tdCantidad.Value := 0 ;
        tDescuentoMat.Value := 0;
        tdCostoNuevo.Value := 0;
    end;
end;

procedure TfrmPedidos.GridPartidasGetCellParams(Sender: TObject; Field: TField;
  AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
     If (Sender as TrxDBGrid).DataSource.DataSet.State = dsBrowse Then
        If anexo_pocompras.RecordCount > 0 Then
        Begin

            Afont.Style := [fsBold];
            AFont.Color := clBlue;

            If (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sStatus').AsString = 'Pendiente' then
            begin
                Afont.Style := [];
                AFont.Color := clRed ;
            end;
        End;
end;

procedure TfrmPedidos.GridPartidasMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
UtGrid3.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmPedidos.GridPartidasMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid3.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmPedidos.frmBarra2btnPrinterClick(Sender: TObject);
begin
    Imprimir1.Click;
end;

procedure TfrmPedidos.tsNumeroActividadExit(Sender: TObject);
begin

  If dbPartidas.ReadOnly = False Then
  Begin
      Connection.qryBusca.Active := False ;
      Connection.qryBusca.SQL.Clear ;
      Connection.qryBusca.SQL.Add('Select sContrato From comentariosxanexo Where sContrato = :Contrato And sNumeroActividad = :Actividad') ;
      connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
      connection.qryBusca.Params.ParamByName('actividad').DataType := ftString ;
      connection.qryBusca.Params.ParamByName('actividad').Value := '' ;
      Connection.qryBusca.Open ;
      If Connection.qryBusca.RecordCount > 0 Then
          imgNotas.Visible := True ;

//      if tsNumeroActividad.Text <> '' then
//         If Not lExisteActividad ( tsNumeroActividad.Text ) then
//              tsNumeroActividad.SetFocus
//         Else
//         Begin
//              tsnMedida.Text := connection.qryBusca.FieldValues['sMedida']  ;
//              tmDescripcion.Text := sDescripcion ;
//         End
  End
end;

procedure TfrmPedidos.tsNumeroActividadEnter(Sender: TObject);
begin
    imgNotas.Visible := False ;

end;

procedure TfrmPedidos.tsNumeroActividadKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then

end;

procedure TfrmPedidos.lblPartidaClick(Sender: TObject);
begin
    if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
    begin
        if lblPartida.Caption = 'ID Material' then
           lblPartida.Caption := 'Material'
        else
           lblPartida.Caption := 'ID Material';
    end;
end;

function TfrmPedidos.lExisteActividad ( sActividad : String ) : Boolean ;
Begin
      connection.qryBusca.Active := False ;
      connection.qryBusca.SQL.Clear ;
      connection.qryBusca.SQL.Add('select mDescripcion, dCantidadAnexo, sMedida, dVentaMN from actividadesxanexo where sContrato = :Contrato ' +
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

procedure TfrmPedidos.Paste1Click(Sender: TObject);
begin
 try
  if grid_entradas.Focused=true then
    begin
      UtGrid.AddRowsFromClip;
    end;
  if grid_insumos.Focused=true then
    begin

      UtGrid2.AddRowsFromClip;
    end;
  if gridpartidas.Focused=true then
    begin
      UtGrid3.AddRowsFromClip;
    end;
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al pegar registro', 0);
  end;
 end;
end;

procedure TfrmPedidos.PgControlChange(Sender: TObject);
begin
    If anexo_ocompras.RecordCount > 0 Then
    Begin
        if Not anexo_ocompras.FieldValues['iFoliopedido'] = null then
           iFolio := anexo_ocompras.FieldValues['iFoliopedido'] ;
        if Not anexo_ocompras.FieldValues['dIdFecha'] = null then
           tdIdFecha.Date := anexo_ocompras.FieldValues['dIdFecha'] ;

        tsIdProveedor.KeyValue := anexo_ocompras.FieldValues['sIdProveedor'] ;

        tsNumeroOrden.Text := anexo_ocompras.FieldValues['sNumeroOrden'] ;
        tmComentarios.Text := anexo_ocompras.FieldValues['mComentarios'] ;
        tsEmbarca.Text     := anexo_ocompras.FieldValues['sLugarEntrega'] ;
        tsEntrega.Text     := anexo_ocompras.FieldValues['sEntrega'] ;
        tsCondiciones.Text := anexo_ocompras.FieldValues['sCondiciones'] ;
        tsPrecios.Text     := anexo_ocompras.FieldValues['sPrecios'] ;
        tsVigencia.Text    := anexo_ocompras.FieldValues['sVigencia'] ;
        tsMoneda.KeyValue  := anexo_ocompras.FieldValues['sMoneda'] ;
        tsVendedor.Text    := anexo_ocompras.FieldValues['sVendedor'] ;
        tsMail.Text        := anexo_ocompras.FieldValues['sMail'] ;

        anexo_pocompras.Active := False ;
        anexo_pocompras.SQL.Clear;
        if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
           anexo_pocompras.SQL.Add(CadenaOrden)
        else
           anexo_pocompras.SQL.Add(CadenaOrden2);
        anexo_pocompras.Params.ParamByName('Contrato').DataType := ftString ;
        anexo_pocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
        if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
        begin
             anexo_pocompras.Params.ParamByName('Convenio').DataType := ftString ;
             anexo_pocompras.Params.ParamByName('Convenio').Value    := global_convenio;
        end;
        anexo_pocompras.Params.ParamByName('Folio').DataType    := ftInteger ;
        anexo_pocompras.Params.ParamByName('Folio').Value       := anexo_ocompras.FieldValues['iFolioPedido'] ;
        anexo_pocompras.Open ;
        If anexo_pocompras.RecordCount > 0 then
           tdCantidad.Value := anexo_pocompras.FieldValues['dCantidad']
        Else
            tdCantidad.Value := 0 ;
    End
    Else
    Begin
        tiFolio.text           := '' ;
        tdIdFecha.Date         := Date ;
        tsIdProveedor.KeyValue := '' ;
        tsNumeroOrden.Text     := '' ;
        tmComentarios.Text     := '' ;
        tdCantidad.Value       := 0 ;

        anexo_pocompras.Active := False ;
        anexo_pocompras.SQL.Clear;

        if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
           anexo_pocompras.SQL.Add(CadenaOrden)
        else
           anexo_pocompras.SQL.Add(CadenaOrden2);
        anexo_pocompras.Params.ParamByName('Contrato').DataType := ftString ;
        anexo_pocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
        if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
        begin
             anexo_pocompras.Params.ParamByName('Convenio').DataType := ftString ;
             anexo_pocompras.Params.ParamByName('Convenio').Value    := global_convenio;
        end;
        anexo_pocompras.Params.ParamByName('Folio').DataType    := ftInteger ;
        anexo_pocompras.Params.ParamByName('Folio').Value       := 0 ;
        anexo_pocompras.Open ;
    End
end;

procedure TfrmPedidos.frxEntradaGetValue(const VarName: String;
  var Value: Variant);
   Var
     sCadena : String ;
     iva  : Currency ;
     dAcumuladoOrden, dAcumuladoGral, dMonto,
     iValorNumerico, tNumerico   : LongInt  ;
     Resultado        : Real     ;
     dIVA : double;

     zConsulta:TZQuery;
     sSQL:string;
begin

  If CompareText(VarName, 'FechaInicio') = 0 then
      Value := FechaInicio.Date ;
  If CompareText(VarName, 'FechaFinal') = 0 then
      Value := FechaFinal.Date ;

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
         connection.zCommand.Active := False;
         connection.zCommand.SQL.Clear;
         connection.zCommand.SQL.Add('select dIVA, sDescripcion from tiposdemoneda where sIdMoneda =:Moneda');
         connection.zCommand.ParamByName('Moneda').AsString := moneda.FieldValues['sIdMoneda'];
         connection.zCommand.Open;

         if connection.zCommand.RecordCount > 0 then
            dIVA := (connection.zCommand.FieldValues['dIVA'] / 100)
         else
            dIVA := 1;

         iVa              := (Montototal * dIVA) ;
         MontoTotal       := Montototal + iva ;
         iValorNumerico   := Trunc(Montototal) ;
         sCadena := xIntToLletres (iValorNumerico) ;
         Resultado := roundto(MontoTotal - iValorNumerico, -2) ;
         Resultado := Resultado * 100;
         iValorNumerico := Trunc(Resultado);
         sCadena := sCadena + ' '+IntToStr(iValorNumerico) + '/100 '+ connection.zCommand.FieldValues['sDescripcion'];
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
  zConsulta.Params.ParamByName('fecha').Value := anexo_ocompras.FieldValues['dIdFecha'];
  zConsulta.Open;
  if zConsulta.RecordCount>0 then begin
      If CompareText(VarName, 'REALIZO_PUESTO') = 0 then
          Value := zConsulta.FieldValues['sPuesto14'] ;
      If CompareText(VarName, 'REVISO_PUESTO') = 0 then
          Value := zConsulta.FieldValues['sPuesto15'] ;
      If CompareText(VarName, 'AUTORIZO_PUESTO') = 0 then
          Value := zConsulta.FieldValues['sPuesto16'] ;
      If CompareText(VarName, 'REALIZO_FIRMA') = 0 then
          Value := zConsulta.FieldValues['sFirmante14'] ;
      If CompareText(VarName, 'REVISO_FIRMA') = 0 then
          Value := zConsulta.FieldValues['sFirmante15'] ;
      If CompareText(VarName, 'AUTORIZO_FIRMA') = 0 then
          Value := zConsulta.FieldValues['sFirmante16'] ;
  end
  else
  begin
      If CompareText(VarName, 'REALIZO_PUESTO') = 0 then
          Value := '' ;
      If CompareText(VarName, 'REVISO_PUESTO') = 0 then
          Value := '';
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

procedure TfrmPedidos.imgNotasDblClick(Sender: TObject);
begin
    ComentariosAdicionales.Click
end;

procedure TfrmPedidos.ComentariosAdicionalesClick(Sender: TObject);
begin
    global_partida := rxPartidas.FieldValues['sNumeroActividad'] ;
    Application.CreateForm(TfrmComentariosxAnexo, frmComentariosxAnexo);
    frmComentariosxAnexo.show ;
end;
procedure TfrmPedidos.Copy1Click(Sender: TObject);
begin
  if grid_entradas.Focused=true then
    begin
      UtGrid.CopyRowsToClip;
    end;
  if grid_insumos.Focused=true then
    begin
      if grid_insumos.datasource.DataSet.IsEmpty=false  then
      if grid_insumos.DataSource.DataSet.RecordCount>0 then
      UtGrid2.CopyRowsToClip;
    end;
  if gridpartidas.Focused=true then
    begin
      UtGrid3.CopyRowsToClip;
    end;
end;

procedure TfrmPedidos.dbPartidasClick(Sender: TObject);
begin
  try
     insumos.Active := False;
     insumos.SQL.Clear;
     insumos.SQL.Add('select a.iFolioRequisicion as Requisicion, a.dCantidad as dCantidad, i.*, SubStr(i.mDescripcion, 1, 255) as sDescripcion, a.iItem, a.sNumeroActividad from anexo_prequisicion a  '+
                     'inner join insumos i on (i.sContrato = a.sContrato and a.sIdInsumo = i.sIdInsumo ) '+
                     'where a.sContrato =:Contrato and a.iFolioRequisicion =:Folio and a.sNumeroActividad =:Partida ');
     insumos.ParamByName('Contrato').AsString := global_contrato;
     insumos.ParamByName('Folio').AsInteger   := rxPartidas.FieldValues['iFolioRequisicion'];
     insumos.ParamByName('Partida').AsString  := rxPartidas.FieldValues['sNumeroActividad'];
     insumos.Open;
  except
    on e : exception do begin
       UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Ordenes de Compra de Materiales', 'Al seleccionar partida', 0);
    end;
  end;

end;

procedure TfrmPedidos.dbPartidasKeyPress(Sender: TObject; var Key: Char);
begin

  If Key = #13 Then
   tdcantidad.SetFocus         
end;

procedure TfrmPedidos.dbRequisicionEnter(Sender: TObject);
var
    x, i, a    : Integer ;
    cad, texto, inicio : string;
begin
    try
       dbrequisicion.Color:=global_color_entrada;
       if  OpcButton1 <> 'Edit' then
       begin
         inicio := dbRequisicion.Text;
         if requisiciones.RecordCount > 0 then
         begin
             requisiciones.First;
             dbRequisicion.Clear;
             while not requisiciones.Eof do
             begin
                  dbRequisicion.Items.Add(IntToStr(requisiciones.FieldValues['iFolioRequisicion'])+ '.- '+requisiciones.FieldValues['sNumeroSolicitud']  );
                  requisiciones.Next;
             end;
         end;

         cad := inicio;
           // if (anexo_ocompras.FieldValues['sFolioRequisicion'] <> null) then
           //    cad := anexo_ocompras.FieldValues['sFolioRequisicion']

            if cad <> '' then
            begin
              x := 1;
              i := length(cad);
              while x = 1 do
              begin
                  x     := pos(',',cad);
                  texto := cad;
                  if x > 0 then
                  begin
                      texto := copy(cad,1, x - 1);
                      cad   := copy(cad,x + 1, i);
                      x     := 1;
                  end;

                  if texto <> '' then
                  begin
                      a := dbRequisicion.Items.IndexOf(texto);
                      dbRequisicion.Checked[a] := True ;
                  end;
             end;
           end;
        end;

    Except
    end;
end;

procedure TfrmPedidos.dbRequisicionExit(Sender: TObject);
begin
dbrequisicion.Color:=global_color_salida;
end;

procedure TfrmPedidos.dbRequisicionKeyPress(Sender: TObject; var Key: Char);
begin
  if tsidproveedor.Enabled=true then
  begin
      If Key = #13 Then
         tsidproveedor.SetFocus
  end;

end;

procedure TfrmPedidos.tsNumeroOrdenEnter(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmPedidos.tsNumeroOrdenExit(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_salida
end;

procedure TfrmPedidos.tsPreciosEnter(Sender: TObject);
begin
    tsPrecios.Color := global_color_entrada;
end;

procedure TfrmPedidos.tsPreciosExit(Sender: TObject);
begin
    tsPrecios.Color := global_color_salida;
end;

procedure TfrmPedidos.tsPreciosKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsVigencia.SetFocus
end;

procedure TfrmPedidos.anexo_ocomprasCalcFields(DataSet: TDataSet);
begin
    if anexo_ocompras.RecordCount > 0 then
    begin
        Connection.qryBusca.Active := False ;
        Connection.qryBusca.SQL.Clear ;
        Connection.qryBusca.SQL.Add('Select Sum(r.dCantidad * (r.dCosto - r.dDescuento)) as dMontoMN From anexo_ppedido r ' +
                                    'Where r.sContrato = :Contrato And r.iFolioPedido = :Folio Group By r.iFolioPedido');
        connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        connection.qryBusca.Params.ParamByName('Contrato').Value := global_Contrato ;
        connection.qryBusca.Params.ParamByName('Folio').DataType := ftInteger ;
        connection.qryBusca.Params.ParamByName('Folio').Value := anexo_ocompras.FieldValues['iFolioPedido'] ;
        Connection.qryBusca.Open ;

        If Connection.qryBusca.RecordCount > 0 Then
           anexo_ocomprasdMontoMN.Value := ((Connection.qryBusca.FieldValues['dMontoMN'] - anexo_ocompras.FieldValues['dDescuento']) +
                                            ((Connection.qryBusca.FieldValues['dMontoMN'] - anexo_ocompras.FieldValues['dDescuento']) * anexo_ocompras.FieldValues['dIva']/100))
        else
           anexo_ocomprasdMontoMN.Value := 0;

        if (anexo_ocompras.State <> dsEdit) or (anexo_ocompras.State <> dsInsert ) then
           dbRequisicion.Text := anexo_ocompras.FieldValues['sFolioRequisicion'];
    end;
end;

procedure TfrmPedidos.anexo_pocomprasCalcFields(  DataSet: TDataSet);
begin
    anexo_pocomprasdMontoMN.Value := (anexo_pocomprasdCantidad.Value * anexo_pocomprasdCosto.Value) - anexo_pocomprasdDescuento.Value;
    anexo_pocomprassDescripcion.Text := MidStr(anexo_pocompras.FieldValues['mDescripcion'], 1 , 295) ;
end;

procedure TfrmPedidos.ReporteCalcFields(DataSet: TDataSet);
begin
  Connection.qryBusca.Active := False ;
  Connection.qryBusca.SQL.Clear ;
  Connection.qryBusca.SQL.Add('Select Sum(p.dCantidad) as dCantidad From anexo_ppedido p ' +
                                   'INNER JOIN anexo_ocompras a ON (a.sContrato = p.sContrato And a.iFolioPedido = p.iFolioPedido And a.dIdFecha <= :Fecha) ' +
                                   'Where p.sContrato = :Contrato And p.sNumeroActividad = :Actividad Group By p.sNumeroActividad ') ;
  connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
  connection.qryBusca.Params.ParamByName('Contrato').Value := Global_Contrato ;
  connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString ;
  connection.qryBusca.Params.ParamByName('Actividad').Value := Reporte.FieldValues['sNumeroActividad'] ;
  connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
  connection.qryBusca.Params.ParamByName('Fecha').Value := anexo_ocompras.FieldValues['dIdFecha'] ;
  Connection.qryBusca.Open ;
  If Connection.qryBusca.RecordCount > 0 Then
     ReportedAcumulado.Value := Connection.qryBusca.FieldValues['dCantidad'] ;
end;

procedure TfrmPedidos.tsCondicionesEnter(Sender: TObject);
begin
    tsCondiciones.Color := global_color_entrada;
end;

procedure TfrmPedidos.tsCondicionesExit(Sender: TObject);
begin
    tsCondiciones.Color := global_color_salida;
end;

procedure TfrmPedidos.tsCondicionesKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsEntrega.SetFocus
end;

procedure TfrmPedidos.tsEmbarcaEnter(Sender: TObject);
begin
    tsEmbarca.Color := global_color_entrada;
end;

procedure TfrmPedidos.tsEmbarcaExit(Sender: TObject);
begin
    tsEmbarca.Color := global_color_salida;
end;

procedure TfrmPedidos.tsEmbarcaKeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 Then
        tsCondiciones.SetFocus
end;

procedure TfrmPedidos.tsentregaEnter(Sender: TObject);
begin
     tsentrega.Color := global_color_entrada
end;

procedure TfrmPedidos.tsentregaExit(Sender: TObject);
begin
      tsentrega.Color := global_color_salida
end;

procedure TfrmPedidos.tsentregaKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsPrecios.SetFocus
end;

procedure TfrmPedidos.tsFormaPagoEnter(Sender: TObject);
begin
      tsFormaPago.Color := global_Color_entrada
end;

procedure TfrmPedidos.tsFormaPagoKeyPress(Sender: TObject; var Key: Char);
begin
      If Key = #13 Then
        tsEntrega.SetFocus
end;

procedure TfrmPedidos.tsIdProveedorEnter(Sender: TObject);
begin
    tsIdProveedor.Color := global_Color_entrada
end;

procedure TfrmPedidos.tsIdProveedorExit(Sender: TObject);
begin
    tsIdProveedor.Color := global_Color_salida
end;

procedure TfrmPedidos.tsIdProveedorKeyPress(Sender: TObject;
  var Key: Char);
begin
if tsnumeroorden.Enabled=true then
   begin
    If Key = #13 Then
        tsNumeroOrden.SetFocus
   end;
end;

procedure TfrmPedidos.FormActivate(Sender: TObject);
begin
    Proveedores.Refresh ;
end;


procedure TfrmPedidos.btnProveedoresClick(Sender: TObject);
begin
    global_frmActivo := 'frm_pedidos';
    Application.CreateForm(TfrmProveedores, frmProveedores);
    frmProveedores.show;
end;

procedure TfrmPedidos.Grid_EntradasCellClick(Column: TColumn);
begin
   If anexo_ocompras.RecordCount > 0 Then
      Begin
          iFolio := anexo_ocompras.FieldValues['iFoliopedido'] ;
          tiFolio.Text := anexo_ocompras.FieldValues['sOrdenCompra'] ;
          tdIdFecha.Date := anexo_ocompras.FieldValues['dIdFecha'] ;
          tsIdProveedor.KeyValue := anexo_ocompras.FieldValues['sIdProveedor'] ;
          tsNumeroOrden.Text := anexo_ocompras.FieldValues['sNumeroOrden'] ;
          tmComentarios.Text := anexo_ocompras.FieldValues['mComentarios'] ;

          anexo_pocompras.Active := False ;
          anexo_pocompras.SQL.Clear;
          if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
             anexo_pocompras.SQL.Add(CadenaOrden)
          else
             anexo_pocompras.SQL.Add(CadenaOrden2);
          anexo_pocompras.Params.ParamByName('Contrato').DataType := ftString ;
          anexo_pocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
          if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
          begin
               anexo_pocompras.Params.ParamByName('Convenio').DataType := ftString ;
               anexo_pocompras.Params.ParamByName('Convenio').Value    := global_convenio;
          end;
          anexo_pocompras.Params.ParamByName('Folio').DataType    := ftInteger ;
          anexo_pocompras.Params.ParamByName('Folio').Value       := anexo_ocompras.FieldValues['iFolioPedido'] ;
          anexo_pocompras.Open ;

          dbRequisicion.Text := anexo_ocompras.FieldValues['sFolioRequisicion'];
      end;
  end ;
procedure TfrmPedidos.Grid_EntradasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmPedidos.Grid_EntradasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmPedidos.Grid_EntradasTitleClick(Column: TColumn);
begin
 UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmPedidos.Grid_InsumosCellClick(Column: TColumn);
begin
if grid_insumos.datasource.DataSet.IsEmpty=false  then
  begin
  if grid_insumos.DataSource.DataSet.RecordCount>0 then
    begin
      if insumos.RecordCount > 0 then
      begin
          if anexo_ocompras.FieldValues['sFolioRequisicion'] <> '' then
             tdCantidad.Value   := insumos.FieldValues['dCantidad']
          else
              tdCantidad.Value := 0;
          tdCostoNuevo.Value := insumos.FieldValues['dNuevoPrecio'];
      end;
    end;
  end;
end;

procedure TfrmPedidos.Grid_InsumosKeyPress(Sender: TObject; var Key: Char);
begin
       if Key = #13 then
         tdCantidad.SetFocus;
end;

procedure TfrmPedidos.Grid_InsumosMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
if grid_insumos.datasource.DataSet.IsEmpty=false  then
if grid_insumos.DataSource.DataSet.RecordCount>0 then
UtGrid2.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmPedidos.Grid_InsumosMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
if grid_insumos.datasource.DataSet.IsEmpty=false  then
if grid_insumos.DataSource.DataSet.RecordCount>0 then
UtGrid2.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmPedidos.Grid_InsumosTitleClick(Column: TColumn);
begin
if grid_insumos.datasource.DataSet.IsEmpty=false  then
if grid_insumos.DataSource.DataSet.RecordCount>0 then
 UtGrid2.DbGridTitleClick(Column);
end;

procedure TfrmPedidos.validaciones;
begin
     //Verificamos el estatus de la orden de compra..
     valida := False;
     connection.zCommand.Active := False;
     connection.zCommand.SQL.Clear;
     connection.zCommand.SQL.Add('select iFolioPedido from anexo_pedidos where sContrato =:Contrato and iFolioPedido =:Pedido and sStatus = "AUTORIZADO" ');
     connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
     connection.zCommand.ParamByName('Pedido').AsString   := anexo_ocompras.FieldValues['iFolioPedido'];
     connection.zCommand.Open;

     if connection.zCommand.RecordCount > 0 then
     begin
          messageDLG(' La Orden de Compra '+anexo_ocompras.FieldValues['sOrdenCompra']+ ' se encuentra en estatus de Autorizado', mtInformation, [mbOk], 0);
          valida := True;
     end;      
end;

procedure TfrmPedidos.validaPedido;
begin
     connection.QryBusca.Active := False;
     connection.QryBusca.SQL.Clear;
     connection.QryBusca.SQL.Add('select iFolioPedido from anexo_ppedido where sContrato =:Contrato and iFolioPedido =:Folio ');
     connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
     connection.QryBusca.ParamByName('Folio').AsInteger   := iFolio;
     connection.QryBusca.Open;

     if connection.QryBusca.RecordCount > 0 then
        lValidaReq := True
     else
        lValidaReq := False;
end;

procedure TfrmPedidos.Seguimiento_Material(dParamActividad: string);
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
                                          'on (i.sContrato = ra.sContrato and i.sIdInsumo = ra.sIdInsumo) '+
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
                                          'left join anexo_ppedido  ap '+
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

procedure TfrmPedidos.tipocambio;
begin
    if tsMoneda.Text <> '' then
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select * from tiposdecambio where sIdMoneda =:Moneda and dIdFecha <=:Fecha order by dIdFecha DESC ');
        connection.zCommand.ParamByName('Moneda').AsString := moneda.FieldValues['sIdMoneda'];
        connection.zCommand.ParamByName('Fecha').AsDate    := tdIdFecha.Date;
        connection.zCommand.Open;

        If connection.zCommand.RecordCount > 0 then
           tsTipoCambio.Value := connection.zCommand.FieldValues['dCambio']
        else
           tsTipoCambio.Value := 0;
    end;
end;


End.








