unit frm_cuadre_normal;

interface

{CUADRE v1 2014-08-12 - Martin, Saul, Rangel}
{CUADRE v2 2014-08-17 - Martin, Saul, Rangel}
{CUADRE v3 2014-09-23 - Martin, Saul, Rangel}
{CUADRE v4 2014-09-25 - Martin, Saul, Rangel}
{CUADRE v4 2014-10-11 - Martin, Saul, Rangel}

uses

  {$region 'Uses'}

  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ZConnection, DB, ZAbstractRODataset, ZDataset, cxSSTypes, cxSSStyles,
  cxSSFormulas, cxSSHeaders, cxExcelConst, StrUtils, Math, frm_connection, global,
  ComObj, UnitExcel, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxNavigator, cxDBData, JvDialogs, cxDropDownEdit, cxLookupEdit,
  cxDBLookupEdit, cxDBLookupComboBox, DBClient, ImgList, ZAbstractDataset,
  ComCtrls, dxtree, NxCollection, cxGridLevel, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxClasses, cxGridCustomView, cxGrid,
  cxSSheet, AdvSmoothProgressBar, cxLabel, cxTextEdit, cxMaskEdit, cxCalc,
  StdCtrls, cxRadioGroup, cxGroupBox, cxButtons, ExtCtrls, DateUtils,Utilerias,
  frm_bitacoradepartamental_2, Types, ZDbcIntfs,


  cxGraphics,
  cxControls, cxLookAndFeels, cxLookAndFeelPainters, dxSkinsCore, AdvMenus,
  AdvStickyPopupMenu, dxSkinsdxBarPainter, dxBar, AdvSmoothPopup, cxStyles,
  dxSkinscxPCPainter, cxListView, cxContainer, cxEdit, dxSkinBlack, dxSkinBlue,
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
  dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint, dxSkinXmas2008Blue, Menus,
  cxProgressBar, cxCheckBox, cxMemo, cxHeader, cxTL, cxTLdxBarBuiltInMenu,
  cxInplaceContainer, LabelEdit, CurvyControls, cxSpinEdit, cxTimeEdit,
  cxListBox, ExcelXP, OleServer, dxGalleryControl,
  dxColorGallery, dxColorEdit, DBCtrls, cxDBEdit;

  {$endregion}

(*{ Martin Samuel }*)

//Tipos de datos para las clases
type
  TCategoriaIndex = type Integer;
  TActividadIndex = type Integer;
  TActividadPadreIndex = type Integer;
  TTipoRecursoIndex = type Integer;
  TIdActividad = type Integer;
  TMoeIndex = type Integer;
  TFolioIndex = type Integer;
  TExcelColIndex = type Integer;
  TInicioFolio = type Integer;
  TFinFolio = type Integer;
  TFilaExcel = type Integer;
  TExcelColAlias = type string;
  TExcelRangeAlias = type string;
  TExcelInstance = type Variant;
  TExcelRow = type integer;

//Martin Samuel
type
  TExcelFila = class( TObject )
    public
      Row,
      OldRow : TExcelRow;
      Top : integer;
      Height : integer;

    function NextRow( count : TExcelRow = 1 ):TExcelRow;
    function sRow:string;

  end;

//Clase simple para un folio
type
  TSimpleFolio = class( TObject )
    Inicio : TInicioFolio;
    Fin : TFinFolio;

    function StrInicio:string;
    function StrFin:string;

    constructor Create;
  end;


//Clase personalizada para columnas en excel
type
  TColumnaExcel = class( TObject )
    iColumna : TExcelColIndex;
    sColumna : TExcelColAlias;
    function Columna():TExcelColAlias;
    function _Columna():TExcelColAlias;overload;
    function _Columna( Increment : TExcelColIndex ):TExcelColAlias;overload;
    function Columna_():TExcelColAlias;overload;
    function Columna_( Increment : TExcelColIndex ):TExcelColAlias;overload;

    constructor Create;
  end;


//Clase para para las categorias de personal y equipo ( MOE )
type
  TCategoria = class( TObject )
    private
      sIdRecurso,
      sPernocta,
      sPlataforma,
      sIdCategoria,
      sAplicaPernocta : string;

      iCol : TExcelColIndex;
      iInicio,
      iSolicitado,
      iAbordo,
      iFin,
      iItemOrden : integer;

      eListo,
      eExiste : Boolean;

      sSuma : string;
  end;


//Clase para actividades padres( Generales )
type
  TActividadPadre = class( TObject )
    private
      IdActividad  : string;
      Fila,
      FilaInicio,
      FilaFin : TFilaExcel;
      ActividadCount : Integer;

      IndexActividad : TActividadIndex;

      function GetChildCountRange( Columna : TExcelColAlias ):TExcelRangeAlias;
  end;


//Puntero a Actividad Padre
type
  PActividadPadre = ^TActividadPadre;


//Clase para las actividades( Horarios )
type
  TActividad = class( TObject )
    private
      sWbs,
      sIdActividad,
      sHInicio,
      sHFin : string;

      dDuracion : Double;

      IsPadre,
      NuevoCorte : Boolean;
      Padre : PActividadPadre;

      iRow : TFilaExcel;
      iIdActividad,
      iIdDiario,
      iTarea,
      iNodoCorte,
      iHermano,
      iHermanosCount,
      iInicioConjunto : Integer;

      IndexPadre : TActividadPadreIndex;

      function Delete( Fecha : TDate ):TActividadIndex;
      constructor Create;
  end;


//Clase para los Folios
type
  TFolio = class ( TObject )
    private
      sFolio,
      sPernocta,
      sPlataforma : string;

      iInicio : TInicioFolio;
      iFin : TFinFolio;

      Fila : TFilaExcel;

      eExiste : Boolean;

      ACTIVIDADES : array of TActividad;
      ACTIVIDADES_PADRES : array of TActividadPadre;
      MOE : array of TCategoria;

      function BuscarActividad( Actividad : string ) : Boolean;
      function BuscarCategoria( iColumn, Pos: Integer ): TMoeIndex;overload;
      function BuscarCategoria( iColumn: Integer ): TMoeIndex;overload;
      function BuscarActividadPorNodoIndex( NodoIndex : Integer ):TActividadIndex;overload;
      function IndexPadre( Actividad : string ):TActividadPadreIndex;

      procedure AvanzarActividades( IndexFrom : Integer ; Increment : Integer = 1 );
      procedure RetrocederActividades( IndexFrom : TActividadIndex );
      procedure InsertarActividad( Actividad : TActividad ; Indice : Integer );
      procedure SetPFMOE( Pernocta, Plataforma : string );
      procedure CleanNodes();
      procedure UpdateRange();
      procedure UpdateActRows();
      procedure IncrementRow();
      procedure ActualizarCountActividades();

  end;

//Clase para el Cuadre en general ( Almacena todas las demas classes y asi las demas )
type
  TCuadre = class( TObject )
  private
    Fecha : string;
    Cambios,
    Guardado : Boolean;

    CATEGORIA : array[0..1] of array of TFolio;
    INSERTADOS : array[0..1] of integer;
    CUADRAR : array[0..1] of Boolean;

    Inicio,
    Fin : TExcelRow;

    function BuscaFolio( Hoja, Fila : integer ) : TFolioIndex;
    function BuscarActividad( Hoja, Fila : integer ): TActividadIndex;
    function GetIdActividad( Hoja, Fila : integer ): TIdActividad;

    procedure SaveToExcel();
    procedure UpdateAllRanges();
    procedure RetrocederFoliosDesde( Indice : TFolioIndex );
    procedure AvanzarFoliosDesde( Indice : TFolioIndex );
  const
    TIPO : array[0..1] of string = ('Personal', 'Equipo');
  end;


type
  TfrmCuadreNormal = class(TForm)

  {$REGION 'Componentes'}
  
    dsMOE: TDataSource;
    dsActividades: TDataSource;
    qrMOE_Sol: TZReadOnlyQuery;
    qrActividades: TZReadOnlyQuery;
    pnlDatos: TPanel;
    dsOT: TDataSource;
    qrOT: TZQuery;
    dsFolios: TDataSource;
    qrFolios: TZQuery;
    dsReportes: TDataSource;
    qrReportes: TZQuery;
    pnlCuadre: TPanel;
    grpParams: TcxGroupBox;
    rbRedondear: TcxRadioButton;
    rbTruncar: TcxRadioButton;
    clcRT: TcxCalcEdit;
    Label6: TLabel;
    Label7: TLabel;
    clcJornadas: TcxCalcEdit;
    imgPop: TcxImageList;
    pnlEstructura: TNxHeaderPanel;
    Panel2: TPanel;
    cxButton1: TcxButton;
    cxButton2: TcxButton;
    imgTree: TcxImageList;
    Libro: TcxSpreadSheetBook;
    CdPivote: TClientDataSet;
    CdTmpActividades: TClientDataSet;
    CdHorarios: TClientDataSet;
    dsCortes: TDataSource;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1Column1: TcxGridDBColumn;
    cxGrid1DBTableView1Column2: TcxGridDBColumn;
    cxGrid1DBTableView1Column3: TcxGridDBColumn;
    cxGrid1DBTableView1Column4: TcxGridDBColumn;
    cxGrid1DBTableView1Column5: TcxGridDBColumn;
    cxGrid1DBTableView1Column6: TcxGridDBColumn;
    cxGrid1DBTableView1Column7: TcxGridDBColumn;
    cxGrid1DBTableView1Column8: TcxGridDBColumn;
    cxGrid1DBTableView1Column9: TcxGridDBColumn;
    cxGrid2: TcxGrid;
    cxGridDBTableView1: TcxGridDBTableView;
    cxGridDBColumn1: TcxGridDBColumn;
    cxGridDBColumn2: TcxGridDBColumn;
    cxGridDBColumn3: TcxGridDBColumn;
    cxGridDBColumn4: TcxGridDBColumn;
    cxGridDBColumn5: TcxGridDBColumn;
    cxGridDBColumn6: TcxGridDBColumn;
    cxGridDBColumn7: TcxGridDBColumn;
    cxGridDBColumn8: TcxGridDBColumn;
    cxGridDBColumn9: TcxGridDBColumn;
    cxGridLevel1: TcxGridLevel;
    CdResult: TClientDataSet;
    zActividadesCortes: TZQuery;
    dlgExcel: TJvSaveDialog;
    pnlGuardar: TNxHeaderPanel;
    rbExcel: TcxRadioButton;
    rbDatabase: TcxRadioButton;
    btnSave: TcxButton;
    dlgOpenExcel: TJvOpenDialog;
    cxGBDatos: TcxGroupBox;
    dsTmpActividades: TDataSource;
    popHoja: TPopupMenu;
    Foliosexistentesenelcuadre1: TMenuItem;
    AbrirCuadreexistente1: TMenuItem;
    Cortarestaactividad1: TMenuItem;
    cxGBBotonera: TcxGroupBox;
    CxBtnCut: TcxButton;
    CxBtnCancel: TcxButton;
    cxGroupBox1: TcxGroupBox;
    CxTextEdtFolio: TcxTextEdit;
    cxMemoDescripcion: TcxMemo;
    CxTextEdthoraInicio: TcxTextEdit;
    CxTextEdtHoraTermino: TcxTextEdit;
    CxLbl1: TcxLabel;
    CxLbl2: TcxLabel;
    CxLbl3: TcxLabel;
    CxLblDescripcin: TcxLabel;
    cxGBCortes: TcxGroupBox;
    CxGridActividades: TcxGrid;
    CxGridActividadesCortes: TcxGridDBTableView;
    CxColumnIdActividad: TcxGridDBColumn;
    CxColumnDescripcion: TcxGridDBColumn;
    CxColumnFolio: TcxGridDBColumn;
    CxColumnHoraInicio: TcxGridDBColumn;
    CxColumnHoraTermino: TcxGridDBColumn;
    CxColumnSeleccionar: TcxGridDBColumn;
    CxLevelActividades: TcxGridLevel;
    CdCortesSugeridos: TClientDataSet;
    dsCortesSugeridos: TDataSource;
    pnl1: TPanel;
    pnl2: TPanel;
    grp1: TcxGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    cbbOts: TcxLookupComboBox;
    cbbReportes: TcxLookupComboBox;
    CxBtnCortar: TcxButton;
    cxGroupBox2: TcxGroupBox;
    cxGBMoe: TcxGroupBox;
    btnPintar: TcxButton;
    cxHeader1: TcxHeader;
    CxTextEdtActividad: TcxTextEdit;
    CxLbl6: TcxLabel;
    cxMemo1: TcxMemo;
    cxHeader3: TcxHeader;
    cxHeader4: TcxHeader;
    cxHeader5: TcxHeader;
    CxLbl5: TcxLabel;
    CxLbl7: TcxLabel;
    CxLbl8: TcxLabel;
    CxTextEdtinicio: TcxTextEdit;
    CxTextEdtTermino: TcxTextEdit;
    CxTextEdtDuracion: TcxTextEdit;
    btnGuardar: TcxButton;
    CxLbl9: TcxLabel;
    CxTextEdtSolicitado: TcxTextEdit;
    CxLbl10: TcxLabel;
    CxTextEdtaBordo: TcxTextEdit;
    CxTextEdt1: TcxTextEdit;
    txtHH: TcxTextEdit;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    Panel1: TPanel;
    CxLblFolio: TcxLabel;
    CxLblCategoria: TcxLabel;
    cxLabel3: TcxLabel;
    cxLabel4: TcxLabel;
    prgFolios: TAdvSmoothProgressBar;
    prgActividades: TAdvSmoothProgressBar;
    lblEstado: TcxLabel;
    pnlCortes: TNxHeaderPanel;
    Panel3: TPanel;
    grpActividadesCortes: TcxGroupBox;
    grpCorte: TcxGroupBox;
    tmRestaHoras: TcxTimeEdit;
    cxLabel5: TcxLabel;
    btnCortar: TcxButton;
    GenerarCorte1: TMenuItem;
    mDescripcionCorte: TcxMemo;
    CurvyPanel1: TCurvyPanel;
    lstCortes: TdxTreeView;
    icnsCortes: TcxImageList;
    btnCancelarCorte: TcxButton;
    lblResultado: TcxLabel;
    lblInicio: TcxLabel;
    lblFin: TcxLabel;
    lblCorteI: TcxLabel;
    lblCorteF: TcxLabel;
    cxLabel8: TcxLabel;
    qrPernoctas: TZReadOnlyQuery;
    qrPlataformas: TZReadOnlyQuery;
    cbbPlataformas: TcxLookupComboBox;
    cxLabel7: TcxLabel;
    cxLabel6: TcxLabel;
    cbbPernoctas: TcxLookupComboBox;
    dsPernoctas: TDataSource;
    dsPlataformas: TDataSource;
    cxLabel9: TcxLabel;
    cxLabel10: TcxLabel;
    cxLabel11: TcxLabel;
    chkAnida: TcxCheckBox;
    Eliminarestaactividad1: TMenuItem;
    cxButton3: TcxButton;
    lstEstructura: TcxListView;
    ExportarCuadrevirtualaExcel1: TMenuItem;
    cbbVista: TcxComboBox;
    Label1: TLabel;
    ExportarCuadre1: TMenuItem;
    ExcelApplication1: TExcelApplication;
    ExcelWorksheet1: TExcelWorksheet;
    ExcelWorkbook1: TExcelWorkbook;
    zAjuste: TZQuery;
    Ajustes1: TMenuItem;
    grpAjuste: TcxGroupBox;
    cxLabel12: TcxLabel;
    cxLabel13: TcxLabel;
    btnAplicaAjuste: TcxButton;
    lstCuentas: TcxListBox;
    clcAjuste: TcxCalcEdit;
    CxLbl4: TcxLabel;
    dxColores: TdxColorGallery;
    cxCmdColor: TcxButton;
    cxCmdAplicar: TcxButton;
    zqFolios: TZQuery;
    ds_Folios: TDataSource;
    tsFolio: TDBLookupComboBox;
    cxMuestraColor: TcxTextEdit;
    cxLabel14: TcxLabel;
    chkImprime: TCheckBox;
    Cuadreen01: TMenuItem;
    tGrupo: TGroupBox;
    cxNec: TcxLabel;
    cxInicio: TcxLabel;
    cxFin: TcxLabel;
    CuadreenCeros1: TMenuItem;
    chkOrdenado: TCheckBox;
    cbbAplicaPernocta: TcxComboBox;
    cxlbl11: TcxLabel;

    {$ENDREGION}

  {$REGION 'Procedimientos Formulario'}
    
    procedure FormCreate(Sender: TObject);
    procedure btnPintarClick(Sender: TObject);
    procedure Foliosexistentesenelcuadre1Click(Sender: TObject);
    procedure popHojaPopup(Sender: TObject);
    procedure cbbOtsPropertiesChange(Sender: TObject);
    procedure LibroEndEdit(Sender: TObject);
    procedure LibroClearCells(Sender: TcxSSBookSheet; const ACellRect: TRect; var UseDefaultStyle, CanClear: Boolean);
    procedure CxBtnCortarClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnGuardarClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure AbrirCuadreexistente1Click(Sender: TObject);
    procedure CortarestaActividad1Click(Sender: TObject);
    procedure grp1Click(Sender: TObject);
    procedure CxColumnSeleccionarPropertiesEditValueChanged(Sender: TObject);
    procedure CxBtnCancelClick(Sender: TObject);
    procedure LibroSetSelection(Sender: TObject; ASheet: TcxSSBookSheet);
    procedure cbbReportesPropertiesCloseUp(Sender: TObject);
    procedure btnActividadesClick(Sender: TObject);
    procedure GenerarCorte1Click(Sender: TObject);
    procedure tmInicioPropertiesEditValueChanged(Sender: TObject);
    procedure lstCortesChange(Sender: TObject; Node: TTreeNode);
    procedure cbbPernoctasPropertiesChange(Sender: TObject);
    procedure cbbPlataformasPropertiesChange(Sender: TObject);
    procedure tmRestaHorasPropertiesChange(Sender: TObject);
    procedure btnCortarClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Eliminarestaactividad1Click(Sender: TObject);
    procedure cxButton3Click(Sender: TObject);
    procedure ExportarCuadrevirtualaExcel1Click(Sender: TObject);
    procedure cbbVistaPropertiesCloseUp(Sender: TObject);
    procedure ExportarCuadre1Click(Sender: TObject);
    procedure Ajustes1Click(Sender: TObject);
    procedure cxButton4Click(Sender: TObject);
    procedure lstCuentasClick(Sender: TObject);
    procedure btnAplicaAjusteClick(Sender: TObject);
    procedure clcAjustePropertiesChange(Sender: TObject);
    procedure cxCmdColorClick(Sender: TObject);
    procedure cxCmdAplicarClick(Sender: TObject);
    procedure zqFoliosAfterScroll(DataSet: TDataSet);
    procedure tsFolioExit(Sender: TObject);
    procedure dxColoresItemClick(Sender: TObject; AItem: TdxGalleryControlItem);
    procedure tsFolioClick(Sender: TObject);
    procedure chkImprimeEnter(Sender: TObject);
    procedure Cuadreen01Click(Sender: TObject);
    procedure cxNecClick(Sender: TObject);
    procedure CuadreenCeros1Click(Sender: TObject);
    procedure cbbAplicaPernoctaPropertiesChange(Sender: TObject);

    {$ENDREGION}

  private
    { Private declarations }

    gForm: TForm;

    COLUMNAS: array[1..1400] of string;
    sArchivo,
    sAInicio,
    sAFin,
    sOLdCorte,
    DescripcionCorte : string;

    iTotalFilas : integer;

    Cuadre : TCuadre;
    ActividadCorte : TActividad;

    ListaCampos: TStringList;
    ePintando,
    ErrorCorte : Boolean;
    lCeros : Boolean;

    function ColumnaNombre(Numero: Integer): String;//Martin
    function sfnRestaHoras(sParamHorasMax, sParamHorasMin: string): string;//Inteligent
    function sfnSumaHoras(sParamHorasMax, sParamHorasMin: string): string;//Inteligent
    function rfnDecimal(sParamCantidad: string): Real;//Inteligent
    function detectarCruces(var cdDatosBuscar: TClientDataSet; var cdDatosDestino: TClientDataSet; idActividad: Integer; HoraInicio: Double; HoraTermino:Double): TStringList;//Saul
    function Redondear(numero : real ; cifrasSig : integer) : real;//Martin
    function Truncar(numero : Real; cifras : integer) : Real;//Martin

    procedure zQueryCopy(var ZDataset: TZQuery; var cdDataset: TClientDataSet); //Saul
    procedure getHM(cadena: string; var h, m: Double); //Martin
    procedure DefinirCortes(DatasetOrigen: TClientDataSet;var DatasetDestino: TClientDataSet; var Campos: TStringList); //Saul
    procedure HorasToDecimal(var DataActividades: TZReadOnlyQuery; var DataDestino: TClientDataSet; inverso: Boolean); //Saul
    procedure Escribe();//Martin
    procedure ComprobarSuma( Edit_Clean : Boolean );//Martin
    procedure GenerarExcel( Personal, Equipo : Boolean );//Martin
    procedure GuardaEnBD();//Martin
    procedure RegenerarSumasMOE(var Excel : Variant; IndiceTipo : TCategoriaIndex );//Martin
    procedure ValidaCategorias(var excel, hoja : Variant ; var zqMoe : TZReadOnlyQuery; iTipo : Integer);//Martin}
    procedure ValidaFolios();//Martin
    procedure VentanaCortes();//Martin
    procedure CargarActividades_ListView( Lista : TdxTreeView ; Folio : Integer; Actividad : string);//Martin
    procedure CortarActividad( Actividad, Inicio, Fin, InicioCorte, FinCorte : string ; Fila : Integer );//Martin/
    procedure EliminarHorario( Fila : TFilaExcel; hFinal : string ; IndiceF : TFolioIndex ; IndiceA : TActividadIndex; IndicePadre : TActividadPadreIndex );//Martin
    procedure Diagrama();//Martin
    procedure CambiaVistaCuadre( Tipo : string );
    procedure ConsultarActividades;

  public
    { Public declarations }
  end;

var
  frmCuadreNormal: TfrmCuadreNormal;

const
  TIPO : array[0..1] of string = ('Personal', 'Equipo');

  EXCEL_VISIBLE : Boolean = False;

  {$REGION 'Inserta Personal y Equipo'}

  SQLINSERT : array[0..1] of string = ('insert into bitacoradepersonal '+
                                          '(sContrato, '+
                                          'dIdFecha, '+
                                          'iIdDiario, '+
                                          'iItemOrden, '+
                                          'sIdPersonal, '+
                                          'sTipoObra, '+
                                          'sDescripcion, '+
                                          'sIdPernocta, '+
                                          'sIdPlataforma, '+
                                          'sHoraInicio, '+
                                          'sHoraFinal, '+
                                          'dCantidad, '+
                                          'sTipoPernocta, '+
                                          'sWbs, '+
                                          'sNumeroActividad, '+
                                          'dCantHH, '+
                                          'dAjuste, '+
                                          'sNumeroOrden, '+
                                          'iIdActividad, '+
                                          'iIdTarea, '+
                                          'sAgrupaPersonal, '+
                                          'sHoraInicioG, '+
                                          'sHoraFinalG, '+
                                          'lImprime,lAplicaPernocta ) '+#10+

                                    'values (:orden, '+
                                        ':fecha, '+
                                        ':iddiario, '+
                                        ':ItemOrden, '+
                                        ':idrecurso, '+
                                        ':TipoObra, '+
                                        ':descripcion, '+
                                        ':pernocta, '+
                                        ':plataforma, '+
                                        ':hinicio, '+
                                        ':hfinal, '+
                                        ':cantidad, '+
                                        '"4.1.", '+
                                        ':wbs, '+
                                        ':actividad, '+
                                        ':cantidadhh, '+
                                        ':Ajuste, '+
                                        ':folio, '+
                                        ':idactividad, '+
                                        ':tarea, '+
                                        ':Categoria, '+
                                        ':hinicio, '+
                                        ':hfinal, '+
                                        ':Imprime,:AplicaPernocta ) '

                                        ,

                                        'insert into bitacoradeequipos (sContrato, '+
                                        'dIdFecha, '+
                                        'iIdDiario, '+
                                        'iItemOrden, '+
                                        'sIdEquipo, '+
                                        'sDescripcion, '+
                                        'sIdPernocta, '+
                                        'sIdPlataforma, '+
                                        'sTipoObra, '+
                                        'sHoraInicio, '+
                                        'sHoraFinal, '+
                                        'dCantidad, '+
                                        'sWbs, '+
                                        'sNumeroActividad, '+
                                        'dCantHH, '+
                                        'dAjuste, '+
                                        'sNumeroOrden, '+
                                        'iIdActividad, '+
                                        'iIdTarea, '+
                                        'sHoraInicioG, '+
                                        'sHoraFinalG, '+
                                        'lImprime ) '+#10+

                                    'values (:orden, '+
                                        ':fecha, '+
                                        ':iddiario, '+
                                        ':ItemOrden, '+
                                        ':idrecurso, '+
                                        ':descripcion, '+
                                        ':pernocta, '+
                                        ':plataforma, '+
                                        ':TipoObra, '+
                                        ':hinicio, '+
                                        ':hfinal, '+
                                        ':cantidad, '+
                                        ':wbs, '+
                                        ':actividad, '+
                                        ':cantidadhh, '+
                                        ':Ajuste, '+
                                        ':folio, '+
                                        ':idactividad, '+
                                        ':tarea, '+
                                        ':hinicio, '+
                                        ':hfinal, '+
                                        ':Imprime )');

  {$ENDREGION}

  {$REGION 'Consultas MOE'}
  sSQLMOE : array [0..1] of string = ('select mr.iIdMoe, '+
                                     'mr.sIdRecurso, '+
                                     'mr.sDescripcion, '+
                                     'ax.sTierra, '+
                                     'mr.dCantidad as iSolicitado, '+
                                     'mra.dCantidad as iAbordo, p.sIdTipoPersonal, p.iItemOrden, '+
                                     'if ( lower( mr.sDescripcion ) = "tiempo extra", "Si" , "No") as sTE '+#10+

                                   'from moerecursos mr '+
                                   'inner join moe m '+
                                     'on ( m.sContrato = :orden '+
                                       'and mr.iIdMoe = m.iIdMoe ) '+ #10+

                                   'inner join moerecursos_abordo mra '+
                                     'on ( mra.iIdMoe = m.iIdMoe '+
                                       'and mra.sIdRecurso = mr.sIdRecurso '+
                                       'and mra.eTipoRecurso = "Personal" ) '+ #10+

                                   'inner join personal p '+
                                     'on (  p.sContrato = :contrato '+
                                       'and mr.eTipoRecurso = "Personal" '+
                                       'and p.sIdPersonal = mr.sIdRecurso ) '+ #10+

                                   'inner join anexos ax '+
                                      'on ( ax.sAnexo = p.sAnexo '+
                                        '&& lower( ax.sTipo ) = "personal" ) '+#10+

                                   'where m.dIdFecha = ( select max( m1.dIdFecha ) '+
                                                       'from moe m1 '+
                                                       'where m1.sContrato = :orden '+
                                                       'and m1.dIdFecha <= :fecha ) group by p.sIdPersonal '+
                                                       ' order by p.iItemOrden '

                                   ,
                                    'select mr.iIdMoe, '+
                                    'mr.sIdRecurso, '+
                                    'mr.sDescripcion, '+
                                    'mr.dCantidad as iSolicitado, '+
                                    'mra.dCantidad as iAbordo, e.iItemOrden '+#10+

                                  'from moerecursos mr '+
                                  'inner join moe m '+
                                    'on ( m.sContrato = :orden '+
                                      'and mr.iIdMoe = m.iIdMoe ) '+#10+

                                  'inner join moerecursos_abordo mra '+
                                    'on ( mra.iIdMoe = m.iIdMoe '+
                                      'and mra.sIdRecurso = mr.sIdRecurso '+
                                      'and mra.eTipoRecurso = "Equipo" ) '+#10+

                                  'inner join equipos e '+
                                    'on (  e.sContrato = :contrato '+
                                      'and mr.eTipoRecurso = "Equipo" '+
                                      'and e.sIdEquipo = mr.sIdRecurso ) '+

                                  'where m.dIdFecha = ( select max( m1.dIdFecha ) '+
                                                      'from moe m1 '+
                                                      'where m1.sContrato = :orden '+
                                                      'and m1.dIdFecha <= :fecha ) '+
                                                      'group by e.sIdEquipo  order by e.iItemOrden ');

  {$ENDREGION}

implementation

{$R *.dfm}

//Martin Samuel
function TExcelFila.sRow:string;
begin
  Result := IntToStr( Row );
end;


//Martin Samuel
function TExcelFila.NextRow( count : TExcelRow = 1 ):TExcelRow;
begin
  Inc( Row, count );
  Result := Row;
end;


//Martin
constructor TSimpleFolio.Create;
begin
  inherited Create;
  Inicio := -1;
  Fin := -1;
end;


//Martin
function TSimpleFolio.StrInicio:string;
begin
  Result := IntToStr( Inicio );
end;


//Martin
function TSimpleFolio.StrFin:string;
begin
  Result := IntToStr( Fin );
end;


//Martin
constructor TColumnaExcel.Create;
begin
  inherited Create;
  iColumna := 1;
  sColumna := ColumnaNombre( iColumna );
end;


//Martin
function TColumnaExcel.Columna():TExcelColAlias;
begin
  sColumna := ColumnaNombre( iColumna );
  Result := sColumna;
end;


//Martin
function TColumnaExcel._Columna:TExcelColAlias;
begin
  Result := ColumnaNombre( iColumna - 1 );
end;


//Martin
function TColumnaExcel.Columna_:TExcelColAlias;
begin
  Result := ColumnaNombre( iColumna + 1 );
end;


//Martin
function TColumnaExcel._Columna( Increment : TExcelColIndex ):TExcelColAlias;
begin
  Result := ColumnaNombre( iColumna - Increment );
end;


//Martin
function TColumnaExcel.Columna_( Increment : TExcelColIndex ):TExcelColAlias;
begin
  Result := ColumnaNombre( iColumna + Increment );
end;


//Martin
constructor TActividad.Create;
begin
  inherited Create;

  iNodoCorte := -1;
  iHermano := -1;
  iInicioConjunto := -1;
  IsPadre := False;
  IndexPadre := -1;
end;


//Martin
procedure TFolio.SetPFMOE(Pernocta: string; Plataforma: string);
var
  IndexM : TMoeIndex;
begin
  for IndexM := 0 to Length( MOE ) - 1 do
  begin
    MOE[IndexM].sPernocta := Pernocta;
    MOE[IndexM].sPlataforma := Plataforma;
  end;
end;

procedure TFolio.UpdateRange;
begin
  iInicio := Fila + 2;
  iFin := ACTIVIDADES[ Length( ACTIVIDADES ) - 1 ].iRow;     
end;


procedure TFolio.ActualizarCountActividades;
var
  APIndex : TActividadPadreIndex;
  AIndex : TActividadIndex;
begin
  AIndex := 0;
  for APIndex := 0 to Length( ACTIVIDADES_PADRES ) - 1 do
  begin
    ACTIVIDADES_PADRES[ APIndex ].ActividadCount := 0;
    while ( AIndex < Length( ACTIVIDADES ) ) and ( ACTIVIDADES[ AIndex ].sIdActividad = ACTIVIDADES_PADRES[ APIndex ].IdActividad ) do
    begin
      if ( ACTIVIDADES[ AIndex ].IsPadre ) then
      begin
        Inc( AIndex );
        Continue;
      end;
      Inc( ACTIVIDADES_PADRES[ APIndex ].ActividadCount );
      Inc( AIndex );
    end;

  end;

end;

procedure TFolio.UpdateActRows;
var
  AIndex : TActividadIndex;
  Row : TExcelRow;
begin

  Row := Fila + 2;
  for AIndex:= 0 to Length( ACTIVIDADES ) - 1 do
  begin
    ACTIVIDADES[ AIndex ].iRow := Row;
      if ACTIVIDADES[ AIndex ].IsPadre then
        ACTIVIDADES_PADRES[ ACTIVIDADES[ AIndex ].IndexPadre ].Fila := Row;

    Inc( Row );
  end;
  
end;


procedure TFolio.IncrementRow;
begin
  Fila := Fila + 1;
  UpdateRange;
  UpdateActRows;
end;

//Martin
function TActividadPadre.GetChildCountRange( Columna : TExcelColAlias ):TExcelRangeAlias;
var
  ACount : Integer;
begin
  Result := Columna + IntToStr( Fila + 1 ) + ':' + Columna + IntToStr( ActividadCount + Fila );
end;


//Martin
function TActividad.Delete( Fecha : TDate ):TActividadIndex;
begin
  if iHermano <> -1 then
  begin
    Connection.QryBusca.Active:=False;
    Connection.QryBusca.SQL.Text := 'select * from bitacoradeactividades '+
                                    'where sContrato = :contrato and ' +
                                    'dIdFecha = :fecha '+
                                    'and sIdTipoMovimiento = "ED" '+
                                    'and iHermano = :Hermano '+
                                    'and sHoraFinal = :HoraFinal';
    Connection.QryBusca.ParamByName('Contrato').AsString := Global_Contrato;
    Connection.QryBusca.ParamByName('Fecha').AsDate := Fecha;
    Connection.QryBusca.ParamByName('Hermano').AsInteger := iHermano;
    Connection.QryBusca.ParamByName('HoraFinal').AsString := sHInicio;
    Connection.QryBusca.Open;

    if Connection.QryBusca.RecordCount=1 then
    begin
      connection.zCommand.Active := False;
      connection.zCommand.SQL.text := 'update bitacoradeactividades '+
                                      'set sHoraFinal = :HoraFinal '+
                                      'where sContrato = :Contrato '+
                                      'and didfecha = :Fecha '+
                                      'and iIdDiario = :Diario '+
                                      'and iIdtarea = :tarea '+
                                      'and iIdActividad = :Actividad';
      Connection.zCommand.ParamByName('Contrato').AsString := Connection.QryBusca.FieldByName('sContrato').AsString;
      Connection.zCommand.ParamByName('Fecha').AsDate := Fecha;
      Connection.zCommand.ParamByName('Diario').AsInteger := Connection.QryBusca.FieldByName('iIdDiario').AsInteger;
      Connection.zCommand.ParamByName('tarea').AsInteger := Connection.QryBusca.FieldByName('iIdtarea').AsInteger;
      Connection.zCommand.ParamByName('Actividad').AsInteger := Connection.QryBusca.FieldByName('iIdActividad').AsInteger;
      Connection.zCommand.ParamByName('HoraFinal').AsString:= sHFin;
      Connection.zCommand.ExecSQL;
    end
    else
    begin
      Connection.QryBusca.Active:=False;
      Connection.QryBusca.SQL.Text := 'select * from bitacoradeactividades '+
                                      'where sContrato = :contrato and ' +
                                      'dIdFecha = :fecha '+
                                      'and sIdTipoMovimiento="ED" '+
                                      'and iHermano = :Hermano '+
                                      'and sHoraInicio=:HoraInicio';
      Connection.QryBusca.ParamByName('Contrato').AsString := global_Contrato;
      Connection.QryBusca.ParamByName('Fecha').AsDate := Fecha;
      Connection.QryBusca.ParamByName('Hermano').AsInteger := iHermano;
      Connection.QryBusca.ParamByName('HoraInicio').AsString := sHFin;
      Connection.QryBusca.Open;

      if Connection.QryBusca.RecordCount=1 then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.text := 'update bitacoradeactividades '+
                                        'set sHoraInicio=:HoraInicio where '+
                                        'sContrato = :Contrato '+
                                        'and didfecha = :Fecha '+
                                        'and iIdDiario = :Diario '+
                                        'and iIdtarea = :tarea '+
                                        'and iIdActividad = :Actividad';
        Connection.zCommand.ParamByName('Contrato').AsString := Connection.QryBusca.FieldByName('sContrato').AsString;
        Connection.zCommand.ParamByName('Fecha').AsDate := Fecha;
        Connection.zCommand.ParamByName('Diario').AsInteger := Connection.QryBusca.FieldByName('iIdDiario').AsInteger;
        Connection.zCommand.ParamByName('tarea').AsInteger := Connection.QryBusca.FieldByName('iIdtarea').AsInteger;
        Connection.zCommand.ParamByName('Actividad').AsInteger := Connection.QryBusca.FieldByName('iIdActividad').AsInteger;
        Connection.zCommand.ParamByName('HoraInicio').AsString := sHInicio;
        Connection.zCommand.ExecSQL;
      end
      else
      begin
        if Messagedlg('El Registro que sera eliminado cuenta con un Agrupador, pero no se encontro su Horario Sucesor ni Predecesor.'+#13 + #10+
                  'Desea Continuar?', mtConfirmation,[mbyes,Mbno],0) =MrNo then
        begin
          //connection.zconnection.startback;

          raise Exception.Create('notdel');
        end;
      end;
    end;
  end;

  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('DELETE FROM bitacoradepersonal WHERE sContrato = :contrato and ' +
    'dIdFecha = :fecha and iIdDiario = :diario and iIdTarea=:Tarea and iIdActividad=:ACtividad');
  Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
  Connection.zCommand.Params.ParamByName('contrato').value    := Global_Contrato;
  Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
  Connection.zCommand.Params.ParamByName('fecha').value       := Fecha;
  Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
  Connection.zCommand.Params.ParamByName('diario').value      := iIdDiario;
  Connection.zCommand.Params.ParamByName('Tarea').AsInteger      := iTarea;
  Connection.zCommand.Params.ParamByName('ACtividad').AsInteger      := iIdActividad;
  connection.zCommand.ExecSQL;

  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('DELETE FROM bitacoradeequipos WHERE sContrato = :contrato and ' +
    'dIdFecha = :fecha and iIdDiario = :diario and iIdTarea=:Tarea and iIdActividad=:ACtividad');
  Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
  Connection.zCommand.Params.ParamByName('contrato').value    := Global_Contrato;
  Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
  Connection.zCommand.Params.ParamByName('fecha').value       := Fecha;
  Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
  Connection.zCommand.Params.ParamByName('diario').value      := iIdDiario;
  Connection.zCommand.Params.ParamByName('Tarea').AsInteger      := iTarea;
  Connection.zCommand.Params.ParamByName('ACtividad').AsInteger      := iIdActividad;
  connection.zCommand.ExecSQL;

  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('DELETE FROM bitacoradeactividades WHERE sContrato = :contrato and ' +
    'dIdFecha = :fecha and iIdDiario = :diario and iIdTarea=:Tarea and iIdActividad=:ACtividad');
  Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
  Connection.zCommand.Params.ParamByName('contrato').value    := Global_Contrato;
  Connection.zCommand.Params.ParamByName('fecha').DataType    := ftDate;
  Connection.zCommand.Params.ParamByName('fecha').value       := Fecha;
  Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
  Connection.zCommand.Params.ParamByName('diario').value      := iIdDiario;
  Connection.zCommand.Params.ParamByName('Tarea').AsInteger      := iTarea;
  Connection.zCommand.Params.ParamByName('ACtividad').AsInteger      := iIdActividad;       
  connection.zCommand.ExecSQL;
end;


//Martin
function TCuadre.BuscaFolio( Hoja, Fila : integer ) : TFolioIndex;
var
  iFCount : Integer;
  eEncontrado : Boolean;
begin
  for iFCount := 0 to Length( CATEGORIA[Hoja] ) - 1 do
  begin
    if ( Fila >= CATEGORIA[Hoja, iFCount].iInicio - 5 ) and ( Fila <= CATEGORIA[Hoja, iFCount].iFin ) then
    begin
      eEncontrado := True;
      Break;
    end
    else
      eEncontrado := False;
  end;

  if eEncontrado then
    Result := iFCount
  else
    Result := -1;
end;


//Martin
function TCuadre.BuscarActividad( Hoja, Fila : integer ): TActividadIndex;
var
  iACount,
  iFIndex : Integer;
  eEncontrado : Boolean;
begin
  iFIndex := BuscaFolio(Hoja, Fila);
  if iFIndex >= 0 then
  begin
    for iACount := 0 to Length( CATEGORIA[Hoja, iFIndex].ACTIVIDADES ) - 1 do
    begin
      if CATEGORIA[Hoja, iFIndex].ACTIVIDADES[iACount].iRow = Fila then
      begin
        eEncontrado := True;
        Break;
      end
      else
        eEncontrado := False;
    end;
  end
  else
    eEncontrado := False;

  if eEncontrado then
    Result := iACount
  else
    Result := -1;
end;


//Martin
function TCuadre.GetIdActividad(Hoja: Integer; Fila: Integer):TIdActividad;
var
  Index : integer;
begin
  Index := BuscarActividad(Hoja, Fila);
  if Index >= 0 then
    Result := CATEGORIA[Hoja, BuscaFolio(Hoja, Fila)].ACTIVIDADES[Index].iIdActividad
  else
    Result := -1;
end;


//Martin
procedure TCuadre.SaveToExcel;
var
  Excel : TExcelInstance;

  IndexT : Integer;
  IndexF : TFolioIndex;
  IndexAP : TActividadPadreIndex;
  IndexA : TActividadIndex;

  Fila : TExcelFila;

const
  Color : array[ 0..1 ] of integer = ( 43, 44 );
  Color2  : array[ 0..1 ] of integer = ( 45, 46 );

begin
  Excel := CreateOleObject('Excel.Application');
  Excel.WorkBooks.Add;
  Excel.DisplayAlerts := False;
  Excel.ScreenUpdating := True;
  Excel.Visible :=  True;
  Fila := TExcelFila.Create;
  Fila.Row := 1;
  Excel.Columns['A:A'].ColumnWidth := 50;
  Excel.Workbooks[1].Sheets.Add;

  for IndexT := 0 to Length( CATEGORIA ) - 1 do
  begin
    Excel.Workbooks[1].Sheets[ IndexT + 1 ].Select;
    Excel.Workbooks[1].Sheets[ IndexT + 1 ].Name := TIPO[ IndexT];

    for IndexF := 0 to Length( CATEGORIA[ IndexT ] ) - 1 do
    begin
      Fila.Row := CATEGORIA[ IndexT, IndexF ].Fila;
      Excel.Range[ 'A'+Fila.sRow ].Value := CATEGORIA[ IndexT, IndexF ].sFolio;

      Fila.Row := CATEGORIA[ IndexT, IndexF ].iInicio;
      Excel.Range['A'+Fila.sRow+':A'+inttostr( CATEGORIA[ IndexT, IndexF ].iFin )].Interior.ColorIndex := color[ IndexF mod 2 ];

      for IndexA := 0 to Length( CATEGORIA[ IndexT, IndexF ].ACTIVIDADES ) - 1 do
      begin
        if CATEGORIA[ IndexT, IndexF ].ACTIVIDADES[ IndexA ].IsPadre then
        begin
          IndexAP := CATEGORIA[ IndexT, IndexF ].ACTIVIDADES[ IndexA ].IndexPadre;
          Fila.Row := CATEGORIA[ IndexT, IndexF ].ACTIVIDADES_PADRES[ IndexAP ].Fila;
          Excel.Range[ 'B'+Fila.sRow ].Value := CATEGORIA[ IndexT, IndexF ].ACTIVIDADES_PADRES[ IndexAP ].IdActividad;
          Excel.Range[ 'B'+Fila.sRow+':B'+IntToStr( Fila.Row + CATEGORIA[ IndexT, IndexF ].ACTIVIDADES_PADRES[ IndexAP ].ActividadCount ) ].Interior.ColorIndex := Color2[ IndexAP mod 2 ];
          Fila.Row := Fila.Row - 1;
        end
        else
        begin
          Fila.Row := CATEGORIA[ IndexT, IndexF ].ACTIVIDADES[ IndexA ].iRow;
          Excel.Range['C'+Fila.sRow].Value := CATEGORIA[ IndexT, IndexF ].ACTIVIDADES[ IndexA ].sIdActividad;
          Excel.Range['D'+Fila.sRow].Value := CATEGORIA[ IndexT, IndexF ].ACTIVIDADES[ IndexA ].sHInicio;
          Excel.Range['E'+Fila.sRow].Value := CATEGORIA[ IndexT, IndexF ].ACTIVIDADES[ IndexA ].sHFin;
          if CATEGORIA[ IndexT, IndexF ].ACTIVIDADES[ IndexA ].NuevoCorte then
            Excel.Range['F'+Fila.sRow].Interior.ColorIndex := 37;

        end;

      end;

    end;

  end;
end;


//Martin
procedure TCuadre.UpdateAllRanges;
var
  TIndex,
  FIndex : TFolioIndex;
begin

  for TIndex := 0 to Length( CATEGORIA ) - 1 do
  begin
    for FIndex := 0 to Length( CATEGORIA[ TIndex ] ) - 1 do
    begin
      CATEGORIA[ TIndex, FIndex ].UpdateRange;
      CATEGORIA[ TIndex, FIndex ].UpdateActRows;
    end;

  end;

  if Length( CATEGORIA ) > 0 then
  begin    
    Inicio := CATEGORIA[ 0, 0 ].iInicio;
    Fin := CATEGORIA[ 0, Length( CATEGORIA[ 0 ] ) - 1].iFin;
  end;
      
end;


//Martin
procedure TCuadre.RetrocederFoliosDesde(Indice: TFolioIndex);
var
  IndexT,
  IndexF : TFolioIndex;
begin
  for IndexT := 0 to Length( CATEGORIA ) - 1 do
  begin
    for IndexF := Indice to Length( CATEGORIA[ IndexT ] ) - 1 do
    begin
      CATEGORIA[ IndexT, IndexF ].Fila := CATEGORIA[ IndexT, IndexF ].Fila - 1;
      CATEGORIA[ IndexT, IndexF ].UpdateRange;
      CATEGORIA[ IndexT, IndexF ].ActualizarCountActividades;
      CATEGORIA[ IndexT, IndexF ].UpdateActRows;
    end;
      
  end;
  
end;


//Martin
procedure TCuadre.AvanzarFoliosDesde(Indice: TFolioIndex);
var
  IndexT,
  IndexF : TFolioIndex;
begin
  for IndexT := 0 to Length( CATEGORIA ) - 1 do
  begin
    for IndexF := Indice to Length( CATEGORIA[ IndexT ] ) - 1 do
    begin
      CATEGORIA[ IndexT, IndexF ].Fila := CATEGORIA[ IndexT, IndexF ].Fila + 1;
      CATEGORIA[ IndexT, IndexF ].UpdateActRows;
      CATEGORIA[ IndexT, IndexF ].UpdateRange;
    end;
      
  end;
  
end;

//Martin
function TFolio.BuscarActividad(Actividad: string) : Boolean;
var
  x : integer;
  Encontrado : Boolean;
begin
  for x := 0 to Length( ACTIVIDADES ) - 1 do
  begin
    if ACTIVIDADES[x].sIdActividad = Actividad then
    begin
      Encontrado := True;
      Break;
    end
    else
      Encontrado := False;
  end;

  Result := Encontrado;
end;


//Martin
function TFolio.BuscarActividadPorNodoIndex(NodoIndex: Integer) : TActividadIndex;
var
  index : Integer;
  encontrado : Boolean;
begin
  encontrado := False;
  for index := 0 to Length( ACTIVIDADES ) - 1 do
  begin
    if ACTIVIDADES[index].iNodoCorte = NodoIndex then
    begin
      encontrado := True;
      Break;
    end;
  end;

  if encontrado then
    Result := index
  else
    Result := -1;
end;


//Martin
function TFolio.BuscarCategoria(iColumn, Pos: Integer):TMoeIndex;
var
  iIndex : integer;
  eFound : Boolean;
begin
  for iIndex := 0 to Length( MOE ) - 1 do
  begin
    if ( iColumn >= MOE[iIndex].iInicio ) and ( iColumn <= MOE[iIndex].iFin ) then
    begin
      eFound := True;
      Break;
    end
    else
      eFound := False;
  end;

  if eFound then
    Result := iIndex
  else
    Result := -1;
end;


//Martin
function TFolio.BuscarCategoria(iColumn : Integer):TMoeIndex;
var
  iIndex : integer;
  eFound : Boolean;
begin
  for iIndex := 0 to Length( MOE ) - 1 do
  begin
    if ( MOE[iIndex].iCol = iColumn ) then
    begin
      eFound := True;
      Break;
    end
    else
      eFound := False;
  end;

  if eFound then
    Result := iIndex
  else
    Result := -1;
end;


//Martin
function TFolio.IndexPadre( Actividad : string ):TActividadPadreIndex;
var
  AIndex : TActividadPadreIndex;
  Found : Boolean;
begin
  Found := False;
  for AIndex := 0 to Length( ACTIVIDADES_PADRES ) - 1 do
  begin
    if ACTIVIDADES_PADRES[ AIndex ].IdActividad = Actividad then
    begin
      Found := True;
      Break;
    end;
  end;

  if Found then
    Result := AIndex
  else
    Result := -1;

end;


//Martin
procedure TFolio.CleanNodes();
var
  iCount : Integer;
begin
  for iCount := 0 to Length( ACTIVIDADES ) -1 do
  begin
    ACTIVIDADES[iCount].iNodoCorte := -1;
  end;
end;


(* Inicia formulario del cuadre *)


//Martin
procedure TfrmCuadreNormal.ComprobarSuma( Edit_Clean : Boolean );
var
  iValor,
  iFCount,
  iMoeIndex,
  Total, inicio : Integer;

  iSuma,
  iResta : Double;

  Rango : TcxSSCellObject;

begin
  if ePintando then
    Exit;

  try
    if qrActividades.RecordCount = 0 then
      Exit;

    Rango := Libro.ActiveSheet.GetCellObject( Libro.ActiveSheet.ActiveCell.X, Libro.ActiveSheet.ActiveCell.Y );

    if Edit_Clean then
    begin
      if Length( VarToStr( Rango.CellValue ) ) > 0 then
      begin
        try
          iValor := StrToInt( VarToStr( Rango.CellValue ) );
          iResta := 0;
        except
          raise Exception.Create('Valor invalido');
        end;
      end;
    end
    else
    begin
      Rango := Libro.ActiveSheet.GetCellObject( Libro.ActiveSheet.ActiveCell.X+1, Libro.ActiveSheet.ActiveCell.Y );
      iValor := 0;
      iResta := StrToFloat( VarToStr( Rango.CellValue ) );
    end;

    if iValor < 0 then
      raise Exception.Create('No se permiten valores negativos');

    if ( Libro.ActiveSheet.ActiveCell.X > 7 ) and ( Libro.ActiveSheet.ActiveCell.Y > 5 ) then
    begin   
        for inicio := Libro.ActiveSheet.SelectionRect.Left to Libro.ActiveSheet.SelectionRect.Right do
        begin
            if Libro.ActiveSheet.GetCellObject(inicio, 5).CellValue = 'CANT.' then
            begin
                Rango := Libro.ActiveSheet.GetCellObject( inicio, 0 );
                Rango.Text;
                iSuma := StrToFloat( VarToStr( Rango.CellValue ) );
                iSuma := iSuma - iResta;
                iMoeIndex := inicio + 1;

                if iMoeIndex = -1 then
                begin
                  raise Exception.Create('Error en la aplicación informe al administrador del sistema');
                  Exit;
                end;

                Cuadre.Cambios := True;
                Cuadre.Guardado := False;

                case CompareValue(iSuma, Libro.ActiveSheet.GetCellObject(inicio, 1).CellValue, 0.000002 ) of
                  LessThanValue    : Rango.Style.Brush.BackgroundColor := 41;
                  EqualsValue      : Rango.Style.Brush.BackgroundColor := 42;
                  GreaterThanValue : Rango.Style.Brush.BackgroundColor := 45;
                end;
            end;
        end;
    end;
  except
    on e:Exception do
    begin
      if ( e.Message = 'Valor invalido' ) or ( e.Message = 'No se permiten valores negativos' ) then
      begin
        Rango := Libro.ActiveSheet.GetCellObject( Libro.ActiveSheet.ActiveCell.X, Libro.ActiveSheet.ActiveCell.Y );
        Rango.Text := '';
      end;

      MessageDlg( e.Message, mtInformation, [mbOK], 0 );
    end;
  end;
end;

//Martin
procedure TFolio.AvanzarActividades( IndexFrom : Integer ; Increment : Integer = 1 );
var
  iACount : Integer;
begin
  for iACount := Length(ACTIVIDADES) - 1 downto IndexFrom do
  begin
    ACTIVIDADES[ iACount ].sWbs           := ACTIVIDADES[ iACount - 1 ].sWbs;
    ACTIVIDADES[ iACount ].sIdActividad   := ACTIVIDADES[ iACount - 1 ].sIdActividad;
    ACTIVIDADES[ iACount ].sHInicio       := ACTIVIDADES[ iACount - 1 ].sHInicio;
    ACTIVIDADES[ iACount ].sHFin          := ACTIVIDADES[ iACount - 1 ].sHFin;
    ACTIVIDADES[ iACount ].dDuracion      := ACTIVIDADES[ iACount - 1 ].dDuracion;
    ACTIVIDADES[ iACount ].iRow           := ACTIVIDADES[ iACount - 1 ].iRow + 1;
    ACTIVIDADES[ iACount ].iIdActividad   := ACTIVIDADES[ iACount - 1 ].iIdActividad;
    ACTIVIDADES[ iACount ].iIdDiario      := ACTIVIDADES[ iACount - 1 ].iIdDiario;
    ACTIVIDADES[ iACount ].iTarea         := ACTIVIDADES[ iACount - 1 ].iTarea;
    ACTIVIDADES[ iACount ].iNodoCorte     := ACTIVIDADES[ iACount - 1 ].iNodoCorte;
    ACTIVIDADES[ iACount ].iHermano       := ACTIVIDADES[ iACount - 1 ].iHermano;
    ACTIVIDADES[ iACount ].IsPadre        := ACTIVIDADES[ iACount - 1 ].IsPadre;
    ACTIVIDADES[ iACount ].iHermanosCount := ACTIVIDADES[ iACount - 1 ].iHermanosCount;
    ACTIVIDADES[ iACount ].IndexPadre     := ACTIVIDADES[ iACount - 1 ].IndexPadre;
    ACTIVIDADES[ iACount ].NuevoCorte     := ACTIVIDADES[ iACount - 1 ].NuevoCorte;
  end;
end;

//Martin
procedure TFolio.RetrocederActividades( IndexFrom : TActividadIndex );
var
  iACount : Integer;
begin
  for iACount := IndexFrom to Length(ACTIVIDADES) - 2 do
  begin
    ACTIVIDADES[ iACount ].sWbs           := ACTIVIDADES[iACount + 1].sWbs;
    ACTIVIDADES[ iACount ].sIdActividad   := ACTIVIDADES[iACount + 1].sIdActividad;
    ACTIVIDADES[ iACount ].sHInicio       := ACTIVIDADES[iACount + 1].sHInicio;
    ACTIVIDADES[ iACount ].sHFin          := ACTIVIDADES[iACount + 1].sHFin;
    ACTIVIDADES[ iACount ].dDuracion      := ACTIVIDADES[iACount + 1].dDuracion;
    ACTIVIDADES[ iACount ].iRow           := ACTIVIDADES[iACount + 1].iRow - 1;
    ACTIVIDADES[ iACount ].iIdActividad   := ACTIVIDADES[iACount + 1].iIdActividad;
    ACTIVIDADES[ iACount ].iIdDiario      := ACTIVIDADES[iACount + 1].iIdDiario;
    ACTIVIDADES[ iACount ].iTarea         := ACTIVIDADES[iACount + 1].iTarea;
    ACTIVIDADES[ iACount ].iNodoCorte     := ACTIVIDADES[iACount + 1].iNodoCorte;
    ACTIVIDADES[ iACount ].iHermano       := ACTIVIDADES[iACount + 1].iHermano;
    ACTIVIDADES[ iACount ].IsPadre        := ACTIVIDADES[iACount + 1].IsPadre;
    ACTIVIDADES[ iACount ].iHermanosCount := ACTIVIDADES[iACount + 1].iHermanosCount;
    ACTIVIDADES[ iACount ].IndexPadre     := ACTIVIDADES[iACount + 1].IndexPadre;
    ACTIVIDADES[ iACount ].NuevoCorte     := ACTIVIDADES[iACount + 1].NuevoCorte;
  end;
end;


//Martin
procedure TFolio.InsertarActividad(Actividad: TActividad; Indice: Integer);
begin
  ACTIVIDADES[Indice].sWbs := Actividad.sWbs;
  ACTIVIDADES[Indice].sIdActividad := Actividad.sIdActividad;
  ACTIVIDADES[Indice].sHInicio := Actividad.sHInicio;
  ACTIVIDADES[Indice].sHFin := Actividad.sHFin;
  ACTIVIDADES[Indice].dDuracion := Actividad.dDuracion;
  ACTIVIDADES[Indice].iRow := Actividad.iRow;
  ACTIVIDADES[Indice].iIdActividad := Actividad.iIdActividad;
  ACTIVIDADES[Indice].iIdDiario := Actividad.iIdDiario;
  ACTIVIDADES[Indice].iTarea := Actividad.iTarea;
  ACTIVIDADES[Indice].iNodoCorte := -1;
  ACTIVIDADES[Indice].iHermano := Actividad.iHermano;
  ACTIVIDADES[Indice].iInicioConjunto := Actividad.iInicioConjunto;
  ACTIVIDADES[Indice].IsPadre := Actividad.IsPadre;
  ACTIVIDADES[Indice].iHermanosCount := Actividad.iHermanosCount;
  ACTIVIDADES[Indice].IsPadre := Actividad.IsPadre;
  ACTIVIDADES[Indice].NuevoCorte := Actividad.NuevoCorte;

end;


procedure TfrmCuadreNormal.CortarestaActividad1Click(Sender: TObject);
var
  LocidActividad: Integer;
  LocHoraInit: Double;
  LocHoraFin: Double;
  locPartMinInit: Double;
  locPartMinFin: Double;
  ListaIdActividades: TStringList;
  Consulta: string;
  i: Integer;
begin
  try
    ListaIdActividades := TStringList.Create;
    try
      //localizar donde se está haciendo click (Actividad)
      LocidActividad := Cuadre.GetIdActividad( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );
      //Localizar en zQuery de Martin el idActividad seleccionado
      if (LocidActividad > -1) and (qrActividades.Locate('iidActividad', LocidActividad, [])) then
      begin
        //Asignar Valores a los text para que el usuario pendjo no se pierda y sepa que actividad etá cortando
        CxTextEdtFolio.Text := qrActividades.FieldByName('iidActividad').asString;
        cxMemoDescripcion.Text := '';
        cxMemoDescripcion.Text  :=  qrActividades.FieldByName('mDescripcion').AsString;
        CxTextEdthoraInicio.Text := qrActividades.FieldByName('sHoraInicio').AsString;
        CxTextEdtHoraTermino.Text := qrActividades.FieldByName('sHoraFinal').AsString;

        getHM(qrActividades.FieldByName('sHoraInicio').AsString, LocHoraInit, locPartMinInit);//Hora inicio
        getHM(qrActividades.FieldByName('sHoraFinal').AsString, LocHoraFin, locPartMinFin);//Hora Termino
        LocHoraInit := LocHoraInit * 60; //Multiplicar por 60 para convertir horas en minutos
        LocHoraFin := LocHoraFin * 60; //Multiplicar por 60 para convertir horas en minutos
        locPartMinInit := LocHoraInit + locPartMinInit; //Sumar los minutos restantes
        locPartMinFin := LocHoraFin + locPartMinFin; //Sumar los minutos restantes

        //Ahora toca convertir los minutos en formato decimal para poder
        //Crear una linea de numerica y poder partirla, posteriormente regresar a su formato original
        LocHoraInit := (locPartMinInit * 100)/60;
        LocHoraFin := (locPartMinFin * 100)/60;

        //Llenar el DataSet
        HorasToDecimal(qrActividades, CdTmpActividades, False);

        //Deetectar donde hay cruces de acuerdo al registro seleccionado
        ListaIdActividades := detectarCruces(CdTmpActividades,CdTmpActividades, qrActividades.FieldByName('iidActividad').AsInteger, LocHoraInit, LocHoraFin);

        //Despues de Ubicar los registros hay que regresar las horas a su formato orginal
        HorasToDecimal(TZReadOnlyQuery(cdTmpActividades), CdCortesSugeridos, true);

        //Crear el filtro en base a los id regresados en la lista anterior
        Consulta := '';
        for i := 0 to ListaIdActividades.Count - 1 do
        begin
          if i = ListaIdActividades.Count - 1 then
            Consulta := Consulta + 'idActividad = ' + QuotedStr(ListaIdActividades[i])
          else
            Consulta := Consulta + 'idActividad = ' + QuotedStr(ListaIdActividades[i]) + ' OR ';
        end;
        //Filtrar el dataset de acuerdo a los Id de Actividades cruzadas con la seleccionada
        CdCortesSugeridos.Filtered := False;
        CdCortesSugeridos.Filter := Consulta;
        CdCortesSugeridos.Filtered := True;

        //Hora de la verdad y ver los horarios que se cruzan de acuerdo a la actividad seleccionada
        try
          if Assigned(FindComponent('FrmCortes')) then
            gForm.Destroy;

          gForm := TForm.Create(self);
          gForm.Name := 'FrmCortes';
          gForm.Caption := 'Cortes personalizados';
          gForm.Position := poScreenCenter;
          gForm.BorderStyle := bsSingle;
          gForm.Width := 625;
          gForm.Height := 230;
          cxGBDatos.Parent := gForm;
          cxGBDatos.align := AlClient;
          cxGBDatos.Visible := True;
          gForm.ShowModal;
        finally
          cxGBDatos.Visible := False;
          cxGBDatos.Align := alNone;
          cxGBDatos.Parent := Self;
        end;
      end;
    finally
      CdCortesSugeridos.Filtered := False;
      ListaIdActividades.Destroy;
    end;
  except
    on e: Exception do
      MessageDlg('Ha Ocurrido un error inesperado, informar al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TfrmCuadreNormal.Cuadreen01Click(Sender: TObject);
begin
    If MessageDlg('Desea eliminar el Cuadre Actual de Personal y Equipo ?' , mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
        //Personal
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Text := 'delete from bitacoradepersonal where sContrato = :orden and didfecha =:fecha';
        connection.zCommand.ParamByName('orden').AsString := global_contrato;
        connection.zCommand.ParamByName('fecha').AsDate := StrToDate(cbbReportes.Text);
        connection.zCommand.ExecSQL;

        //Equipo
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Text := 'delete from bitacoradeequipos where sContrato = :orden and didfecha =:fecha';
        connection.zCommand.ParamByName('orden').AsString := global_contrato;
        connection.zCommand.ParamByName('fecha').AsDate := StrToDate(cbbReportes.Text);
        connection.zCommand.ExecSQL;

        cbbReportes.Properties.OnCloseUp(sender);
    end;
end;

procedure TfrmCuadreNormal.CuadreenCeros1Click(Sender: TObject);
var
    lExiste : boolean;
begin
    lCeros  := False;
    lExiste := False;

    if lExiste = False then
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Text := 'select sIdPersonal from  bitacoradepersonal where sContrato = :orden and didfecha =:fecha';
        connection.zCommand.ParamByName('orden').AsString := global_contrato;
        connection.zCommand.ParamByName('fecha').AsDate := StrToDate(cbbReportes.Text);
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
           lExiste := True;
    end;

    if lExiste = False then
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Text := 'select * from bitacoradeequipos where sContrato = :orden and didfecha =:fecha';
        connection.zCommand.ParamByName('orden').AsString := global_contrato;
        connection.zCommand.ParamByName('fecha').AsDate := StrToDate(cbbReportes.Text);
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
           lExiste := True;
    end;

    if lExiste then
    begin
        If MessageDlg('Desea Poner en 0 el Cuadre Actual de Personal y Equipo ?' , mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
            //Personal
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Text := 'Update bitacoradepersonal set dCantidad = 0, dCantHH = 0, dAjuste = 0 where sContrato = :orden and didfecha =:fecha';
            connection.zCommand.ParamByName('orden').AsString := global_contrato;
            connection.zCommand.ParamByName('fecha').AsDate := StrToDate(cbbReportes.Text);
            connection.zCommand.ExecSQL;

            //Equipo
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Text := 'Update bitacoradeequipos set dCantidad = 0, dCantHH = 0, dAjuste = 0 where sContrato = :orden and didfecha =:fecha';
            connection.zCommand.ParamByName('orden').AsString := global_contrato;
            connection.zCommand.ParamByName('fecha').AsDate := StrToDate(cbbReportes.Text);
            connection.zCommand.ExecSQL;

            cbbReportes.Properties.OnCloseUp(sender);
        end;
    end
    else
    begin
        lCeros := True;
        GuardaEnBD;
        cbbReportes.Properties.OnCloseUp(sender);
        lCeros := False;
    end;
end;

//Saul
procedure TfrmCuadreNormal.CxBtnCancelClick(Sender: TObject);
begin
  if Assigned(FindComponent('frmCortes')) then
    TForm(FindComponent('FrmCortes')).Close;
end;

procedure TfrmCuadreNormal.CxBtnCortarClick(Sender: TObject);
var
  HoraAnterior: Double;
  tuvoCortes: Boolean;
  noReg, indice: Integer;
begin
  //CdResult.FieldList.Text;

  noReg := -9;
  indice := 1;
  HoraAnterior := -9;
  tuvoCortes := False;

  //HorasToDecimal(qrActividades, CdPivote, false); //Convertir las horas a decimal
  //DefinirCortes(cdPivote, CdHorarios, ListaCampos); //DefinirCortes

  CdTmpActividades.EmptyDataSet;
  CdPivote.First;
  while Not CdPivote.Eof do
  begin
    CdHorarios.First;
    HoraAnterior := CdHorarios.FieldByName('Horario').AsFloat;
    while Not CdHorarios.Eof do
    begin
      noReg := -9;
      //Recorrer los horarios usando de pivote cdPivote
      if (CdHorarios.FieldByName('Horario').AsFloat > CdPivote.FieldByName('HoraInicio').AsFloat) and (CdHorarios.FieldByName('Horario').AsFloat <= CdPivote.FieldByName('HoraTermino').AsFloat) then
      begin
        //Si es inicio o en medio
        if CdHorarios.RecNo > 0 then
        begin
          if (CdHorarios.FieldByName('Horario').AsFloat <> CdPivote.FieldByName('HoraTermino').AsFloat) then
          begin
            CdTmpActividades.Append;
            CdTmpActividades.FieldByName('IdPadre').AsInteger := CdPivote.FieldByName('idActividad').AsInteger;
            CdTmpActividades.FieldByName('idActividad').AsInteger := CdPivote.FieldByName('idActividad').AsInteger;
            CdTmpActividades.FieldByName('Actividad').AsString := cdPivote.FieldByName('Actividad').AsString;
            if CdPivote.FieldByName('Cortes').AsBoolean = true then
              CdTmpActividades.FieldByName('HoraInicio').AsFloat := HoraAnterior
            else
              CdTmpActividades.FieldByName('HoraInicio').AsFloat := CdPivote.FieldByName('HoraInicio').AsFloat;
            CdTmpActividades.FieldByName('HoraTermino').AsFloat := CdHorarios.FieldByName('Horario').AsFloat;
            CdTmpActividades.FieldByName('Duracion').AsFloat := CdHorarios.FieldByName('Horario').AsFloat - HoraAnterior;
            CdTmpActividades.FieldByName('wbs').AsString := CdPivote.FieldByName('wbs').AsString;
            CdTmpActividades.FieldByName('NumeroOrden').AsString := CdPivote.FieldByName('NumeroOrden').AsString;
            CdTmpActividades.FieldByName('Descripcion').AsString := CdPivote.FieldByName('Descripcion').AsString;
            cdTmpActividades.FieldByName('Hermano').AsInteger := CdPivote.FieldbyName('Hermano').AsInteger;
            CdTmpActividades.FieldByName('idDiario').AsString := CdPivote.FieldByName('idDiario').AsString;

            CdTmpActividades.Post;

            noReg := CdPivote.RecNo;
            CdPivote.Edit;
            CdPivote.FieldByName('cortes').AsBoolean := True;
            CdPivote.Post;
            HoraAnterior := CdHorarios.FieldByName('Horario').AsFloat;
          end   //Si es el fin
          else if (CdHorarios.FieldByName('Horario').AsFloat = CdPivote.FieldByName('HoraTermino').AsFloat) and (CdPivote.FieldByName('cortes').AsBoolean = True)  then
          begin
            CdTmpActividades.Append;
            CdTmpActividades.FieldByName('idActividad').AsInteger := CdTmpActividades.RecordCount + indice;
            CdTmpActividades.FieldByName('IdPadre').AsInteger := CdPivote.FieldByName('idActividad').AsInteger;
            CdTmpActividades.FieldByName('Actividad').AsString := CdPivote.FieldByName('Actividad').AsString;
            CdTmpActividades.FieldByName('HoraInicio').AsFloat := HoraAnterior;
            CdTmpActividades.FieldByName('HoraTermino').AsFloat := CdHorarios.FieldByName('Horario').AsFloat;
            CdTmpActividades.FieldByName('Duracion').AsFloat := CdHorarios.FieldByName('Horario').AsFloat - HoraAnterior;
            CdTmpActividades.FieldByName('wbs').AsString := CdPivote.FieldByName('wbs').AsString;
            CdTmpActividades.FieldByName('NumeroOrden').AsString := CdPivote.FieldByName('NumeroOrden').AsString;
            CdTmpActividades.FieldByName('Descripcion').AsString := CdPivote.FieldByName('Descripcion').AsString;
            cdTmpActividades.FieldByName('Hermano').AsInteger := CdPivote.FieldByName('Hermano').AsInteger;
            CdTmpActividades.FieldByName('idDiario').AsString := CdPivote.FieldByName('idDiario').AsString;

            CdTmpActividades.Post;

            noReg := CdPivote.RecNo;
            CdPivote.Edit;
            CdPivote.FieldByName('cortes').AsBoolean := True;
            CdPivote.Post;
            HoraAnterior := CdHorarios.FieldByName('Horario').AsFloat;
          end;
        end;
      end;
      CdHorarios.Next;
    end;
    if noReg <> -9 then
      CdPivote.RecNo := noReg;
    CdPivote.Next;
  end;

  //Revisar los que tuvieron cortes y los que no
  //Insertar los que no tuvieron Cortes
  CdPivote.First;
  while Not CdPivote.Eof do
  begin
    if CdPivote.FieldByName('Cortes').AsBoolean = False then
    begin
      CdTmpActividades.Append;
      CdTmpActividades.FieldByName('idActividad').AsInteger := CdPivote.FieldByName('idActividad').AsInteger;
      CdTmpActividades.FieldByName('Actividad').AsString := CdPivote.FieldByName('Actividad').AsString;
      CdTmpActividades.FieldByName('HoraInicio').AsFloat := CdPivote.FieldByName('HoraInicio').AsInteger;
      CdTmpActividades.FieldByName('HoraTermino').AsFloat := CdPivote.FieldByName('HoraTermino').AsInteger;
      CdTmpActividades.FieldByName('Duracion').AsFloat := CdPivote.FieldByName('HoraTermino').AsFloat - CdPivote.FieldByName('HoraInicio').AsInteger;
      CdTmpActividades.FieldByName('wbs').AsString := CdPivote.FieldByName('wbs').AsString;
      CdTmpActividades.FieldByName('NumeroOrden').AsString := CdPivote.FieldByName('NumeroOrden').AsString;
      CdTmpActividades.FieldByName('idDiario').AsString := CdPivote.FieldByName('idDiario').AsString;
      CdTmpActividades.FieldByName('Descripcion').AsString := CdPivote.FieldByName('Descripcion').AsString;
      cdTmpActividades.FieldByName('Hermano').AsInteger := cdPivote.FieldByName('Hermano').AsInteger;
      CdTmpActividades.Post;
    end;
    CdPivote.Next;
  end;
  zQueryCopy(TzQuery(qrActividades), CdResult);
  //HorasToDecimal(TZReadOnlyQuery(cdtmpActividades), CdResult, True);
  //AQUI METE TU ZQUERY PINCHE MARTIN
  //TU ASS SAUL JAJAJA
  //zQueryCopy(zActividadesCortes, CdResult);
end;

procedure TfrmCuadreNormal.cxButton3Click(Sender: TObject);
var
  IndexF : TFolioIndex;
  Folio : string;

  sPernocta,
  sPlataforma : string;
begin
  IndexF := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );
  Folio := Cuadre.CATEGORIA[Libro.ActivePage, IndexF].sFolio;
  sPernocta := cbbPernoctas.EditValue;
  if Libro.ActivePage = 0 then
    sPlataforma := cbbPlataformas.EditValue
  else
    sPlataforma := '*';

  if ( Length( Trim( sPernocta ) ) > 0 ) and ( Length( Trim( sPlataforma ) ) > 0 ) then
  begin
    if IndexF >= 0 then
    begin
      if MessageDlg('¿Desea asignar la pernocta y plataforma seleccionadas a las categorias del folio:'+Folio+' ?', mtconfirmation, [mbyes, mbcancel], 0) = mryes then
      begin
        Cuadre.CATEGORIA[Libro.ActivePage, IndexF].SetPFMOE( cbbPernoctas.EditValue, cbbPlataformas.EditValue );
        MessageDlg('Listo!', mtconfirmation, [mbok], 0);
      end;

    end;
    
  end;

end;

procedure TfrmCuadreNormal.cxButton4Click(Sender: TObject);
var
  Folio : Integer;
begin

  Folio := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveCell.Y + 1 );

  if Folio >= 0 then
  begin

    with zAjuste do
    begin
      Active := False;
      SQL.Text := 'delete from bitacoradepernocta where sContrato = :orden and sNumeroOrden = :folio and dIdFecha = :fecha';
      ParamByName( 'orden' ).AsString := global_contrato;
      ParamByName( 'fecha' ).AsString := FormatDateTime( 'YYYY-MM-DD', global_fecha );
      ParamByName( 'folio' ).AsString := Cuadre.CATEGORIA[ Libro.ActivePage, Folio ].sFolio;
      ExecSQL;

      SQL.Text := 'insert into bitacoradepernocta values ( :orden, :fecha, :folio, 1, 0, :cantidad )';
      ParamByName( 'orden' ).AsString := global_contrato;
      ParamByName( 'fecha' ).AsString := FormatDateTime( 'YYYY-MM-DD', global_fecha );
      ParamByName( 'folio' ).AsString :=  Cuadre.CATEGORIA[ Libro.ActivePage, Folio ].sFolio;
      ParamByName( 'cantidad' ).AsString := clcAjuste.Text;
      ExecSQL;

    end;

  end;
end;

procedure TfrmCuadreNormal.cxCmdAplicarClick(Sender: TObject);
var
    fila, columna, total, indice, i : integer;
    sFolio, sDatos : string;
    Rango : Variant;
begin
   dxColores.Visible := False;

   {sDatos := libro.ActiveSheet.GetCellObject(10,5).CellValue;
   total := 10;
   //obtenermos por pestaña el total de columnas..
   while sDatos <> '' do
   begin
       sDatos := libro.ActiveSheet.GetCellObject(total,5).CellValue;
       inc(total);
   end;


   fila    := 6;
   qrActividades.First;
   while not qrActividades.Eof do
   begin
       if libro.ActiveSheet.GetCellObject(0,fila).CellValue = tsFolio.KeyValue then
       begin
           if qrActividades.FieldByName('sNumeroOrden').AsString = tsFolio.KeyValue then
           begin
               for i := 0 to total do
               begin
                   libro.ActiveSheet.GetCellObject(i,fila).Style.Brush.BackgroundColor := dxColores.ColorValue;
               end;
           end;
       end;
       inc(fila,1);
       qrActividades.Next;
   end;     }

   //Actualizamos el color en el folio de trabajo seleccionado.
   connection.zCommand.Active := False;
   connection.zCommand.SQL.Clear;
   connection.zCommand.SQL.Add('update bitacoradeactividades set sColor =:Color where sContrato =:Contrato and sNumeroOrden =:Folio and dIdFecha =:Fecha ');
   connection.zCommand.ParamByName('contrato').AsString := global_contrato;
   connection.zCommand.ParamByName('Folio').AsString    := tsFolio.KeyValue;
   connection.zCommand.ParamByName('Fecha').AsDate      := StrToDate(cbbReportes.Text);
   connection.zCommand.ParamByName('Color').AsString    := ColorToString(dxColores.ColorValue);
   connection.zCommand.ExecSQL;

   zqFolios.Refresh;
end;

procedure TfrmCuadreNormal.cxCmdColorClick(Sender: TObject);
begin
    if dxColores.Visible = False then
       dxColores.Visible := True
    else
       dxColores.Visible := False;
end;

procedure TfrmCuadreNormal.CxColumnSeleccionarPropertiesEditValueChanged(
  Sender: TObject);
begin
//  CdCortesSugeridos.Edit;
//  CdCortesSugeridos.FieldByName('Incluir').AsBoolean := not CdCortesSugeridos.FieldByName('Incluir').AsBoolean;
//  CdCortesSugeridos.Post;
end;

procedure TfrmCuadreNormal.cxNecClick(Sender: TObject);
begin

end;

//Saul
procedure TfrmCuadreNormal.DefinirCortes(DatasetOrigen: TClientDataSet;
  var DatasetDestino: TClientDataSet; var Campos: TStringList);
var
  Encontrado: Boolean;
  i: Integer;
function BuscarValor(Valor: Double): Boolean;
begin
  DatasetDestino.First;
  while not DatasetDestino.Eof do
  begin
    if DatasetDestino.FieldByName('Horario').AsFloat = Valor then
    begin
      Encontrado := True;
      Break;
    end
    else
      Encontrado := False;

    DatasetDestino.Next;
  end;
  Result := Encontrado;
end;
begin
  if (DatasetOrigen.Active) and (DatasetOrigen.RecordCount > 0) then
  begin
    DatasetDestino.EmptyDataset;
    DatasetOrigen.First;
    while Not DatasetOrigen.Eof do
    begin
      //Buscar duplicados
      for I := 0 to Campos.Count - 1 do
        if BuscarValor(DatasetOrigen.FieldByName(Campos[i]).AsFloat) = False then
        begin
          DatasetDestino.Append;
          DatasetDestino.FieldByName('Horario').AsFloat := DatasetOrigen.FieldByName(Campos[i]).AsFloat;
          DatasetDestino.Post;
        end;
      DatasetOrigen.Next;
    end;
    DatasetDestino.IndexFieldNames := 'horario';
  end;
end;



function TfrmCuadreNormal.detectarCruces(var cdDatosBuscar,
  cdDatosDestino: TClientDataSet; idActividad: Integer; HoraInicio,
  HoraTermino: Double): TStringList;
var
  Cursor: TCursor;
  Listita: TStringList;
begin
  try
    Listita := TStringList.Create;
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    try
      cdDatosBuscar.First;
      while not cdDatosBuscar.Eof do
      begin
        if cdDatosBuscar.FieldByName('idActividad').AsInteger <> idActividad then
          if
            (cdDatosBuscar.FieldByName('HoraInicio').AsFloat > HoraInicio) and (cdDatosBuscar.FieldByName('HoraInicio').AsFloat < HoraTermino) or
            (cdDatosBuscar.FieldByName('HoraTermino').AsFloat > HoraInicio) and (cdDatosBuscar.FieldByName('horaTermino').AsFloat < HoraTermino) then
          begin
            Listita.Add(cdDatosBuscar.FieldByName('idActividad').AsString);
          end;
        cdDatosBuscar.Next;
      end;
      Result := Listita;
    finally
      Screen.Cursor := Cursor;
    end;
  Except
    on e: exception do
      MessageDlg('Ha ocurrido un error inesperado, informar al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TfrmCuadreNormal.AbrirCuadreexistente1Click(Sender: TObject);
begin
  try
    if dlgOpenExcel.Execute() then
    begin
      Libro.LoadFromFile(dlgOpenExcel.FileName);
    end;
  except
    on e:Exception do
      MessageDlg('No se puede abrir el archivo especificado error: '+e.Message, mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmCuadreNormal.Ajustes1Click(Sender: TObject);
var
  Formulario : TForm ;
  Folio : Integer;
  ZCuentas : TZReadOnlyQuery;
begin

  Folio := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveCell.Y + 1 );

  if (folio >= 0) and ( Length( Cuadre.CATEGORIA[ Libro.ActivePage ] ) > 0 ) then
  begin
    try

      ZCuentas := TZReadOnlyQuery.Create( nil );
      ZCuentas.Connection := connection.zConnection;
      ZCuentas.Active := False;

      Formulario := TForm.Create( nil );
      Formulario.BorderStyle := bsDialog;
      Formulario.Caption := EmptyStr;
      Formulario.Width := 300;
      Formulario.Height := 300;
      Formulario.position := poScreenCenter;

      grpAjuste.Parent := Formulario;
      grpAjuste.Align := alClient;
      grpAjuste.Visible := True;

      ZCuentas.SQL.Text := 'select sIdPernocta from cuentas';
      ZCuentas.Open;
      ZCuentas.First;

      lstCuentas.Items.Clear;
      while not ZCuentas.Eof do
      begin
        lstCuentas.Items.Add( ZCuentas.FieldByName( 'sIdPernocta' ).AsString );
        ZCuentas.Next;
      end;

      Formulario.ShowModal;

      grpAjuste.Parent := Self;
      grpAjuste.Align := alNone;
      grpAjuste.Visible := False;
      grpAjuste.Left := 0;
      grpAjuste.Top := 0;
      grpAjuste.Width := 0;
      grpAjuste.Height := 0;

    finally
      Formulario.Free;
    end;
  end;
end;

procedure TfrmCuadreNormal.btnActividadesClick(Sender: TObject);
var
  Continuar : Boolean;
begin
  Continuar := True;
  if Cuadre.Cambios then
  begin
    if MessageDlg('No se han guardado cambios, ¿Desea continuar?', mtConfirmation, [mbYes, mbNo], 0) =  mrYes then
      Continuar := True
    else
      Continuar := False;
  end;

  if Continuar then
  begin
    ;
  end;
end;

procedure TfrmCuadreNormal.btnAplicaAjusteClick(Sender: TObject);
var
  ZAjuste : TZQuery;
  IdCuenta : string;
  FIndex : Integer;
begin

  try
    ZAjuste := TZQuery.Create( nil );
    ZAjuste.Connection := connection.zConnection;
    ZAjuste.Active := False;

    FIndex := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveCell.Y + 1 );

    if FIndex >= 0 then
    begin

      with ZAjuste do
      begin
        Active := False;
        SQL.Text := 'select '+
                      '(select count( sIdCuenta ) from bitacoradepernocta '+
                          'where '+
                            'sContrato = :orden and '+
                            'dIdFecha = :fecha and '+
                            'sNumeroOrden = :folio and '+
                            'sIdCuenta = (select sIdCuenta '+
                                         'from cuentas '+
                                         'where sIdPernocta = :pernocta ) limit 1 '+
                        ') as ExisteAjuste, '+
                      '(select sIdCuenta from cuentas where sIdPernocta = :pernocta ) as sIdCuenta;';
        ParamByName( 'orden' ).AsString := global_contrato;
        ParamByName( 'fecha' ).AsString := FormatDateTime( 'YYYY-MM-DD', StrToDate( Cuadre.Fecha ) );
        ParamByName( 'folio' ).AsString := Cuadre.CATEGORIA[ Libro.ActivePage, FIndex ].sFolio;
        ParamByName( 'pernocta' ).AsString := lstCuentas.Items[ lstCuentas.ItemIndex ];
        Open;

        IdCuenta := FieldByName( 'sIdCuenta' ).AsString;
      end;

      if ZAjuste.FieldByName( 'ExisteAjuste' ).AsInteger > 0 then
      begin
        with ZAjuste do
        begin
          Active := False;
          SQL.Text := 'update bitacoradepernocta set dCantidad = :ajuste '+
                      'where '+
                        'sContrato = :orden and '+
                        'dIdFecha = :fecha and '+
                        'sNumeroOrden = :folio and '+
                        'sIdCuenta = :cuenta; ';

        end;
      end
      else
      begin
        with ZAjuste do
        begin
          Active := False;
          SQL.Text := 'insert into bitacoradepernocta '+
                      'values( :orden, :fecha, :folio, :cuenta, 0, :ajuste ); ';

        end;  
      end;

      ZAjuste.ParamByName( 'ajuste' ) .AsFloat := clcAjuste.Value;
      ZAjuste.ParamByName( 'orden' ).AsString := global_contrato;
      ZAjuste.ParamByName( 'fecha' ).AsString := FormatDateTime( 'YYYY-MM-DD', StrToDate( Cuadre.Fecha ) );
      ZAjuste.ParamByName( 'folio' ).AsString := Cuadre.CATEGORIA[ Libro.ActivePage, FIndex ].sFolio; 
      ZAjuste.ParamByName( 'cuenta' ).AsString := IdCuenta;
      ZAjuste.ExecSQL;

      btnAplicaAjuste.Enabled := False;

    end;
    
  finally
    ZAjuste.Free;
  end;

end;

procedure TfrmCuadreNormal.btnCortarClick(Sender: TObject);
var
  indice,
  folio : Integer;
begin
  if not (lblCorteI.Caption <> '--:--' ) and ( lblCorteF.Caption <> '--:--' ) then
    Exit;

  DescripcionCorte := mDescripcionCorte.Lines.Text;

  folio := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );

  if folio = -1 then
    exit;

  indice := Cuadre.CATEGORIA[Libro.ActivePage, folio].BuscarActividadPorNodoIndex( lstCortes.Selected.Index );

  if indice = -1 then
    exit;

    
  if MessageDlg('¿Confirma la actualizacion de la Actividad y generar una nueva?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    CortarActividad( Cuadre.CATEGORIA[Libro.ActivePage, folio].ACTIVIDADES[indice].sIdActividad, lblInicio.Caption, lblFin.Caption, lblCorteI.Caption, lblCorteF.Caption, Cuadre.CATEGORIA[Libro.ActivePage, folio].ACTIVIDADES[indice].iRow );
    if not ErrorCorte then
      btnCancelarCorte.Click();
  end;
end;

procedure TfrmCuadreNormal.btnGuardarClick(Sender: TObject);
var
   Form : TForm;
   Respuesta : Boolean;
begin
   if ObtenerEstatusReporte(global_contrato, StrToDate(cbbReportes.Text)) <> 'Pendiente' then
   begin
       MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
       exit;
   end;

  if qrActividades.RecordCount > 0 then
  begin
    if Length( Trim( cbbReportes.Text ) ) > 0 then
    begin
      if cbbReportes.Text <> Cuadre.Fecha then
      begin
        if MessageDlg( 'La fecha especificada no coindide con la fecha generada del cuadre'+#10+
                    'Este cuadre se guardara en el dia: +'+ Cuadre.Fecha + ' Desea continuar?', mtConfirmation, [mbYes, mbNo], 0 ) = mrYes then
        begin
          try
            GuardaEnBD;
          except
          end;
        end;
      end
      else
      begin
        try
          GuardaEnBD;
        except
        end;
      end;
    end;
  end;
end;

procedure TfrmCuadreNormal.btnPintarClick(Sender: TObject);
var
  Pintar : Boolean;
begin
  try
    //connection.zconnection.startTransaction;
    try
      Pintar := true;
      if (Cuadre.Cambios) then
      begin
        if MessageDlg('No se han guardado cambios, ¿Desea continuar?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          Pintar := True;
        end
        else
          Pintar := False;
      end;

      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('select lOrdenaxHorario from reportediario where sContrato =:contrato and dIdFecha =:Fecha and sOrden =:Orden ');
      connection.zCommand.ParamByName('contrato').AsString := global_contrato_barco;
      connection.zCommand.ParamByName('Fecha').AsDate      := StrToDate(cbbReportes.Text);
      connection.zCommand.ParamByName('Orden').AsString    := global_contrato;
      connection.zCommand.Open;


      qrPernoctas.Active := False;
      qrPernoctas.Open;
      qrPlataformas.Active := False;
      qrPlataformas.Open;

      if ( qrPlataformas.RecordCount = 0 ) then
      begin
        Pintar := False;
        raise Exception.Create('No se encontraron plataformas');
      end;

      if ( qrPernoctas.RecordCount = 0 ) then
      begin
        Pintar := False;
        raise Exception.Create('No se encontraron pernoctas');
      end;

      if Pintar then
      begin
        if Length( Trim( cbbReportes.Text ) ) > 0 then
        begin

          btnPintar.Enabled := False;

          btnGuardar.Enabled := False;
          cbbReportes.Enabled := False;
          grp1.Enabled := False;
          Application.ProcessMessages;

          Libro.Enabled := False;
          Cuadre.Fecha := cbbReportes.Text;

          ConsultarActividades;

          connection.QryBusca.Active := False;
          connection.QryBusca.SQL.Text := sSQLMOE[0];
          connection.QryBusca.ParamByName('contrato').AsString := global_Contrato_Barco;
          connection.QryBusca.ParamByName('orden').AsString := global_contrato;
          connection.QryBusca.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
          connection.QryBusca.Open;

          connection.QryBusca2.Active := False;
          connection.QryBusca2.SQL.Text := sSQLMOE[1];
          connection.QryBusca2.ParamByName('contrato').AsString := global_Contrato_Barco;
          connection.QryBusca2.ParamByName('orden').AsString := global_contrato;
          connection.QryBusca2.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
          connection.QryBusca2.Open;

          if (connection.QryBusca.RecordCount = 0) and (connection.QryBusca2.RecordCount = 0) then
            raise Exception.Create('No hay MOE Solicitado para Personal y Equipo');

          if qrActividades.RecordCount > 0 then
          begin
            CxBtnCortar.Click;
            GenerarExcel(True, True);

            Libro.&Protected := False;
            Libro.LoadFromFile(sArchivo);
            Libro.&Protected := True;     //*  estaba protegido

            if FileExists( sArchivo ) then
              DeleteFile( sArchivo );

            Libro.Enabled := True;
          end
          else
            raise Exception.Create('No se encontraron Folios y Actividades en el dia especificado');

        end;

      end;
    except
      on e:Exception do
        MessageDlg( 'No se puede continuar ' + e.Message, mtInformation, [mbOK], 0 )
    end;
  finally
    btnPintar.Enabled := True;
    btnGuardar.Enabled := True;
    cbbReportes.Enabled := True;
    grp1.Enabled := True;
//    //connection.zConnection.Commit;
  end;
end;

procedure TfrmCuadreNormal.btnSaveClick(Sender: TObject);
begin
  if rbExcel.Checked then
  begin
    dlgExcel.FileName := 'Cuadre.xls';
    if dlgExcel.Execute() then
    begin
      libro.&Protected := False;
      Libro.SaveToFile(dlgExcel.FileName);
      Libro.&Protected := True;
      MessageDlg( 'Archivo guardado', mtInformation, [mbOK], 0 );
    end;
  end;
  if rbDatabase.Checked then
  begin
    
  end;
end;

procedure TfrmCuadreNormal.cbbAplicaPernoctaPropertiesChange(Sender: TObject);
var
   activeCelda: TcxSSCellObject;
begin
    if cbbAplicaPernocta.ItemIndex<>-1 then
    begin
      if Libro.ActiveSheet.ActiveCell.Y>5 then
      begin
          activeCelda := Libro.ActiveSheet.GetCellObject(3, Libro.ActiveSheet.ActiveCell.Y);
          if activeCelda.Text<>cbbAplicaPernocta.Text then
          begin
            activeCelda.Style.Locked := False;
            activeCelda.Text :=cbbAplicaPernocta.Text;
            activeCelda.Style.Locked := True;
          end;
       //   chkImprime.Font.Color := clRed;
      end;
    end;
end;


procedure TfrmCuadreNormal.cbbOtsPropertiesChange(Sender: TObject);
begin
  if Length( Trim( cbbOts.EditText ) ) > 0 then
  begin
    qrReportes.Active := False;
    qrReportes.ParamByName('orden').AsString := cbbOts.EditValue;
    qrReportes.Open;
  end;
end;

procedure TfrmCuadreNormal.cbbPernoctasPropertiesChange(Sender: TObject);
var
  IndexF,
  IndexM : Integer;
begin
  try
    if Length( Trim( cbbPernoctas.Text ) ) = 0 then
      raise Exception.Create('Clean');

    IndexF := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );
    if IndexF = -1 then
      raise Exception.Create('Clean');

    IndexM := Cuadre.CATEGORIA[Libro.ActivePage, IndexF].BuscarCategoria( Libro.ActiveSheet.ActiveCell.X + 1, 0 );

    if IndexM = -1 then
      raise Exception.Create('Clean');

    Cuadre.CATEGORIA[Libro.ActivePage, Indexf].MOE[IndexM].sPernocta := VarToStr( cbbPernoctas.EditValue );
  except
    on e:Exception do
    begin
      if e.Message = 'Clean' then
        cbbPernoctas.EditText := '';
    end;
  end;
end;

procedure TfrmCuadreNormal.cbbPlataformasPropertiesChange(Sender: TObject);
var
  indexf,
  indexm : Integer;
begin
  try
    if Length( Trim( cbbPlataformas.Text ) ) = 0 then
      raise Exception.Create('Clean');

    IndexF := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );
    if IndexF = -1 then
      raise Exception.Create('Clean');

    IndexM := Cuadre.CATEGORIA[Libro.ActivePage, IndexF].BuscarCategoria( Libro.ActiveSheet.ActiveCell.X + 1, 0 );

    if IndexM = -1 then
      raise Exception.Create('Clean');

    Cuadre.CATEGORIA[Libro.ActivePage, Indexf].MOE[IndexM].sPlataforma := VarToStr( cbbPlataformas.EditValue );
  except
    on e:Exception do
    begin
      if e.Message = 'Clean' then
        cbbPlataformas.EditText := '';
    end;
  end;
end;

procedure TfrmCuadreNormal.cbbReportesPropertiesCloseUp(Sender: TObject);
begin
  if Length( Trim( cbbReportes.EditText ) ) > 0 then
  begin
    qrFolios.Active := False;
    qrFolios.ParamByName('orden').AsString := global_contrato;
    try
      qrFolios.ParamByName('fecha').AsDate := StrToDate(cbbReportes.Text); 
    except
      ;
    end;
    qrFolios.Open;

    zqFolios.Active := False;
    zqFolios.ParamByName('orden').AsString := global_contrato;
    zqFolios.ParamByName('fecha').AsDate   := StrToDate(cbbReportes.Text);
    zqFolios.Open;

    tsFolio.KeyValue := zqFolios.FieldByName('sNumeroOrden').AsString;

    Application.ProcessMessages;

    btnPintar.Click;
  end;
end;

procedure TfrmCuadreNormal.cbbVistaPropertiesCloseUp(Sender: TObject);
begin
  if ( Cuadre.CUADRAR[ 0 ] ) or ( Cuadre.CUADRAR[ 1 ] ) then
  begin
    if Length( Cuadre.CATEGORIA ) > 0 then
    begin
      if MessageDlg( 'Desea cambiar el modo de vista a: ' + cbbVista.Text, mtConfirmation, [ mbYes, mbCancel ], 0 ) = mrYes then
      begin
        CambiaVistaCuadre( LowerCase( cbbVista.Text ) );
      end;
    end;
  end;
end;

procedure TfrmCuadreNormal.chkImprimeEnter(Sender: TObject);
var
   activeCelda: TcxSSCellObject;
begin
    if chkImprime.Checked then
    begin
        if Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X, 5).CellValue = 'CANT.' then
        begin
            activeCelda := Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X, QrActividades.RecordCount + 24);
            activeCelda.Style.Locked := False;
            activeCelda.Text := 'No';
            activeCelda.Style.Locked := True;
            chkImprime.Font.Color := clRed;
        end;
    end
    else
    begin
        if Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X, 5).CellValue = 'CANT.' then
        begin
            activeCelda := Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X, QrActividades.RecordCount + 24);
            activeCelda.Style.Locked := False;
            activeCelda.Text := 'Si';
            activeCelda.Style.Locked := True;
            chkImprime.Font.Color := clBlack;
        end;
    end;
end;

procedure TfrmCuadreNormal.clcAjustePropertiesChange(Sender: TObject);
begin
  btnAplicaAjuste.Enabled := True;
end;

procedure TfrmCuadreNormal.Diagrama();
var
  x, y : Integer;

  item : TListItem;
  Folio : TTreeNode;

  Form : TForm;

  IndexF : TFolioIndex;
  IndexAP : TActividadPadreIndex;
  IndexA : TActividadIndex;
begin
  lstEstructura.Items.Clear;
  if Length( Cuadre.CATEGORIA[Libro.ActivePage] ) > 0 then
  begin

    for IndexF := 0 to Length( Cuadre.CATEGORIA[ Libro.ActivePage ] ) - 1 do
    begin
      item := lstEstructura.Items.Add;
      item.Caption := Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].sFolio;
      item.SubItems.Add( ' : ' + IntToStr( Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].iInicio ) + '-' + IntToStr( Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].iFin ) );

      for IndexA := 0 to Length( Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES ) - 1 do
      begin
        if Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES[ IndexA ].IsPadre then
        begin
          IndexAP := Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES[ IndexA ].IndexPadre;
          item := lstEstructura.Items.Add;
          item.Caption := '';
          item.SubItems.Add( Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES_PADRES[ IndexAP ].IdActividad );
          item.SubItems.Add( IntToStr( Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES_PADRES[ IndexAP ].Fila ) );
          item.SubItems.Add( IntToStr( Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES_PADRES[ IndexAP ].ActividadCount ) );
        end
        else
        begin
          item := lstEstructura.Items.Add;
          item.Caption := '';
          item.SubItems.Add('');
          item.SubItems.Add(Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES[ IndexA ].sHInicio);
          item.SubItems.Add(Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES[ IndexA ].sHFin);
          item.SubItems.Add(IntToStr( Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES[ IndexA ].iRow ) );
        end;
        
      end;
        
    end;

    Form := TForm.Create(nil);
    Form.BorderStyle := bsDialog;
    form.Position := poScreenCenter;
    Form.BorderIcons := [];
    Form.Width := 710;
    form.Height := 635;
    Form.Caption := '';
    pnlEstructura.Parent := Form;
    pnlEstructura.Align := alClient;
    pnlEstructura.Visible := True;
    Form.ShowModal;
    pnlEstructura.Parent := frmCuadreNormal;
    pnlEstructura.Align := alNone;
    pnlEstructura.Width := 0;
    pnlEstructura.Height := 0;
    pnlEstructura.Left := 0;
    pnlEstructura.Top := 0;
    pnlEstructura.Visible := False;
    Form.Free;
  end;
end;

procedure TfrmCuadreNormal.dxColoresItemClick(Sender: TObject;
  AItem: TdxGalleryControlItem);
begin
   cxMuestraColor.StyleDisabled.Color := dxColores.ColorValue;
end;

procedure TfrmCuadreNormal.Foliosexistentesenelcuadre1Click(Sender: TObject);
begin
  Diagrama();
end;

procedure TfrmCuadreNormal.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmCuadreNormal.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  if (Cuadre.Cambios) then
  begin
    if MessageDlg('¿Hay cambios desea realmente salir?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      CanClose := True;
    end
    else
    begin
      CanClose := False;
    end;
  end
  else
  begin
    CanClose := True;
  end;
end;

procedure TfrmCuadreNormal.FormCreate(Sender: TObject);
var
  i, x, y, z : integer;
  Cursor: TCursor;
  sArchivo : string;
begin
  try
    connection.zConnection.TransactIsolationLevel := tiSerializable; 

    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    try
      GetTempPath(SizeOf(global_TempPath), global_TempPath);

      qrOT.Active := false;
      qrOT.Open;

      qrReportes.Active := False;
      qrReportes.ParamByName('orden').AsString := global_contrato;
      qrReportes.Open;

      Libro.Pages[0].Caption := 'Personal';
      Libro.Pages[1].Caption := 'Equipo';

      Cuadre := TCuadre.Create();

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

      CdTmpActividades.FieldDefs.Add('idDiario', ftInteger, 0, False);
      CdTmpActividades.FieldDefs.Add('idActividad', ftInteger, 0, True);
      CdTmpActividades.FieldDefs.Add('IdPadre', ftInteger, 0, False);
      CdTmpActividades.FieldDefs.Add('Wbs', ftString, 100, False);
      CdTmpActividades.FieldDefs.Add('NumeroOrden', ftString, 50, False);
      CdTmpActividades.FieldDefs.Add('Duracion', ftFloat, 0, True);
      CdTmpActividades.FieldDefs.Add('HoraInicio', ftFloat, 0, True);
      CdTmpActividades.FieldDefs.Add('HoraTermino', ftFloat, 0, True);
      CdTmpActividades.FieldDefs.Add('Actividad', ftString, 50, True);
      CdTmpActividades.FieldDefs.Add('Descripcion', ftMemo, 0, False);
      CdTmpActividades.FieldDefs.Add('Cortes', ftBoolean, 0, True);
      CdTmpActividades.FieldDefs.Add('Incluir', ftBoolean, 0, true);
      CdTmpActividades.FieldDefs.Add('Tarea', ftInteger, 0, True);
      CdTmpActividades.FieldDefs.Add('Hermano', ftInteger, 0, True);
      CdTmpActividades.FieldDefs.Add('Actividades', ftInteger, 0, True);
      CdTmpActividades.FieldDefs.Add('HermanosCount', ftInteger, 0, True);

      CdPivote.FieldDefs.Add('idDiario', ftInteger, 0, False);
      CdPivote.FieldDefs.Add('idActividad', ftInteger, 0, True);
      CdPivote.FieldDefs.Add('IdPadre', ftInteger, 0, False);
      cdPivote.FieldDefs.Add('Wbs', ftString, 100, False);
      CdPivote.FieldDefs.Add('NumeroOrden', ftString, 50, False);
      CdPivote.FieldDefs.Add('Duracion', ftFloat, 0, True);
      CdPivote.FieldDefs.Add('HoraInicio', ftFloat, 0, True);
      CdPivote.FieldDefs.Add('HoraTermino', ftFloat, 0, True);
      CdPivote.FieldDefs.Add('Actividad', ftString, 50, True);
      CdPivote.FieldDefs.Add('Descripcion', ftMemo, 0, False);
      CdPivote.FieldDefs.Add('Cortes', ftBoolean, 0, True);
      CdPivote.FieldDefs.Add('incluir', ftBoolean, 0, True);
      CdPivote.FieldDefs.Add('Tarea', ftInteger, 0, True);
      CdPivote.FieldDefs.Add('Hermano', ftInteger, 0, True);
      CdPivote.FieldDefs.Add('Actividades', ftInteger, 0, True);
      CdPivote.FieldDefs.Add('HermanosCount', ftInteger, 0, True);

      CdCortesSugeridos.FieldDefs.Add('idDiario', ftInteger, 0, False);
      CdCortesSugeridos.FieldDefs.Add('idActividad', ftInteger, 0, True);
      CdCortesSugeridos.FieldDefs.Add('IdPadre', ftInteger, 0, False);
      CdCortesSugeridos.FieldDefs.Add('Wbs', ftString, 100, False);
      CdCortesSugeridos.FieldDefs.Add('NumeroOrden', ftString, 50, False);
      CdCortesSugeridos.FieldDefs.Add('Duracion', ftFloat, 0, True);
      CdCortesSugeridos.FieldDefs.Add('HoraInicio', ftString, 5, True);
      CdCortesSugeridos.FieldDefs.Add('HoraTermino', ftString, 5, True);
      CdCortesSugeridos.FieldDefs.Add('Actividad', ftString, 50, True);
      CdCortesSugeridos.FieldDefs.Add('Descripcion', ftMemo, 0, False);
      CdCortesSugeridos.FieldDefs.Add('Cortes', ftBoolean, 0, True);
      CdCortesSugeridos.FieldDefs.Add('incluir', ftBoolean, 0, True);
      CdCortesSugeridos.FieldDefs.Add('Tarea', ftInteger, 0, True);
      CdCortesSugeridos.FieldDefs.Add('Hermano', ftInteger, 0, True);
      CdCortesSugeridos.FieldDefs.Add('Actividades', ftInteger, 0, True);
      CdCortesSugeridos.FieldDefs.Add('HermanosCount', ftInteger, 0, True);

      CdResult.FieldDefs.Add('idDiario', ftInteger, 0, False);
      CdResult.FieldDefs.Add('idActividad', ftInteger, 0, True);
      CdResult.FieldDefs.Add('IdPadre', ftInteger, 0, False);
      CdResult.FieldDefs.Add('Wbs', ftString, 100, False);
      CdResult.FieldDefs.Add('NumeroOrden', ftString, 50, False);
      CdResult.FieldDefs.Add('Duracion', ftFloat, 0, True);
      CdResult.FieldDefs.Add('HoraInicio', ftString, 5, True);
      CdResult.FieldDefs.Add('HoraTermino', ftString, 5, True);
      CdResult.FieldDefs.Add('Actividad', ftString, 50, True);
      CdResult.FieldDefs.Add('Descripcion', ftMemo, 0, False);
      CdResult.FieldDefs.Add('Cortes', ftBoolean, 0, True);
      CdResult.FieldDefs.Add('incluir', ftBoolean, 0, True);
      CdResult.FieldDefs.Add('Tarea', ftInteger, 0, True);
      CdResult.FieldDefs.Add('Hermano', ftInteger, 0, True);
      CdResult.FieldDefs.Add('Actividades', ftInteger, 0, True);
      CdResult.FieldDefs.Add('HermanosCount', ftInteger, 0, True);

      CdHorarios.FieldDefs.Add('Horario', ftFloat, 0, True);
      Cuadre.Cambios := False;
      Cuadre.Guardado := False;

      if FileExists(global_TempPath+'inteliCuad_intelcode.xls' ) then
        DeleteFile( global_TempPath+'inteliCuad_intelcode.xls' );

      if FileExists( global_TempPath+'inteliCuad_intelcode_Save.xls' ) then
        DeleteFile( global_TempPath+'inteliCuad_intelcode_Save.xls' );

      Self.Caption := 'Cuadre - ' + FormatDateTime( 'YYYY-MM-DD', global_fecha ) + global_estado_reporte ;

    finally
      Screen.Cursor := Cursor;
    end;
  except
    on e: Exception do
      MessageDlg('Ha ocurrido un error inesperado, informa al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TfrmCuadreNormal.FormShow(Sender: TObject);
begin
  dxColores.Visible := False;
  ListaCampos := TStringList.Create;
  ListaCampos.Add('HoraInicio');
  ListaCampos.Add('HoraTermino');
  CdTmpActividades.CreateDataSet;
  cdCortesSugeridos.CreateDataset;
  CdCortesSugeridos.Open;
  CdPivote.CreateDataSet;
  CdHorarios.CreateDataSet;
  CdResult.CreateDataSet;
  CdTmpActividades.Open;
  CdHorarios.Open;
  CdPivote.Open;
  CdResult.Open;

  try
    cbbReportes.EditValue := global_fecha;
  finally
    ;
  end;

  zqFolios.Active := False;
  zqFolios.ParamByName('orden').AsString := global_contrato;
  zqFolios.ParamByName('fecha').AsDate   := StrToDate(cbbReportes.Text);
  zqFolios.Open;
  tsFolio.KeyValue := zqFolios.FieldByName('sNumeroOrden').AsString;
end;

//Martin
procedure TfrmCuadreNormal.getHM(cadena: string; var h, m: Double);
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

procedure TfrmCuadreNormal.grp1Click(Sender: TObject);
begin

end;

//Saul
procedure TfrmCuadreNormal.HorasToDecimal(var DataActividades: TZReadOnlyQuery; var DataDestino: TClientDataSet; inverso: Boolean);
const
  HorasInicio: array[1..4] of String = ('00:00', '01:00', '05:00', '03:00');
  HorasFin: array[1..4] of String = ('12:00', '10:00', '11:00', '08:00');
var
  Cursor: TCursor;
  i: Integer;
  HoraI, MinutoI, HoraDecI: Double;
  HoraF, MinutoF, HoraDecF: Double;
  HoraiDex: Double;
  HoratDex: Double;
  HorasMinTotales: Double;
  TotalMinutos: Double;
function DexToHora(HoraDex: Double): string;
var
  partHoras: Double;
  PartMin: Double;
  HoraString: string;
  minString: string;
begin
  partHoras := 0;
  PartMin := 0;
  TotalMinutos := HoraDex;
  //HoraDex := (HoraDex/60)*100;
  HoraDex := horaDex/60;
  partHoras := Truncar(HoraDex, 0);
  HorasMinTotales := partHoras * 60;//Aqui obtengo el total de minutos por lkas horas obtenidas en partHoras
  PartMin := TotalMinutos - HorasMinTotales;
  //PartMin := HoraDex - partHoras;
  //PartMin := (PartMin / 60) * 100;
  Partmin := (partmin*100);
//  PartMin := Truncar(partmin*60/100,0);
  if partmin < 10 then
    minString := '0' + FloatToStr(partmin)
  else
    minString := FloatToStr(partmin);

  if partHoras < 10 then
    HoraString := '0' + FloatToStr(Truncar(partHoras,0)) + ':'  + minString
  else
    HoraString := FloatToStr(Truncar(partHoras, 0)) + ':'  + minString;

  Result := HoraString;
end;
begin
  try
    //Después de cargar las actividades hay que darle el formato en decimal
    DataDestino.EmptyDataSet;
    DataActividades.First;
    while Not DataActividades.Eof do
    begin
      HoraI := 0;
      HoraF := 0;
      MinutoI := 0;
      MinutoF := 0;


      if inverso = False then
      begin
        DataDestino.Append;
        //Convertir de Formato 00:00 a Decimal
        getHM(DataActividades.FieldByName('sHoraInicio').AsString, HoraI, MinutoI);//Hora inicio
        getHM(DataActividades.FieldByName('sHoraFinal').AsString, HoraF, MinutoF);//Hora Termino
        HoraI := HoraI * 60; //Multiplicar por 60 para convertir horas en minutos
        HoraF := HoraF * 60; //Multiplicar por 60 para convertir horas en minutos
        HoraI := HoraI + MinutoI; //Sumar los minutos restantes
        HoraF := HoraF + MinutoF; //Sumar los minutos restantes

        //Ahora toca convertir los minutos en formato decimal para poder
        //Crear una linea de numerica y poder partirla, posteriormente regresar a su formato original
        HoraDecI := (HoraI * 100)/60;
        HoraDecF := (HoraF * 100)/60;
        // Lo pongo así por la diferencia en nombre de los campos
        DataDestino.FieldByName('HoraInicio').AsFloat := HoraDecI;
        DataDestino.FieldByName('HoraTermino').AsFloat := HoraDecF;
        DataDestino.FieldByName('idDiario').AsInteger := DataActividades.FieldByName('iIdDiario').AsInteger;
        DataDestino.FieldByName('idActividad').AsInteger := DataActividades.FieldByName('iidactividad').AsInteger;
        DataDestino.FieldByName('wbs').asString := DataActividades.FieldByName('sWbs').AsString;
        DataDestino.FieldByName('IdPadre').AsInteger := DataActividades.FieldByName('iidactividad').AsInteger;
        DataDestino.FieldByName('Actividad').AsString := DataActividades.FieldByName('sNumeroActividad').AsString;
        DataDestino.FieldByName('NumeroOrden').AsString := DataActividades.FieldByName('sNumeroOrden').AsString;
        DataDestino.FieldByName('Descripcion').AsString := DataActividades.FieldByName('mDescripcion').AsString;
        DataDestino.FieldByName('Hermano').AsInteger := DataActividades.FieldByName('iHermano').AsInteger;
        DataDestino.FieldByName('Actividades').AsInteger := DataActividades.FieldByName('iActividades').AsInteger;
        DataDestino.FieldByName('HermanosCount').AsInteger := DataActividades.FieldByName('iHermanosCount').AsInteger;
        DataDestino.FieldByName('incluir').AsBoolean := False;
        DataDestino.FieldByName('Duracion').AsFloat := HoraDecF - HoraDecI;
        DataDestino.FieldByName('Cortes').AsBoolean := False;
        DataDestino.Post;
      end
      else
      begin
        DataDestino.Append;
        //Convertir de decimal a formato 00:00
        HOraiDex := DataActividades.FieldByName('HoraInicio').AsFloat;
        HoratDex := DataActividades.FieldByName('HoraTermino').AsFloat;
        HOraiDex := (HoraiDex *60/100);
        HoratDex := (HoratDex *60/100);
        DataDestino.FieldByName('HoraInicio').AsString := DexToHora(HoraiDex);
        DataDestino.FieldByName('HoraTermino').AsString := DexToHora(HoratDex);
        DataDestino.FieldByName('idDiario').AsInteger := DataActividades.FieldByName('IdDiario').AsInteger;
        DataDestino.FieldByName('idActividad').AsInteger := DataActividades.FieldByName('idactividad').AsInteger;
        DataDestino.FieldByName('wbs').asString := DataActividades.FieldByName('Wbs').AsString;
        DataDestino.FieldByName('IdPadre').AsInteger := DataActividades.FieldByName('idactividad').AsInteger;
        DataDestino.FieldByName('Actividad').AsString := DataActividades.FieldByName('Actividad').AsString;
        DataDestino.FieldByName('NumeroOrden').AsString := DataActividades.FieldByName('NumeroOrden').AsString;
        DataDestino.FieldByName('Descripcion').AsString := DataActividades.FieldByName('Descripcion').AsString;
        DataDestino.FieldByName('Hermano').AsInteger := DataActividades.FieldByName('Hermano').AsInteger;
        DataDestino.FieldByName('Actividades').AsInteger := DataActividades.FieldByName('Actividades').AsInteger;
        DataDestino.FieldByName('HermanosCount').AsInteger := DataActividades.FieldByName('HermanosCount').AsInteger;
        DataDestino.FieldByName('incluir').AsBoolean := DataActividades.FieldByName('Incluir').AsBoolean;
        DataDestino.FieldByName('Duracion').AsFloat := 0;
        DataDestino.FieldByName('Cortes').AsBoolean := False;
        DataDestino.Post;
      end;

      DataActividades.Next;
    end;
  except
    on e: Exception do
    begin
      if DataDestino.State in [dsInsert, dsEdit] then
        DataDestino.Cancel;
      MessageDlg('Ha ocurrido un error inesperado, informar al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
    end;
  end;
end;

procedure TfrmCuadreNormal.LibroClearCells(Sender: TcxSSBookSheet;
  const ACellRect: TRect; var UseDefaultStyle, CanClear: Boolean);
begin
  ComprobarSuma( False );
end;

procedure TfrmCuadreNormal.LibroEndEdit(Sender: TObject);
begin
  ComprobarSuma( True );
end;

procedure TfrmCuadreNormal.LibroSetSelection(Sender: TObject; ASheet: TcxSSBookSheet);
var
  activeCelda: TcxSSCellObject;

  iIndexf,
  iIndexm,
  iindexa,
  x, inicio   : integer;
begin
  cbbPlataformas.Visible := Libro.ActivePage = 0;
  cxLabel7.Visible := cbbPlataformas.Visible;

  if ( Libro.ActiveSheet.ActiveCell.X >= 0 ) and ( Libro.ActiveSheet.ActiveCell.Y > 4 ) then
  begin
    try
        if qrActividades.RecordCount > 0 then
        begin

          if Libro.ActiveSheet.GetCellObject(0,Libro.ActiveSheet.ActiveCell.y).CellValue = '' then
             iindexf := -1;
          if iindexf = -1 then
          begin
              tGrupo.Caption   := ' FOLIO ';
              Cxlbl6.Caption   := 'Descripción';
              CxNec.Caption    := 'NEC';
              CxInicio.Caption := 'INICIO';
              CxFin.Caption    := 'FIN';
              exit;
          end;

          try
            if Libro.ActiveSheet.GetCellObject(8,Libro.ActiveSheet.ActiveCell.y).CellValue = '' then
               iindexf := -1;
            if iindexa = -1 then
              exit;
          Except
          end;

          cxMemo1.Text := '';

          activeCelda := Libro.ActiveSheet.GetCellObject(0,Libro.ActiveSheet.ActiveCell.y);
          tsFolio.KeyValue := VarToStr( activeCelda.CellValue );

          activeCelda := Libro.ActiveSheet.GetCellObject(9,0);
          txtHH.Text  := VarToStr( activeCelda.CellValue );

          activeCelda := Libro.ActiveSheet.GetCellObject(8, Libro.ActiveSheet.ActiveCell.y);
          CxTextEdtActividad.Text := activeCelda.CellValue;

          activeCelda := Libro.ActiveSheet.GetCellObject(1, Libro.ActiveSheet.ActiveCell.y);
          cxMemo1.Text := activeCelda.CellValue;

          activeCelda := Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.x, 0);
          CxTextEdt1.Text := activeCelda.CellValue;

          activeCelda := Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.x, 1);
          CxTextEdtSolicitado.Text := activeCelda.CellValue;

           activeCelda := Libro.ActiveSheet.GetCellObject(3,Libro.ActiveSheet.ActiveCell.y);
          cbbAplicaPernocta.Text:=activeCelda.CellValue;


          if (Libro.ActiveSheet.ActiveCell.Y > 5) and (Libro.ActiveSheet.GetCellObject(0,Libro.ActiveSheet.ActiveCell.y).CellValue <> '' ) then
          begin
              activeCelda := Libro.ActiveSheet.GetCellObject(10, Libro.ActiveSheet.ActiveCell.y);
              Cxlbl6.Caption := 'Descripción:   '+FloatToStr(activeCelda.CellValue);

              activeCelda := Libro.ActiveSheet.GetCellObject(5, Libro.ActiveSheet.ActiveCell.y);
              CxNec.Caption := activeCelda.CellValue ;

              activeCelda := Libro.ActiveSheet.GetCellObject(7, Libro.ActiveSheet.ActiveCell.y);
              if VarToStr(activeCelda.CellValue) = '1' then
                 CxInicio.Caption := '24:00'
              else
                 CxInicio.Caption := formatDateTime('hh:mm',activeCelda.CellValue);

              activeCelda := Libro.ActiveSheet.GetCellObject(8, Libro.ActiveSheet.ActiveCell.y);
              if VarToStr(activeCelda.CellValue) = '1' then
                 CxFin.Caption := '24:00'
              else
                 CxFin.Caption := formatDateTime('hh:mm',activeCelda.CellValue);

              activeCelda := Libro.ActiveSheet.GetCellObject(0, Libro.ActiveSheet.ActiveCell.y);
              tGrupo.Caption := ' FOLIO: ' + activeCelda.CellValue + ' ';
          end
          else
          begin
              tGrupo.Caption   := ' FOLIO ';
              Cxlbl6.Caption    := 'Descripción';
              CxNec.Caption    := 'NEC';
              CxInicio.Caption := 'INICIO';
              CxFin.Caption    := 'FIN';
          end;


          if iIndexf >= 0 then
          begin
            CxLblFolio.Caption := Libro.ActiveSheet.GetCellObject(0,Libro.ActiveSheet.ActiveCell.y).CellValue;

            if qrMOE_Sol.RecordCount > 0 then
            begin
                if ( Libro.ActiveSheet.ActiveCell.X > 8 ) and ( Libro.ActiveSheet.ActiveCell.Y > 2 ) then
                begin
                  activeCelda := Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X,3);
                  CxLblCategoria.Caption := 'Categoria: '+ VarToStr(Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X, 4).CellValue) + ' - ' + VarToStr( activeCelda.CellValue );

    //              if Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X, 5).CellValue = 'CANT.' then
    //              begin
    //                  if Libro.ActiveSheet.ActiveCell.Y > 5 then
    //                  begin
    //                      if Libro.ActiveSheet.GetCellObject(0,Libro.ActiveSheet.ActiveCell.y).CellValue <> '' then
    //                      begin
    //                         activeCelda := Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X, Libro.ActiveSheet.ActiveCell.Y);
    //                         activeCelda.Style.Locked := False;
    //                      end;
    //                  end;
    //              end;

                  //Estos es para desactivar las celdas de la columna completa x la longitud de la selecicon.
                  for inicio := Libro.ActiveSheet.SelectionRect.Left to Libro.ActiveSheet.SelectionRect.Right do
                  begin
                      if Libro.ActiveSheet.GetCellObject(inicio, 5).CellValue = 'CANT.' then
                      begin
                          if Libro.ActiveSheet.ActiveCell.Y > 5 then
                          begin
                              for x := 1 to qrActividades.RecordCount   do
                              begin
                                  if Libro.ActiveSheet.GetCellObject(0, x + 5).CellValue <> '' then
                                  begin
                                     activeCelda := Libro.ActiveSheet.GetCellObject(inicio, x + 5);
                                     activeCelda.Style.Locked := False;
                                  end;
                              end;
                          end;
                      end;
                  end;

                 chkImprime.Checked := False;
                 chkImprime.Font.Color := clBlack;
                 if (Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X, 5).CellValue = 'CANT.') then
                    if Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.X,qrActividades.RecordCount + 24).CellValue = 'Si' then
                       chkImprime.Checked := True
                    else
                       chkImprime.Font.Color := clRed;

                  cbbPernoctas.EditValue   := VarToStr(Libro.ActiveSheet.GetCellObject(2, Libro.ActiveSheet.ActiveCell.y).CellValue);
                  cbbPlataformas.EditValue := VarToStr(Libro.ActiveSheet.GetCellObject(4, Libro.ActiveSheet.ActiveCell.y).CellValue);
                end
                else
                  CxLblCategoria.Caption := '';
            end;
          end
          else
            CxLblFolio.Caption := '';
        end;

    Except
    end;

  end
  else
  begin
    cbbPernoctas.EditText := '';
    cbbPlataformas.EditText := '';
    CxTextEdtSolicitado.Text := '';
    CxTextEdtaBordo.Text := '';
    txtHH.Text := '';
  end;
end;

procedure TfrmCuadreNormal.lstCortesChange(Sender: TObject; Node: TTreeNode);
var
  indexa : Integer;
  Celda : TcxSSCellObject;
begin
  if Trim( Node.Text ) <> 'Actividad'  then
  begin


    indexa := Cuadre.CATEGORIA[Libro.ActivePage, Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 )].BuscarActividadPorNodoIndex( Node.Index );

    if indexa >= 0 then
    begin
      Celda := Libro.ActiveSheet.GetCellObject(1,  Cuadre.CATEGORIA[Libro.ActivePage, Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 )].ACTIVIDADES[indexa].iRow - 1);
      mDescripcionCorte.Text := Celda.CellValue;

      sAInicio := Cuadre.CATEGORIA[ Libro.ActivePage, Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 )].ACTIVIDADES[indexa].sHInicio;
      sAFin := Cuadre.CATEGORIA[ Libro.ActivePage, Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 )].ACTIVIDADES[indexa].sHFin;
      lblInicio.Caption := sAInicio;
      lblFin.Caption := sAFin;
      tmRestaHoras.Text := sAFin;
    end;
  end
  else
  begin
    lblInicio.Caption := '--:--';
    lblFin.Caption := '--:--';
    mDescripcionCorte.Text := '';
  end;

end;

procedure TfrmCuadreNormal.lstCuentasClick(Sender: TObject);
var
  ZAjuste : TZReadOnlyQuery;
  FIndex : Integer;
begin
  FIndex := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveCell.Y + 1 );

  if FIndex >= 0 then
  begin
    ZAjuste := TZReadOnlyQuery.Create( nil );
    ZAjuste.Connection := connection.zConnection;
    ZAjuste.Active := False;
    ZAjuste.SQL.Text := 'select dCantidad from bitacoradepernocta '+
                       'where '+
                          'sContrato = :orden and '+
                          'dIdFecha = :fecha and '+
                          'sNumeroOrden = :folio and '+
                          'sIdCuenta = (select sIdCuenta from cuentas where sIdPernocta = :pernocta);';

    ZAjuste.ParamByName( 'orden' ).AsString := global_contrato;
    ZAjuste.ParamByName( 'fecha' ).AsString := FormatDateTime( 'YYYY-MM-DD', StrToDate( Cuadre.Fecha ) );
    ZAjuste.ParamByName( 'folio' ).AsString := Cuadre.CATEGORIA[ Libro.ActivePage, FIndex ].sFolio;
    ZAjuste.ParamByName( 'pernocta' ).AsString := lstCuentas.Items[ lstCuentas.ItemIndex ];
    ZAjuste.Open;

    clcAjuste.Value := ZAjuste.FieldByName( 'dCantidad' ).AsFloat;
    btnAplicaAjuste.Enabled := False;
  end;
end;

procedure TfrmCuadreNormal.popHojaPopup(Sender: TObject);
begin
  Foliosexistentesenelcuadre1.Enabled := Length( Cuadre.CATEGORIA[Libro.ActivePage] ) > 0
end;

procedure TfrmCuadreNormal.Eliminarestaactividad1Click(Sender: TObject);
var
  Continuar : Boolean;

  IFIndex : TFolioIndex;
  IAIndex : TActividadIndex;
  IAPIndex : TActividadPadreIndex;
begin
  Continuar := True;
  if Cuadre.Cambios then
  begin
    if MessageDlg('No se han guardado cambios, ¿Desea continuar?', mtInformation, [mbYes, mbNo], 0) = mrYes then
      Continuar := True
    else
      Continuar := False;      
  end;

  if Continuar then
  begin
    IFIndex := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );

    if IFIndex >= 0 then
    begin
      IAIndex := Cuadre.BuscarActividad( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );
      IAPIndex := Cuadre.CATEGORIA[ Libro.ActivePage, IFIndex ].IndexPadre( Cuadre.CATEGORIA[Libro.ActivePage, IFIndex].ACTIVIDADES[IAIndex].sIdActividad );
      
      if IAIndex >= 0 then
      begin
        try
          if IAPIndex = -1 then
            raise Exception.Create( 'No se ha encontrado la actividad padre.' ); 

          with Cuadre.CATEGORIA[Libro.ActivePage, IFIndex].ACTIVIDADES[IAIndex] do
          begin
            if IsPadre then
              raise Exception.Create( 'Accion invalida a un agrupador' );

            if ( iIdActividad = iHermano ) or ( iHermano = -1 ) then
              raise Exception.Create('No se puede eliminar el horario original');
          end;
          
          if MessageDlg( '¿Desea eliminar el horario?', mtConfirmation, [mbYes, mbCancel], 0 ) = mrYes then
          begin
              //connection.zconnection.startTransaction;
              Cuadre.CATEGORIA[Libro.ActivePage, iFIndex].ACTIVIDADES[IAIndex].Delete( StrToDate(Cuadre.Fecha) );
              EliminarHorario( Cuadre.CATEGORIA[Libro.ActivePage,iFIndex].ACTIVIDADES[IAIndex].iRow,
                               Cuadre.CATEGORIA[Libro.ActivePage, iFIndex].ACTIVIDADES[IAIndex].sHFin,
                               IFIndex,
                               IAIndex,
                               IAPIndex );
              //connection.zConnection.Commit;
              Cuadre.UpdateAllRanges;
          end;
        except
          on e:Exception do
          begin
            if e.Message <> 'notdel' then
            begin
              if connection.zConnection.InTransaction then
                //connection.zconnection.startback;
              MessageDlg('Ha ocurrido un error al eliminar el horario:'+#10+e.Message, mtError, [mbOK], 0);
            end;

          end;

        end;

      end;

    end;

  end;

end;

procedure TfrmCuadreNormal.Escribe;
var
  iFila,
  iColumna,
  iCol,
  iFCount,
  iACount,
  iMCount,
  iPCount,
  iJornadas : Integer;

  sFactorHoras,
  sFolio : string;

  iHoras,
  iMinutos,
  dDuracion : Double;

  Rect : TRect;
  Rango : TcxSSCellObject;

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
  ePintando := True;
  Libro.&Protected := False;
  Libro.Visible := False;

  with Libro do
  begin
    for iPCount := 0 to PageCount - 1 do
    begin
      Caption := TIPO[ iPCount ];
      qrMOE_Sol.Active := False;
      if iPCount = 0 then
      begin
        qrMOE_Sol.SQL.Text := 'select mr.*, p.sIdTipoPersonal, p.iItemOrden '+
                              'from moerecursos mr '+
                              'inner join moe m '+
                              '  on ( m.sContrato = :orden '+
                              '    and mr.iIdMoe = m.iIdMoe ) '+#10+

                              'inner join personal p '+
                              '  on (  p.sContrato = :contrato '+
                              '    and mr.eTipoRecurso = "Personal" '+
                              '    and p.sIdPersonal = mr.sIdRecurso ) '+#10+

                              'where m.dIdFecha = ( select max( m1.dIdFecha ) '+
                              '                    from moe m1 '+
                              '                    where m1.sContrato = :orden '+
                              '                    and m1.dIdFecha <= :fecha )';
      end
      else
      begin
        qrMOE_Sol.SQL.Text := 'select mr.*, e.iItemOrden '+
                              'from moerecursos mr '+
                              'inner join moe m '+
                              '  on ( m.sContrato = :orden '+
                              '    and mr.iIdMoe = m.iIdMoe ) '+#10+

                              'inner join equipos e '+
                              '  on (  e.sContrato = :contrato '+
                              '    and mr.eTipoRecurso = "Equipo" '+
                              '    and e.sIdEquipo = mr.sIdRecurso ) '+#10+

                              'where m.dIdFecha = ( select max( m1.dIdFecha ) '+
                              '                    from moe m1 '+
                              '                    where m1.sContrato = :orden '+
                              '                    and m1.dIdFecha <= :fecha )';
      end;

      qrMOE_Sol.ParamByName('orden').AsString := global_contrato;
      qrMOE_Sol.ParamByName('contrato').AsString := global_Contrato_Barco;

      try
        qrMOE_Sol.ParamByName('fecha').AsDate := StrToDate( cbbReportes.EditText );
      except
        ;
      end;
      qrMOE_Sol.Open;

      iFila := 3;
      qrFolios.First;
      Pages[iPCount].ClearAll;
      SetLength( Cuadre.CATEGORIA[iPCount], qrFolios.RecordCount );
      iFCount := 0;

      lblEstado.Caption := 'Cargando ' + TIPO[iPCount];
      prgFolios.Maximum := qrFolios.RecordCount;
      prgFolios.Position := 0;
      Application.ProcessMessages;

      while not qrFolios.Eof do
      begin
        iJornadas := qrFolios.FieldByName('ijornadas').AsInteger;
        Cuadre.CATEGORIA[iPCount, iFCount] := TFolio.Create();
        Cuadre.CATEGORIA[iPCount, iFCount].sFolio := qrFolios.FieldByName('snumeroorden').asstring;
        SetLength( Cuadre.CATEGORIA[iPCount, iFCount].MOE, qrMOE_Sol.RecordCount );

        iColumna := 4;
        qrMOE_Sol.First;
        Pages[iPCount].Rows.Size[ifila - 2] := 35;
        Pages[iPCount].Rows.Size[iFila - 1] := 35;
        Pages[iPCount].Rows.Size[ifila - 0] := 75;

        Rect.Left := 0;
        Rect.Top := ifila + 1;
        Rect.Right := 3;
        Rect.Bottom := iFila + 1;
        Pages[iPCount].SetMergedState(Rect, True);

        Rango := Pages[iPCount].GetCellObject(0, iFila + 1);
        Rango.Style.Locked := True;
        Rango.Text := qrFolios.FieldByName('sNumeroOrden').AsString;
        Rango.Style.Brush.BackgroundColor := 43;
        Rango.Style.HorzTextAlign := haCENTER;

        iMCount := 0;
        while not qrMOE_Sol.Eof do
        begin
          {CANTIDAD SOLITADA}
          Rect.Left := iColumna;
          Rect.Top := iFila - 2;
          Rect.Right := iColumna + 1;
          Rect.Bottom := iFila - 2;
          Pages[iPCount].SetMergedState(Rect, True);

          Rango := Pages[iPCount].GetCellObject(iColumna, iFila - 2);
          Rango.Text := qrMOE_Sol.FieldByName('dCantidad').AsString;
          Rango.Style.HorzTextAlign := haCENTER;
          Rango.Style.WordBreak := True;
          Rango.Style.Brush.BackgroundColor := 47;

          {DESCRIPCION DEL RECURSO}
          Rect.Left := iColumna;
          Rect.Top := iFila;
          Rect.Right := iColumna + 1;
          Rect.Bottom := iFila;
          Pages[iPCount].SetMergedState(Rect, True);

          Rango := Pages[iPCount].GetCellObject(iColumna, iFila);
          Rango.Text := qrMOE_Sol.FieldByName('sDescripcion').AsString;
          Rango.Style.HorzTextAlign := haCENTER;
          Rango.Style.WordBreak := True;
          Rango.Style.Brush.BackgroundColor := 47;

          {ID DEL RECURSO}
          Rect.Left := iColumna;
          Rect.Top := iFila + 1;
          Rect.Right := iColumna + 1;
          Rect.Bottom := iFila + 1;
          Pages[iPCount].SetMergedState(Rect, True);

          Rango := Pages[iPCount].GetCellObject(iColumna, iFila + 1);
          Rango.Text := qrMOE_Sol.FieldByName('sIdRecurso').AsString;
          Rango.Style.HorzTextAlign := haCENTER;
          Rango.Style.WordBreak := True;
          Rango.Style.Brush.BackgroundColor := 47;

          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount] := TCategoria.Create();
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sIdRecurso := qrMOE_Sol.FieldByName('sidrecurso').AsString;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iSolicitado := qrMOE_Sol.FieldByName('dcantidad').AsInteger;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iAbordo := qrMOE_Sol.FieldByName('dcantidad').AsInteger;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iCol := iColumna;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sSuma := COLUMNAS[iColumna]+inttostr( iFila - 3 );
          Inc( iMCount );
          Inc(iColumna, 2);
          qrMOE_Sol.Next;
        end;

        CdResult.Filtered := False;
        CdResult.Filter := 'NumeroOrden = ' + QuotedStr(qrFolios.FieldByName('snumeroorden').AsString) ;
        CdResult.Filtered := True;

        Inc(ifila, 2);
        iColumna := 0;
        Rango := Pages[iPCount].GetCellObject(iColumna, iFila);
        Rango.Text := 'ACTIVIDAD';
        Rango.Style.HorzTextAlign := haCENTER;
        Rango.Style.WordBreak := True;
        Rango.Style.Brush.BackgroundColor := 43;
        Inc(iColumna);

        Rango := Pages[iPCount].GetCellObject(iColumna, iFila);
        Rango.Text := 'INICIA';
        Rango.Style.HorzTextAlign := haCENTER;
        Rango.Style.WordBreak := True;
        Rango.Style.Brush.BackgroundColor := 43;
        Inc(iColumna);

        Rango := Pages[iPCount].GetCellObject(iColumna, iFila);
        Rango.Text := 'FINALIZA';
        Rango.Style.HorzTextAlign := haCENTER;
        Rango.Style.WordBreak := True;
        Rango.Style.Brush.BackgroundColor := 43;
        Inc(iColumna);

        Rango := Pages[iPCount].GetCellObject(iColumna, iFila);
        Rango.Text := 'DURACION';
        Rango.Style.HorzTextAlign := haCENTER;
        Rango.Style.WordBreak := True;
        Rango.Style.Brush.BackgroundColor := 43;
        Inc(iColumna);

        Inc(iFila);

        Cuadre.CATEGORIA[iPCount, iFCount].iInicio := iFila+1;
        SetLength( Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES, CdResult.RecordCount );
        iACount := 0;

        prgActividades.Maximum := CdResult.RecordCount;
        prgActividades.Position := 0;
        CdResult.First;

        while not CdResult.Eof do
        begin
          iColumna := 0;

          Rango := Pages[iPCount].GetCellObject(iColumna, iFila);
          Rango.Text := CdResult.FieldByName('Actividad').AsString;

          Inc(iColumna);

          Rango := Pages[iPCount].GetCellObject(iColumna, iFila);
          Rango.Style.Format := $23;
          Rango.Text := CdResult.FieldByName('HoraInicio').AsString;
          Rango.Style.HorzTextAlign := haCENTER;
          Inc(iColumna);

          Rango := Pages[iPCount].GetCellObject(iColumna, iFila);
          Rango.Style.Format := $23;
          Rango.Text := CdResult.FieldByName('HoraTermino').AsString;
          Rango.Style.HorzTextAlign := haCENTER;
          Inc(iColumna);

          sFactorHoras := sfnRestaHoras(CdResult.FieldByName('HoraTermino').AsString, CdResult.FieldByName('HoraInicio').AsString);
          getHM(sFactorHoras, iHoras, iMinutos);
          iHoras := iHoras / 24;
          iMinutos := ( iMinutos / 24 ) / 60;

          Rango := Pages[iPCount].GetCellObject(iColumna, iFila);

          if rbRedondear.Checked then
            dDuracion := ( Redondear((iHoras + iminutos), Round( clcRT.Value ) ) );
          if rbTruncar.Checked then
            dDuracion := ( Truncar((iHoras + iminutos), Round( clcRT.Value ) ) );
          Rango.Style.HorzTextAlign := haCENTER;
          Rango.Text := FloatToStr( dDuracion );

          iCol := 4;
          for iColumna := 3 to qrMOE_Sol.RecordCount + 2 do
          begin
            Rango := Pages[iPCount].GetCellObject(iCol, iFila);
            Rango.Style.Locked := False;

            Rango := Pages[iPCount].GetCellObject(iCol+1, iFila);
            if iPCount = 0 then
               Rango.Text := '=('+COLUMNAS[icol+1]+inttostr(ifila + 1)+'*D'+inttostr(ifila + 1)+')*'+ IntToStr( iJornadas )
            else
              Rango.Text := '=('+COLUMNAS[icol+1]+inttostr(ifila + 1)+'*D'+inttostr(ifila + 1)+')';
            inc(icol, 2);
          end;

          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount] := TActividad.Create();
          with Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount] do
          begin
            iIdActividad := CdResult.FieldByName('idactividad').AsInteger;
            sIdActividad := CdResult.FieldByName('actividad').AsString;
            sWbs         := CdResult.FieldByName('wbs').AsString;
            sHInicio     := CdResult.FieldByName('horainicio').AsString;
            sHFin        := CdResult.FieldByName('horatermino').AsString;
            iIdDiario    := CdResult.FieldByName('iddiario').AsInteger;
            dDuracion    := dDuracion;
            iRow         := iFila + 1;
          end;

          Inc( iACount );
          Inc( iFila );
          CdResult.Next;

          prgActividades.Position := prgActividades.Position + 2;
          prgActividades.Refresh;
        end;

        CdResult.Filtered := False;

        Cuadre.CATEGORIA[iPCount, iFCount].iFin := iFila;

        Inc(iFila, 6);
        qrFolios.Next;
        Inc(iFCount);
      end;

      {SUMAS}
      for iFCount := 0 to Length( Cuadre.CATEGORIA[iPCount] ) - 1 do
      begin
        qrMOE_Sol.First;
        iFila := Cuadre.CATEGORIA[iPCount, iFCount].iInicio - 7;
        iCol := 4;
        while not qrMOE_Sol.Eof do
        begin
          Rect.Left := iCol;
          Rect.Right := iCol + 1;
          Rect.Top := iFila;
          Rect.Bottom := iFila;

          Pages[iPCount].SetMergedState(Rect, True);
          Rango := Pages[iPCount].GetCellObject(iCol, iFila);
          Rango.Text := '=SUM('+COLUMNAS[iCol+2]+inttostr( Cuadre.CATEGORIA[iPCount, iFCount].iInicio)+':'+COLUMNAS[iCol+2]+inttostr( Cuadre.CATEGORIA[iPCount, iFCount].iFin)+')';
          Rango.Style.HorzTextAlign := haCENTER;
          Rango.Style.Brush.BackgroundColor := 41;
          qrMOE_Sol.Next;
          Inc(icol, 2);
        end;
      end;

      prgFolios.Position := prgFolios.Position + 2;
      prgFolios.Refresh;
    end;
  end;

  Libro.&Protected := True;
  ePintando := False;
  Libro.Visible := True;
  Libro.ActivePage := 0;
  lblEstado.Caption := 'Listo';
end;

procedure TfrmCuadreNormal.ExportarCuadre1Click(Sender: TObject);

var
  Excel,
  Book,
  Hoja,
  Rango : TExcelInstance;

  iHoja : Integer;

begin
 if dlgExcel.Execute then
  begin
    {$REGION 'Inicia'}

    sArchivo := dlgExcel.FileName;
    if not Length( Cuadre.CATEGORIA ) > 0 then
      Exit;

    if FileExists( sArchivo ) then
    begin
      DeleteFile( sArchivo );
    end;
    Libro.SaveToFile(sArchivo);
  
    try
      Excel := CreateOleObject('Excel.Application');
      Book := Excel.Workbooks.Open( sArchivo );
      Hoja := Book.Sheets[1];
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
      Excel.Visible := False;
    except
      MessageDlg('No se puede continuar verifique tener instalada la suite de Microsoft Office', mtInformation, [mbOK], 0);
      Exit;
    end;

    {$ENDREGION}

    for iHoja := 1 to 2 do
    begin
      Book.Sheets[ iHoja ].Select;
      Excel.ActiveSheet.Unprotect;
    end;
    Book.Sheets[ 1 ].Select;
    Book.Save;

  end;

end;

procedure TfrmCuadreNormal.ExportarCuadrevirtualaExcel1Click(Sender: TObject);
begin
  Cuadre.SaveToExcel;
end;

function TfrmCuadreNormal.ColumnaNombre(Numero: Integer): String;
Var
  Valor, NumLetras: Integer;
  Cad: String;
Begin
  NumLetras := 26;
  Cad := '';
  Valor := Numero Mod NumLetras;
  if Valor = 0 then
    Valor := 26;

  if Numero - Valor > 0 then
    Cad := Char( 64 + Trunc( ( Numero - Valor) / NumLetras) );

  Cad := Cad + Char(64 + Valor);

  Result := Cad;
End;

function TfrmCuadreNormal.sfnRestaHoras(sParamHorasMax, sParamHorasMin: string): string;
var
  nHorasMax, nMinutosMax: Real;
  nHorasMin, nMinutosMin: Real;
  nHorasResult, nMinutosResult: Real;
  sHoras, sMinutos: string;
begin
  sParamHorasMax := Trim(sParamHorasMax);
  sParamHorasMin := Trim(sParamHorasMin);

  nHorasMax := rfnDecimal(MidStr(sParamHorasMax, 1, 2));
  nHorasMin := rfnDecimal(MidStr(sParamHorasMin, 1, 2));

  nMinutosMax := rfnDecimal(MidStr(sParamHorasMax, 4, 2));
  nMinutosMin := rfnDecimal(MidStr(sParamHorasMin, 4, 2));

  if nMinutosMax >= nMinutosMin then
    nMinutosResult := nMinutosMax - nMinutosMin
  else
  begin
    nHorasMax := nHorasMax - 1;
    nMinutosResult := (60 + nMinutosMax) - nMinutosMin;
  end;

  nHorasResult := nHorasMax - nHorasMin;

  Str(nHorasResult: 2: 0, sHoras);
  sHoras := Trim(sHoras);
  if nHorasResult >= 10 then
    sHoras := sHoras + ':'
  else
    sHoras := '0' + sHoras + ':';

  Str(nMinutosResult: 2: 0, sMinutos);
  sMinutos := Trim(sMinutos);
  if nMinutosResult >= 10 then
    sfnRestaHoras := sHoras + sMinutos
  else
    sfnRestaHoras := sHoras + '0' + sMinutos;
end;

function TfrmCuadreNormal.sfnSumaHoras(sParamHorasMax, sParamHorasMin: string): string;
var
  nHorasMax, nMinutosMax: Real;
  nHorasMin, nMinutosMin: Real;
  nHorasResult, nMinutosResult: Real;
  sHoras, sMinutos: string;
begin
  sParamHorasMax := Trim(sParamHorasMax);
  sParamHorasMin := Trim(sParamHorasMin);
  nHorasMax := rfnDecimal(MidStr(sParamHorasMax, 1, Pos(':', sParamHorasMax) - 1));
  nMinutosMax := rfnDecimal(MidStr(sParamHorasMax, Pos(':', sParamHorasMax) + 1, 2));

  nHorasMin := rfnDecimal(MidStr(sParamHorasMin, 1, 2));
  nMinutosMin := rfnDecimal(MidStr(sParamHorasMin, 4, 2));

  nMinutosResult := nMinutosMax + nMinutosMin;
  nHorasResult := nHorasMax + nHorasMin;

  if nMinutosResult >= 60 then
  begin
    nHorasResult := nHorasResult + 1;
    nMinutosResult := nMinutosResult - 60;
  end;

  Str(nHorasResult: 2: 0, sHoras);
  sHoras := Trim(sHoras);
  if nHorasResult >= 10 then
    sHoras := sHoras + ':'
  else
    sHoras := '0' + sHoras + ':';

  Str(nMinutosResult: 2: 0, sMinutos);
  sMinutos := Trim(sMinutos);
  if nMinutosResult >= 10 then
    sfnSumaHoras := sHoras + sMinutos
  else
    sfnSumaHoras := sHoras + '0' + sMinutos;
end;

procedure TfrmCuadreNormal.tmInicioPropertiesEditValueChanged(Sender: TObject);
begin
  lblResultado.Caption := 'El resultado de la actividad será: ';
end;

procedure TfrmCuadreNormal.tmRestaHorasPropertiesChange(Sender: TObject);
begin
  if Trim( lstCortes.Selected.Text ) <> 'Actividad' then
  begin
    try
      if ( StrToTime( tmRestaHoras.Text ) < StrToTime( sAInicio ) ) and ( StrToTime( tmRestaHoras.Text ) > StrToTime( sAFin ) ) then
      begin
        tmRestaHoras.Text := sOLdCorte;;
      end;
    except
      tmRestaHoras.Text := sOLdCorte;
    end;
    
    lblFin.Caption := tmRestaHoras.Text;    
    if tmRestaHoras.Text = '00:00' then
    begin
      lblCorteI.Caption := '--:--';
      lblCorteF.Caption := '--:--';
    end
    else
    begin
      lblCorteI.Caption := lblFin.Caption;
      lblCorteF.Caption := sAFin;
    end;

    sOLdCorte := tmRestaHoras.Text;
  end;
end;

function TfrmCuadreNormal.rfnDecimal(sParamCantidad: string): Real;
var
  Code: Integer;
  Resultado: Real;
begin
  Val(sParamCantidad, Resultado, Code);
  if Code <> 0 then
    Resultado := 0;
  rfnDecimal := Resultado;
end;

function TfrmCuadreNormal.Redondear(numero : real ; cifrasSig : integer) : real;
var
  p10 : extended;
begin
  if (cifrasSig = 2) then
    result := round(numero * 100) / 100
  else
  begin
    p10 := Power(10, cifrasSig);
    result := round(numero * p10) / p10;
  end;
end;

function TfrmCuadreNormal.Truncar(numero: Real; cifras: Integer) : Real;
var
  x, y : Integer;
  cadena, cad : string;
begin
  cadena := FloatToStr( numero );
  cad := '';
  for x := 1 to Length( cadena ) do
  begin
    if cadena[x] = '.' then
      Break
    else
      cad := cad + cadena[x];
  end;

  cad := cad + '.';

  for y := x+1 to x+1+cifras - 1 do
  begin
    cad := cad + cadena[y];
  end;

  Result := StrToFloat( cad );
end;

procedure TfrmCuadreNormal.tsFolioClick(Sender: TObject);
begin
   if zqFolios.FieldByName('sColor').AsString <> '' then
   begin
       zqFolios.Locate('sNumeroOrden', tsFolio.Text, []);
       cxMuestraColor.StyleDisabled.Color := StringToColor(zqFolios.FieldByName('sColor').AsString);
       dxColores.ColorValue := StringToColor(zqFolios.FieldByName('sColor').AsString);
   end
   else
   begin
      cxMuestraColor.Style.Color := 2;
      cxMuestraColor.StyleDisabled.Color := 2;
   end;
end;

procedure TfrmCuadreNormal.tsFolioExit(Sender: TObject);
begin
   if zqFolios.FieldByName('sColor').AsString <> '' then
   begin
       zqFolios.Locate('sNumeroOrden', tsFolio.Text, []);
       cxMuestraColor.StyleDisabled.Color := StringToColor(zqFolios.FieldByName('sColor').AsString);
       dxColores.ColorValue := StringToColor(zqFolios.FieldByName('sColor').AsString);
   end
   else
   begin
       cxMuestraColor.Style.Color := 2;
       cxMuestraColor.StyleDisabled.Color := 2;
   end;
end;

procedure TfrmCuadreNormal.zqFoliosAfterScroll(DataSet: TDataSet);
begin
    if zqFolios.FieldByName('sColor').AsString <> '' then
    begin
       dxColores.ColorValue := StringToColor(zqFolios.FieldByName('sColor').AsString);
       cxMuestraColor.StyleDisabled.Color := dxColores.ColorValue;
    end
    else
    begin
       cxMuestraColor.Style.Color := 2;
       cxMuestraColor.StyleDisabled.Color := 2;
    end;
end;

procedure TfrmCuadreNormal.zQueryCopy(var ZDataset: TZQuery;
  var cdDataset: TClientDataSet);
begin
  try
    cdDataset.EmptyDataSet;
    ZDataset.First;
    while not ZDataset.Eof do
    begin
      cdDataset.Append;
      cdDataset.FieldByName('HoraInicio').asString := ZDataset.FieldByName('sHoraInicio').asString;
      cdDataset.FieldByName('HoraTermino').asString := ZDataset.FieldByName('sHoraFinal').asString;
      cdDataset.FieldByName('idDiario').asString := ZDataset.FieldByName('iIdDiario').asString;
      cdDataset.FieldByName('idActividad').asString := ZDataset.FieldByName('iidactividad').asString;
      cdDataset.FieldByName('wbs').asString := ZDataset.FieldByName('sWbs').AsString;
      cdDataset.FieldByName('IdPadre').asString := ZDataset.FieldByName('iidactividad').asString;
      cdDataset.FieldByName('Actividad').AsString := ZDataset.FieldByName('sNumeroActividad').AsString;
      cdDataset.FieldByName('NumeroOrden').AsString := ZDataset.FieldByName('sNumeroOrden').AsString;
      cdDataset.FieldByName('Tarea').AsInteger := ZDataset.FieldByName('iIdTarea').AsInteger;
      cdDataset.FieldByName('Descripcion').AsString := ZDataset.FieldByName('mDescripcion').AsString;
      cdDataset.FieldByName('Hermano').asInteger := ZDataSet.FieldByName('iHermano').AsInteger;
      cdDataset.FieldByName('Actividades').asInteger := ZDataSet.FieldByName('iActividades').AsInteger;
      cdDataset.FieldByName('HermanosCount').asInteger := ZDataSet.FieldByName('iHermanosCount').AsInteger;
      cdDataset.FieldByName('incluir').AsBoolean := False;
      cdDataset.FieldByName('Duracion').AsFloat := 0;
      cdDataset.FieldByName('Cortes').AsBoolean := False;
      cdDataset.Post;
      ZDataset.Next;
    end;

  except
    on e: Exception do
    begin
      If ZDataset.State in [dsInsert, dsEdit] then
        ZDataset.Cancel;
      ZDataset.EmptyDataSet;
      MessageDlg('Ha ocurrido un error inesperado, informar al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
    end;
  end;
end;

//Martin
procedure TfrmCuadreNormal.GenerarCorte1Click(Sender: TObject);
var
  Celda : TcxSSCellObject;
  IFIndex,
  IAIndex : Integer;
  Continuar : Boolean;
begin
  chkAnida.Checked := True;
  if Length( Cuadre.CATEGORIA[Libro.ActivePage] ) > 0 then
  begin
    Continuar := True;
    if Cuadre.Cambios then
    begin
      if MessageDlg('No se han guardado cambios, ¿Desea continuar?', mtInformation, [mbYes, mbNo], 0) = mrYes then
        Continuar := True
      else
        Continuar := False;
    end;

    if Continuar then
    begin
      IFIndex := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );

      if IFIndex >= 0 then
      begin
        Cuadre.CATEGORIA[Libro.ActivePage, IFIndex].CleanNodes();
        IAIndex := Cuadre.BuscarActividad( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );

        if IAIndex >= 0 then
        begin
          CargarActividades_ListView( lstCortes, IFIndex, Cuadre.CATEGORIA[Libro.ActivePage, iFIndex].ACTIVIDADES[IAindex].sIdActividad );
          lstCortes.Items.Item[ Cuadre.CATEGORIA[Libro.ActivePage, IFIndex].ACTIVIDADES[IAIndex].iNodoCorte + 1].Selected := True;
          VentanaCortes();
        end;
      end;
    end;
  end;
end;

procedure TfrmCuadreNormal.GenerarExcel(Personal: Boolean; Equipo: Boolean);
var
  Excel,
  Libro,
  Hoja,
  Rango : Variant;

  iPCount : TTipoRecursoIndex;
  iFCount : TFolioIndex;
  iMCount : TCategoriaIndex;
  iACount : TActividadIndex;
  IAPCount : TActividadPadreIndex;
  iHoja,
  iPages,
  iFila,
  iColumna,
  iJornadas,
  iCol,
  iActividad,
  iInicio,
  iFin,
  iSumaCount,
  iIndexH,
  iFilaH,
  iIndexRH,
  Actividades,
  iRegistros : Integer;

  sFactorHoras,
  sFolio,
  sSumaFolios,
  ActividadAnterior : string;

  iHoras,
  iMinutos,
  dDuracion : Double;

  iSuma,
  iResta : double;

  Cursor : TCursor;

  zqExiste : TZReadOnlyQuery;

  lEncabezado : Boolean;

  function Fecha( Cadena : string ):string;
  var
    x : Integer;
    cad : string;
  begin
    cad := '';
    cad := cad + Copy(Cadena, 7, 10) + '-';
    cad := cad + Copy(Cadena, 4, 2) + '-';
    cad := cad + Copy(Cadena, 1, 2);

    Result := cad;
  end;

  {$REGION 'Consulta Personal y Equipo Existente'}
const

  sSQLEXISTE : array[0..1] of string = ('select ot.sNumeroOrden, '+
                               'ba.iIdActividad, '+
                               'ba.iIdDiario, '+
                               'ba.sNumeroActividad, '+
                               'ba.mDescripcion, '+
                               'mr.sIdRecurso, '+
                               'p.sDescripcion, '+
                               'bp.dCantidad, '+
                               'bp.dCantHH, '+
                               'bp.sIdPernocta, '+
                               'bp.sIdPlataforma, '+
                               'bp.lImprime '+ #10 +

                        'from bitacoradepersonal bp '+
                        'inner join bitacoradeactividades ba '+
                          'on ( bp.sContrato = ba.sContrato '+
                            'and bp.iIdActividad = ba.iIdActividad '+
                            'and bp.sNumeroActividad = ba.sNumeroActividad '+
                            'and bp.iIdDiario = ba.iIdDiario '+
                            'and bp.sNumeroOrden = ba.sNumeroOrden '+
                            'and ba.sIdTipoMovimiento = "ED" '+

                          ') '+ #10 +

                        'inner join personal p '+
                          'on ( p.sContrato = :contrato '+
                            'and p.sIdPersonal = bp.sIdPersonal '+
                          ') '+ #10 +

                        'inner join moerecursos mr '+
                          'on ( mr.iIdMoe = ( select m.iIdMoe '+
                                             'from moe m '+
                                             'where m.sContrato = :orden '+
                                             'and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                                                                  'from moe m1 '+
                                                                  'where m1.sContrato = :orden '+
                                                                  'and m1.dIdFecha <= :fecha '+
                                                                ') '+
                                            ') '+
                               'and mr.sIdRecurso = bp.sIdPersonal '+
                          ') '+ #10 +

                        'inner join ordenesdetrabajo ot '+
                          'on ( ot.sNumeroOrden = ba.sNumeroOrden ) '+ #10 +

                        'inner join plataformas pl '+
                          'on ( ot.sIdPlataforma = pl.sIdPlataforma ) '+ #10 +

                        'inner join pernoctan pr '+
                          'on ( ot.sIdPernocta = pr.sIdPernocta ) '+ #10 +

                        'where bp.sContrato = :orden '+
                        'and bp.dIdFecha = :fecha group by ba.iIdActividad, ba.sNumeroActividad, p.sIdPersonal '+
                        'order by p.iItemOrden '

                        ,

                        'select ot.sNumeroOrden, '+
                                 'ba.iIdActividad, '+
                                 'ba.iIdDiario, '+
                                 'ba.sNumeroActividad, '+
                                 'ba.mDescripcion, '+
                                 'mr.sIdRecurso, '+
                                 'e.sDescripcion, '+
                                 'be.dCantidad, '+
                                 'be.dCantHH, '+
                                 'be.sIdPlataforma, '+
                                 'be.sIdPernocta, '+
                                 'be.lImprime '+ #10 +

                          'from bitacoradeequipos be '+
                          'inner join bitacoradeactividades ba '+
                            'on ( be.sContrato = ba.sContrato '+
                              'and be.iIdActividad = ba.iIdActividad '+
                              'and be.sNumeroActividad = ba.sNumeroActividad '+
                              'and be.iIdDiario = ba.iIdDiario '+
                              'and be.sNumeroOrden = ba.sNumeroOrden '+
                              'and ba.sIdTipoMovimiento = "ED" '+
                            ') '+ #10 +

                          'inner join equipos e '+
                            'on ( e.sContrato = :contrato '+
                              'and e.sIdEquipo = be.sIdEquipo '+
                            ') '+ #10 +

                          'inner join moerecursos mr '+
                            'on ( mr.iIdMoe = ( select m.iIdMoe '+
                                               'from moe m '+
                                               'where m.sContrato = :orden '+
                                               'and m.dIdFecha = ( select max( m1.dIdFecha ) '+
                                                                    'from moe m1 '+
                                                                    'where m1.sContrato = :orden '+
                                                                    'and m1.dIdFecha <= :fecha '+
                                                                  ') '+
                                              ') '+
                                 'and mr.sIdRecurso = be.sIdEquipo '+
                          ') '+ #10 +

                          'inner join ordenesdetrabajo ot '+
                            'on ( ot.sNumeroOrden = ba.sNumeroOrden ) '+ #10 +

                          'inner join plataformas pl '+
                            'on ( ot.sIdPlataforma = pl.sIdPlataforma ) '+ #10 +

                          'inner join pernoctan pr '+
                            'on ( ot.sIdPernocta = pr.sIdPernocta ) '+ #10 +

                          'where be.sContrato = :orden '+
                          'and be.dIdFecha = :fecha ' +
                          'group by ba.iIdActividad, ba.sNumeroActividad, e.sIdEquipo '+
                          'order by e.iItemOrden ');

  {$ENDREGION}

begin
  try
    zqExiste := TZReadOnlyQuery.Create(nil);
    zqExiste.Connection := connection.zConnection;

    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;

    qrFolios.First;
    qrActividades.First;

    {$REGION 'Crear excel'}

    try
      Excel := CreateOleObject('Excel.Application');
      Libro := Excel.Workbooks.Add;

      while Libro.Sheets.Count > 1 do
        Libro.Sheets[1].Delete;

      Libro.Sheets.Add;
      Hoja := Libro.Sheets[1];
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
      Excel.Visible :=  EXCEL_VISIBLE;
      Excel.Workbooks[1].Sheets[1].Name := 'Personal';
      Excel.Workbooks[1].Sheets[2].Name := 'Equipo';

    except
      on e:Exception do
      begin
        MessageDlg('No se puede continuar verifique tener instalada la suite de Microsoft Office', mtInformation, [mbOK], 0);
        Exit;
      end;
    end;

    {$ENDREGION}

    {$REGION 'Escribe'}

    iPages := 1;
    if Personal and Equipo then
      iPages := 2;

    iPCount := 0;
    for iHoja := 1 to iPages do
    begin
      {$REGION 'Encabezado Excel'}
      prgFolios.Maximum := 0;
      prgFolios.Position := 0;
      prgFolios.Refresh;
      prgActividades.Maximum := 0;
      prgActividades.Position := 0;
      prgActividades.Refresh;
      Application.ProcessMessages;

      zqExiste.Active := False;
      zqExiste.SQL.Clear;
      zqExiste.SQL.Text := (sSQLEXISTE[iPCount]);
      zqExiste.ParamByName('contrato').AsString := global_Contrato_Barco;
      zqExiste.ParamByName('orden').AsString := global_contrato;
      zqExiste.ParamByName('fecha').AsDate := StrToDate(Cuadre.Fecha);
      zqExiste.Open;

      iFila    := 4;
      iColumna := 10;
      Libro.Sheets[iHoja].select;

      iRegistros  := qrActividades.RecordCount;
      iTotalFilas := qrActividades.RecordCount + 6;

      Excel.Columns['A:A'].columnwidth := 10.80;
      Excel.Columns['B:B'].columnwidth := 4;
      Excel.Columns['B:B'].wraptext := false;
      Excel.Columns['C:C'].columnwidth := 4;
      Excel.Columns['D:D'].columnwidth := 8;
      Excel.Columns['E:E'].columnwidth := 4;
      Excel.Columns['F:G'].columnwidth := 4.14;
      Excel.Columns['H:H'].columnwidth := 5.86;
      Excel.Columns['I:I'].columnwidth := 7.12;
      Excel.Columns['J:J'].columnwidth := 9;
     // Excel.Columns['I:I'].columnwidth := 9;
     // Excel.Columns['J:K'].columnwidth := 6.29;
     Excel.Columns['K:L'].columnwidth := 6.29;
      qrMOE_Sol.Active := False;
      qrMOE_Sol.SQL.Text := sSQLMOE[iPCount];
      qrMOE_Sol.ParamByName('orden').AsString := global_contrato;
      qrMOE_Sol.ParamByName('contrato').AsString := global_Contrato_Barco;
      qrMOE_Sol.ParamByName('fecha').AsDate := StrToDate( cbbReportes.Text );
      qrMOE_Sol.Open;

      iFCount := 0;
      qrFolios.First;

      if qrMOE_Sol.RecordCount = 0 then
      begin

        Rango := Excel.Range['A1:E4'];
        Rango.MergeCells := True;
        Rango.Value := 'No hay MOE para esta categoria';
        Rango.WrapText := True;
        Rango.HorizontalAlignment := xlCenter;
        Rango.verticalAlignment := xlCenter;
        Rango.Font.Size := 25;
        Rango.Interior.ColorIndex := 43;

        Continue;
      end;

      lEncabezado := True;

      iJornadas := qrFolios.FieldByName('ijornadas').AsInteger;
      iMCount  := 0;
      IAPCount := 0;
     // iColumna := 12;
     iColumna := 13;
      qrMOE_Sol.First;
      lblEstado.Caption := 'Cargando ' + TIPO[iPCount];
      prgFolios.Maximum := qrFolios.RecordCount;
      prgFolios.Position := 0;
      Application.ProcessMessages;

      Excel.Rows[ifila].rowheight := 30;

      Rango := Excel.range['I'+inttostr(ifila - 3)+':I'+inttostr(ifila - 3)];
      Rango.mergecells := True;
      Rango.value := 'TOTAL';

      Rango := Excel.range['A'+inttostr(ifila - 2)+':L'+inttostr(ifila - 1)];
      Rango.interior.colorindex := 44;
      Rango.horizontalalignment := xlright;

      Rango := Excel.range['A'+inttostr(ifila - 2)+':L'+inttostr(ifila - 2)];
      Rango.mergecells := True;
      Rango.value := 'SOLICITADO';

      Rango := Excel.range['A'+inttostr(ifila - 1)+':L'+inttostr(ifila - 1)];
      Rango.MergeCells := true;
      Rango.value := 'A BORDO';


      while not qrMOE_Sol.Eof do
      begin
          if iHoja = 2 then
            Excel.Columns[ColumnaNombre(iColumna)].ColumnWidth := 20;

          Rango := Excel.range[ColumnaNombre(iColumna)+inttostr(ifila)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila)];
          Rango.MergeCells := True;
          Rango.Value      := qrMOE_Sol.FieldByName('sDescripcion').AsString;

          Rango := Excel.Range[ColumnaNombre(iColumna)+inttostr(ifila+1)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila+1)];
          Rango.MergeCells := True;
          Rango.Value := qrMOE_Sol.FieldByName('sIdRecurso').asstring;

          Rango := Excel.Range[ColumnaNombre(iColumna)+inttostr(ifila+2)+':'+ColumnaNombre(iColumna)+inttostr(ifila+2)];
          Rango.MergeCells := True;
          Rango.Value := 'CANT.';

          Rango := Excel.Range[ColumnaNombre(iColumna+1)+inttostr(ifila+2)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila+2)];
          Rango.MergeCells := True;
          Rango.Value := 'FRACC.';

          Rango := Excel.Range[ColumnaNombre(iColumna)+inttostr(ifila-1)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila-1)];
          Rango.MergeCells := True;
          Rango.Value := qrMOE_Sol.FieldByName('iabordo').asinteger;

          Rango := Excel.Range[ColumnaNombre(iColumna)+inttostr(ifila-2)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila-2)];
          Rango.MergeCells := True;
          Rango.Value := qrMOE_Sol.FieldByName('isolicitado').asinteger;


          Excel.Range[ColumnaNombre(iColumna)+inttostr(ifila-1)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila-2)].NumberFormat := '0.00';

          Inc( iMCount );
          Inc(iColumna, 2);
          qrMOE_Sol.Next;
      end;

      Rango := Excel.Range[ColumnaNombre(10)+inttostr(ifila - 3)];
      Rango.Formula := '=SUM('+'M'+inttostr( iFila -3 ) + ':' + COLUMNAS[ iColumna]+ IntToStr( iFila -3) + ')';
      {$ENDREGION}

      Rango := Excel.Range[ColumnaNombre(1)+inttostr(ifila-2)+':'+ColumnaNombre(icolumna-1)+inttostr(ifila+1)];
      Rango.Interior.Color := $00E2B48D;
      Rango.Horizontalalignment := xlcenter;
      Rango.verticalalignment := xlCenter;
      Rango.Wraptext := True;

      Rango := Excel.Range['A'+inttostr(ifila)+':L'+inttostr(ifila)];
      Rango.MergeCells := True;
      Rango.Horizontalalignment := xlcenter;
      Rango.verticalalignment := xlCenter;
      Rango.Wraptext := True;
      Rango.Value := 'CUADRE DE RECURSOS';
      Rango.font.Size := 15;
      Rango.font.bold := true;

      Rango := Excel.Range['A'+inttostr(2)+':M'+inttostr(ifila +1)];
      Rango.Interior.Color := $00E2B48D;

      Inc(ifila,2);

      Rango := Excel.range['A'+inttostr(ifila)+':'+ColumnaNombre(icolumna-1)+inttostr(ifila)];
      Rango.interior.color := $00E8DEB7;
      Rango.HorizontalAlignment := xlcenter;

      Rango := Excel.range['A'+inttostr(ifila)];
      Rango.Value := 'FOLIO';
      Rango := Excel.range['B'+inttostr(ifila)];
      Rango.Value := 'DESCRIPCION';
      Rango := Excel.range['C'+inttostr(ifila)];
      Rango.Value := 'PERNOCTA';
      Rango := Excel.range['D'+inttostr(ifila)];
      Rango.Value := '$_PERN';
      Rango := Excel.range['E'+inttostr(ifila)];
      Rango.Value := 'PLATAFORMA';
      Rango := Excel.range['F'+inttostr(ifila)];
      Rango.Value := 'NEC';
      Rango := Excel.range['G'+inttostr(ifila)];
      Rango.Value := 'F.T.';
      Rango := Excel.range['H'+inttostr(ifila)];
      Rango.Value := 'INICIO';
      Rango := Excel.range['I'+inttostr(ifila)];
      Rango.Value := 'FIN';
      Rango := Excel.range['J'+inttostr(ifila)];
      Rango.Value := 'FRAC. DIA';
      Rango := Excel.range['K'+inttostr(ifila)];
      Rango.Value := 'ACT.';
      Rango := Excel.range['L'+inttostr(ifila)];
      Rango.Value := 'ID.';

      Inc(iFila);


      iACount := 0;
      qrActividades.First;
      ActividadAnterior := '*_*';
      prgActividades.Maximum := qrActividades.RecordCount;
      prgActividades.Position := 0;
      iActividad := qrActividades.FieldByName('iIdactividad').AsInteger;
      iInicio := iFila;

      while not qrActividades.Eof do
      begin

        if ActividadAnterior <> qrActividades.FieldByName('sNumeroActividad').AsString then
        begin
          Actividades       := iFila + qrActividades.FieldByName('iActividades').AsInteger;
          ActividadAnterior := qrActividades.FieldByName('sNumeroActividad').AsString;
          Inc( IAPCount );

          if lEncabezado then
          begin
              Rango := Excel.Rows[iFila];
              if LowerCase( cbbVista.Text ) = 'horarios' then
                Rango.EntireRow.Hidden := True;

              iCol := 13;
              for iColumna := 4 to qrMOE_Sol.RecordCount + 3 do
              begin
                Rango := Excel.Range[ColumnaNombre(iCol)+inttostr(ifila - 6)];
                Rango.Formula := '=SUM('+COLUMNAS[ iCol +1]+inttostr( iFila ) + ':' + COLUMNAS[ iCol +1]+ IntToStr( iFila + iRegistros  - 1) + ')';

                Rango := Excel.Range[COLUMNAS[ iCol ]+inttostr( 1 )+':'+COLUMNAS[ iCol +1 ]+inttostr(1)];
                Rango.Horizontalalignment := xlcenter;
                Rango.verticalalignment   := xlCenter;
                Rango.numberformat        := '0.00';

                if zqExiste.RecordCount = 0 then
                begin
                    Rango := Excel.Range[ ColumnaNombre( iCol ) + IntToStr(qrActividades.RecordCount + 25)];
                    Rango.Value := 'Si';
                end;

                //Combinar celda de total
                Rango := Excel.Range[ColumnaNombre(iCol)+inttostr(ifila-6)+':'+ColumnaNombre(iCol+1)+IntToStr( iFila-6)];
                Rango.MergeCells := True;

                //Lineas separadoras
                Rango := Excel.Range[ColumnaNombre(iCol)+inttostr(ifila-6)+':'+ColumnaNombre(iCol+1)+IntToStr( iFila + iRegistros  - 1)];
                Rango.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                Rango.Borders[xlEdgeLeft].Weight       := xlThin;
                Rango.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                Rango.Borders[xlEdgeTop].Weight        := xlThin;
                Rango.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                Rango.Borders[xlEdgeBottom].Weight     := xlThin;
                Rango.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                Rango.Borders[xlEdgeRight].Weight      := xlThin;

                inc(icol, 2);
                qrMOE_Sol.Next;
              end;
              lEncabezado := False;
          end;
          Inc( iACount );

          Continue;
        end;

        iActividad := qrActividades.FieldByName('iIdActividad').AsInteger;
        Rango := Excel.Range['A'+inttostr(ifila)+':G'+inttostr(ifila)];
        Rango.numberformat := '@';

        Excel.range['A'+inttostr(ifila)].value := qrActividades.FieldByName('sNumeroOrden').AsString;
        Excel.range['B'+inttostr(ifila)].value := qrActividades.FieldByName('mDescripcion').AsString;
        Excel.range['C'+inttostr(ifila)].value := qrActividades.FieldByName('sIdPernocta').AsString;
        Excel.range['D'+inttostr(ifila)].value := qrActividades.FieldByName('aplicapernocta').AsString;

        Excel.range['E'+inttostr(ifila)].value := qrActividades.FieldByName('sIdPlataforma').AsString;
        Excel.range['F'+inttostr(ifila)].value := qrActividades.FieldByName('sIdClasificacion').AsString;
        Excel.range['G'+inttostr(ifila)].value := qrActividades.FieldByName('sTipoObra').AsString;
        Excel.range['H'+inttostr(ifila)].value := qrActividades.FieldByName('sHoraInicio').AsString;
        Excel.range['I'+inttostr(ifila)].value := qrActividades.FieldByName('sHoraFinal').AsString;
        Rango := Excel.Range['K'+inttostr(ifila)+':K'+inttostr(ifila)];
        Rango.numberformat := '@';
        Excel.range['K'+inttostr(ifila)].value := qrActividades.FieldByName('sNumeroActividad').AsString;
        Excel.range['L'+inttostr(ifila)].value := qrActividades.FieldByName('iIdActividad').AsInteger;
        Excel.range['J'+inttostr(ifila)].formula := '=I'+inttostr(ifila)+' - ' +'H'+inttostr(ifila);
        Excel.Range['J'+inttostr(ifila)].numberformat := '0.00';
        dDuracion := Excel.Range['G'+inttostr(ifila)].text;


        Rango := Excel.Rows[iFila];
        if LowerCase( cbbVista.Text ) = 'actividades' then
          Rango.EntireRow.Hidden := True;

        Excel.Range['B'+inttostr(ifila) + ':B' + inttostr(ifila)].wraptext := false;

        qrMOE_Sol.First;
        iCol := 14;
        for iColumna := 4 to qrMOE_Sol.RecordCount + 3 do
        begin
          Rango := Excel.Range[ColumnaNombre(iCol)+inttostr(ifila)];

          if iPCount = 0 then
          begin
            if qrMOE_Sol.FieldByName('sTE').AsString = 'Si' then
              Rango.Formula := '=('+COLUMNAS[icol-1]+inttostr(ifila)+'*J'+inttostr(ifila)+')*'+ IntToStr( 24 )
            else
            if qrMOE_Sol.FieldByName('sTierra').AsString = 'Si' then
              Rango.Formula := '=('+COLUMNAS[icol-1]+inttostr(ifila)+'*J'+inttostr(ifila)+'.)*'+ IntToStr( 3 )
            else
            begin
                 Rango.Formula := '=('+COLUMNAS[icol-1]+inttostr(ifila)+'*J'+inttostr(ifila)+')*'+ IntToStr( iJornadas )
            end;
          end
          else
            Rango.Formula := '=('+COLUMNAS[icol-1]+inttostr(ifila)+'*J'+inttostr(ifila)+')';

          inc(icol, 2);
          qrMOE_Sol.Next;
        end;

        Inc( iACount );

        Inc( iFila );
        qrActividades.Next;

        prgActividades.Position := prgActividades.Position + 2;
        prgActividades.Refresh;
      end;

      qrActividades.Filtered := False;
      Excel.Range['B'+inttostr(ifila) + ':B' + inttostr(ifila)].wraptext := false;



      {SUMAS}
//      for iFCount := 0 to qrMOE_Sol.RecordCount - 1 do
//      begin
        if zqExiste.RecordCount > 0 then
        begin
           qrActividades.First;
           iFila := 7;
           //Recorer Actividades y Categorias
           while not qrActividades.Eof do
           begin
              zqExiste.Filtered := False;
              zqExiste.Filter := 'sNumeroOrden = ' + QuotedStr( qrActividades.FieldByName('sNumeroOrden').AsString )
                              + ' AND iIdActividad = ' + IntToStr( qrActividades.FieldByName('iIdActividad').AsInteger) ;
              zqExiste.Filtered := True;
              iCol := 13;

              if zqExiste.RecordCount > 0 then
              begin
                qrMOE_Sol.First;
                while not qrMOE_Sol.Eof do
                begin
                  if zqExiste.Locate('sIdRecurso', qrMOE_Sol.FieldByName('sIdRecurso').AsString, [] ) then
                  begin
                      Excel.Range[ ColumnaNombre( iCol ) + IntToStr( iFila )].Value := zqExiste.FieldByName('dCantidad').AsString;

                      //Impresion de la categoria si alguna no se imprime..
                      Rango := Excel.Range[ ColumnaNombre( iCol ) + IntToStr(qrActividades.RecordCount + 25)];
                      Rango.Value := zqExiste.FieldByName('lImprime').AsString;
                  end;
                  Inc(iCol, 2);
                  qrMOE_Sol.Next;
                end;

              end;
              inc(iFila);
              qrActividades.Next;
           end;
        end;
     // end;

      zqExiste.Filtered := False;
      if zqExiste.RecordCount > 0 then
      begin
         iCol := 13;
         qrMOE_Sol.First;
         while not qrMOE_Sol.Eof do
         begin
             iSuma  := Excel.Range[ ColumnaNombre( iCol ) + IntToStr( 1 )].Value;
             iResta := Excel.Range[ ColumnaNombre( iCol ) + IntToStr( 2 )].Value;

             Rango  := Excel.Range[ ColumnaNombre( iCol ) + IntToStr( 1 )];
             if xround(iSuma,6) = xRound(iResta,6) then
                Rango.Interior.Color := $000FD37F
             else
                if xround(iSuma, 6) > xround(iResta,6) then
                   Rango.Interior.Color := $00266DEE
                else
                   Rango.Interior.Color := clWhite;

             inc(iCol,2);
             qrMOE_sol.Next;
         end;
      end;

      iFila := 7;
      qrActividades.first;
      while not qrActividades.Eof do
      begin
          Rango  :=  Excel.range['A'+inttostr(ifila)+':A'+inttostr(ifila)];
          if Rango.Value = qrActividades.FieldByName('sNumeroOrden').AsString then
          begin
              if qrActividades.FieldByName('sColor').AsString <> '' then
              begin
                  Rango  :=  Excel.range['A'+inttostr(ifila)+':'+ColumnaNombre( iCol -1 )+inttostr(ifila)];
                  Rango.Interior.Color := StringToColor(qrActividades.FieldByName('sColor').AsString);
              end;
              inc(iFila);
          end;
          qrActividades.Next;
      end;

      Excel.Columns['B'].EntireColumn.Hidden := True;
      Excel.Columns['C'].EntireColumn.Hidden := True;
      //Excel.Columns['D'].EntireColumn.Hidden := True;
      Excel.Columns['E'].EntireColumn.Hidden := True;
      Excel.Columns['L'].EntireColumn.Hidden := True;

      Excel.Rows[qrActividades.RecordCount + 25].rowheight := 0;
      Excel.Range['M7'].Select;
      Excel.ActiveWindow.FreezePanes := True;
      Libro.ActiveSheet.Protect( True, True, True );
      Inc(iPCount);

      prgFolios.Position := prgFolios.Position + 2;
      prgFolios.Refresh;
    end;
    prgFolios.Position := prgFolios.Maximum;
    prgFolios.Refresh;
    GetTempPath(SizeOf(global_TempPath), global_TempPath);
    sArchivo := global_TempPath+'inteliCuad_intelcode.xls';
    Libro.SaveAs( sArchivo, 56 );
    Excel.Quit;
    lblEstado.Caption := 'Listo';
    {$ENDREGION}

  finally
    Screen.Cursor := Cursor;
  end;
end;

procedure TfrmCuadreNormal.ValidaFolios;
var
  iPCount,
  iFCount : integer;

  zqFolios : TZReadOnlyQuery;

  sSQL : string;
begin
  try
    sSQL := 'select ot.sNumeroOrden as sIdFolio, '+
                   'ot.sNumeroOrden, '+
                   'ot.iJornadas, '+
                   'ot.sIdPernocta, '+
                   'ot.sIdPlataforma, '+
                   'ba.sColor, '+
                   'ot.lAplicaJornada '+
            'from bitacoradeactividades ba '+#10+

            'inner join reportediario rd '+
              'on ( rd.sOrden = ba.sContrato '+
                'and rd.dIdFecha = ba.dIdFecha ) '+#10+

            'inner join ordenesdetrabajo ot '+
              'on ( ot.sContrato = ba.sContrato '+
                'and ot.sNumeroOrden = ba.sNumeroOrden ) '+#10+
    
            'inner join estatus e '+
              'on ( ot.cIdStatus = e.cIdStatus '+
                'and e.sDescripcion = "PROCESO" ) '+#10+
    
            'inner join contratos c '+
              'on ( c.sContrato = ba.sContrato ) '+#10+
  
            'where ba.sContrato = :orden '+
            'and ba.sIdTipoMovimiento = "ED" '+
            'and ba.dIdFecha = :fecha '+#10+
            'group by ot.sNumeroOrden';

  
    zqFolios := TZReadOnlyQuery.Create(nil);
    zqFolios.Connection := connection.zConnection;
    zqFolios.Active := False;
    zqFolios.SQL.Text := sSQL;
    zqFolios.ParamByName('orden').AsString := global_contrato;
    zqFolios.ParamByName('fecha').AsDate   := StrToDate( Cuadre.Fecha );
    zqFolios.Open;
    zqFolios.First;

    if zqFolios.RecordCount = 0 then
      raise Exception.Create('No existen folios registrados en el dia '+ cbbReportes.EditText);

    for iPCount := 0 to Length( Cuadre.CATEGORIA ) - 1 do
    begin

      if Cuadre.CUADRAR[ iPCount ] then
      begin

        for iFCount := 0 to Length( Cuadre.CATEGORIA[iPCount] ) - 1 do
        begin
          if zqFolios.Locate('sNumeroOrden', Cuadre.CATEGORIA[iPCount, iFCount].sFolio, [] ) then
          begin
            Cuadre.CATEGORIA[iPCount, iFCount].eExiste := True;
            Cuadre.CATEGORIA[iPCount, iFcount].sPernocta := zqFolios.FieldByName('sIdPernocta').AsString;
            Cuadre.CATEGORIA[iPCount, iFcount].sPlataforma := zqFolios.FieldByName('sIdPlataforma').AsString;
          end
          else
            Cuadre.CATEGORIA[iPCount, iFCount].eExiste := False;
        end;

      end;

    end;
        
  finally
    zqFolios.Free;
  end;
end;

procedure TfrmCuadreNormal.ValidaCategorias(var excel, hoja : Variant ; var zqMoe : TZReadOnlyQuery; iTipo : Integer);
var
  iPCount,
  iFCount,
  iMCount,
  iAcount,
  iFila,
  iColumna,
  iHoja : Integer;

  dSuma : Double;
begin
  zqMoe.First;
  iPCount := 0;

  for iFCount := 0 to Length( Cuadre.CATEGORIA[iTipo] ) - 1 do
  begin
    if Cuadre.CATEGORIA[iTipo, iFCount].eExiste then{ Existe el folio }
    begin
      for iMCount := 0 to Length( Cuadre.CATEGORIA[iTipo, iFCount].MOE ) - 1 do
      begin
        if iTipo = 1 then
            Cuadre.CATEGORIA[iTipo, iFCount].MOE[iMCount].sPlataforma := '*';

        if zqMoe.Locate( 'sidrecurso', Cuadre.CATEGORIA[iTipo, iFCount].MOE[iMCount].sIdRecurso, [] ) then{ Existe la categoria }
        begin
          Cuadre.CATEGORIA[iTipo, iFCount].MOE[iMCount].eExiste := True;
          dSuma := excel.range[columnanombre( Cuadre.CATEGORIA[iTipo, iFCount].MOE[iMCount].iCol )+inttostr( Cuadre.CATEGORIA[iTipo, iFCount].iInicio - 5 )].Text;
          if ( Length( Trim( Cuadre.CATEGORIA[iTipo, iFCount].MOE[iMCount].sPlataforma ) ) > 0 ) and ( Length( Trim( Cuadre.CATEGORIA[iTipo, iFCount].MOE[iMCount].sPernocta ) ) > 0 ) then
          begin
            if ( dSuma > 0 ) and (dSuma <= Cuadre.CATEGORIA[iTipo, iFCount].MOE[iMCount].iAbordo )  then
            begin
              Cuadre.CATEGORIA[iTipo, iFCount].MOE[iMCount].eListo := True;
            end
            else
              Cuadre.CATEGORIA[iTipo, iFCount].MOE[iMCount].eListo := False;
          end;

        end;

      end;

    end;

  end;

end;

//Martin
procedure TfrmCuadreNormal.GuardaEnBD;
var
  Excel,
  Book,
  Hoja,
  Rango : Variant;

  iFila,
  iColumna,
  iCol,
  iPCount,
  iFCount,
  iMCount,
  iACount,
  iOk,
  iHoja,
  iActividad,
  iPivote,
  iMIndex,
  iIdDiario,
  iTarea,
  iItemOrden : Integer;

  sArchivo,
  sWbs,
  sNumeroActividad,
  sIdRecurso,
  sDescripcion,
  sInicio,
  sFin,
  sIdPernocta,
  sIdPlataforma,
  sIdCategoria,
  sNumeroOrden,
  sTipoObra : string;

  dCantidad,
  dCantidadHH : Real48;

  dAjuste : double;

  zqActividades,
  zqMoe, zqCuadreAjustes : TZReadOnlyQuery;

  zqSave : TZQuery;
  sImprime  : string;

const
  TABLA : array[0..1] of string = ('Personal', 'Equipos');

begin
  try
    try
      //connection.zconnection.startTransaction;

      zqMoe := TZReadOnlyQuery.Create(nil);
      zqMoe.Connection := connection.zConnection;

      zqCuadreAjustes := TZReadOnlyQuery.Create(nil);
      zqCuadreAjustes.Connection := connection.zConnection;

      {$REGION 'Inicia'}
      GetTempPath(SizeOf(global_TempPath), global_TempPath);
      sArchivo := global_TempPath+'inteliCuad_intelcode_Save.xls';
      if not Length( Cuadre.CATEGORIA ) > 0 then
        Exit;

//      if FileExists( sArchivo ) then
//      begin
//        DeleteFile( sArchivo );
//      end;
      Libro.SaveToFile(sArchivo);
      zqActividades := TZReadOnlyQuery.Create(nil);
      zqActividades.Connection := connection.zConnection;
      zqSave := TZQuery.Create(nil);
      zqSave.Connection := connection.zConnection;
      try
        Excel := CreateOleObject('Excel.Application');
        Book := Excel.Workbooks.Open( sArchivo );
        Hoja := Book.Sheets[1];
        Excel.DisplayAlerts := False;
        Excel.ScreenUpdating := True;
        Excel.Visible := EXCEL_VISIBLE;
      except
        MessageDlg('No se puede continuar verifique tener instalada la suite de Microsoft Office', mtInformation, [mbOK], 0);
        Exit;
      end;

      {$ENDREGION}

      ValidaFolios();

      {$REGION 'Actividades'}

      zqActividades := TZReadOnlyQuery.Create(nil);
      zqActividades.Connection := connection.zConnection;
      zqActividades.SQL.Add( 'select ba.iIdActividad, '+
                                       'ba.iIdDiario, '+
                                       'ba.sWbs, '+
                                       'ba.sNumeroActividad, '+
                                       'ba.sHoraInicio, '+
                                       'ba.sHoraFinal, '+
                                       'ba.sNumeroOrden '+

                                'from bitacoradeactividades ba '+

                                'inner join reportediario rd '+
                                  'on ( rd.sOrden = ba.sContrato '+
                                    'and rd.dIdFecha = ba.dIdFecha ) '+

                                'inner join ordenesdetrabajo ot '+
                                  'on ( ot.sContrato = ba.sContrato '+
                                    'and ot.sNumeroOrden = ba.sNumeroOrden ) '+

                                'where ba.sContrato = :orden '+
                                'and ba.dIdFecha = :fecha '+
                                'and ba.sIdTipoMovimiento = "ED" ');
      zqActividades.ParamByName('orden').AsString := global_contrato;
      zqActividades.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
      zqActividades.Open;
      zqActividades.First;

      if zqActividades.RecordCount = 0 then
        raise Exception.Create('No se encontraron actividades reportadas en el dia especificado');

      {$ENDREGION}

      ConsultarActividades;

      iPCount := 0;
      for iHoja := 1 to 2 do
      begin
          prgFolios.Maximum := 100;
          prgFolios.Position := 0;
          prgFolios.Refresh;
          prgActividades.Maximum := 0;
          prgActividades.Position := 0;
          prgActividades.Refresh;
          Application.ProcessMessages;


          Hoja := Book.Sheets[iHoja];
          Hoja.Select;
          Excel.ActiveSheet.Unprotect;

          {$REGION 'Moe'}

          zqMoe.Active := False;
          zqMoe.SQL.Text := sSQLMOE[iPCount];
          zqMoe.ParamByName('orden').AsString     := global_contrato;
          zqMoe.ParamByName('contrato').AsString  := global_Contrato_Barco;
          zqMoe.ParamByName('fecha').AsDate       := StrToDate( Cuadre.Fecha );
          zqMoe.Open;
          zqMoe.First;

          if zqMoe.RecordCount = 0 then
            raise Exception.Create('No se encontro un moe vigente a la fecha especificada')  ;

          ValidaCategorias(Excel, Hoja, zqMoe, iPCount);

          {$ENDREGION}

          {$REGION 'Vaciar tabla'}
          zqCuadreAjustes.Active := False;
          zqCuadreAjustes.SQL.Clear;
          if iPCount = 0 then          
             zqCuadreAjustes.SQL.Text := 'select sIdPersonal, sNumeroActividad, sHoraInicio, sHoraFinal, sNumeroOrden, sIdPernocta, dAjuste from bitacorade'+lowercase( TABLA[ iPCount ] )+' where sContrato = :orden and didfecha =:fecha and dAjuste <> 0 '
          else   
             zqCuadreAjustes.SQL.Text := 'select sIdEquipo, sNumeroActividad, sHoraInicio, sHoraFinal, sNumeroOrden, sIdPernocta, dAjuste from bitacorade'+lowercase( TABLA[ iPCount ] )+' where sContrato = :orden and didfecha =:fecha and dAjuste <> 0 ';
          zqCuadreAjustes.ParamByName('orden').AsString := global_contrato;
          zqCuadreAjustes.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
          zqCuadreAjustes.Open;          

          zqSave.Active := False;
          zqSave.SQL.Text := 'delete from bitacorade'+lowercase( TABLA[ iPCount ] )+' where sContrato = :orden and didfecha =:fecha';
          zqSave.ParamByName('orden').AsString := global_contrato;
          zqSave.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
          zqSave.ExecSQL;

          {$ENDREGION}

          iFila    := 7;
          prgActividades.Maximum := qrActividades.RecordCount;
          prgActividades.Position := 0;
          qrActividades.First;
          while not qrActividades.Eof do
          begin
              iActividad        := qrActividades.FieldByName('iIdActividad').AsInteger;
              sWbs              := qrActividades.FieldByName('sWbs').AsString;
              iIdDiario         := qrActividades.FieldByName('iIdDiario').AsInteger;
              sNumeroActividad  := qrActividades.FieldByName('sNumeroActividad').AsString;

              {$REGION 'Valida Actividad'}

              zqActividades.Filtered := False;
              zqActividades.Filter := 'iidactividad = '+inttostr(iactividad)+
                                      ' AND swbs ='+QuotedStr(sWbs)+
                                      ' AND snumeroactividad='+QuotedStr(sNumeroActividad)+
                                      ' AND iiddiario = '+inttostr(iIdDiario);
              zqActividades.Filtered := True;

              if zqActividades.RecordCount = 0 then
              begin
                Inc( iFila );
                Continue;
              end;

              {$ENDREGION}
              sInicio := qrActividades.FieldByName('sHoraInicio').AsString;
              sFin    := qrActividades.FieldByName('sHoraFinal').AsString;
              iTarea  := qrActividades.FieldByName('iIdTarea').AsInteger;

              zqMoe.First;
              iCol:=13;//iCol := 12;
              while not zqMoe.Eof do
              begin
                    sIdRecurso    := zqMoe.FieldByName('sIdRecurso').AsString;
                    sIdPernocta   := qrActividades.FieldByName('sIdPernocta').AsString;
                    sIdPlataforma := qrActividades.FieldByName('sIdPlataforma').AsString;
                    sIdCategoria  := zqMoe.FieldByName('sIdRecurso').AsString;
                    iItemOrden    := zqMoe.FieldByName('iItemOrden').AsInteger;
                    sTipoObra     := qrActividades.FieldByName('sTipoObra').AsString;
                    sFolio        := qrActividades.FieldByName('sNumeroOrden').AsString;

                    try
                      dCantidad   := Excel.Range[columnanombre(icol)+inttostr(ifila)].Value;
                    except
                      dCantidad := 0;
                    end;

                    dAjuste := 0;
                    //Ahora identificamos si esta categoria tiene ajuste..
                    if zqCuadreAjustes.RecordCount > 0 then
                    begin
                        zqCuadreAjustes.Filtered := False;
                        if iPCount = 0 then                        
                        begin
                           zqCuadreAjustes.Filter := ' sIdPersonal = '+ QuotedStr(sIdRecurso) +
                                                  ' and sNumeroOrden = ' + QuotedStr(sFolio) +
                                                  ' and sNumeroActividad = '+ QuotedStr(sNumeroActividad) +
                                                  ' and sHoraInicio = '+ QuotedStr(sInicio) +
                                                  ' and sHoraFinal = '+ QuotedStr(sFin) +
                                                  ' and sIdpernocta = '+ QuotedStr(sIdPernocta);
                        end
                        else
                        begin
                           zqCuadreAjustes.Filter := ' sIdEquipo = '+ QuotedStr(sIdRecurso) +
                                                  ' and sNumeroOrden = ' + QuotedStr(sFolio) +
                                                  ' and sNumeroActividad = '+ QuotedStr(sNumeroActividad) +
                                                  ' and sHoraInicio = '+ QuotedStr(sInicio) +
                                                  ' and sHoraFinal = '+ QuotedStr(sFin) +
                                                  ' and sIdpernocta = '+ QuotedStr(sIdPernocta);
                        end;
                        zqCuadreAjustes.Filtered := True;

                        if zqCuadreAjustes.RecordCount > 0 then
                           dAjuste := zqCuadreAjustes.FieldByName('dAjuste').AsFloat;

                        zqCuadreAjustes.Filtered := False;   
                    end;


                    //Aqui es para ver si se imprime o no, si no se imprime sin tener valores se tiene que guardar en 0 en bitacora.
                    sImprime :=  Excel.Cells[qrActividades.RecordCount + 25, iCol].Value;
                    if trim(sImprime) = '' then
                       sImprime := 'Si';

                    if (dCantidad > 0) or (sImprime = 'No') or (lCeros = True) then
                    begin
                        dCantidadHH := Excel.Cells[iFila, iCol+1].Value;

                        zqSave.Active := False;
                        zqSave.SQL.Text := SQLINSERT[iPCount];
                        zqSave.ParamByName('orden').AsString := global_contrato;
                        zqSave.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
                        zqSave.ParamByName('iddiario').AsInteger := iIdDiario;
                        zqSave.ParamByName('idrecurso').AsString := sIdRecurso;
                        zqSave.ParamByName('descripcion').AsString := zqMoe.FieldByName('sDescripcion').AsString;
                        zqSave.ParamByName('pernocta').AsString := sIdPernocta;
                        zqSave.ParamByName('hinicio').AsString := sInicio;
                        zqSave.ParamByName('hfinal').AsString := sFin;
                        zqSave.ParamByName('cantidad').AsFloat := dCantidad;
                        zqSave.Params.ParamByName('cantidadhh').DataType := ftFloat;
                        zqSave.Params.ParamByName('cantidadhh').Value := Excel.Cells[iFila, iCol+1].Value;
                        zqSave.ParamByName('Ajuste').AsFloat  := dAjuste;
                        zqSave.ParamByName('idactividad').AsInteger := iActividad;
                        zqSave.ParamByName('wbs').asstring := sWbs;
                        zqSave.ParamByName('actividad').asstring := sNumeroActividad;
                        zqSave.ParamByName('tarea').AsInteger := iTarea;
                        zqSave.ParamByName('folio').AsString := sFolio;
                        zqSave.ParamByName('ItemOrden').AsInteger := iItemOrden;
                        zqSave.ParamByName('plataforma').AsString := sIdPlataforma;
                        if iPCount = 0 then
                        begin
                            zqSave.ParamByName('Categoria').AsString  := sIdCategoria;
                        end;
                        zqSave.ParamByName('TipoObra').AsString   := sTipoObra;
                        zqSave.ParamByName('Imprime').AsString    := sImprime;
                        if iHoja=1 then
                          zqSave.ParamByName('aplicapernocta').AsString := Excel.Cells[iFila,4].Value;

                        zqSave.ExecSQL;


                        //Excel.Range[ColumnaNombre(icol)+inttostr(ifila)+':'+ColumnaNombre(icol+1)+inttostr(ifila)].interior.colorindex := 43;
                        Inc( iOk );
                    end;

                  Inc(iCol, 2);
                  zqMoe.Next;
              end;

              prgActividades.Position := prgActividades.Position + 2;
              prgActividades.Refresh;

              Inc( iFila );
              qrActividades.Next;
          end;

          Cuadre.INSERTADOS[iPCount] := iOk;
          Inc(iPCount);
          iOk := 0;

          prgFolios.Position := prgFolios.Position + 50;
          prgFolios.Refresh;
      end;

    except
      on e:Exception do
      begin
        MessageDlg('Ha ocurrido un error durante el proceso '+#10+'favor de informar al administrador del sistema el siguiente error: '+#10+e.Message, mtInformation, [mbOK], 0);
        //connection.zconnection.startback;
        Cuadre.INSERTADOS[0] := 0;
        Cuadre.INSERTADOS[1] := 0;
        Exit;
      end;

    end;

  finally
    MessageDlg('Proceso Terminado con éxito!', mtInformation, [mbOK], 0);

    //if ( Cuadre.INSERTADOS[0] > 0 ) or ( Cuadre.INSERTADOS[1] > 0 ) then
      //connection.zConnection.Commit;

    //chkOrdenado.Checked := False;
    ConsultarActividades;

    Book.SaveAs(sArchivo);
    Libro.&Protected := False;
    Libro.LoadFromFile(sArchivo);
    Libro.&Protected := True;     //*Protegido
    Excel.Quit;

    if FileExists( sArchivo ) then
    begin
      DeleteFile( sArchivo );
    end;

    zqSave.Free;
    zqActividades.Free;
    zqMoe.Free;
    zqCuadreAjustes.Free;

    Cuadre.Guardado := True;
    Cuadre.Cambios := False;

    Kardex('Reporte Diario', 'Guarda cuadre del dia '+ Cuadre.Fecha , '', '', '', '', '' , 'Tarifa Diaria', 'Cuadre');
  end;
end;


procedure TfrmCuadreNormal.VentanaCortes();
var
  Form : TForm;
begin
  Form := TForm.Create(nil);
  Form.BorderIcons := [biSystemMenu];
  Form.BorderStyle := bsSizeable;
  Form.Position := poScreenCenter;
  Form.Width := 750;
  Form.Height := 350;
  pnlCortes.Parent := Form;
  pnlCortes.Align := alClient;
  pnlCortes.Visible :=  True;
  Form.ShowModal;
  pnlCortes.Parent := frmCuadreNormal;
  pnlCortes.Align := alNone;
  pnlCortes.Visible :=  False;
  pnlCortes.Width := 0;
  pnlCortes.Height := 0;
  pnlCortes.Left := 0;
  pnlCortes.Top := 0;
  Form.Free;
  Application.ProcessMessages;
end;

procedure TfrmCuadreNormal.CargarActividades_ListView( Lista : TdxTreeView ; Folio : Integer; Actividad : string);
var
  IACount : Integer;

  Item,
  Agrupador : TTreeNode;

  Inicio,
  Fin : string;
begin
  Lista.Items.Clear;
  Agrupador := Lista.Items.Add( nil, 'Actividad' );
  Agrupador.ImageIndex := 0;
  for IACount := 0 to Length( Cuadre.CATEGORIA[Libro.ActivePage, Folio].ACTIVIDADES ) - 1 do
  begin
    if Cuadre.CATEGORIA[Libro.ActivePage, Folio].ACTIVIDADES[IACount].IsPadre then
      Continue;

    if Cuadre.CATEGORIA[Libro.ActivePage, Folio].ACTIVIDADES[IACount].sIdActividad = Actividad then
    begin
      Inicio := Cuadre.CATEGORIA[Libro.ActivePage, Folio].ACTIVIDADES[IACount].sHInicio;
      Fin := Cuadre.CATEGORIA[Libro.ActivePage, Folio].ACTIVIDADES[IACount].sHFin;

      Item := Lista.Items.AddChild( Agrupador, Actividad + ' | ' + Inicio + ' - ' + Fin);
      Item.ImageIndex := 1;
      Cuadre.CATEGORIA[Libro.ActivePage, Folio].ACTIVIDADES[IACount].iNodoCorte := Item.Index;
    end;
  end;
end;

procedure TfrmCuadreNormal.CortarActividad(Actividad: string; Inicio: string; Fin: string; InicioCorte: string; FinCorte: string; Fila: Integer);
var
  sFolio,
  sWbs,
  sNumeroActividad,
  sFactorHoras,
  sFactorHorasAc,
  TipoActividad,
  Estatus : string;

  IndexF : TFolioIndex;
  IndexA : TActividadIndex;
  IndexAP : TActividadPadreIndex;

  IdActividad,
  IndexO,
  IdDiario,
  IdDiarioNota,
  iMCount,
  Count,
  CountA,
  iHoja,
  iAnida,
  TareaAct,
  IdActividadHermano : Integer;

  iHoras,
  iMinutos,
  iHorasAc,
  iMinutosAc,
  dDuracionAc,
  dDuracion : Double;

  Old,
  Corte : TActividad;

  Rango : TcxSSCellObject;
  Rect : TRect;

  Excel, Book, Hoja , Rang: Variant;

const
  SQL_UPDATE : string = ('update bitacoradeactividades set sHoraFinal = :n_hora, iHermano = :hermano where iIdActividad = :idactividad '+
                                                                               'and sContrato = :orden '+
                                                                               'and dIdFecha = :fecha '+
                                                                               'and sWbs = :wbs '+
                                                                               'and sNumeroActividad = :actividad '+
                                                                               'and sIdTipoMovimiento = "ED" ');

  SQL_INSERT : string = ('insert into bitacoradeactividades ( sContrato, dIdFecha, iIdDiario, '+
                                                            'sIdTurno, sNumeroOrden, sWbs, '+
                                                            'sNumeroActividad, sIdTipoMovimiento, '+
                                                            'sHoraInicio, sHoraFinal, mDescripcion,'+
                                                            'iIdDiarioNota, sIdClasificacion, iIdTarea, '+
                                                            'iHermano, eTipoActividad, eAplicaVolumenes, iIdTarea_act, eEstatus, sIdConvenio ) '+

                        'values ( :orden, :fecha, :diario, :turno, :folio, :wbs, :actividad, "ED", :inicio, :fin, :descripcion, :nota, "TE", :diario, :hermano, :tipoActividad, :aplica, :tarea_act, :estatus, :convenio )');

  SQL_BUSCAR_TIPO_ACTIVIDAD : string = 'select eTipoActividad from bitacoradeactividades where iIdActividad = :IdActividad';


  HERMANO_NULL : Shortint = -1;

begin

  try

    //connection.zconnection.startTransaction;

    {$REGION 'Valida'}

    IndexF := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );

    if indexF = -1 then
      raise Exception.Create('Indice del folio no encontrado.');

    IndexA := Cuadre.CATEGORIA[Libro.ActivePage, IndexF].BuscarActividadPorNodoIndex( lstCortes.Selected.Index );

    if IndexA = -1 then
      raise Exception.Create('Indice de la actividad no encontrada');

    IndexAP := Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].IndexPadre( Cuadre.CATEGORIA[ Libro.ActivePage, IndexF ].ACTIVIDADES[ IndexA ].sIdActividad );

    if IndexAP = -1 then
      raise Exception.Create('Indice de la actividad padre no encontrada, no se puede continuar');

    if Cuadre.CATEGORIA[Libro.ActivePage, IndexF].ACTIVIDADES[IndexA].iHermano >= 0 then
      iAnida := Cuadre.CATEGORIA[Libro.ActivePage, IndexF].ACTIVIDADES[IndexA].iHermano
    else
      iAnida := Cuadre.CATEGORIA[Libro.ActivePage, IndexF].ACTIVIDADES[IndexA].iIdActividad;

    {$ENDREGION}

    {$REGION 'Asigna Valores'}

    sFolio := Cuadre.CATEGORIA[Libro.ActivePage, IndexF].sFolio;
    sWbs := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].sWbs;
    sNumeroActividad := Actividad;
    IdActividad := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].iIdActividad;
    IdActividadHermano := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].iHermano;

    if IdActividadHermano = HERMANO_NULL then
      IdActividadHermano := IdActividad;


    {$ENDREGION}

    {$REGION 'Busca IdDiario para la nota'}

    with connection.QryBusca do
    begin
      Active := False;
      SQL.Text := 'select iIdDiario from bitacoradeactividades where sContrato = :orden '+
                  'and dIdFecha = :fecha '+
                  'and sNumeroOrden = :folio '+
                  'and sWbs = :wbs '+
                  'and sNumeroActividad = :actividad '+
                  'and sIdTipoMovimiento = "E" ';

      ParamByName('orden').asstring := global_contrato;
      ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
      ParamByName('folio').AsString := sFolio;
      ParamByName('wbs').AsString := sWbs;
      ParamByName('actividad').AsString := sNumeroActividad;
      Open;
    end;

    if connection.QryBusca.RecordCount = 0 then
      raise Exception.Create('No se encontro iIdDiario de nota para la nueva actividad');

    IdDiarioNota := connection.QryBusca.FieldByName('iIdDiario').AsInteger;

    {$ENDREGION}

    {$REGION 'Buscar Actividad a punto actualizar - por seguridad'}

    with connection.QryBusca do
    begin
      Active := False;
      SQL.Text := 'select iidactividad from bitacoradeactividades '+
                  'where scontrato = :orden '+
                  'and didfecha = :fecha '+
                  'and iiddiario = :diario '+
                  'and snumeroorden = :folio '+
                  'and swbs = :wbs '+
                  'and snumeroactividad = :actividad '+
                  'and shorainicio = :inicio '+
                  'and shorafinal = :final '+
                  'and sidtipomovimiento = "ED"';
      ParamByName('orden').AsString := global_contrato;
      ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
      ParamByName('diario').AsInteger := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].iIdDiario;
      ParamByName('folio').AsString := sFolio;
      ParamByName('wbs').AsString := sWbs;
      ParamByName('inicio').AsString := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].sHInicio;
      ParamByName('final').AsString := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].sHFin;
      ParamByName('actividad').AsString := sNumeroActividad;
      Open;
    end;

    Estatus := '';
    if connection.QryBusca.RecordCount = 0 then
    begin
      MessageDlg('Alguien ya ha registrado realizado un corte a esta actividad no concuenrdan los horarios originales' + #10 +
                 'Se recomienda actualizar el cuadre' , mtInformation, [mbOK], 0);
      raise Exception.Create('');
    end;


    with connection.QryBusca do
    begin
      Active := False;
      SQL.Text := 'select iIdTarea_act, eEstatus from bitacoradeactividades where iIdActividad = :idactividad';
      ParamByName( 'idactividad' ).AsInteger := IdActividadHermano;
      Open;
    end;

    Estatus := connection.QryBusca.FieldByName( 'eEstatus' ).AsString;
    TareaAct := connection.QryBusca.FieldByName('iIdTarea_act').AsInteger;

    {$ENDREGION}

    {$region 'Busca el tipo de actividad'}

    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Text := SQL_BUSCAR_TIPO_ACTIVIDAD;
    connection.QryBusca.ParamByName( 'IdActividad' ).AsInteger := IdActividad;
    connection.QryBusca.Open;

    TipoActividad := connection.QryBusca.FieldByName( 'eTipoActividad' ).AsString;

    {$endregion}

    {$REGION 'Actualiza la actividad con su nueva hora final'}

    with connection.zCommand do
    begin
      Active := False;
      sql.Text := SQL_UPDATE;
      ParamByName('n_hora').AsString := Fin;
      ParamByName('idactividad').AsInteger := IdActividad;
      ParamByName('actividad').AsString := sNumeroActividad;
      ParamByName('orden').asstring := global_contrato;
      ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
      ParamByName('wbs').AsString := sWbs;
      ParamByName('hermano').asinteger := iAnida;
      ExecSQL;
    end;

    {$ENDREGION}

    {$REGION 'Buscar Actividad a punto de crear por seguridad'}

    Application.CreateForm( TfrmBitacoraDepartamental_2, frmBitacoraDepartamental_2 );
    IdDiario := frmBitacoraDepartamental_2.MaximoItem;
    frmBitacoraDepartamental_2.Free;

    with connection.QryBusca do
    begin
      Active := False;
      SQL.Text := 'select iidactividad from bitacoradeactividades '+
                  'where scontrato = :orden '+
                  'and didfecha = :fecha '+
                  'and iiddiario = :diario '+
                  'and snumeroorden = :folio '+
                  'and swbs = :wbs '+
                  'and snumeroactividad = :actividad '+
                  'and shorainicio = :inicio '+
                  'and shorafinal = :final '+
                  'and sidtipomovimiento = "ED" '+
                  'and eTipoActividad = :tipoActividad';
      ParamByName( 'orden' ).AsString := global_contrato;
      ParamByName( 'fecha' ).AsDate := StrToDate( Cuadre.Fecha );
      ParamByName( 'diario' ).AsInteger := IdDiario;
      ParamByName( 'folio' ).AsString := sFolio;
      ParamByName( 'wbs' ).AsString := sWbs;
      ParamByName( 'inicio' ).AsString := InicioCorte;
      ParamByName( 'final' ).AsString := FinCorte;
      ParamByName( 'actividad' ).AsString := sNumeroActividad;
      ParamByName( 'tipoactividad' ).AsString := TipoActividad;
      Open;
    end;

    if connection.QryBusca.RecordCount > 0 then
    begin
      raise exception.Create('Alguien ya ha registrado la actividad ' + sNumeroActividad + ' con horario de '+ InicioCorte+ ' a ' + FinCorte+ ' registrada' + #10 + 'Se recomienda actualizar el cuadre');
    end;

    {$ENDREGION}

    {$REGION 'Crea la nueva actividad'}

    with connection.zCommand do
    begin
      Active := False;
      sql.Text := SQL_INSERT;
      ParamByName( 'orden' ).asstring := global_contrato;
      ParamByName( 'fecha' ).AsDate := StrToDate( Cuadre.Fecha );
      ParamByName( 'diario' ).AsInteger := IdDiario;
      ParamByName( 'turno' ).AsString := global_turno;
      ParamByName( 'folio' ).AsString := sFolio;
      ParamByName( 'wbs' ).AsString := sWbs;
      ParamByName( 'actividad' ).AsString := sNumeroActividad;
      ParamByName( 'inicio' ).AsString := InicioCorte;
      ParamByName( 'fin' ).AsString := FinCorte;
      ParamByName( 'descripcion' ).AsString := mDescripcionCorte.Text;
      ParamByName( 'nota' ).AsInteger := IdDiarioNota;
      if TipoActividad = '' then
         ParamByName( 'tipoActividad' ).AsString := 'desconocido'
      else
         ParamByName( 'tipoActividad' ).AsString := TipoActividad;
      ParamByName( 'aplica' ).AsString := 'No';
      ParamByName( 'tarea_act' ).AsInteger := TareaAct;
      ParamByName( 'estatus' ).AsString := Estatus;
      ParamByName( 'convenio' ).AsString := global_convenio;
      if chkAnida.Checked then
        ParamByName('hermano').AsInteger := iAnida
      else
        ParamByName('hermano').AsInteger := -1;

      ExecSQL;
    end;

    {$ENDREGION}

    {$REGION 'Buscar PK de la actividad ya creada'}

    with connection.QryBusca do
    begin
      Active := False;
      SQL.Text := 'select iidactividad from bitacoradeactividades '+
                  'where scontrato = :orden '+
                  'and didfecha = :fecha '+
                  'and iiddiario = :diario '+
                  'and snumeroorden = :folio '+
                  'and swbs = :wbs '+
                  'and snumeroactividad = :actividad '+
                  'and shorainicio = :inicio '+
                  'and shorafinal = :final '+
                  'and sidtipomovimiento = "ED" '+
                  'and eTipoActividad = :tipoactividad';
      ParamByName('orden').AsString := global_contrato;
      ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
      ParamByName('diario').AsInteger := IdDiario;
      ParamByName('folio').AsString := sFolio;
      ParamByName('wbs').AsString := sWbs;
      ParamByName('inicio').AsString := InicioCorte;
      ParamByName('final').AsString := FinCorte;
      ParamByName('actividad').AsString := sNumeroActividad;
      ParamByName( 'tipoactividad' ).AsString := TipoActividad;
      Open;
    end;

    {$ENDREGION}

    {$REGION 'Inserta nueva actividad en el Cuadre virtualizado '}

    Corte := TActividad.Create();
    Corte.sIdActividad := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].sIdActividad;
    Corte.sWbs := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].sWbs;
    Corte.iIdActividad := connection.QryBusca.FieldByName('iidactividad').AsInteger;
    Corte.sHInicio := InicioCorte;
    Corte.sHFin := FinCorte;
    Corte.iIdDiario := IdDiario;
    if IndexA = Length( Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES ) - 1  then
      Corte.iRow := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].iRow + 1
    else
      Corte.iRow := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa + 1].iRow;
    Corte.iNodoCorte := -1;
    Corte.iTarea := IdDiario;
    Corte.IsPadre := False;
    Corte.NuevoCorte := True;
    if chkAnida.Checked then
    begin
      Corte.iHermano := iAnida;
      if Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].iInicioConjunto < 0 then
        Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].iInicioConjunto := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].iRow;
      Corte.iInicioConjunto := Cuadre.CATEGORIA[Libro.ActivePage, indexf].ACTIVIDADES[indexa].iInicioConjunto;
    end;

    {$ENDREGION}

    {$REGION 'Ordena Cuadre Virtualizado y Actividades Padre y Crea la nueva actividad'}

    Cuadre.AvanzarFoliosDesde( IndexF + 1 );

    for iHoja := 0 to Length( Cuadre.CATEGORIA ) - 1 do
    begin
      Cuadre.CATEGORIA[ iHoja, indexf ].ACTIVIDADES[indexa].sHInicio := Inicio;
      Cuadre.CATEGORIA[ iHoja, indexf ].ACTIVIDADES[indexa].sHFin := Fin;

      SetLength( Cuadre.CATEGORIA[iHoja, IndexF].ACTIVIDADES, Length( Cuadre.CATEGORIA[iHoja, indexf].ACTIVIDADES ) + 1 );
      Cuadre.CATEGORIA[ iHoja, indexf ].ACTIVIDADES[ Length( Cuadre.CATEGORIA[iHoja, indexf].ACTIVIDADES ) - 1 ] := TActividad.Create();
      Cuadre.CATEGORIA[ iHoja, IndexF ].AvanzarActividades( IndexA + 1 );
      Cuadre.CATEGORIA[ iHoja, IndexF ].InsertarActividad( Corte, IndexA + 1 );

      Cuadre.CATEGORIA[ iHoja, indexf ].UpdateRange;
      Cuadre.CATEGORIA[ iHoja, indexf ].UpdateActRows;

      if chkAnida.Checked then
        Cuadre.CATEGORIA[iHoja, indexf].ACTIVIDADES[indexa].iHermano := iAnida;

      
      Cuadre.CATEGORIA[iHoja, indexf].ActualizarCountActividades;
    end;

    {$ENDREGION}

    {$REGION 'Genera excel con nueva actividad'}

    {$REGION 'Inicializa'}

    GetTempPath(SizeOf(global_TempPath), global_TempPath);
    sArchivo := global_TempPath+'inteliCuad_intelcode.xls';

    if FileExists( sArchivo ) then
      DeleteFile( sArchivo );

    Libro.SaveToFile( sArchivo );

    try
      Excel := CreateOleObject('Excel.Application');
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
      Excel.Visible := EXCEL_VISIBLE;
      Excel.Workbooks.Open( sArchivo );
      Book := Excel.Workbooks[1];
    except
      on e:Exception do
      begin
        MessageDlg('No se puede continuar verifique tener instalada la suite de Microsoft Office', mtInformation, [mbOK], 0);
        try
          Excel.Quit;
        except
          ;
        end;

        Exit;
      end;
    end;

    {$ENDREGION}

    {$region 'Buscar hermanos'}

    with connection.QryBusca do
    begin
      active := False;
      sql.text := 'select '+
                    'snumeroactividad, '+
                    'min( shorainicio ) as sInicioConjunto , '+
                    'max(shorafinal) as sFinConjunto, '+
                    'count( ihermano ) as iRows '+#10+
                  'from bitacoradeactividades '+#10+
                  'where sContrato = :orden '+
                  '&& didfecha = :fecha '+
                  '&& snumeroorden = :folio '+
                  '&& sidtipomovimiento = "ED" '+
                  '&& ihermano = :hermano ';
      parambyname('orden').AsString := global_contrato;
      parambyname('fecha').AsDate := StrToDate( Cuadre.Fecha );
      parambyname('folio').asstring := Cuadre.CATEGORIA[0, IndexF].sFolio;
      parambyname('hermano').AsInteger := Cuadre.CATEGORIA[0, indexf].ACTIVIDADES[indexa].iHermano;
      open;
    end;

    {$endregion}

    for iHoja := 1 to 2 do
    begin

      Hoja := Book.Sheets[iHoja];
      Hoja.Select;
      Hoja.Unprotect;

      {$REGION 'Calcula Duracion'}
      
      {$REGION 'Nueva Actividad'}

      sFactorHorasAc := sfnRestaHoras(Fin, Inicio);
      getHM(sFactorHorasAc, iHorasAc, iMinutosAc);
      iHorasAc := iHorasAc / 24;
      iMinutosAc := ( iMinutosAc / 24 ) / 60;

      if rbRedondear.Checked then
        dDuracionAc := ( Redondear((iHorasAc + iminutosAc), Round( clcRT.Value ) ) );
      if rbTruncar.Checked then
        dDuracionAc := ( Truncar((iHorasAc + iminutosAc), Round( clcRT.Value ) ) );

      {$ENDREGION}

      {$REGION 'Actividad Anterior'}

      sFactorHoras := sfnRestaHoras(Corte.sHFin, Corte.sHInicio);
      getHM(sFactorHoras, iHoras, iMinutos);
      iHoras := iHoras / 24;
      iMinutos := ( iMinutos / 24 ) / 60;

      if rbRedondear.Checked then
        dDuracion := ( Redondear((iHoras + iminutos), Round( clcRT.Value ) ) );
      if rbTruncar.Checked then
        dDuracion := ( Truncar((iHoras + iminutos), Round( clcRT.Value ) ) );

      {$ENDREGION}

      {$ENDREGION}

      if chkAnida.Checked then
      begin
        with connection.qrybusca do
        begin
          Rang := Excel.Range['A'+inttostr(Corte.iInicioConjunto)+':A'+inttostr(Corte.iRow)];
          Rang.MergeCells := False;
        end;
      end;

      Excel.Range['C'+inttostr(Corte.iRow - 1)].Value := Inicio;
      Excel.Range['D'+inttostr(Corte.iRow - 1)].Value := Fin;
      Excel.Range['E'+inttostr(Corte.iRow - 1)].Formula := '=D'+inttostr(Corte.iRow - 1)+' - C'+inttostr(Corte.iRow - 1);

      Excel.Rows[Corte.iRow].insert;
      Excel.Rows[Corte.iRow - 1].Copy;
      Excel.Rows[Corte.iRow].PasteSpecial;
      Excel.Range['A'+inttostr(Corte.iRow)+':'+'E'+inttostr(Corte.iRow)].Interior.colorindex := 35;
      Excel.Range['A'+inttostr(Corte.iRow)+':'+'E'+inttostr(Corte.iRow)].Value := '';
      Excel.Range['A'+inttostr(Corte.iRow)].Value := Corte.sIdActividad;
      Excel.Range['B'+inttostr(Corte.iRow)].WrapText := False;
      Excel.Range['B'+inttostr(Corte.iRow)].Value := DescripcionCorte;
      Excel.Range['C'+inttostr(Corte.iRow)].Value := Corte.sHInicio;
      Excel.Range['D'+inttostr(Corte.iRow)].Value := Corte.sHFin;
      Excel.Range['E'+inttostr(Corte.iRow)].Formula := '=D'+inttostr(Corte.iRow)+' - C'+inttostr(Corte.iRow);
      Excel.Range['E'+inttostr(Corte.iRow)].NumberFormat := '0.00';
      dDuracion := Excel.Cells[Corte.iRow, 5].Text;
      dDuracionAc := Excel.Cells[Corte.iRow - 1, 5].Text;

      Cuadre.CATEGORIA[iHoja - 1, IndexF].ACTIVIDADES[IndexA].dDuracion := dDuracionAc;
      Cuadre.CATEGORIA[iHoja - 1, IndexF].ACTIVIDADES[IndexA + 1].dDuracion := dDuracion;

      iMCount := 6;
      for Count := 0 to Length( Cuadre.CATEGORIA[iHoja - 1, indexf].MOE ) - 1 do
      begin
        Excel.Range[ColumnaNombre(iMCount)+inttostr(Corte.iRow)].Value := '';
        Excel.Range[ColumnaNombre(iMCount)+inttostr(Corte.iRow)].interior.colorindex := 2;
        Excel.Range[ColumnaNombre(iMCount+1)+inttostr(Corte.iRow)].interior.colorindex := 15;
        Inc( iMCount, 2 );
      end;

      if chkAnida.Checked then
      begin
        with connection.qrybusca do
        begin
          Rang := Excel.Range['A'+inttostr(Corte.iInicioConjunto)+':A'+inttostr(Corte.iInicioConjunto + Fieldbyname('iRows').AsInteger - 1)];
          Rang.MergeCells := False;
          Rang.MergeCells := True;
          Rang.Value := Corte.sIdActividad + #10 + Fieldbyname('sInicioConjunto').asstring+ ' - ' + Fieldbyname('sFinConjunto').asstring;
        end;
      end;

      Excel.Range['A1'].Select;
      RegenerarSumasMOE(Excel, iHoja - 1);
      Hoja.Protect(True, True, True);
    end;

    Hoja := Book.Sheets[1];
    Hoja.Select;
    Book.SaveAs( sArchivo );
    Libro.LoadFromFile( sArchivo );

    Excel.Quit;

    if FileExists( sArchivo ) then
      DeleteFile( sArchivo );

  {$ENDREGION}

    //connection.zConnection.Commit;
    Cuadre.UpdateAllRanges;
    ErrorCorte := False;
  except
    on e:Exception do
    begin
      MessageDlg('Ha ocurrido el sigueinte error al crear la nueva actividad: '+e.Message+#10+ 'Se revertiran los cambios', mtInformation, [mbOk], 0);
      ErrorCorte := True;
      //connection.zconnection.startback;

      try
        Excel.Quit;
      except
      end;

    end;
  end;
end;

procedure tfrmCuadreNormal.RegenerarSumasMOE(var Excel: Variant; IndiceTipo : TCategoriaIndex);
var
  IndexT : TCategoriaIndex;
  IndexF,
  IndexFI : TFolioIndex;
  IndexA : TActividadIndex;
  IndexAP : TActividadPadreIndex;
  IndexM : TMoeIndex;

  Col : TColumnaExcel;
  Folio : TSimpleFolio;
  
  Inicio : TInicioFolio;
  Fin : TFinFolio;

  Rango : Variant;
  Formula : string;
begin
  Col := TColumnaExcel.Create();
  Folio := TSimpleFolio.Create;
  Formula := '';
  IndexT := IndiceTipo;
  
  for IndexM := 0 to Length( Cuadre.CATEGORIA[IndexT, 0].MOE ) - 1 do
  begin
    Formula := '=(';
    for IndexF := 0 to Length( Cuadre.CATEGORIA[IndexT] ) - 1 do
    begin
      Col.iColumna := Cuadre.CATEGORIA[IndexT, IndexF].MOE[IndexM].iCol;
      for IndexAP := 0 to Length( Cuadre.CATEGORIA[ IndexT, IndexF ].ACTIVIDADES_PADRES ) - 1 do
        Formula := Formula + Col.Columna_ + IntToStr( Cuadre.CATEGORIA[ IndexT, IndexF ].ACTIVIDADES_PADRES[ IndexAP ].Fila ) + ' + ';
    end;

    Formula := Trim( Formula );
    Formula[ Length(Formula) ] := ')';

    for IndexF := 0 to Length( Cuadre.CATEGORIA[IndexT] ) - 1 do
    begin
      Folio.Inicio := Cuadre.CATEGORIA[IndexT, IndexF].Fila + 2;
      Rango := Excel.Range[Col.Columna+InttoStr(Folio.Inicio - 5)];
      Rango.Formula := Formula;
    end;
         
  end;

end;

procedure TfrmCuadreNormal.EliminarHorario( Fila : TFilaExcel; hFinal : string ; IndiceF : TFolioIndex ; IndiceA : TActividadIndex; IndicePadre : TActividadPadreIndex );
var
  Excel,
  Book,
  Hoja,
  Rango : Variant;

  iHoja : integer;

  IndexF : TFolioIndex;
  IndexAP : TActividadPadreIndex;
begin

  {$REGION 'Inicia'}
  GetTempPath(SizeOf(global_TempPath), global_TempPath);
  sArchivo := global_TempPath+'inteliCuad_intelcode_Save.xls';
  if not Length( Cuadre.CATEGORIA ) > 0 then
    Exit;

  if FileExists( sArchivo ) then
  begin
    DeleteFile( sArchivo );
  end;
  Libro.SaveToFile(sArchivo);
  
  try
    Excel := CreateOleObject('Excel.Application');
    Book := Excel.Workbooks.Open( sArchivo );
    Hoja := Book.Sheets[1];
    Excel.DisplayAlerts := False;
    Excel.ScreenUpdating := True;
    Excel.Visible := EXCEL_VISIBLE;
  except
    MessageDlg('No se puede continuar verifique tener instalada la suite de Microsoft Office', mtInformation, [mbOK], 0);
    Exit;
  end;

  {$ENDREGION}

  Cuadre.RetrocederFoliosDesde( IndiceF + 1 );

  for iHoja := 1 to 2 do
  begin
    Hoja := Book.Sheets[iHoja];
    Hoja.Select;
    Excel.Activesheet.Unprotect;
    Excel.Rows[Fila].Delete;
    Excel.Range['D'+inttostr( Cuadre.CATEGORIA[iHoja - 1, IndiceF].ACTIVIDADES[IndiceA].iRow - 1 )].Value := hFinal;
    Cuadre.CATEGORIA[ iHoja - 1, IndiceF ].RetrocederActividades( IndiceA ) ;
    Cuadre.CATEGORIA[ iHoja - 1, IndiceF ].ACTIVIDADES[IndiceA - 1].sHFin := hFinal;

    SetLength( Cuadre.CATEGORIA[iHoja - 1, IndiceF].ACTIVIDADES, Length( Cuadre.CATEGORIA[iHoja - 1, IndiceF].ACTIVIDADES ) -1 );
    Hoja.Protect(True, True, True);
  end;

  Cuadre.UpdateAllRanges;

  Hoja := Book.Sheets[1];
  Hoja.Select;
  Book.SaveAs( sArchivo );
  Libro.LoadFromFile( sArchivo );

  Excel.Quit;

  if FileExists( sArchivo ) then
    DeleteFile( sArchivo );
    
end;


//Martin
procedure TfrmCuadreNormal.CambiaVistaCuadre(Tipo: string);
var
  Excel,
  Book,
  Hoja,
  Rango : TExcelInstance;

  IndexP,
  IndexF : TFolioIndex;

  IndexA : TActividadIndex;
  IndexAP : TActividadPadreIndex;

begin
  try
    try

      {$REGION 'Inicia'}
      GetTempPath(SizeOf(global_TempPath), global_TempPath);
      sArchivo := global_TempPath+'inteliCuad_intelcode_Save.xls';
      if not Length( Cuadre.CATEGORIA ) > 0 then
        Exit;

      if FileExists( sArchivo ) then
      begin
        DeleteFile( sArchivo );
      end;
      Libro.SaveToFile(sArchivo);
  
      try
        Excel := CreateOleObject('Excel.Application');
        Book := Excel.Workbooks.Open( sArchivo );
        Hoja := Book.Sheets[1];
        Excel.DisplayAlerts := False;
        Excel.ScreenUpdating := True;
        Excel.Visible := EXCEL_VISIBLE;
      except
        MessageDlg('No se puede continuar verifique tener instalada la suite de Microsoft Office', mtInformation, [mbOK], 0);
        Exit;
      end;

    

      {$ENDREGION}

      if Tipo = 'horarios' then
      begin
        for IndexP := 0 to Length( Cuadre.CATEGORIA ) - 1 do
        begin
          Hoja := Book.Sheets[ IndexP + 1 ].Select;
          Excel.ActiveSheet.Unprotect;

          for IndexF := 0 to Length( Cuadre.CATEGORIA[ IndexP ] ) - 1 do
          begin
            Rango := Excel.Rows[ IntToStr( Cuadre.CATEGORIA[ IndexP, IndexF ].iInicio )+':'+IntToStr( Cuadre.CATEGORIA[ IndexP, IndexF ].iFin ) ];
            Rango.EntireRow.Hidden := False;

            for IndexAP := 0 to Length( Cuadre.CATEGORIA[ IndexP, IndexF ].ACTIVIDADES_PADRES ) - 1 do
            begin
              Rango := Excel.Rows[ Cuadre.CATEGORIA[ IndexP, IndexF ].ACTIVIDADES_PADRES[ IndexAP ].Fila ];
              Rango.EntireRow.Hidden := True;
            end;

          end;

        end;
        
      end;

      if Tipo = 'actividades' then
      begin
        for IndexP := 0 to Length( Cuadre.CATEGORIA ) - 1 do
        begin
          Hoja := Book.Sheets[ IndexP + 1 ].Select;
          Excel.ActiveSheet.Unprotect;

          for IndexF := 0 to Length( Cuadre.CATEGORIA[ IndexP ] ) - 1 do
          begin
            Rango := Excel.Rows[ IntToStr( Cuadre.CATEGORIA[ IndexP, IndexF ].iInicio )+':'+IntToStr( Cuadre.CATEGORIA[ IndexP, IndexF ].iFin ) ];
            Rango.EntireRow.Hidden := True;

            for IndexAP := 0 to Length( Cuadre.CATEGORIA[ IndexP, IndexF ].ACTIVIDADES_PADRES ) - 1 do
            begin
              Rango := Excel.Rows[ Cuadre.CATEGORIA[ IndexP, IndexF ].ACTIVIDADES_PADRES[ IndexAP ].Fila ];
              Rango.EntireRow.Hidden := False;
            end;
          end;

        end;
        
      end;

      if Tipo = 'general' then
      begin
        for IndexP := 0 to Length( Cuadre.CATEGORIA ) - 1 do
        begin
          Hoja := Book.Sheets[ IndexP + 1 ].Select;
          Excel.ActiveSheet.Unprotect;

          for IndexF := 0 to Length( Cuadre.CATEGORIA[ IndexP ] ) - 1 do
          begin
            Rango := Excel.Rows[ IntToStr( Cuadre.CATEGORIA[ IndexP, IndexF ].iInicio )+':'+IntToStr( Cuadre.CATEGORIA[ IndexP, IndexF ].iFin ) ];
            Rango.EntireRow.Hidden := False
          end;

        end;
        
      end;

      for IndexP := 1 to Book.Sheets.Count - 1 do
      begin
        Book.Sheets[ IndexP ].Select;
        Book.ActiveSheet.Protect( True, True, True );
      end;
      
    except
      on e:Exception do
        MessageDlg( 'Ha ocurrido un error al cambiar la vista', mtInformation, [mbOK], 0 );
    end;
  finally
    Book.Save;
    Libro.LoadFromFile( sArchivo );
    Book.Close;
    Excel.Quit;
  end;
end;

procedure TfrmCuadreNormal.ConsultarActividades;
begin
          qrActividades.Active := False;
          qrActividades.SQL.Clear;
          qrActividades.SQL.Add('select '+
                  'ba.iIdActividad, '+
                  'ba.iIdDiario, '+
                  'ba.sWbs, '+
                  'ba.sNumeroOrden, '+
                  'ba.sIdClasificacion, '+
                  'ba.sTipoObra, '+
                  'ba.sNumeroActividad, '+
                  'ba.sHoraInicio, '+
                  'ba.sHoraFinal, '+
                  'ot.sIdPlataforma, '+
                  'ot.sIdPernocta, '+
                  'ba.mDescripcion, '+
                  'ba.iIdTarea, '+
                  'ba.iHermano, '+
                  'ba.sColor, '+

                  '( select count( b.sNumeroActividad ) '+
                  '  from bitacoradeactividades b '+
                  '  where b.sContrato = ba.sContrato '+
                  '  && b.dIdFecha = ba.dIdFecha '+
                  '  && b.sNumeroOrden = ba.sNumeroOrden '+
                  '  && b.sNumeroActividad = ba.sNumeroActividad '+
                  '  && b.sIdTipoMovimiento = "ED" ) as iActividades, '+

                  '( select count( bb.iHermano > 0 ) '+
                  '  from bitacoradeactividades bb    '+
                  '  where bb.sContrato = ba.sContrato '+
                  '  && bb.dIdFecha = ba.dIdFecha '+
                  '  && bb.sNumeroOrden = ba.sNumeroOrden '+
                  '  && bb.sNumeroActividad = ba.sNumeroActividad '+
                  '  && bb.iHermano = ba.iHermano '+
                  '  && bb.iHermano >= 0 '+
                  '  && bb.sIdTipoMovimiento = "ED"  ) as iHermanosCount, '+

                  'ifnull((select bp.laplicapernocta from' + #13#10 +
                  'bitacoradepersonal bp' + #13#10 +
                  'where ba.sContrato=bp.sContrato and ba.sNumeroOrden=bp.sNumeroOrden' + #13#10 +
                  'and ba.iIdDiario=bp.iIdDiario and ba.swbs=bp.swbs' + #13#10 +
                  'and ba.snumeroactividad=bp.sNumeroActividad' + #13#10 + 
                  'and ba.iIdActividad=bp.iIdActividad limit 1 ),"Si") as aplicapernocta '  +

                'from bitacoradeactividades ba '+

                'inner join reportediario rd     '+
                '  on ( rd.sOrden = ba.sContrato '+
                '    and rd.dIdFecha = ba.dIdFecha ) '+

                'inner join ordenesdetrabajo ot        '+
                '  on ( ot.sContrato = ba.sContrato    '+
                '    and ot.sNumeroOrden = ba.sNumeroOrden ) '+

                'inner join pernoctan pc '+
                '  on ( ot.sIdPernocta = pc.sIdPernocta ) '+

                'inner join plataformas pl '+
                '  on ( ot.sIdPlataforma = pl.sIdPlataforma ) '+

                'where ba.sContrato = :orden '+
                'and ba.dIdFecha = :fecha ');




          if connection.zCommand.FieldByName('lOrdenaxHorario').AsString = 'No' then
          begin
             qrActividades.SQL.Add('and ba.sIdTipoMovimiento = "ED" ');
            {if chkOrdenado.Checked  then
                qrActividades.SQL.Add('order by ba.iIdActividad ;')
             else  }
                qrActividades.SQL.Add('order by ot.iOrden, ba.sNumeroActividad, ba.sHoraInicio;');
          end
          else
          begin
             qrActividades.SQL.Add('and ba.sIdTipoMovimiento = "ED" ');
            { if chkOrdenado.Checked  then
                qrActividades.SQL.Add('order by ba.iIdActividad ;')
             else }
                qrActividades.SQL.Add('order by ot.iOrden, ba.sHoraInicio, ba.sNumeroActividad ;');
          end;
          qrActividades.ParamByName('orden').AsString := global_contrato;
          qrActividades.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
          qrActividades.Open;
end;


end.
