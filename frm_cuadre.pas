unit frm_cuadre;

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
  AdvStickyPopupMenu, dxSkinsdxBarPainter, dxBar, AdvSmoothPopup,
  cxStyles, dxSkinscxPCPainter, cxListView, cxContainer, cxEdit, dxSkinBlack,
  dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom,
  dxSkinDarkSide, dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle, dxSkinFoggy,
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
  cxInplaceContainer, LabelEdit, CurvyControls, cxSpinEdit,
  cxTimeEdit, cxListBox, ExcelXP, OleServer;

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
      sIdCategoria : string;

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
  TfrmCuadre = class(TForm)

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
    CxLbl4: TcxLabel;
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
            
    Cuadre : TCuadre;
    ActividadCorte : TActividad;

    ListaCampos: TStringList;
    ePintando,
    ErrorCorte : Boolean;

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

  public
    { Public declarations }
  end;

var
  frmCuadre: TfrmCuadre;

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
                                          'sNumeroOrden, '+
                                          'iIdActividad, '+
                                          'iIdTarea, '+
                                          'sAgrupaPersonal, '+
                                          'sHoraInicioG, '+
                                          'sHoraFinalG ) '+#10+

                                    'values (:orden, '+
                                        ':fecha, '+
                                        ':iddiario, '+
                                        ':ItemOrden, '+
                                        ':idrecurso, '+
                                        '"PU", '+
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
                                        ':folio, '+
                                        ':idactividad, '+
                                        ':tarea, '+
                                        ':Categoria, '+
                                        ':hinicio, '+
                                        ':hfinal ) '

                                        ,

                                        'insert into bitacoradeequipos (sContrato, '+
                                        'dIdFecha, '+
                                        'iIdDiario, '+
                                        'iItemOrden, '+
                                        'sIdEquipo, '+
                                        'sDescripcion, '+
                                        'sIdPernocta, '+
                                        'sTipoObra, '+
                                        'sHoraInicio, '+
                                        'sHoraFinal, '+
                                        'dCantidad, '+
                                        'sWbs, '+
                                        'sNumeroActividad, '+
                                        'dCantHH, '+
                                        'sNumeroOrden, '+
                                        'iIdActividad, '+
                                        'iIdTarea, '+
                                        'sHoraInicioG, '+
                                        'sHoraFinalG ) '+#10+

                                    'values (:orden, '+
                                        ':fecha, '+
                                        ':iddiario, '+
                                        ':ItemOrden, '+
                                        ':idrecurso, '+
                                        ':descripcion, '+
                                        ':pernocta, '+
                                        '"PU", '+
                                        ':hinicio, '+
                                        ':hfinal, '+
                                        ':cantidad, '+
                                        ':wbs, '+
                                        ':actividad, '+
                                        ':cantidadhh, '+
                                        ':folio, '+
                                        ':idactividad, '+
                                        ':tarea, '+
                                        ':hinicio, '+
                                        ':hfinal )');

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
                                                       'and m1.dIdFecha <= :fecha ) group by p.sIdPersonal'

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
                                                      'and m1.dIdFecha <= :fecha )');

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
  Excel.Visible :=  False;
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
procedure TfrmCuadre.ComprobarSuma( Edit_Clean : Boolean );
var
  iValor,
  iFCount,
  iMoeIndex,
  Total : Integer;

  iSuma,
  iResta : Double;

  Rango : TcxSSCellObject;

begin
  if ePintando then
    Exit;

  try
    if Length( Cuadre.CATEGORIA[Libro.ActivePage] ) = 0 then
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

    for iFCount := 0 to Length( Cuadre.CATEGORIA[Libro.ActivePage] ) - 1 do
    begin
      if ( Libro.ActiveSheet.ActiveCell.Y + 1 >= Cuadre.CATEGORIA[Libro.ActivePage,iFCount].iInicio ) and ( Libro.ActiveSheet.ActiveCell.Y <= Cuadre.CATEGORIA[Libro.ActivePage,iFCount].iFin ) then
      begin
        Rango := Libro.ActiveSheet.GetCellObject( Libro.ActiveSheet.ActiveCell.X, Cuadre.CATEGORIA[Libro.ActivePage,iFCount].iInicio - 6 );
        Rango.Text;
        iSuma := StrToFloat( VarToStr( Rango.CellValue ) );
        iSuma := iSuma - iResta;
        iMoeIndex :=Cuadre.CATEGORIA[Libro.ActivePage,iFCount].BuscarCategoria( Libro.ActiveSheet.ActiveCell.X + 1);

        if iMoeIndex = -1 then
        begin
          raise Exception.Create('Error en la aplicación informe al administrador del sistema');
          Exit;
        end;

        Cuadre.Cambios := True;
        Cuadre.Guardado := False;


        case CompareValue(iSuma, Cuadre.CATEGORIA[Libro.ActivePage,iFCount].MOE[iMoeIndex].iAbordo, 0.2 ) of
          LessThanValue    : Rango.Style.Brush.BackgroundColor := 41;
          EqualsValue      : Rango.Style.Brush.BackgroundColor := 42;
          GreaterThanValue : Rango.Style.Brush.BackgroundColor := 45;
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


procedure TfrmCuadre.CortarestaActividad1Click(Sender: TObject);
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

//Saul
procedure TfrmCuadre.CxBtnCancelClick(Sender: TObject);
begin
  if Assigned(FindComponent('frmCortes')) then
    TForm(FindComponent('FrmCortes')).Close;
end;

procedure TfrmCuadre.CxBtnCortarClick(Sender: TObject);
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

procedure TfrmCuadre.cxButton3Click(Sender: TObject);
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

procedure TfrmCuadre.cxButton4Click(Sender: TObject);
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

procedure TfrmCuadre.CxColumnSeleccionarPropertiesEditValueChanged(
  Sender: TObject);
begin
//  CdCortesSugeridos.Edit;
//  CdCortesSugeridos.FieldByName('Incluir').AsBoolean := not CdCortesSugeridos.FieldByName('Incluir').AsBoolean;
//  CdCortesSugeridos.Post;
end;

//Saul
procedure TfrmCuadre.DefinirCortes(DatasetOrigen: TClientDataSet;
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



function TfrmCuadre.detectarCruces(var cdDatosBuscar,
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

procedure TfrmCuadre.AbrirCuadreexistente1Click(Sender: TObject);
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

procedure TfrmCuadre.Ajustes1Click(Sender: TObject);
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

procedure TfrmCuadre.btnActividadesClick(Sender: TObject);
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

procedure TfrmCuadre.btnAplicaAjusteClick(Sender: TObject);
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

procedure TfrmCuadre.btnCortarClick(Sender: TObject);
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

procedure TfrmCuadre.btnGuardarClick(Sender: TObject);
var
   Form : TForm;
   Respuesta : Boolean;
begin
  if Length( Cuadre.CATEGORIA[Libro.ActivePage] ) > 0 then
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

procedure TfrmCuadre.btnPintarClick(Sender: TObject);
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

          qrActividades.Active := False;
          qrActividades.ParamByName('orden').AsString := global_contrato;
          qrActividades.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
          qrActividades.Open;

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
            Libro.&Protected := True;

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

procedure TfrmCuadre.btnSaveClick(Sender: TObject);
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

procedure TfrmCuadre.cbbOtsPropertiesChange(Sender: TObject);
begin
  if Length( Trim( cbbOts.EditText ) ) > 0 then
  begin
    qrReportes.Active := False;
    qrReportes.ParamByName('orden').AsString := cbbOts.EditValue;
    qrReportes.Open;
  end;
end;

procedure TfrmCuadre.cbbPernoctasPropertiesChange(Sender: TObject);
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

procedure TfrmCuadre.cbbPlataformasPropertiesChange(Sender: TObject);
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

procedure TfrmCuadre.cbbReportesPropertiesCloseUp(Sender: TObject);
begin
  if Length( Trim( cbbReportes.EditText ) ) > 0 then
  begin
    qrFolios.Active := False;
    qrFolios.ParamByName('orden').AsString := global_contrato;
    try
      qrFolios.ParamByName('fecha').AsDate := cbbReportes.EditValue;
    except
      ;
    end;
    qrFolios.Open;

    Application.ProcessMessages;

    btnPintar.Click;
  end;
end;

procedure TfrmCuadre.cbbVistaPropertiesCloseUp(Sender: TObject);
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

procedure TfrmCuadre.clcAjustePropertiesChange(Sender: TObject);
begin
  btnAplicaAjuste.Enabled := True;
end;

procedure TfrmCuadre.Diagrama();
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
    pnlEstructura.Parent := frmCuadre;
    pnlEstructura.Align := alNone;
    pnlEstructura.Width := 0;
    pnlEstructura.Height := 0;
    pnlEstructura.Left := 0;
    pnlEstructura.Top := 0;
    pnlEstructura.Visible := False;
    Form.Free;
  end;
end;

procedure TfrmCuadre.Foliosexistentesenelcuadre1Click(Sender: TObject);
begin
  Diagrama();
end;

procedure TfrmCuadre.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmCuadre.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
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

procedure TfrmCuadre.FormCreate(Sender: TObject);
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

      btnGuardar.Visible := LowerCase( global_estado_reporte ) = 'pendiente';
      Self.Caption := 'Cuadre - ' + FormatDateTime( 'YYYY-MM-DD', global_fecha ) + global_estado_reporte ;

    finally
      Screen.Cursor := Cursor;
    end;
  except
    on e: Exception do
      MessageDlg('Ha ocurrido un error inesperado, informa al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TfrmCuadre.FormShow(Sender: TObject);
begin
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
end;

//Martin
procedure TfrmCuadre.getHM(cadena: string; var h, m: Double);
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

procedure TfrmCuadre.grp1Click(Sender: TObject);
begin

end;

//Saul
procedure TfrmCuadre.HorasToDecimal(var DataActividades: TZReadOnlyQuery; var DataDestino: TClientDataSet; inverso: Boolean);
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

procedure TfrmCuadre.LibroClearCells(Sender: TcxSSBookSheet;
  const ACellRect: TRect; var UseDefaultStyle, CanClear: Boolean);
begin
  ComprobarSuma( False );
end;

procedure TfrmCuadre.LibroEndEdit(Sender: TObject);
begin
  ComprobarSuma( True );
end;

procedure TfrmCuadre.LibroSetSelection(Sender: TObject; ASheet: TcxSSBookSheet);
var
  activeCelda: TcxSSCellObject;

  iIndexf,
  iIndexm,
  iindexa : integer;
begin
  cbbPlataformas.Visible := Libro.ActivePage = 0;
  cxLabel7.Visible := cbbPlataformas.Visible;


  if ( Libro.ActiveSheet.ActiveCell.X > 4 ) and ( Libro.ActiveSheet.ActiveCell.Y > 3 ) then
  begin
    if Length( Cuadre.CATEGORIA[Libro.ActivePage] ) > 0 then
    begin
      iindexf := Cuadre.BuscaFolio(Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1);
      if iindexf = -1 then
        exit;

      iindexa := Cuadre.BuscarActividad( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );
      if iindexa = -1 then
        exit;

      cxMemo1.Text := '';

      activeCelda := Libro.ActiveSheet.GetCellObject(0, Libro.ActiveSheet.ActiveCell.y);
      CxTextEdtActividad.Text := activeCelda.CellValue;

      activeCelda := Libro.ActiveSheet.GetCellObject(1, Libro.ActiveSheet.ActiveCell.y);
      cxMemo1.Text := activeCelda.CellValue;

      CxTextEdtinicio.Text := Cuadre.CATEGORIA[Libro.ActivePage, iindexf].ACTIVIDADES[iindexa].sHInicio;

      CxTextEdtTermino.Text := Cuadre.CATEGORIA[Libro.ActivePage, iindexf].ACTIVIDADES[iindexa].sHFin;

      activeCelda := Libro.ActiveSheet.GetCellObject(4, Libro.ActiveSheet.ActiveCell.y);
      CxTextEdtDuracion.Text := activeCelda.CellValue;

      activeCelda := Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.x, 0);
      CxTextEdt1.Text := activeCelda.CellValue;

      activeCelda := Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.x, 1);
      CxTextEdtSolicitado.Text := activeCelda.CellValue;

      activeCelda := Libro.ActiveSheet.GetCellObject(Libro.ActiveSheet.ActiveCell.x, 2);
      CxTextEdtaBordo.Text := activeCelda.CellValue;

      iIndexf := Cuadre.BuscaFolio( Libro.ActivePage, Libro.ActiveSheet.ActiveCell.Y + 1 );

      if iIndexf >= 0 then
      begin
        CxLblFolio.Caption := Cuadre.CATEGORIA[ Libro.ActivePage, iindexf ].sFolio;
        iIndexm := Cuadre.CATEGORIA[Libro.ActivePage, iIndexf].BuscarCategoria( Libro.ActiveSheet.ActiveCell.X + 1 , 0);

        if Length(Cuadre.CATEGORIA[Libro.ActivePage, iIndexF].MOE) > 0 then
        begin
          if iIndexm >= 0 then
          begin
            activeCelda := Libro.ActiveSheet.GetCellObject( Cuadre.CATEGORIA[Libro.ActivePage, iIndexf].MOE[iIndexm].iCol - 1, 3 );
            CxLblCategoria.Caption := 'Categoria: '+ Cuadre.CATEGORIA[Libro.ActivePage, iindexf].MOE[iindexm].sIdRecurso + ' - ' + VarToStr( activeCelda.CellValue );

            activeCelda := Libro.ActiveSheet.GetCellObject( Cuadre.CATEGORIA[Libro.ActivePage, iIndexf].MOE[iIndexm].iCol, Libro.ActiveSheet.ActiveCell.Y );
              txtHH.Text := VarToStr( activeCelda.CellValue );

            cbbPernoctas.EditValue := Cuadre.CATEGORIA[Libro.ActivePage, iIndexf].MOE[iIndexm].sPernocta;
            cbbPlataformas.EditValue := Cuadre.CATEGORIA[Libro.ActivePage, iIndexf].MOE[iIndexm].sPlataforma;
          end
          else
            CxLblCategoria.Caption := '';
        end;
      end
      else
        CxLblFolio.Caption := '';
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

procedure TfrmCuadre.lstCortesChange(Sender: TObject; Node: TTreeNode);
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

procedure TfrmCuadre.lstCuentasClick(Sender: TObject);
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

procedure TfrmCuadre.popHojaPopup(Sender: TObject);
begin
  Foliosexistentesenelcuadre1.Enabled := Length( Cuadre.CATEGORIA[Libro.ActivePage] ) > 0
end;

procedure TfrmCuadre.Eliminarestaactividad1Click(Sender: TObject);
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

procedure TfrmCuadre.Escribe;
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

procedure TfrmCuadre.ExportarCuadre1Click(Sender: TObject);

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

    for iHoja := 1 to Book.Sheets.Count - 1 do
    begin
      Book.Sheets[ iHoja ].Select;
      Excel.ActiveSheet.Unprotect;
    end;

    Book.Save;

  end;

end;

procedure TfrmCuadre.ExportarCuadrevirtualaExcel1Click(Sender: TObject);
begin
  Cuadre.SaveToExcel;
end;

function TfrmCuadre.ColumnaNombre(Numero: Integer): String;
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

function TfrmCuadre.sfnRestaHoras(sParamHorasMax, sParamHorasMin: string): string;
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

function TfrmCuadre.sfnSumaHoras(sParamHorasMax, sParamHorasMin: string): string;
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

procedure TfrmCuadre.tmInicioPropertiesEditValueChanged(Sender: TObject);
begin
  lblResultado.Caption := 'El resultado de la actividad será: ';
end;

procedure TfrmCuadre.tmRestaHorasPropertiesChange(Sender: TObject);
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

function TfrmCuadre.rfnDecimal(sParamCantidad: string): Real;
var
  Code: Integer;
  Resultado: Real;
begin
  Val(sParamCantidad, Resultado, Code);
  if Code <> 0 then
    Resultado := 0;
  rfnDecimal := Resultado;
end;

function TfrmCuadre.Redondear(numero : real ; cifrasSig : integer) : real;
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

function TfrmCuadre.Truncar(numero: Real; cifras: Integer) : Real;
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

procedure TfrmCuadre.zQueryCopy(var ZDataset: TZQuery;
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
procedure TfrmCuadre.GenerarCorte1Click(Sender: TObject);
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

procedure TfrmCuadre.GenerarExcel(Personal: Boolean; Equipo: Boolean);
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
  Actividades : Integer;

  sFactorHoras,
  sFolio,
  sSumaFolios,
  ActividadAnterior : string;

  iHoras,
  iMinutos,
  dDuracion : Double;

  Cursor : TCursor;

  zqExiste : TZReadOnlyQuery;

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
                               'bp.sIdPlataforma '+ #10 +

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
                        'and bp.dIdFecha = :fecha group by ba.iIdActividad, ba.sNumeroActividad, p.sIdPersonal '

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
                                 'be.sIdPernocta '+ #10 +

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
                          'and be.dIdFecha = :fecha');

  {$ENDREGION}

begin
  try
    zqExiste := TZReadOnlyQuery.Create(nil);
    zqExiste.Connection := connection.zConnection;

    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;

    qrFolios.First;
    CdResult.First;

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
      prgFolios.Maximum := 0;
      prgFolios.Position := 0;
      prgFolios.Refresh;
      prgActividades.Maximum := 0;
      prgActividades.Position := 0;
      prgActividades.Refresh;
      Application.ProcessMessages;

      zqExiste.Active := False;
      zqExiste.SQL.Text := sSQLEXISTE[iPCount];
      zqExiste.ParamByName('contrato').AsString := global_Contrato_Barco;
      zqExiste.ParamByName('orden').AsString := global_contrato;
      zqExiste.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
      zqExiste.Open;

      iFila := 4;
      iColumna := 5;
      Libro.Sheets[iHoja].select;

      Excel.Columns['B:B'].columnwidth := 45;
      Excel.Columns['B:B'].wraptext := false;

      qrMOE_Sol.Active := False;
      qrMOE_Sol.SQL.Text := sSQLMOE[iPCount];
      qrMOE_Sol.ParamByName('orden').AsString := global_contrato;
      qrMOE_Sol.ParamByName('contrato').AsString := global_Contrato_Barco;
      qrMOE_Sol.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
      qrMOE_Sol.Open;

      SetLength( Cuadre.CATEGORIA[iPCount], qrFolios.RecordCount );
      iFCount := 0;
      qrFolios.First;

      Cuadre.CUADRAR[ IPCount ] := True;

      if qrMOE_Sol.RecordCount = 0 then
      begin

        Cuadre.CUADRAR[ IPCount ] := False;

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

      while not qrFolios.Eof do
      begin
        iJornadas := qrFolios.FieldByName('ijornadas').AsInteger;
        Cuadre.CATEGORIA[iPCount, iFCount] := TFolio.Create();
        Cuadre.CATEGORIA[iPCount, iFCount].sFolio := qrFolios.FieldByName('snumeroorden').asstring;
        SetLength( Cuadre.CATEGORIA[iPCount, iFCount].MOE, qrMOE_Sol.RecordCount );
        iMCount := 0;
        IAPCount := 0;
        iColumna := 6;
        Excel.Rows[ifila].rowheight := 65;

        qrMOE_Sol.First;
        lblEstado.Caption := 'Cargando ' + TIPO[iPCount];
        prgFolios.Maximum := qrFolios.RecordCount;
        prgFolios.Position := 0;
        Application.ProcessMessages;

        Rango := Excel.range['A'+inttostr(ifila - 2)+':E'+inttostr(ifila - 1)];
        Rango.interior.colorindex := 44;
        Rango.horizontalalignment := xlright;

        Rango := Excel.range['A'+inttostr(ifila - 2)+':E'+inttostr(ifila - 2)];
        Rango.mergecells := True;
        Rango.value := 'SOLICITADO';

        Rango := Excel.range['A'+inttostr(ifila - 1)+':E'+inttostr(ifila - 1)];
        Rango.MergeCells := true;
        Rango.value := 'A BORDO';

        while not qrMOE_Sol.Eof do
        begin
          if iHoja = 2 then
            Excel.Columns[ColumnaNombre(iColumna)].ColumnWidth := 20;

          Rango := Excel.range[ColumnaNombre(iColumna)+inttostr(ifila)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila)];
          Rango.MergeCells := True;
          Rango.Value := qrMOE_Sol.FieldByName('sDescripcion').AsString;

          Rango := Excel.Range[ColumnaNombre(iColumna)+inttostr(ifila+1)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila+1)];
          Rango.MergeCells := True;
          Rango.Value := qrMOE_Sol.FieldByName('sIdRecurso').asstring;

          Rango := Excel.Range[ColumnaNombre(iColumna)+inttostr(ifila-1)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila-1)];
          Rango.MergeCells := True;
          Rango.Value := qrMOE_Sol.FieldByName('iabordo').asinteger;

          Rango := Excel.Range[ColumnaNombre(iColumna)+inttostr(ifila-2)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila-2)];
          Rango.MergeCells := True;
          Rango.Value := qrMOE_Sol.FieldByName('isolicitado').asinteger;

          Excel.Range[ColumnaNombre(iColumna)+inttostr(ifila-1)+':'+ColumnaNombre(iColumna+1)+inttostr(ifila-2)].NumberFormat := '0.00';

          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount] := TCategoria.Create();
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sIdRecurso    := qrMOE_Sol.FieldByName('sidrecurso').AsString;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iSolicitado   := qrMOE_Sol.FieldByName('isolicitado').AsInteger;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iAbordo       := qrMOE_Sol.FieldByName('iabordo').AsInteger;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iCol          := iColumna;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iInicio       := iColumna;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iFin          := iColumna + 1;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sSuma         := COLUMNAS[iColumna]+inttostr( iFila - 3 );
          if iPCount = 0 then
             Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sIdCategoria  := qrMOE_Sol.FieldByName('sIdTipoPersonal').AsString;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iItemOrden    := qrMOE_Sol.FieldByName('iItemOrden').AsInteger;

          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sPernocta   := qrFolios.FieldByName('sIdPernocta').AsString;
          Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sPlataforma := qrFolios.FieldByName('sIdPlataforma').AsString;

          Inc( iMCount );
          Inc(iColumna, 2);
          qrMOE_Sol.Next;
        end;

        Rango := Excel.Range[ColumnaNombre(5)+inttostr(ifila-2)+':'+ColumnaNombre(icolumna-1)+inttostr(ifila+1)];
        Rango.Interior.ColorIndex := 44;
        Rango.Horizontalalignment := xlcenter;
        Rango.verticalalignment := xlCenter;
        Rango.Wraptext := True;

        Rango := Excel.Range['A'+inttostr(ifila)+':E'+inttostr(ifila)];
        Rango.MergeCells := True;
        Rango.Interior.ColorIndex := 44;
        Rango.Horizontalalignment := xlcenter;
        Rango.verticalalignment := xlCenter;
        Rango.Wraptext := True;
        Rango.Value := qrFolios.FieldByName('snumeroorden').AsString;
        Rango.font.Size := 15;
        Rango.font.bold := true;

        Cuadre.CATEGORIA[ iPCount, iFCount ].Fila := iFila;

        CdResult.Filtered := False;
        CdResult.Filter := 'NumeroOrden = ' + QuotedStr( Trim( qrFolios.FieldByName('snumeroorden').AsString ) ) ;
        CdResult.Filtered := True;

        Inc(ifila);
        iColumna := 1;

        Rango := Excel.range['A'+inttostr(ifila)+':E'+inttostr(ifila)];
        Rango.interior.colorindex := 45;
        Rango.HorizontalAlignment := xlcenter;

        Rango := Excel.range['A'+inttostr(ifila)];
        Rango.Value := 'ACTIVIDAD';
        Rango := Excel.range['B'+inttostr(ifila)];
        Rango.Value := 'DESCRIPCION';
        Rango := Excel.range['C'+inttostr(ifila)];
        Rango.Value := 'INICIO';
        Rango := Excel.range['D'+inttostr(ifila)];
        Rango.Value := 'TERMINO';
        Rango := Excel.range['E'+inttostr(ifila)];
        Rango.Value := 'DURACION';

        Inc(iFila);
        Cuadre.CATEGORIA[iPCount, iFCount].iInicio := iFila;
        SetLength( Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES, CdResult.RecordCount );
        iACount := 0;
        CdResult.First;
        ActividadAnterior := '*_*';
        prgActividades.Maximum := CdResult.RecordCount;
        prgActividades.Position := 0;
        iActividad := CdResult.FieldByName('idactividad').AsInteger;
        iInicio := iFila;

        while not CdResult.Eof do
        begin

          if ActividadAnterior <> CdResult.FieldByName('Actividad').AsString then
          begin
            Actividades := iFila + CdResult.FieldByName('Actividades').AsInteger;
            ActividadAnterior := CdResult.FieldByName('Actividad').AsString;
            SetLength( Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES, Length( Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES ) + 1 );
            Cuadre.CATEGORIA[ iPCount, IFCount ].ACTIVIDADES[ iACount ] := TActividad.Create;
            Cuadre.CATEGORIA[ iPCount, IFCount ].ACTIVIDADES[ iACount ].sIdActividad := ActividadAnterior;
            Cuadre.CATEGORIA[ iPCount, IFCount ].ACTIVIDADES[ iACount ].IsPadre := True;
            Cuadre.CATEGORIA[ iPCount, IFCount ].ACTIVIDADES[ iACount ].iRow := iFila;
            Cuadre.CATEGORIA[ iPCount, IFCount ].ACTIVIDADES[ iACount ].IndexPadre := IAPCount;

            SetLength( Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES_PADRES, Length( Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES_PADRES ) + 1 );
            Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES_PADRES[ IAPCount ] := TActividadPadre.Create;
            Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES_PADRES[ IAPCount ].IdActividad := ActividadAnterior;
            Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES_PADRES[ IAPCount ].Fila := iFila;
            Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES_PADRES[ IAPCount ].ActividadCount := CdResult.FieldByName('Actividades').AsInteger;
            Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES_PADRES[ IAPCount ].IndexActividad := iACount;

            Cuadre.CATEGORIA[ iPCount, IFCount ].ACTIVIDADES[ iACount ].Padre := @Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES_PADRES[ IAPCount ];
            Inc( IAPCount );

            Rango := Excel.Range['C'+inttostr( ifila )+':E'+inttostr( iFila )];
            Rango.MergeCells := True;
            Rango.NumberFormat := '@';
            Rango.Value := ActividadAnterior;
            Rango.HorizontalAlignment := xlCenter;
            Rango.VerticalAlignment := xlCenter;
            Rango.Font.Bold := True;

            Rango := Excel.Rows[iFila];
            if LowerCase( cbbVista.Text ) = 'horarios' then
              Rango.EntireRow.Hidden := True;

            iCol := 7;
            for iColumna := 4 to qrMOE_Sol.RecordCount + 3 do
            begin
              Rango := Excel.Range[ColumnaNombre(iCol)+inttostr(ifila)];
              Rango.Formula := '=SUM('+COLUMNAS[ iCol ]+inttostr( iFila + 1 ) + ':' + COLUMNAS[ iCol ]+ IntToStr( Actividades ) + ')';

              inc(icol, 2);
              qrMOE_Sol.Next;
            end;

            Excel.Range[ 'A'+inttostr( ifila ) + ':'+COLUMNAS[ iCol - 2 ] + IntToStr( ifila ) ].interior.colorindex := 42;

            Inc( iFila );
            Inc( iACount );

            Continue;
          end;

          iActividad := CdResult.FieldByName('idactividad').AsInteger;
          Rango := Excel.Range['A'+inttostr(ifila)+':B'+inttostr(ifila)];
          Rango.numberformat := '@';

          {$REGION  'Redondeo del sistema'}

          {sFactorHoras := sfnRestaHoras(CdResult.FieldByName('HoraTermino').AsString, CdResult.FieldByName('HoraInicio').AsString);
          getHM(sFactorHoras, iHoras, iMinutos);
          iHoras := iHoras / 24;
          iMinutos := ( iMinutos / 24 ) / 60;

          if rbRedondear.Checked then
            dDuracion := ( Redondear((iHoras + iminutos), Round( clcRT.Value ) ) );
          if rbTruncar.Checked then
            dDuracion := ( Truncar((iHoras + iminutos), Round( clcRT.Value ) ) );}

          {$ENDREGION}

          Excel.range['A'+inttostr(ifila)].value := CdResult.FieldByName('Actividad').AsString;
          Excel.range['B'+inttostr(ifila)].value := CdResult.FieldByName('Descripcion').AsString;
          Excel.range['C'+inttostr(ifila)].value := CdResult.FieldByName('HoraInicio').AsString;
          Excel.range['D'+inttostr(ifila)].value := CdResult.FieldByName('HoraTermino').AsString;
          Excel.range['E'+inttostr(ifila)].formula := '=D'+inttostr(ifila)+' - ' +'C'+inttostr(ifila);
          Excel.Range['E'+inttostr(ifila)].numberformat := '0.00';
          dDuracion := Excel.Range['E'+inttostr(ifila)].text;

          Rango := Excel.Rows[iFila];
          if LowerCase( cbbVista.Text ) = 'actividades' then
            Rango.EntireRow.Hidden := True;

          qrMOE_Sol.First;
          iCol := 7;
          for iColumna := 4 to qrMOE_Sol.RecordCount + 3 do
          begin
            Rango := Excel.Range[ColumnaNombre(iCol)+inttostr(ifila)];

            if iPCount = 0 then
            begin
              if qrMOE_Sol.FieldByName('sTE').AsString = 'Si' then
                Rango.Formula := '=('+COLUMNAS[icol-1]+inttostr(ifila)+'*E'+inttostr(ifila)+')*'+ IntToStr( 24 )
              else
              if qrMOE_Sol.FieldByName('sTierra').AsString = 'Si' then
              begin
                Rango.Formula := '=('+COLUMNAS[icol-1]+inttostr(ifila)+'*E'+inttostr(ifila)+'.)*'+ IntToStr( 3 ) ;
             //   showmessage('=('+COLUMNAS[icol-1]+inttostr(ifila)+'*E'+inttostr(ifila)+'.)*'+ IntToStr( 3 ));
              end
              else
                Rango.Formula := '=('+COLUMNAS[icol-1]+inttostr(ifila)+'*E'+inttostr(ifila)+')*'+ IntToStr( iJornadas )
            end
            else
              Rango.Formula := '=('+COLUMNAS[icol-1]+inttostr(ifila)+'*E'+inttostr(ifila)+')';

            inc(icol, 2);
            qrMOE_Sol.Next;
          end;

          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount] := TActividad.Create();
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iIdActividad   := CdResult.FieldByName('idactividad').AsInteger;
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].sIdActividad   := CdResult.FieldByName('actividad').AsString;
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].sWbs           := CdResult.FieldByName('wbs').AsString;
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].sHInicio       := CdResult.FieldByName('horainicio').AsString;
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].sHFin          := CdResult.FieldByName('horatermino').AsString;
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iIdDiario      := CdResult.FieldByName('iddiario').AsInteger;
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iTarea         := CdResult.FieldByName('Tarea').AsInteger;
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].dDuracion      := dDuracion;
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iRow           := iFila;
          Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iHermanosCount := CdResult.FieldByName('HermanosCount').AsInteger;
          if CdResult.FieldByName('Hermano').AsInteger >= 0 then
            Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iHermano   := CdResult.FieldByName('Hermano').AsInteger;

          Inc( iACount );
          Inc( iFila );
          CdResult.Next;

          prgActividades.Position := prgActividades.Position + 2;
          prgActividades.Refresh;
        end;

        CdResult.Filtered := False;
        Cuadre.CATEGORIA[iPCount, iFCount].iFin := iFila - 1;
        Cuadre.CATEGORIA[ iPCount, IFCount ].UpdateRange;
        Cuadre.CATEGORIA[ IPCount, IFCount ].ActualizarCountActividades;
        Excel.Range['B'+inttostr( Cuadre.CATEGORIA[iPCount, iFCount].iInicio ) + ':B' + inttostr( Cuadre.CATEGORIA[iPCount, iFCount].iFin )].wraptext := false;

        Excel.range['A'+inttostr(ifila-1)+':'+columnanombre( Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount-1].iCol + 1 )+inttostr(ifila)].horizontalalignment := xlCenter;
        Excel.range['A'+inttostr(ifila-1)+':'+columnanombre( Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount-1].iCol + 1 )+inttostr(ifila)].verticalalignment := xlCenter;

        {$region 'Busca hermanos para hacer Merge a las celdas'}

        iindexh := 0;
        while iindexh < Length( Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES ) - 1 do
        begin
          iFilaH := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iIndexH].iRow;
          if ( Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iIndexH].iHermanosCount > 1 )then
          begin
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
              parambyname('folio').asstring := qrFolios.FieldByName('snumeroorden').AsString;
              parambyname('hermano').AsInteger := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iIndexH].iHermano;
              open;

              if recordcount > 0 then
              begin
                (* Buscar el ultimo hermano
                   Obtener el final(TActividad.iRow) de rango al seleccionar *)

                {$region 'Recorrer hermanos'}


                iIndexRH := iIndexH;

                while (iIndexRH < Length( Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES ) - 1 ) and ( Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iIndexRH].iHermano = Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iIndexH].iHermano ) do
                begin
                  Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iIndexRH].iInicioConjunto := iFilaH;
                  Inc(iIndexRH);
                end;
                iIndexRH := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iIndexH].iRow;
                {$endregion}

                Rango := Excel.Range['A'+inttostr(iFilaH)+':A'+inttostr(ifilaH + fieldbyname('iRows').AsInteger - 1)];
                Rango.MergeCells := True;
                Rango.NumberFormat := '@';
                Rango.Value :=  FieldByName('snumeroactividad').AsString + #10 + FieldByName('sInicioConjunto').asstring + ' - ' + FieldByName('sFinConjunto').asstring;
                Rango.VerticalAlignment := xlCenter;
                Rango.HorizontalAlignment := xlCenter;
              end;

              Inc( iindexh, fieldbyname('iRows').AsInteger - 1 );
            end;

          end;

          inc(iindexh);
        end;
        
        {$endregion}

        Inc(iFila, 6);
        qrFolios.Next;
        Inc(iFCount);
      end;

      {SUMAS}
      for iFCount := 0 to Length( Cuadre.CATEGORIA[iPCount] ) - 1 do
      begin
        if zqExiste.RecordCount > 0 then
        begin
          //Recorer Actividades y Categorias
          for iACount := 0 to Length( Cuadre.CATEGORIA[iPCount, iFcount].ACTIVIDADES ) - 1 do
          begin
            try
              if not Cuadre.CATEGORIA[ iPCount, iFCount ].ACTIVIDADES[ iACount ].IsPadre then
              begin
                zqExiste.Filtered := False;
                zqExiste.Filter := 'sNumeroOrden = ' + QuotedStr( Cuadre.CATEGORIA[iPCount, iFcount].sFolio )
                                + ' AND iIdActividad = ' + IntToStr( Cuadre.CATEGORIA[iPCount, iFcount].ACTIVIDADES[iAcount].iIdActividad ) ;
                zqExiste.Filtered := True;
                iCol := 6;
                iFila := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iRow;
                if zqExiste.RecordCount > 0 then
                begin
                  qrMOE_Sol.First;
                  iMCount := 0;
                  while not qrMOE_Sol.Eof do
                  begin
                    if zqExiste.Locate('sIdRecurso', qrMOE_Sol.FieldByName('sIdRecurso').AsString, [] ) then
                    begin
                      if iPCount = 0 then
                        Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sPlataforma := zqExiste.FieldByName('sidplataforma').AsString;
                      Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sPernocta := zqExiste.FieldByName('sidpernocta').AsString;
                      Excel.Range[ ColumnaNombre( iCol ) + IntToStr( iFila )].Value := zqExiste.FieldByName('dCantidad').AsString;
                    end;
                    Inc(iCol, 2);
                    Inc( iMCount );
                    qrMOE_Sol.Next;
                  end;

                end;

              end;

            finally
              zqExiste.Filtered := False;
            end;

          end;

        end;

        qrMOE_Sol.First;
        iFila := Cuadre.CATEGORIA[iPCount, iFCount].iInicio - 5;
        iCol := 6;
        while not qrMOE_Sol.Eof do
        begin
          Rango := Excel.range[ColumnaNombre(iCol)+inttostr(ifila)+':'+columnanombre(icol+1)+inttostr(ifila)];
          Rango.MergeCells := True;

          sSumaFolios := '=(';
          for iSumaCount := 0 to Length( Cuadre.CATEGORIA[iPCount] ) - 1 do
          begin
            for IACount := 0 to Length( Cuadre.CATEGORIA[ IPCount, ISumaCount ].ACTIVIDADES ) - 1 do
            begin
              if Cuadre.CATEGORIA[ iPCount, iSumaCount ].ACTIVIDADES[ IACount ].IsPadre then
                sSumaFolios := sSumaFolios + ColumnaNombre(iCol+1) + IntToStr( Cuadre.CATEGORIA[iPCount, iSumaCount].ACTIVIDADES[ iACount ].iRow ) + '+';
            end;
          end;
          sSumaFolios[Length( sSumaFolios )] := ')';

          Rango.Formula := sSumaFolios;
          Rango.NumberFormat := '0.0000';

          try
            for IAPCount := 0 to Length( Cuadre.CATEGORIA[ IPCount, IFCount ].ACTIVIDADES_PADRES ) - 1 do
            begin
              Rango := Excel.Range[ Cuadre.CATEGORIA[ iPCount, iFCount ].ACTIVIDADES_PADRES[ IAPCount ].GetChildCountRange( COLUMNAS[ iCol + 1 ] ) ];
              Rango.interior.colorindex := 15;
              Rango.NumberFormat := '0.00';
              Rango := Excel.range[ Cuadre.CATEGORIA[ iPCount, iFCount ].ACTIVIDADES_PADRES[ IAPCount ].GetChildCountRange( COLUMNAS[ iCol] ) ];
              Rango.NumberFormat := '0';
              Rango.Locked := False;
            end;
          except
          end;

          qrMOE_Sol.Next;
          Inc(icol, 2);
        end;
        Rango := Excel.range['E'+inttostr(ifila)+':'+columnanombre(icol-2)+inttostr(ifila)];
        Rango.horizontalalignment := xlCenter;
      end;

      Excel.Columns['B'].EntireColumn.Hidden := True;
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

procedure TfrmCuadre.ValidaFolios;
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
                   'ot.sIdPlataforma '+
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

procedure TfrmCuadre.ValidaCategorias(var excel, hoja : Variant ; var zqMoe : TZReadOnlyQuery; iTipo : Integer);
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
procedure TfrmCuadre.GuardaEnBD;
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
  sIdCategoria : string;

  dCantidad,
  dCantidadHH : Real48;

  zqActividades,
  zqMoe : TZReadOnlyQuery;
  
  zqSave : TZQuery;

const
  TABLA : array[0..1] of string = ('Personal', 'Equipos');

begin
  try
    try
      //connection.zconnection.startTransaction;

      zqMoe := TZReadOnlyQuery.Create(nil);
      zqMoe.Connection := connection.zConnection;

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
      zqActividades.SQL.Text := 'select ba.iIdActividad, '+
                                       'ba.iIdDiario, '+
                                       'ba.sWbs, '+
                                       'ba.sNumeroActividad, '+
                                       'ba.sHoraInicio, '+
                                       'ba.sHoraFinal, '+
                                       'ba.sNumeroOrden '+#10+

                                'from bitacoradeactividades ba '+

                                'inner join reportediario rd '+
                                  'on ( rd.sOrden = ba.sContrato '+
                                    'and rd.dIdFecha = ba.dIdFecha ) '+#10+
    
                                'inner join ordenesdetrabajo ot '+
                                  'on ( ot.sContrato = ba.sContrato '+
                                    'and ot.sNumeroOrden = ba.sNumeroOrden ) '+#10+
    
                                'where ba.sContrato = :orden '+
                                'and ba.dIdFecha = :fecha '+
                                'and ba.sIdTipoMovimiento = "ED" ';
      zqActividades.ParamByName('orden').AsString := global_contrato;
      zqActividades.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
      zqActividades.Open;
      zqActividades.First;

      if zqActividades.RecordCount = 0 then
        raise Exception.Create('No se encontraron registradas en el dia especificado'); 
      
      {$ENDREGION}

      iPCount := 0;
      for iHoja := 1 to 2 do
      begin
        if Cuadre.CUADRAR[ IPCount ] then
        begin

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

          zqSave.Active := False;
          zqSave.SQL.Text := 'delete from bitacorade'+lowercase( TABLA[ iPCount ] )+' where sContrato = :orden and didfecha =:fecha';
          zqSave.ParamByName('orden').AsString := global_contrato;
          zqSave.ParamByName('fecha').AsDate := StrToDate( Cuadre.Fecha );
          zqSave.ExecSQL;

          {$ENDREGION}

          for iFCount := 0 to Length( Cuadre.CATEGORIA[iPCount] ) - 1 do
          begin
            if Cuadre.CATEGORIA[iPCount, iFCount].eExiste then
            begin
              sFolio :=  Cuadre.CATEGORIA[iPCount, iFcount].sFolio;
              iACount := 0;
              while iACount <= Length( Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES ) - 1 do
              begin

                if not Cuadre.CATEGORIA[ iPCount, IFCount ].ACTIVIDADES[ iACount ].IsPadre then
                begin

                  iActividad        := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iIdActividad;
                  sWbs              := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].sWbs;
                  iIdDiario         := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iIdDiario;
                  sNumeroActividad  := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].sIdActividad;

                  {$REGION 'Valida Actividad'}

                  zqActividades.Filtered := False;
                  zqActividades.Filter := 'iidactividad = '+inttostr(iactividad)+
                                          ' AND swbs ='+QuotedStr(sWbs)+
                                          ' AND snumeroactividad='+QuotedStr(sNumeroActividad)+
                                          ' AND iiddiario = '+inttostr(iIdDiario);
                  zqActividades.Filtered := True;

                  if zqActividades.RecordCount = 0 then
                  begin
                    Inc( iACount );
                    Continue;
                  end;

                  {$ENDREGION}
                
                  while ( iACount <= Length( Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES ) - 1 )
                    and ( ( iActividad = zqActividades.FieldByName('iidactividad').AsInteger )
                        and ( iIdDiario = Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iIdDiario ) )  do
                  begin
                    iActividad := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iIdActividad;
                    sInicio := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].sHInicio;
                    sFin := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].sHFin;
                    iFila := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iRow;
                    iTarea := Cuadre.CATEGORIA[iPCount, iFCount].ACTIVIDADES[iACount].iTarea;

                    for iMCount := 0 to Length( Cuadre.CATEGORIA[iPCount, iFCount].MOE ) - 1 do
                    begin
                      if ( Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].eExiste ) and ( Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].eListo ) then
                      begin
                        sIdRecurso := Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sIdRecurso;
                        iCol := Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iCol;

                        sIdPernocta   := Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sPernocta;
                        sIdPlataforma := Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sPlataforma;
                        sIdCategoria  := Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sIdCategoria;
                        iItemOrden    := Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].iItemOrden;

                        zqMoe.Locate('sidrecurso', Cuadre.CATEGORIA[iPCount, iFCount].MOE[iMCount].sIdRecurso, []);

                        if iMIndex = -1 then
                          Continue;

                        try
                          dCantidad   := Excel.Range[columnanombre(icol)+inttostr(ifila)].Value;
                        except
                          dCantidad := 0;
                        end;

                        if dCantidad = 0 then
                          Continue;

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
                        zqSave.ParamByName('idactividad').AsInteger := iActividad;
                        zqSave.ParamByName('wbs').asstring := sWbs;
                        zqSave.ParamByName('actividad').asstring := sNumeroActividad;
                        zqSave.ParamByName('tarea').AsInteger := iTarea;
                        zqSave.ParamByName('folio').AsString := sFolio;
                        zqSave.ParamByName('ItemOrden').AsInteger := iItemOrden;
                        if iPCount = 0 then
                        begin
                          zqSave.ParamByName('Categoria').AsString := sIdCategoria;
                          zqSave.ParamByName('plataforma').AsString := sIdPlataforma;
                        end;

                        zqSave.ExecSQL;

                        Excel.Range[ColumnaNombre(icol)+inttostr(ifila)+':'+ColumnaNombre(icol+1)+inttostr(ifila)].interior.colorindex := 43;
                        Inc( iOk );

                      end;

                    end;
                    Inc( iACount );

                  end;

                end
                else
                  Inc( IACount );

              end;


            end;

                   
          end;

          Cuadre.INSERTADOS[iPCount] := iOk;
          Inc(iPCount);
          iOk := 0;

        end;

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
    MessageDlg('Listo'+#10+'Personal: '+ IntToStr( Cuadre.INSERTADOS[0] )+ ' registros'+ #10+ 'Equipo: '+ IntToStr( Cuadre.INSERTADOS[1] ) + ' registros', mtInformation, [mbOK], 0);

    if ( Cuadre.INSERTADOS[0] > 0 ) or ( Cuadre.INSERTADOS[1] > 0 ) then
      //connection.zConnection.Commit;

    Book.SaveAs(sArchivo);
    Libro.&Protected := False;
    Libro.LoadFromFile(sArchivo);
    Libro.&Protected := True;
    Excel.Quit;

    if FileExists( sArchivo ) then
    begin
      DeleteFile( sArchivo );
    end;

    zqSave.Free;
    zqActividades.Free;
    zqMoe.Free;

    Cuadre.Guardado := True;
    Cuadre.Cambios := False;

    Kardex('Reporte Diario', 'Guarda cuadre del dia '+ Cuadre.Fecha , '', '', '', '', '' , 'Tarifa Diaria', 'Cuadre');
  end;
end;

procedure TfrmCuadre.VentanaCortes();
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
  pnlCortes.Parent := frmCuadre;
  pnlCortes.Align := alNone;
  pnlCortes.Visible :=  False;
  pnlCortes.Width := 0;
  pnlCortes.Height := 0;
  pnlCortes.Left := 0;
  pnlCortes.Top := 0;
  Form.Free;
  Application.ProcessMessages;
end;

procedure TfrmCuadre.CargarActividades_ListView( Lista : TdxTreeView ; Folio : Integer; Actividad : string);
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

procedure TfrmCuadre.CortarActividad(Actividad: string; Inicio: string; Fin: string; InicioCorte: string; FinCorte: string; Fila: Integer);
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

procedure tfrmCuadre.RegenerarSumasMOE(var Excel: Variant; IndiceTipo : TCategoriaIndex);
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

procedure TfrmCuadre.EliminarHorario( Fila : TFilaExcel; hFinal : string ; IndiceF : TFolioIndex ; IndiceA : TActividadIndex; IndicePadre : TActividadPadreIndex );
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
procedure TfrmCuadre.CambiaVistaCuadre(Tipo: string);
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

end.
