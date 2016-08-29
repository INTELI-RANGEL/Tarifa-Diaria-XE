unit frm_ReportePeriodo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, global, ComCtrls, 
  StdCtrls, ExtCtrls, DBCtrls, db, Menus, OleCtrls,
  frxClass, frxDBSet, Buttons, RxLookup, RxMemDS, utilerias, Newpanel,
  RXCtrls, DateUtils, math, strUtils, ImgList, UnitTBotonesPermisos,
  ZAbstractRODataset, ZDataset, ZAbstractDataset,
  jpeg, ComObj, UnitExcepciones, UFunctionsGHH, DBDateTimePicker, 
  
  RXDBCtrl, ExtDlgs,
  UnitPatrick, UnitExcel,
  cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxContainer,
  cxEdit, ShlObj,
  cxCheckBox, CheckLst, AdvGlowButton, dxSkinsCore, dxSkinBlack, dxSkinBlue,
  dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide,
  dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle, dxSkinFoggy,
  dxSkinGlassOceans, dxSkinHighContrast, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013White,
  dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus,
  dxSkinSilver, dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008,
  dxSkinTheAsphaltWorld, dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010,
  dxSkinWhiteprint, dxSkinXmas2008Blue, dxSkinMetropolis, dxSkinMetropolisDark,
  dxSkinOffice2013DarkGray, dxSkinOffice2013LightGray;
function IsDate(ADate: string): Boolean;
type
  Tcadena = class
    sValor :string;
  end;
type
  TfrmReportePeriodo = class(TForm)
    dbReporte: TfrxDBDataset;
    ds_ordenesdetrabajo: TDataSource;
    rxInactivos: TRxMemoryData;
    rxInactivossContrato: TStringField;
    rxInactivossNumeroReporte: TStringField;
    rxInactivossHoraInicio: TStringField;
    rxInactivossHoraFinal: TStringField;
    rxInactivossTiempoEfectivo: TStringField;
    rxInactivosiPersonal: TIntegerField;
    rxInactivossTiempoMuerto: TStringField;
    rxInactivossTiempoMuertoReal: TStringField;
    rxInactivostmsHoraInicio: TStringField;
    rxInactivostmsHoraFinal: TStringField;
    rxInactivostmsTiempoMuerto: TStringField;
    rxInactivostmiPersonal: TIntegerField;
    rxInactivostmiPersonalOrden: TIntegerField;
    rxInactivostmmDescripcion: TMemoField;
    rxInactivossIdUsuarioAutoriza: TStringField;
    rxInactivosdIdFecha: TDateField;
    rxInactivossDescripcion: TStringField;
    rFotografico: TfrxDBDataset;
    rxReporteFotografico: TRxMemoryData;
    rxReporteFotograficosContrato: TStringField;
    rxReporteFotograficosNumeroReporte: TStringField;
    rxReporteFotograficoiImagen: TIntegerField;
    rxReporteFotograficobImagen: TBlobField;
    rxReporteFotograficosDescripcion: TMemoField;
    rxReporteFotograficosDescripcionCorta: TStringField;
    rxReporteFotograficosNumeroOrden: TStringField;
    rxReporteFotograficodIdFecha: TDateField;
    dsPlataforma: TfrxDBDataset;
    rxReporteFotograficobImagen2: TBlobField;
    rxReporteFotograficosDescripcion2: TMemoField;
    dsPersonalPernocta: TfrxDBDataset;
    dsPersonalPlataforma: TfrxDBDataset;
    rxPlataforma: TRxMemoryData;
    rxPlataformasTitulo: TStringField;
    rxPlataformasCantidad: TStringField;
    rxPernocta: TRxMemoryData;
    StringField1: TStringField;
    StringField2: TStringField;
    rxPlataformaiItem: TIntegerField;
    rxPernoctaiItem: TIntegerField;
    dsPernocta: TfrxDBDataset;
    rxNotas: TRxMemoryData;
    StringField3: TStringField;
    StringField4: TStringField;
    MemoField1: TMemoField;
    rxNotassDescripcion: TStringField;
    dsNotas: TfrxDBDataset;
    rDiario: TfrxReport;
    dsComentarios: TfrxDBDataset;
    rNotas: TfrxReport;
    ordenesdetrabajo: TZReadOnlyQuery;
    QryComentarios: TZReadOnlyQuery;
    QryPlataforma: TZReadOnlyQuery;
    QryPernocta: TZReadOnlyQuery;
    pgInformes: TPageControl;
    tabFotografico: TTabSheet;
    gbParametros: tNewGroupBox;
    tNewGroupBox1: tNewGroupBox;
    tNewGroupBox3: tNewGroupBox;
    tNewGroupBox5: tNewGroupBox;
    GroupReportesDiarios: tNewGroupBox;
    QryPersonalPlataforma: TZReadOnlyQuery;
    QryPersonalPernocta: TZReadOnlyQuery;
    ActividadesxOrden: TZReadOnlyQuery;
    ActividadesxOrdensContrato: TStringField;
    ActividadesxOrdensNumeroOrden: TStringField;
    ActividadesxOrdeniNivel: TIntegerField;
    ActividadesxOrdensSimbolo: TStringField;
    ActividadesxOrdensWbs: TStringField;
    ActividadesxOrdensWbsAnterior: TStringField;
    ActividadesxOrdensNumeroActividad: TStringField;
    ActividadesxOrdensTipoActividad: TStringField;
    ActividadesxOrdenmDescripcion: TMemoField;
    ActividadesxOrdendFechaInicio: TDateField;
    ActividadesxOrdendFechaFinal: TDateField;
    ActividadesxOrdendPonderado: TFloatField;
    ActividadesxOrdendVentaMN: TFloatField;
    ActividadesxOrdendVentaDLL: TFloatField;
    ActividadesxOrdeniColor: TIntegerField;
    ActividadesxOrdensIdConvenio: TStringField;
    ActividadesxOrdensWbsSpace: TStringField;
    ActividadesxOrdendCantidadPeriodo: TFloatField;
    ActividadesxOrdendAcumulado: TFloatField;
    ActividadesxOrdendAcumuladoAnterior: TFloatField;
    dsActividadesxOrden: TfrxDBDataset;
    ActividadesxOrdensMedida: TStringField;
    ActividadesxOrdendCantidad: TFloatField;
    ActividadesxOrdendTotal: TFloatField;
    ActividadesxOrdendTotalAcumulado: TCurrencyField;
    ActividadesxOrdeniItemOrden: TStringField;
    rReporte: TfrxReport;
    GroupBox3: TGroupBox;
    rbDetalleInstalacion: TRadioButton;
    rbComentarios: TRadioButton;
    tsNumeroOrden3: TDBLookupComboBox;
    Label6: TLabel;
    btnImprimeDiarios: TBitBtn;
    GroupBox4: TGroupBox;
    chkConsolidado: TCheckBox;
    chkTiempoMuerto: TCheckBox;
    Label4: TLabel;
    tsNumeroOrden: TDBLookupComboBox;
    btnPrinter: TBitBtn;
    GroupBox5: TGroupBox;
    chkPernocta: TCheckBox;
    chkPlataforma: TCheckBox;
    chkDetalle: TCheckBox;
    btnPersonalProgramado: TBitBtn;
    GroupBox6: TGroupBox;
    GroupBox2: TGroupBox;
    tsOrdenesdeTrabajo: TRxCheckListBox;
    btnAlbum: TBitBtn;
    GroupBox7: TGroupBox;
    Label2: TLabel;
    Label1: TLabel;
    mdReporte: TRxMemoryData;
    dsReporte: TfrxDBDataset;
    mdReportesContrato: TStringField;
    mdReportesNumeroActividad: TStringField;
    mdReportemDescripcion: TMemoField;
    mdReportedCantidad: TFloatField;
    mdReportesMedida: TStringField;
    mdReporteDia1: TFloatField;
    mdReporteDia2: TFloatField;
    mdReporteDia3: TFloatField;
    mdReporteDia4: TFloatField;
    mdReporteDia5: TFloatField;
    mdReporteDia6: TFloatField;
    mdReporteDia7: TFloatField;
    mdReporteDia8: TFloatField;
    mdReporteDia9: TFloatField;
    mdReporteDia10: TFloatField;
    mdReporteDia11: TFloatField;
    mdReporteDia12: TFloatField;
    mdReporteDia13: TFloatField;
    mdReporteDia14: TFloatField;
    mdReporteDia15: TFloatField;
    mdReporteDia16: TFloatField;
    mdReporteDia17: TFloatField;
    mdReporteDia18: TFloatField;
    mdReporteDia19: TFloatField;
    mdReporteDia20: TFloatField;
    mdReporteDia21: TFloatField;
    mdReporteDia22: TFloatField;
    mdReporteDia23: TFloatField;
    mdReporteDia24: TFloatField;
    mdReporteDia25: TFloatField;
    mdReporteDia26: TFloatField;
    mdReporteDia27: TFloatField;
    mdReporteDia28: TFloatField;
    mdReporteDia29: TFloatField;
    mdReporteDia30: TFloatField;
    mdReporteDia31: TFloatField;
    mdReportesWbs: TStringField;
    chkMoneda: TCheckBox;
    Label7: TLabel;
    dbFiltro: TDBLookupComboBox;
    Afectaciones: TZReadOnlyQuery;
    ds_afectaciones: TDataSource;
    dsConfiguracion: TfrxDBDataset;
    dsQryGraficaTiemposMuertos: TfrxDBDataset;
    rbDetalleAvances: TRadioButton;
    mdReportedTotal: TFloatField;
    mdReporteMes: TStringField;
    mdReporteAnio: TStringField;
    chkDLL: TCheckBox;
    mdReportedIdFecha: TDateField;
    mdDatosAux: TRxMemoryData;
    StringField10: TStringField;
    StringField11: TStringField;
    MemoField3: TMemoField;
    FloatField1: TFloatField;
    StringField12: TStringField;
    FloatField2: TFloatField;
    FloatField3: TFloatField;
    FloatField5: TFloatField;
    FloatField6: TFloatField;
    FloatField7: TFloatField;
    FloatField8: TFloatField;
    FloatField9: TFloatField;
    FloatField10: TFloatField;
    FloatField11: TFloatField;
    FloatField12: TFloatField;
    FloatField13: TFloatField;
    FloatField14: TFloatField;
    FloatField15: TFloatField;
    FloatField16: TFloatField;
    FloatField17: TFloatField;
    FloatField18: TFloatField;
    FloatField19: TFloatField;
    FloatField20: TFloatField;
    FloatField21: TFloatField;
    FloatField22: TFloatField;
    FloatField23: TFloatField;
    FloatField24: TFloatField;
    FloatField25: TFloatField;
    FloatField26: TFloatField;
    FloatField27: TFloatField;
    FloatField28: TFloatField;
    FloatField29: TFloatField;
    FloatField30: TFloatField;
    FloatField31: TFloatField;
    FloatField32: TFloatField;
    FloatField33: TFloatField;
    StringField13: TStringField;
    FloatField34: TFloatField;
    StringField14: TStringField;
    StringField15: TStringField;
    DateField1: TDateField;
    mdReportesWbsAnterior: TStringField;
    mdDatosAuxsWbsAnterior: TStringField;
    mdReportedFechaInicio: TDateField;
    mdReportedFechaFinal: TDateField;
    QryGraficaTiemposMuertos: TZReadOnlyQuery;
    chkFases: TCheckBox;
    Timer1: TTimer;
    //rxReporteFotograficosContrato: TStringField;  conflicto de versiones
    ProgressBar1: TProgressBar;
    rxReporteFotograficosNumeroActividad: TStringField;
    rxReporteFotograficosFasePartida: TStringField;
    rxReporteFotograficosIdFolio: TStringField;
    tdFechaInicial: TDateTimePicker;
    tdFechaFinal: TDateTimePicker;
    TabSheet1: TTabSheet;
    dbDetalle: TfrxDBDataset;
    GroupBox1: TGroupBox;
    tNewGroupBox2: tNewGroupBox;
    GroupBox9: TGroupBox;
    tNewGroupBox6: tNewGroupBox;
    GroupBox10: TGroupBox;
    tNewGroupBox4: tNewGroupBox;
    Label8: TLabel;
    Label3: TLabel;
    tsNumeroOrdenActa: TDBLookupComboBox;
    QryPartidasEfectivas: TZReadOnlyQuery;
    ds_ParidasEfectivas: TDataSource;
    Detalle: TZReadOnlyQuery;
    Imagenes: TImageList;
    GroupBox11: TGroupBox;
    bImagen: TImage;
    Grid_Imagenes: TRxDBGrid;
    btnInsert: TBitBtn;
    ds_FotosPartidas: TDataSource;
    FotosPartidas: TZReadOnlyQuery;
    btnPreview: TBitBtn;
    labelFase: TLabel;
    Label5: TLabel;
    cmdNuevo: TBitBtn;
    cmdEliminar: TBitBtn;
    btnImportar: TBitBtn;
    GroupBox8: TGroupBox;
    bImagen2: TImage;
    cmdImprimeActa: TBitBtn;
    ComboActas: TComboBox;
    ds_fotografico_acta: TDataSource;
    btnPreview2: TBitBtn;
    btnUp: TBitBtn;
    btnDown: TBitBtn;
    btnDelete: TBitBtn;
    OpenPicture: TOpenPictureDialog;
    fotografico_acta: TZQuery;
    fotografico_actaiImagen: TIntegerField;
    fotografico_actalImprime: TStringField;
    fotografico_actasNumeroActividad: TStringField;
    fotografico_actasFasePartida: TStringField;
    fotografico_actasContrato: TStringField;
    fotografico_actasNumeroOrden: TStringField;
    fotografico_actasActaFotografica: TStringField;
    fotografico_actadIdFecha: TDateField;
    fotografico_actabImagen: TBlobField;
    fotografico_actasWbs: TStringField;
    grid_bitacorapersonal: TRxDBGrid;
    rxReporteFotograficomDescripcion: TMemoField;
    rxReporteFotograficosTituloOrden: TStringField;
    rxReporteFotograficoContratoPrincipal: TMemoField;
    rbCronologias: TRadioButton;
    rbCronologiasFiltradas: TRadioButton;
    rbNotaCampo: TRadioButton;
    ReportesCheck: TCheckBox;
    TabSheet2: TTabSheet;
    GroupBox12: TGroupBox;
    lbfolios01: TLabel;
    btnimprimir: TBitBtn;
    dbFolio01: TDBLookupComboBox;
    chkCaratula: TCheckBox;
    chkOficio: TCheckBox;
    chk1: TCheckBox;
    chk2: TCheckBox;
    chk4: TCheckBox;
    chk5: TCheckBox;
    chk6: TCheckBox;
    chk7: TCheckBox;
    chk8: TCheckBox;
    grpSeleccionLibro: tNewGroupBox;
    cxCheckBox1: TcxCheckBox;
    GroupBox13: TGroupBox;
    LsContenido: TCheckListBox;
    ZQRFolio: TZReadOnlyQuery;
    ZQRConfiguracion: TZReadOnlyQuery;
    ZQRContenido: TZReadOnlyQuery;
    CheckBox1: TCheckBox;
    BtnImpLb: TAdvGlowButton;
    ZQRperiodo: TZReadOnlyQuery;
    ZQREstimacion: TZReadOnlyQuery;
    ZqrContrato: TZReadOnlyQuery;
    rxReporteFotograficosIdPlataforma: TStringField;
    chk4img: TRadioButton;
    chk2img: TRadioButton;
    tsNumeroActividad: TDBLookupComboBox;

    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure btnPrinterClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure tdFechaInicialEnter(Sender: TObject);
    procedure tdFechaInicialExit(Sender: TObject);
    procedure tdFechaInicialKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaFinalEnter(Sender: TObject);
    procedure tdFechaFinalExit(Sender: TObject);
    procedure tdFechaFinalKeyPress(Sender: TObject; var Key: Char);
    procedure ReporteGetValue(const VarName: String; var Value: Variant);
    procedure frxFotograficoGetValue(const VarName: String;
      var Value: Variant);
    procedure btnPersonalProgramadoClick(Sender: TObject);
    procedure frxGerencialGetValue(const VarName: String;
      var Value: Variant);
    procedure ActualizaPlataforma(Sender: TObject);
    procedure ActualizaPernocta(Sender: TObject);
    procedure frxProgramacionGetValue(const VarName: String;
      var Value: Variant);
    procedure btnImprimeDiariosClick(Sender: TObject);
    procedure rDiarioGetValue(const VarName: String; var Value: Variant);
    procedure chkDetalleClick(Sender: TObject);
    procedure ActividadesxOrdenCalcFields(DataSet: TDataSet);
    procedure dbFiltroKeyPress(Sender: TObject; var Key: Char);
    procedure dbFiltroEnter(Sender: TObject);
    procedure dbFiltroExit(Sender: TObject);
    procedure chkMonedaClick(Sender: TObject);
    procedure chkDLLClick(Sender: TObject);
    procedure rbDetalleInstalacionEnter(Sender: TObject);
    procedure rbDetalleAvancesEnter(Sender: TObject);
    procedure rbComentariosEnter(Sender: TObject);
    procedure rbResumenInstalacionEnter(Sender: TObject);
    procedure rbAnalisisFinancieroEnter(Sender: TObject);
    procedure rbResumenGralEnter(Sender: TObject);
    procedure rbProgramadoEnter(Sender: TObject);
    procedure rbVolumenGeneralEnter(Sender: TObject);
    procedure OnClick3D(Sender: TObject);
    procedure tsNumeroOrden3Enter(Sender: TObject);
    procedure tsNumeroOrden3Exit(Sender: TObject);
    procedure tsNumeroOrden3KeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaInicialChange(Sender: TObject);
    procedure tsNumeroActividadEnter(Sender: TObject);
    procedure tsNumeroActividadChange(Sender: TObject);
    procedure btnPreviewClick(Sender: TObject);
    procedure FotosPartidasAfterScroll(DataSet: TDataSet);
    procedure pgInformesChange(Sender: TObject);
    procedure cmdNuevoClick(Sender: TObject);
    procedure ComboActasKeyPress(Sender: TObject; var Key: Char);
    procedure cmdEliminarClick(Sender: TObject);
    procedure btnInsertClick(Sender: TObject);
    procedure ComboActasExit(Sender: TObject);
    procedure btnPreview2Click(Sender: TObject);
    procedure fotografico_actaAfterScroll(DataSet: TDataSet);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnUpClick(Sender: TObject);
    procedure btnDownClick(Sender: TObject);
    procedure OrdenarFotos(sParamOrden : string);
    procedure ComboActasEnter(Sender: TObject);
    procedure tsNumeroOrdenActaExit(Sender: TObject);
    procedure btnImportarClick(Sender: TObject);
    procedure cmdImprimeActaClick(Sender: TObject);
    procedure fotografico_actaAfterInsert(DataSet: TDataSet);
    procedure tsNumeroOrdenActaKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroActividadKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenActaEnter(Sender: TObject);
    procedure tsNumeroActividadExit(Sender: TObject);
    procedure ImprimeExcel_Cronologias;
    procedure btnAlbumClick(Sender: TObject);
    procedure ReportesCheckClick(Sender: TObject);
    procedure tsNumeroOrden3Click(Sender: TObject);
    procedure btnimprimirClick(Sender: TObject);
    procedure ImprimeExcel_NotaDeCampo;

    procedure imprimirLibro;
    procedure cxCheckBox1Click(Sender: TObject);
    procedure LsContenidoMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure BtnImpLBClick(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);

    function RedimensionarJPG(sFilePath: string): string;
    function numeroDeImagenes(limite : Integer) : boolean;

  private
  sMenuP: String;
    { Private declarations }
    procedure PowerPointAlbum(roqDatos : TRxMemoryData; fileName : string);
    procedure CargarContenidoLB(Cmp: TCheckListBox);
    procedure ImprimeHoja(nombre: string; Lb,excel: Variant);
    procedure NotaCampoExcel(excel, Hoja, Libro: Variant);
    Procedure FormatoNormal(var Excel: Variant;Cadena:string; Align: Integer;Negrita,Ajustar:Boolean);overload;
    Procedure FormatoNormal(var Excel: Variant;Cadena:VAriant; Align: Integer;Negrita,Ajustar:Boolean;Formato:String;Column:Integer=0);overload;
    procedure AjustarTexto(var rangoE: Variant; TotalR: Integer);


  public
    { Public declarations }
  end;

Const
  MaxCol = 31;
  NomMes : Array[1..12] of String = ('ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
                                     'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIMEBRE', 'DICIEMBRE');

var
  frmReportePeriodo : TfrmReportePeriodo;
  sHoraResult : String ;
  iReportes, Dias1, Dias2   : Integer ;
  sOrdenes : WideString ;
  dProgramado, dReal, dPromedio : Real ;
  sContrato, sPernocta, sPlataforma : String ;
  sFechaInicio : String ;
  sConvenioInicio, sConvenioFinal, sActa, sNuevoInicio, sNuevoFinal : String ;
  StringPuesto: TStrings;
  StringNombre: TStrings;
  sPoliza, sFianza : String ;
  lExiste            : Boolean ;
  InicioReal, TerminoReal: TDate;
  BotonPermiso: TBotonesPermisos;
  previndex:Integer;

  TablasConsult: array[1..2,1..4] of string=(('bitacoradepersonal','sIdPersonal','personal','PERSONAL'),
                                             ('bitacoradeequipos','sIdEquipo','equipos','EQUIPO'){,
                                             ('bitacoradebarcoxfases','sIdEmbarcacion','BARCO')});

implementation



{$R *.dfm}




procedure TfrmReportePeriodo.ImprimeExcel_NotaDeCampo;
Var
  Libro, Excel, Hoja: Variant;
  sFolios: String;
  iFila, iColumna, iHojas, iCounter, iLoop: Integer;
  NombreDelExcel, SQLExtra: String;
  TempPath: String;
  Fs: TStream;
  Pic : TJpegImage;
  imgAux: TImage;
  dContrato_Inicio,
  dContrato_Final: TDateTime;
  TmpName: String;

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

begin
  Try
    NombreDelExcel := PGetTempDir + 'TEMP~' + PNombreAleatorio(3) + 'NotaDeCampo.xls';
    Try
      Excel := CreateOleObject('Excel.Application');
    Except
      On E: Exception do begin
        FreeAndNil(Excel);
        ShowMessage(E.Message);
        Exit;
      end;
    End;

    Excel.Visible := True;
    Excel.DisplayAlerts:= False;
    Libro := Excel.Workbooks.Add;

    Excel.WorkBooks[1].WorkSheets[1].Name := 'NOTA DE CAMPO';
    Hoja := Excel.WorkBooks[1].WorkSheets[1];
    Libro.Sheets[1].Select;

    iColumna := 1;
    iFila := 1;

    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 0.75;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
    Inc(iColumna);

{$REGION 'IMAGENES DE CABECERA'}
  //Imagen Izquierda
  Try
    TmpName := '';
    imgAux := TImage.Create(nil);
    if TmpName='' then begin
//      GetTempPath(SizeOf(TempPath), TempPath);
      TempPath := ExtractFilePath(Application.Exename);
      TmpName:=TempPath +'imgtempSln1'+formatdatetime('dddddd hhnnss',now)+'.jpg';
      fs := Connection.contrato.CreateBlobStream(Connection.contrato.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
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
  Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 20, 70, 35);
  //Imagen Derecha
  Try
    TmpName := '';
    imgAux := TImage.Create(nil);
    if TmpName='' then begin
//      GetTempPath(SizeOf(TempPath), TempPath);
      TempPath := ExtractFilePath(Application.Exename);
      TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
  Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 498, 20, 70, 40);
  //Texto de Cabecera
  Excel.Range['A1:L4'].Select;
  //PFormatosExcel_Bordes(Excel, False, True, False, False);

  Excel.Range['A2:L4'].Select;
  PFormatosExcel_H2(Excel, 0, True, 10);
  Excel.Selection.Value := Connection.configuracion.FieldByName('sNombre').AsString;



  Excel.Range['A4:L5'].Select;
  PFormatosExcel_H2(Excel, 0, True, 8);
 // PFormatosExcel_Bordes(Excel, False, True, False, False, -4119);
  Excel.Selection.Value := 'NOTA DE CAMPO';
  {$ENDREGION}

{$REGION 'CABECERA'}
    iFila := 6;
    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := 'CONTRATO:';

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 11, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := Connection.contrato.FieldByName('sContrato').AsString;

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := 'FOLIO:';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := Ordenesdetrabajo.FieldByName('sNumeroOrden').AsString;
    Inc (iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(7)+':'+ColumnaNombre(3)+IntToStr(8)].Select;
    PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := 'DESCRIPCIÓN:';

    Excel.Range[ColumnaNombre(4)+IntToStr(7)+':'+ColumnaNombre(6)+IntToStr(8)].Select;
    PFormatosExcel_H2(Excel, 11, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := Connection.contrato.FieldByName('mDescripcion').AsString;

    Excel.Range[ColumnaNombre(7)+IntToStr(7)+':'+ColumnaNombre(8)+IntToStr(7)].Select;
    PFormatosExcel_H2(Excel, 47, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := 'OBRA:';

    Excel.Range[ColumnaNombre(9)+IntToStr(7)+':'+ColumnaNombre(12)+IntToStr(7)].Select;
    PFormatosExcel_H2(Excel, 47, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := Connection.contrato.FieldByName('sTitulo').AsString;
    Inc (iFila);
     
    Excel.Range[ColumnaNombre(7)+IntToStr(8)+':'+ColumnaNombre(8)+IntToStr(8)].Select;
    PFormatosExcel_H2(Excel, 21, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := 'LOCALIZACIÓN:';

    Excel.Range[ColumnaNombre(9)+IntToStr(8)+':'+ColumnaNombre(12)+IntToStr(8)].Select;
    PFormatosExcel_H2(Excel, 21, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.Value := Connection.contrato.FieldByName('sUbicacion').AsString;

{$ENDREGION}

{$REGION 'IMPRESION ACTIVIDADES'}
    iFila := 12;

    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add(' ' +
                              'SELECT ' +
                              '	sNumeroOrden, ' +
                              '	mDescripcion, ' +
                              '	sNumeroActividad ' +
                              'FROM actividadesxorden ' +
                              'WHERE sNumeroOrden = :Folio');
    Connection.QryBusca.ParamByName('Folio').AsString := Ordenesdetrabajo.FieldByName('sNumeroOrden').AsString;
    Connection.QryBusca.Open;

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PARTIDA';

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'ACTIVIDAD';
    Inc(iFila);

    if Connection.QryBusca.RecordCount > 0 then begin
      while Not Connection.QryBusca.Eof do begin

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := Connection.QryBusca.FieldByName('sNumeroActividad').AsString;

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := Connection.QryBusca.FieldByName('mDescripcion').AsString;

    Inc (iFila);
    Connection.QryBusca.Next;
    end;
      end else begin
    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';
      end;

    Inc (iFila,3);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 11, False, 7, clBlack, 'Arial');
    Excel.Selection.Value := 'PERIODOS DE EJECUCION DE LA ACTIVIDAD';

    Inc(iFila,2);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'FECHA';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'INICIO';

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'TERMINO';

    Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'AFECTACION';

    Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'INTERVALO TIEMPO';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'AVANCE ANTERIOR';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'AVANCE ACTUAL';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'AVANCE ACUMULADO';

    Inc(iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Inc(iFila);

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlRight;
    Excel.Selection.Value := 'DURACION TIEMPO EFECTIVO (HRS):' ;

    Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, True, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.Value := '01:56' ;

    Inc (iFila);

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlRight;
    Excel.Selection.Value := 'DURACION TIEMPO AFECTACIONES (HRS):' ;

    Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, True, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.Value := '00:00' ;

    Inc (iFila);

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlRight;
    Excel.Selection.Value := 'TIEMPO TOTAL (HRS):' ;

    Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 15, True, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.Value := '01:56' ;

    Inc (iFila,2);

{$ENDREGION}

{$REGION 'IMPRESION ACTIVIDADES BARCO'}
    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PARTIDA';

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'ACTIVIDAD';
    Inc(iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';
    Inc(iFila,2);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'MOVIMIENTOS DE EMBARCACION';
    Inc(iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PARTIDA';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'DESCRIPCION';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'UNIDAD';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'CANTIDAD';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PU MN';

    Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PU USD';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'IMP MN';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'IMP USD';
    Inc (iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';
    Inc (iFila);

    Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlRight;
    Excel.Selection.Value := 'IMPORTE BARCO';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';
    Inc (iFila,2);
{$ENDREGION}

{$REGION 'IMPRESION PERSONAL'}
    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PERSONAL';
    Inc(iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PARTIDA';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'DESCRIPCION';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'UNIDAD';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'CANTIDAD';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PU MN';

    Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PU USD';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'IMP MN';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'IMP USD';
    Inc (iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';
    Inc (iFila);

    Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlRight;
    Excel.Selection.Value := 'IMPORTE PERSONAL';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';
    Inc (iFila,2);
{$ENDREGION}

{$REGION 'IMPRESION EQUIPOS'}
    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'EQUIPO';
    Inc(iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PARTIDA';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'DESCRIPCION';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'UNIDAD';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'CANTIDAD';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PU MN';

    Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PU USD';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'IMP MN';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'IMP USD';
    Inc (iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';
    Inc (iFila);

    Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlRight;
    Excel.Selection.Value := 'IMPORTE EQUIPO';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';
    Inc (iFila,2);
{$ENDREGION}

{$REGION 'IMPRESION PERCNOTAS'}
    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PERNOCTAS';
    Inc(iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PARTIDA';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'DESCRIPCION';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'UNIDAD';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'CANTIDAD';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PU MN';

    Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'PU USD';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'IMP MN';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'IMP USD';
    Inc (iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';
    Inc (iFila);

    Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlRight;
    Excel.Selection.Value := 'IMPORTE PERCNOTAS';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';
    Inc (iFila,2);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'MATERIAL';
    Inc (iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'TRAZABILIDAD';

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'DESCRIPCION';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'UNIDAD';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    PFormatosExcel_Rellenar(Excel, $00BBBBBB);
    Excel.Selection.Value := 'CANTIDAD';
    Inc (iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';

    Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '';
    Inc (iFila,2);

    Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlRight;
    Excel.Selection.Value := 'COSTO TOTAL DE LA ACTIVIDAD';

    Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
    Excel.Selection.Value := '$0';


{$ENDREGION}

  Finally
    //;
  End;
end;


procedure TfrmReportePeriodo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    BotonPermiso.free;
    action := cafree ;
end;

procedure TfrmReportePeriodo.tsNumeroActividadChange(Sender: TObject);
begin
    if QryPartidasEfectivas.RecordCount > 0 then
    begin
        FotosPartidas.Active := False;
        FotosPartidas.ParamByName('Contrato').AsString  := global_contrato;
        FotosPartidas.ParamByName('Orden').AsString     := tsNumeroOrdenActa.Text;
        FotosPartidas.ParamByName('Wbs').AsString       := QryPartidasefectivas.FieldValues['sWbs'];
        FotosPartidas.ParamByName('Actividad').AsString := QryPartidasEfectivas.FieldValues['sNumeroActividad'];
        FotosPartidas.Open;

        if FotosPartidas.RecordCount = 0 then
        begin
//            bImagen.Picture.LoadFromFile('');
            messageDLG('No Existen Imagenes Asignadas a la Partida '+QryPartidasEfectivas.FieldValues['sNumeroActividad'], mtInformation, [mbOk],0);
        end;
    end;
end;

procedure TfrmReportePeriodo.tsNumeroActividadEnter(Sender: TObject);
begin
    tsNumeroActividad.Color := global_color_entrada;
    QryPartidasEfectivas.Active := False;
    QryPartidasEfectivas.Params.ParamByName('contrato').DataType := ftString;
    QryPartidasEfectivas.Params.ParamByName('contrato').Value    := global_contrato;
    QryPartidasEfectivas.Params.ParamByName('convenio').DataType := ftString;
    QryPartidasEfectivas.Params.ParamByName('convenio').Value    := global_convenio;
    QryPartidasEfectivas.Params.ParamByName('Orden').DataType    := ftString;
    QryPartidasEfectivas.Params.ParamByName('Orden').Value       := tsNumeroOrdenActa.Text;
    QryPartidasEfectivas.Open;
end;

procedure TfrmReportePeriodo.tsNumeroActividadExit(Sender: TObject);
begin
    tsNumeroActividad.Color := global_color_salida;
end;

procedure TfrmReportePeriodo.tsNumeroActividadKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key = #13 then
       grid_imagenes.SetFocus;
end;

procedure TfrmReportePeriodo.tsNumeroOrden3Click(Sender: TObject);
begin
  if ReportesCheck.Checked = true then
    ReportesCheckClick(nil);
  
end;

procedure TfrmReportePeriodo.tsNumeroOrden3Enter(Sender: TObject);
begin
  tsnumeroorden3.Color:=global_color_entrada
end;

procedure TfrmReportePeriodo.tsNumeroOrden3Exit(Sender: TObject);
begin
  tsnumeroorden3.Color:=global_color_salida
end;

procedure TfrmReportePeriodo.tsNumeroOrden3KeyPress(Sender: TObject;
  var Key: Char);
begin
   If Key = #13 Then
      tsNumeroOrden.SetFocus  
end;

procedure TfrmReportePeriodo.tsNumeroOrdenActaEnter(Sender: TObject);
begin
    tsNumeroOrdenActa.Color := global_color_entrada;
end;

procedure TfrmReportePeriodo.tsNumeroOrdenActaExit(Sender: TObject);
begin
    tsNumeroOrdenActa.Color := global_color_salida;
    //Ahora el llenado del combobox de las actas fotograficas,
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('select sActaFotografica from reportefotografico_acta where sContrato =:Contrato and sNumeroOrden =:Orden group by sActaFotografica ');
    connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
    connection.QryBusca.ParamByName('Orden').AsString    := tsNumeroOrdenActa.Text;
    connection.QryBusca.Open;

    ComboActas.Items.Clear;
    while not connection.QryBusca.Eof do
    begin
        ComboActas.Items.Add(connection.QryBusca.FieldValues['sActaFotografica']);
        connection.QryBusca.Next;
    end;

    if comboActas.Items.Count > 0 then
    begin
       comboActas.ItemIndex := 0;
       comboActas.OnExit(sender);
    end;
end;

procedure TfrmReportePeriodo.ImprimeExcel_Cronologias;
Var
  Libro, Excel, Hoja: Variant;
  iFila, iColumna, iHojas, iCounter, iLoop: Integer;
  NombreDelExcel, SQLExtra: String;

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

begin
  Try
    NombreDelExcel := PGetTempDir + 'TEMP~' + PNombreAleatorio(3) + 'ReporteDiario.xls';
    Try
      Excel := CreateOleObject('Excel.Application');
    Except
      On E: Exception do begin
        FreeAndNil(Excel);
        ShowMessage(E.Message);
        Exit;
      end;
    End;
    Excel.Visible := True;
    Excel.DisplayAlerts:= False;
    Libro := Excel.Workbooks.Add;

    Excel.WorkBooks[1].WorkSheets[1].Name := 'ORDEN 015';
    Hoja := Excel.WorkBooks[1].WorkSheets[1];
    Libro.Sheets[1].Select;

    if rbCronologiasFiltradas.Checked then begin
      SQLExtra := ' AND b.sNumeroOrden = :Orden ';
    end else begin
      SQLExtra := ' ';
    end;
    
    iColumna := 1;
    iFila := 1;

    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 24;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 20;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7; //G
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7;
    Inc(iColumna);
    Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 58;
    Inc(iColumna);

    iColumna := 1;

    If reportesCheck.Checked Then begin
        Connection.QryBusca.SQL.Clear;
    Connection.QryBusca.SQL.Text := 'SELECT b.*, ' +
                                    '	( ' +
                                    '		SELECT ' +
                                    '			(ifnull(sum(ba.dAvance), 0)) ' +
                                    '		FROM ' +
                                    '			bitacoradeactividades AS ba ' +
                                    '		WHERE ' +
                                    '			ba.sContrato = b.sContrato ' +
                                    '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                    '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                    '   AND ba.sIdClasificacion = b.sIdClasificacion ' +
                                    '   AND ba.sIdTipoMovimiento = "ED" ' +
                                    '		AND ( ' +
                                    '			ba.didfecha < b.didfecha ' +
                                    '			OR ( ' +
                                    '				ba.didfecha = b.didfecha ' +
                                    '				AND cast(ba.sHoraInicio AS Time) < cast(b.sHoraInicio AS Time) ' +
                                    '			) ' +
                                    '		) ' +
                                    '	) AS dAvanceAnteriorPorPartida ' +
                                    ' FROM bitacoradeactividades AS b WHERE b.sContrato = :Contrato ' + SQLExtra + ' AND b.sIdTipoMovimiento = "ED" ' +
                                    ' ORDER BY dIdFecha ASC;';
    Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
    if rbCronologiasFiltradas.Checked then begin
      Connection.QryBusca.Params.ParamByName('Orden').AsString := tsNumeroOrden3.Text;
    end;
    Connection.QryBusca.Open;

    Connection.QryBusca.SQL.Clear;
    Connection.QryBusca.SQL.Text := 'SELECT b.*, ' +
                                    '	( ' +
                                    '		SELECT ' +
                                    '			(ifnull(sum(ba.dAvance), 0)) ' +
                                    '		FROM ' +
                                    '			bitacoradeactividades AS ba ' +
                                    '		WHERE ' +
                                    '			ba.sContrato = b.sContrato ' +
                                    '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                    '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                    '   AND ba.sIdClasificacion = b.sIdClasificacion ' +
                                    '   AND ba.sIdTipoMovimiento = "ED" ' +
                                    '		AND ( ' +
                                    '			ba.didfecha < b.didfecha ' +
                                    '			OR ( ' +
                                    '				ba.didfecha = b.didfecha ' +
                                    '				AND cast(ba.sHoraInicio AS Time) < cast(b.sHoraInicio AS Time) ' +
                                    '			) ' +
                                    '		) ' +
                                    '	) AS dAvanceAnteriorPorPartida ' +
                                    ' FROM bitacoradeactividades AS b WHERE b.sContrato = :Contrato ' + SQLExtra + ' AND b.sIdTipoMovimiento = "ED" ' +
                                    ' ORDER BY dIdFecha ASC;';
    Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
    if rbCronologiasFiltradas.Checked then begin
      Connection.QryBusca.Params.ParamByName('Orden').AsString := tsNumeroOrden3.Text;
    end;
    Connection.QryBusca.Open;

     end else begin

    Connection.QryBusca.SQL.Clear;
    Connection.QryBusca.SQL.Text := 'SELECT b.*, ' +
                                    '	( ' +
                                    '		SELECT ' +
                                    '			(ifnull(sum(ba.dAvance), 0)) ' +
                                    '		FROM ' +
                                    '			bitacoradeactividades AS ba ' +
                                    '		WHERE ' +
                                    '			ba.sContrato = b.sContrato ' +
                                    '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                    '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                    '   AND ba.sIdClasificacion = b.sIdClasificacion ' +
                                    '   AND ba.sIdTipoMovimiento = "ED" ' +
                                    '		AND ( ' +
                                    '			ba.didfecha < b.didfecha ' +
                                    '			OR ( ' +
                                    '				ba.didfecha = b.didfecha ' +
                                    '				AND cast(ba.sHoraInicio AS Time) < cast(b.sHoraInicio AS Time) ' +
                                    '			) ' +
                                    '		) ' +
                                    '	) AS dAvanceAnteriorPorPartida ' +
                                    ' FROM bitacoradeactividades AS b WHERE b.sContrato = :Contrato ' + SQLExtra + ' AND b.dIdFecha > :FechaInicial AND b.dIdFecha < :FechaFinal AND b.sIdTipoMovimiento = "ED" ' +
                                    ' ORDER BY dIdFecha ASC;';
    Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
    Connection.QryBusca.Params.ParamByName('FechaInicial').AsDateTime := tdFechaInicial.DateTime;
    Connection.QryBusca.Params.ParamByName('FechaFinal').AsDateTime := tdFechaFinal.DateTime;
    if rbCronologiasFiltradas.Checked then begin
      Connection.QryBusca.Params.ParamByName('Orden').AsString := tsNumeroOrden3.Text;
    end;
    Connection.QryBusca.Open;

     end;

    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    Excel.Selection.Value := 'FECHA';
    PFormatosExcel_Bordes(Excel);
    Inc(iColumna);

    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    Excel.Selection.Value := 'FOLIO';
    PFormatosExcel_Bordes(Excel);
    Inc(iColumna);

    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'PDA';
    Inc(iColumna);

    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'CLAS.';
    Inc(iColumna);

    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'INI.';
    Inc(iColumna);
    
    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'TERM.';
    Inc(iColumna);

    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'ANT';
    Inc(iColumna);

    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'AVANCE';
    Inc(iColumna);

    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'ACUM';
    Inc(iColumna);

    Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow');
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'DESCRIPCIÓN DEL TRABAJO';
    Inc(iColumna);
    Inc(iFila);


    if Connection.QryBusca.RecordCount > 0 then begin
      while not Connection.QryBusca.Eof do begin
        iColumna := 1;
        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 11, clBlack, 'Arial Narrow', '[$-80A]d" de "mmmm" de "aaaa;@');
        Excel.Selection.Value := Connection.QryBusca.FieldByName('dIdFecha').AsDateTime;
        Inc(iColumna);

        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 11, clBlack, 'Arial Narrow', '@');
        Excel.Selection.Value := Trim(Connection.QryBusca.FieldByName('sNumeroOrden').AsString);
        Inc(iColumna);

        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 11, clBlack, 'Arial Narrow', '@');
        Excel.Selection.Value := Trim(Connection.QryBusca.FieldByName('sNumeroActividad').AsString);
        Inc(iColumna);

        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, True, 11, clBlack, 'Arial Narrow', '@');
        Excel.Selection.Value := Trim(Connection.QryBusca.FieldByName('sIdClasificacion').AsString);
        Inc(iColumna);

        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 11, clBlack, 'Arial Narrow', '@');
        Excel.Selection.Value := Trim(Connection.QryBusca.FieldByName('sHoraInicio').AsString);
        Inc(iColumna);

        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 11, clBlack, 'Arial Narrow', '@');
        Excel.Selection.Value := Trim(Connection.QryBusca.FieldByName('sHoraFinal').AsString);
        Inc(iColumna);

        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 11, clBlack, 'Arial Narrow', '00.00%');
        Excel.Selection.Value := Connection.QryBusca.FieldByName('dAvanceAnteriorPorPartida').AsFloat;
        Inc(iColumna);

        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 11, clBlack, 'Arial Narrow', '00.00%');
        Excel.Selection.Value := Connection.QryBusca.FieldByName('dAvance').AsFloat;
        Inc(iColumna);

        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 11, clBlack, 'Arial Narrow', '00.00%');
        Excel.Selection.Formula := '='+ColumnaNombre(iColumna - 2)+IntToStr(iFila)+'+'+ColumnaNombre(iColumna - 1)+IntToStr(iFila);
        Inc(iColumna);

        Excel.Range[ColumnaNombre(iColumna)+IntToStr(iFila)+':'+ColumnaNombre(iColumna)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 11, clBlack, 'Arial Narrow', '@');
        Excel.Selection.Value := Connection.QryBusca.FieldByName('mDescripcion').AsString;
        Excel.Selection.WrapText := False;
        Excel.Selection.HorizontalAlignment := xlLeft;
        Inc(iColumna);
        Inc(iFila);
        Connection.QryBusca.Next;
      end;
    end;
  Finally
//    FreeAndNil(Excel);
  End;
end;

procedure TfrmReportePeriodo.tsNumeroOrdenActaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key = #13 then
       tsNumeroActividad.SetFocus;
end;

procedure TfrmReportePeriodo.tsNumeroOrdenEnter(
  Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmReportePeriodo.tsNumeroOrdenExit(Sender: TObject);
begin
  tsNumeroOrden.Color := global_color_salida
end;

procedure TfrmReportePeriodo.tsNumeroOrdenKeyPress(
  Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        dbfiltro.SetFocus
end;

procedure TfrmReportePeriodo.btnPreview2Click(Sender: TObject);
var
   bS  : TStream;
   Pic : TJpegImage;
   BlobField : tField ;
begin
    If fotografico_acta.RecordCount > 0 then
    Begin
//        bImagen.Picture.LoadFromFile('') ;
        BlobField := fotografico_acta.FieldByName('bImagen') ;
        BS := fotografico_acta.CreateBlobStream(BlobField, bmRead) ;
          //.CreateBlobStream ( BlobField , bmRead ) ;
        If bs.Size > 1 Then
        Begin
            try
                Pic:=TJpegImage.Create;
                try
                   Pic.LoadFromStream(bS);
                   bImagen2.Picture.Graphic := Pic;
                finally
                   Pic.Free;
                end;
            finally
                bS.Free
            End
        End;
    End;
end;

procedure TfrmReportePeriodo.btnPreviewClick(Sender: TObject);
var
   bS  : TStream;
   Pic : TJpegImage;
   BlobField : tField ;
begin
    If FotosPartidas.RecordCount > 0 then
    Begin
//        bImagen.Picture.LoadFromFile('') ;
        BlobField := FotosPartidas.FieldByName('bImagen') ;
        BS := FotosPartidas.CreateBlobStream(BlobField, bmRead) ;
          //.CreateBlobStream ( BlobField , bmRead ) ;
        If bs.Size > 1 Then
        Begin
            try
                Pic:=TJpegImage.Create;
                try
                   Pic.LoadFromStream(bS);
                   bImagen.Picture.Graphic := Pic;
                finally
                   Pic.Free;
                end;
            finally
                bS.Free
            End
        End;
        labelFase.Caption := ' Fase: '+ fotospartidas.FieldValues['sFasePartida'];
    End;
end;

procedure TfrmReportePeriodo.btnPrinterClick(Sender: TObject);
Var
    iCuadrilla     : Integer ;
    sDescripcion   : String ;
    sMuerto,
    sTmpMuerto     : String ;
    sFecha         : String ;
    sOrden         : String ;
    sLinea,
    sLinea2        : String ;
    lImprime       : Boolean ;
    lRegistra      : Boolean ;
    QryGraficaTiemposMuertos,
    QryReportesDiarios,
    QryTiempoMuerto        : tzReadOnlyquery ;
    sDir: string;
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdFechaFinal.Date<tdFechaInicial.Date then
   begin
   showmessage('la fecha final de impresión es menor a la fecha inicial de impresión' );
   tdFechaFinal.SetFocus;
   exit;
   end;
  try
    if not BotonPermiso.imprimir then
    begin
      showmessage('No tiene permisos de impresión');
      exit;
    end;
      
    QryReportesDiarios := tzReadOnlyQuery.Create(Self) ;
    QryReportesDiarios.Connection := connection.zConnection ;

    QryTiempoMuerto := tzReadOnlyQuery.Create(Self) ;
    QryTiempoMuerto.Connection := connection.zConnection ;

    QryGraficaTiemposMuertos := tzReadOnlyQuery.Create(Self) ;
    QryGraficaTiemposMuertos.Connection := connection.zConnection ;

    dsQryGraficaTiemposMuertos.DataSet  := QryGraficaTiemposMuertos ;
    dsQryGraficaTiemposMuertos.UserName := 'dsQryGraficaTiemposMuertos' ;

    sSuperIntendente := '' ;
    sSupervisor := '' ;
    sPuestoSuperintendente := '' ;
    sPuestoSupervisor := '' ;
    sSupervisorTierra := '' ;
    sPuestoSupervisorTierra := '' ;

    if dbFiltro.KeyValue <> null then
       sLinea2 := ' and j.sIdTipoMovimiento =:Filtro ';

    QryGraficaTiemposMuertos.Active := False ;
    QryGraficaTiemposMuertos.SQL.Clear ;

    QryReportesDiarios.Active := False ;
    QryReportesDiarios.SQL.Clear ;

    If chkConsolidado.Checked Then
    Begin
        QryGraficaTiemposMuertos.SQL.Add('Select concat("CONTRATO No. " , r.sContrato) as sNumeroOrden, r.dIdFecha, ' +
                                         '(sum(substr(j.sTiempoMuerto, 1 , 2)) + sum(substr(j.sTiempoMuerto, 4 , 2)) div 60 + ' +
                                         '(sum(substr(j.sTiempoMuerto, 4 , 2)) % 60 ) / 100 ) as iTiempoMuertoReal From reportediario r ' +
                                         'inner join turnos t on (r.sContrato = t.sContrato and r.sIdTurno = t.sIdTurno And t.sOrigenTierra = "No") ' +
                                         'left join jornadasdiarias j on (r.sContrato = j.sContrato And r.sNumeroOrden = j.sNumeroOrden And r.dIdFecha = j.dIdFecha And r.sIdTurno = j.sIdTurno And j.sTipo = "Tiempo Inactivo" '+sLinea2+')'+
                                         'inner join tiposdemovimiento tm on (t.sContrato = tm.sContrato And j.sIdTipoMovimiento = tm.sIdTipoMovimiento) '+
                                         'Where r.sContrato = :Contrato And r.dIdFecha >= :FechaI And ' +
                                         'r.dIdFecha <= :FechaF And r.lStatus = "Autorizado" Group By r.dIdfecha Order By r.dIdFecha') ;

        QryReportesDiarios.SQL.Add('Select r.sNumeroOrden, r.sNumeroReporte, r.dIdFecha, r.sIdTurno, r.sIdUsuarioAutoriza, r.sOperacionInicio, ' +
                                   'r.sOperacionFinal, r.sTiempoAdicional, r.sTiempoEfectivo, r.sTiempoMuerto, r.sTiempoMuertoReal, r.iPersonal From reportediario r ' +
                                   'inner join turnos t on (r.sContrato = t.sContrato and r.sIdTurno = t.sIdTurno And t.sOrigenTierra = "No") ' +
                                   'where r.sContrato = :Contrato And r.dIdFecha >= :FechaInicio And r.dIdFecha <= :FechaFinal And r.lStatus = "Autorizado" ' +
                                   'Order By r.dIdFecha')
    End
    Else
    Begin
        QryGraficaTiemposMuertos.SQL.Add('Select concat("ORDEN DE TRABAJO No. " , r.sNumeroOrden) as sNumeroOrden, r.dIdFecha, ' +
                                         '(sum(substr(j.sTiempoMuerto, 1 , 2)) + sum(substr(j.sTiempoMuerto, 4 , 2)) div 60 + ' +
                                         '(sum(substr(j.sTiempoMuerto, 4 , 2)) % 60 ) / 100 ) as iTiempoMuertoReal From reportediario r ' +
                                         'inner join turnos t on (r.sContrato = t.sContrato and r.sIdTurno = t.sIdTurno And t.sOrigenTierra = "No") ' +
                                         'left join jornadasdiarias j on (r.sContrato = j.sContrato And r.sNumeroOrden = j.sNumeroOrden And r.dIdFecha = j.dIdFecha And r.sIdTurno = j.sIdTurno And j.sTipo = "Tiempo Inactivo" '+sLinea2+')'+
                                         'inner join tiposdemovimiento tm on (t.sContrato = tm.sContrato And j.sIdTipoMovimiento = tm.sIdTipoMovimiento) '+
                                         'Where r.sContrato = :Contrato And r.dIdFecha >= :FechaI And ' +
                                         'r.dIdFecha <= :FechaF And r.lStatus = "Autorizado" And r.sNumeroOrden = :Orden Group By r.dIdfecha Order By r.dIdFecha') ;
        QryGraficaTiemposMuertos.Params.ParamByName('orden').DataType := ftString ;
        QryGraficaTiemposMuertos.Params.ParamByName('orden').Value := tsNumeroOrden.Text ;

        QryReportesDiarios.SQL.Add('Select r.sNumeroOrden, r.sNumeroReporte, r.dIdFecha, r.sIdTurno, r.sIdUsuarioAutoriza, r.sOperacionInicio, ' +
                                   'r.sOperacionFinal, r.sTiempoAdicional, r.sTiempoEfectivo, r.sTiempoMuerto, r.sTiempoMuertoReal, r.iPersonal From reportediario r ' +
                                   'inner join turnos t on (r.sContrato = t.sContrato and r.sIdTurno = t.sIdTurno And t.sOrigenTierra = "No") ' +
                                   'where r.sContrato = :Contrato And r.sNumeroOrden = :Orden And r.dIdFecha >= :FechaInicio And r.dIdFecha <= :FechaFinal And r.lStatus = "Autorizado" ' +
                                   'Order By r.dIdFecha') ;
        QryReportesDiarios.Params.ParamByName('Orden').DataType := ftString ;
        QryReportesDiarios.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
    End ;

    QryReportesDiarios.Params.ParamByName('Contrato').DataType := ftString ;
    QryReportesDiarios.Params.ParamByName('Contrato').Value := global_contrato ;
    QryReportesDiarios.Params.ParamByName('FechaInicio').DataType := ftDate ;
    QryReportesDiarios.Params.ParamByName('FechaInicio').Value := tdFechaInicial.Date ;
    QryReportesDiarios.Params.ParamByName('FechaFinal').DataType := ftDate ;
    QryReportesDiarios.Params.ParamByName('FechaFinal').Value := tdFechaFinal.Date ;
    QryReportesDiarios.Open ;

    rxInactivos.EmptyTable ;
    if dbFiltro.KeyValue <> null then
       sLinea := ' and t.sIdTipoMovimiento =:Filtro ';

    sHoraResult := '00:00' ;
    iReportes := 0 ;
    sFecha := '' ;
    sOrden := '' ;
    While NOT QryReportesDiarios.Eof Do
    Begin
        iReportes := iReportes + 1 ;
        iCuadrilla := 0 ;

        Connection.QryBusca.Active := False ;
        Connection.QryBusca.SQL.Clear ;
        Connection.QryBusca.SQL.Add('Select SUM(b.dCantidad) as dTotal From bitacoradepersonal b ' +
                                    'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha and b.iIdDiario = b2.iIdDiario) ' +
                                    'INNER JOIN turnos t ON (b2.sContrato = t.sContrato And b2.sIdTurno = t.sIdTurno And t.sOrigenTierra = "No")' +
                                    'Where b.sContrato = :Contrato And b.dIdFecha = :Fecha Group By b2.sContrato') ;
        Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.QryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.QryBusca.Params.ParamByName('Fecha').Value := QryReportesDiarios.FieldValues['dIdFecha'] ;
        Connection.QryBusca.Open ;
        If Connection.QryBusca.RecordCount > 0 Then
            iCuadrilla := Connection.QryBusca.FieldValues['dTotal'] ;

        lImprime := False ;
        if dbFiltro.KeyValue = null then
           sHoraResult := sfnSumaHoras ( sHoraResult, QryReportesDiarios.FieldValues['sTiempoMuerto'] );

        //Aqui se cambia para el tiempo muerto del capri
        qryTiempoMuerto.Active := False ;
        qryTiempoMuerto.SQL.Clear ;
        qryTiempoMuerto.SQL.Add('Select t.sIdPernocta, t.sHoraInicio, t.sHoraFinal, t.dPersonal, t.dFrente, t.sTiempoMuerto, t.mDescripcion, tm.sDescripcion From jornadasdiarias t ' +
                             'INNER JOIN tiposdemovimiento tm on (t.sContrato = tm.sContrato And t.sIdTipoMovimiento = tm.sIdTipoMovimiento) ' +
                             'Where t.sContrato = :Contrato And t.sNumeroOrden = :Orden And t.dIdFecha = :fecha And t.sIdTurno = :Turno And t.sTipo = "Tiempo Inactivo"'+ sLinea) ;
        qryTiempoMuerto.Params.ParamByName('Contrato').DataType := ftString ;
        qryTiempoMuerto.Params.ParamByName('Contrato').Value    := global_contrato ;
        qryTiempoMuerto.Params.ParamByName('Orden').DataType    := ftString ;
        qryTiempoMuerto.Params.ParamByName('Orden').Value       := QryReportesDiarios.FieldValues['sNumeroOrden'] ;
        qryTiempoMuerto.Params.ParamByName('Fecha').DataType    := ftDate ;
        qryTiempoMuerto.Params.ParamByName('Fecha').Value       := QryReportesDiarios.FieldValues['dIdFecha'] ;
        qryTiempoMuerto.Params.ParamByName('Turno').DataType    := ftString ;
        qryTiempoMuerto.Params.ParamByName('Turno').Value       := QryReportesDiarios.FieldValues['sIdTurno'] ;
        if dbFiltro.KeyValue <> null then
        begin
            qryTiempoMuerto.Params.ParamByName('Filtro').DataType := ftString ;
            qryTiempoMuerto.Params.ParamByName('Filtro').Value    := dbFiltro.KeyValue ;
        end;
        qryTiempoMuerto.Open ;
        lImprime := True ;
        If qryTiempoMuerto.RecordCount > 0 Then
            While NOT qryTiempoMuerto.Eof Do
            Begin
                If chkTiempoMuerto.Checked Then
                    If QryReportesDiarios.FieldValues['sTiempoMuerto'] <> '00:00' Then
                        lRegistra := True
                    Else
                        lRegistra := False
                Else
                    lRegistra := True ;

                If lRegistra Then
                Begin
                    rxInactivos.Append ;
                    rxInactivos.FieldValues ['sContrato']          := global_contrato ;
                    rxInactivos.FieldValues ['dIdFecha']           := QryReportesDiarios.FieldValues['dIdFecha'] ;
                    rxInactivos.FieldValues ['sNumeroReporte']     := QryReportesDiarios.FieldValues['sNumeroReporte'] ;
                    rxInactivos.FieldValues ['sIdUsuarioAutoriza'] := QryReportesDiarios.FieldValues['sIdUsuarioAUtoriza'] ;
                    rxInactivos.FieldValues ['sHoraInicio']        := QryReportesDiarios.FieldValues['sOperacionInicio'] ;
                    rxInactivos.FieldValues ['sHoraFinal']         := QryReportesDiarios.FieldValues['sOperacionFinal'] ;
                    rxInactivos.FieldValues ['sHoraInicio']        := QryReportesDiarios.FieldValues['sOperacionInicio'] ;
                    rxInactivos.FieldValues ['sTiempoEfectivo']    := QryReportesDiarios.FieldValues['sTiempoEfectivo'] ;
                    rxInactivos.FieldValues ['sTiempoMuerto']      := QryReportesDiarios.FieldValues['sTiempoMuerto'] ;
                    If lImprime Then
                    Begin
                        If QryReportesDiarios.FieldValues['sTiempoMuertoReal'] <> '00:00' Then
                        Begin
                             if dbFiltro.KeyValue <> null then
                                rxInactivos.FieldValues ['sTiempoMuertoReal']  := qryTiempoMuerto.FieldValues['sTiempoMuerto']
                             else
                                 rxInactivos.FieldValues ['sTiempoMuertoReal'] := QryReportesDiarios.FieldValues['sTiempoMuerto'];
                            lImprime := False;
                        End
                    End
                    Else
                    Begin
                        rxInactivos.FieldValues ['sTiempoMuerto']     := '' ;
                        if dbFiltro.KeyValue <> null then
                           rxInactivos.FieldValues ['sTiempoMuertoReal'] := qryTiempoMuerto.FieldValues['sTiempoMuerto']
                        else
                            rxInactivos.FieldValues ['sTiempoMuertoReal'] := '';
                    End ;

                    if dbFiltro.KeyValue <> null then
                       sHoraResult := sfnSumaHoras ( sHoraResult, qryTiempoMuerto.FieldValues['sTiempoMuerto'] ) ;

                    rxInactivos.FieldValues ['tmsDescripcion']   := qryTiempoMuerto.FieldValues['sDescripcion'] ;
                    rxInactivos.FieldValues ['tmsHoraInicio']    := qryTiempoMuerto.FieldValues['sHoraInicio'] ;
                    rxInactivos.FieldValues ['tmsHoraFinal']     := qryTiempoMuerto.FieldValues['sHoraFinal'] ;
                    rxInactivos.FieldValues ['tmsTiempoMuerto']  := qryTiempoMuerto.FieldValues['sTiempoMuerto'] ;
                    rxInactivos.FieldValues ['tmiPersonal']      := qryTiempoMuerto.FieldValues['dFrente'] ;
                    rxInactivos.FieldValues ['tmiPersonalOrden'] := QryReportesDiarios.FieldValues['iPersonal'] ;
                    rxInactivos.FieldValues['iPersonal']         := iCuadrilla ;
                    rxInactivos.FieldValues ['tmmDescripcion']   := qryTiempoMuerto.FieldValues['mDescripcion'] ;
                    rxInactivos.Post;
                End ;
                qryTiempoMuerto.Next
        End;
        QryReportesDiarios.Next ;
    End ;

    { SOLO PATA MOSTRAR ENCABEZADOS SINO EXISEN TIEMPOS MUERTOS }
    if rxInactivos.RecordCount = 0 then
    begin
        rxInactivos.Append ;
        rxInactivos.FieldValues ['sContrato'] := global_contrato ;
        rxInactivos.Post;
    end;

    QryGraficaTiemposMuertos.Params.ParamByName('Contrato').DataType := ftString ;
    QryGraficaTiemposMuertos.Params.ParamByName('Contrato').Value    := global_contrato ;
    QryGraficaTiemposMuertos.Params.ParamByName('FechaI').DataType   := ftDate ;
    QryGraficaTiemposMuertos.Params.ParamByName('FechaI').Value      := tdFechaInicial.Date ;
    QryGraficaTiemposMuertos.Params.ParamByName('FechaF').DataType   := ftDate ;
    QryGraficaTiemposMuertos.Params.ParamByName('FechaF').Value      := tdFechaFinal.Date ;
    if dbFiltro.KeyValue <> null then
    begin
         QryGraficaTiemposMuertos.Params.ParamByName('Filtro').DataType := ftString ;
         QryGraficaTiemposMuertos.Params.ParamByName('Filtro').Value    := dbFiltro.KeyValue ;
    end;
    QryGraficaTiemposMuertos.Open ;

    { MOSTRAR EN CERO LA GRAFICA SINO EXITEN TIEMPOS MUERTOS }
    if QryGraficaTiemposMuertos.RecordCount = 0 then
    begin

        QryGraficaTiemposMuertos.Active := False ;
        QryGraficaTiemposMuertos.SQL.Clear ;
        If chkConsolidado.Checked Then
            QryGraficaTiemposMuertos.SQL.Add('Select concat("CONTRATO No. " , sContrato) as sNumeroOrden, dIdFecha, 0 as iTiempoMuertoReal From reportediario '+
                                             'where sContrato =:Contrato and dIdFecha =:Fecha ')
        else
        begin
            QryGraficaTiemposMuertos.SQL.Add('Select concat("ORDEN DE TRABAJO No. ", sNumeroOrden) as sNumeroOrden, dIdFecha, 0 as iTiempoMuertoReal From reportediario '+
                                             'where sContrato =:Contrato and sNumeroOrden =:Orden and dIdFecha =:Fecha ') ;
            QryGraficaTiemposMuertos.Params.ParamByName('Orden').DataType := ftString ;
            QryGraficaTiemposMuertos.Params.ParamByName('Orden').Value    := tsNumeroOrden.Text ;
        end;
        QryGraficaTiemposMuertos.Params.ParamByName('Contrato').DataType  := ftString ;
        QryGraficaTiemposMuertos.Params.ParamByName('Contrato').Value     := global_contrato ;
        QryGraficaTiemposMuertos.Params.ParamByName('Fecha').DataType     := ftDate ;
        QryGraficaTiemposMuertos.Params.ParamByName('Fecha').Value        := tdFechaFinal.Date ;
        QryGraficaTiemposMuertos.Open;
    end;

    rDiarioFirmas (global_contrato, tsNumeroOrden.Text, 'A',tdFechaFinal.Date, frmReportePeriodo) ;

    if (rxInactivos.RecordCount > 0) or (QryGraficaTiemposMuertos.RecordCount > 0) then
    begin
        rReporte.PreviewOptions.MDIChild  := False ;
        rReporte.PreviewOptions.Modal     := True ;
        rReporte.PreviewOptions.Maximized := lCheckMaximized () ;
        rReporte.PreviewOptions.ShowCaptions := False ;
        rReporte.Previewoptions.ZoomMode  := zmPageWidth ;

        If chkConsolidado.Checked Then
           sDir := 'ReporteConsolidadodeTiemposMuertos.fr3'
           //rReporte.LoadFromFile (global_files + 'ReporteConsolidadodeTiemposMuertos.fr3')
        Else
           sDir := 'ConcentradodeReportesDiarios.fr3';
           //rReporte.LoadFromFile (global_files + 'ConcentradodeReportesDiarios.fr3') ;
        rReporte.LoadFromFile (global_files + sDir) ;
        rReporte.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP)) ;
        if not FileExists(global_files + sDir) then
            showmessage('El archivo de reporte ' + sDir + '.fr3 no existe, notifique al administrador del sistema');
    end
    else begin
        showmessage('No hay datos para imprimir');
    end;

    QryGraficaTiemposMuertos.Destroy ;
    QryReportesDiarios.Destroy;
    qryTiempoMuerto.Destroy ;
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Album Fotografico/Tiempos Muertos ...', 'Al imprimir tiempos muertos', 0);
    end;
  end;
end;

procedure TfrmReportePeriodo.btnUpClick(Sender: TObject);
begin
    OrdenarFotos('Arriba');
end;

procedure TfrmReportePeriodo.btnInsertClick(Sender: TObject);
var
   iItem : integer;
begin
    if comboActas.Text = '' then
    begin
       messageDLG('Seleccione una Acta Fotografica o Cree una nueva!', mtInformation, [mbOk], 0);
       exit;
    end;

    if fotosPartidas.RecordCount > 0 then
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select max(iImagen) as iImagen from reportefotografico_acta '+
                                    'where sContrato =:Contrato and sNumeroOrden =:Orden and sActaFotografica =:Acta group by sContrato');
        connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
        connection.zcommand.Params.ParamByName('Contrato').Value    := Global_Contrato ;
        connection.zcommand.Params.ParamByName('Orden').DataType    := ftString ;
        connection.zcommand.Params.ParamByName('Orden').Value       := tsNumeroOrdenActa.Text;
        connection.zcommand.Params.ParamByName('Acta').DataType     := ftString ;
        connection.zcommand.Params.ParamByName('Acta').Value        := ComboActas.Text ;
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
           iItem := connection.zCommand.FieldValues['iImagen'] + 1
        else
           iItem := 1;

        connection.zCommand.Active := False ;
        connection.zCommand.SQL.Clear ;
        connection.zcommand.SQL.Add ( 'Insert Into reportefotografico_acta (sContrato, sNumeroOrden, sActaFotografica, dIdFecha, iImagen, bImagen, sWbs, sNumeroActividad, sFasePartida) ' +
                                      'Values (:Contrato, :Orden, :Acta, :Fecha, :Item, :Imagen, :Wbs, :Actividad, :Fase)') ;
        connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
        connection.zcommand.Params.ParamByName('Contrato').Value    := Global_Contrato ;
        connection.zcommand.Params.ParamByName('Orden').DataType    := ftString ;
        connection.zcommand.Params.ParamByName('Orden').Value       := tsNumeroOrdenActa.Text;
        connection.zcommand.Params.ParamByName('Acta').DataType     := ftString ;
        connection.zcommand.Params.ParamByName('Acta').Value        := ComboActas.Text ;
        connection.zcommand.Params.ParamByName('Fecha').DataType    := ftDate ;
        connection.zcommand.Params.ParamByName('Fecha').Value       := date ;
        connection.zcommand.Params.ParamByName('Item').DataType     := ftInteger ;
        connection.zcommand.Params.ParamByName('Item').Value        := iItem ;
        connection.zcommand.Params.ParamByName('Imagen').DataType   := ftBlob;
        connection.zcommand.Params.ParamByName('Imagen').Value      := fotosPartidas.FieldValues['bImagen'];
        connection.zcommand.Params.ParamByName('Wbs').DataType      := ftString ;
        connection.zcommand.Params.ParamByName('Wbs').Value         := fotosPartidas.FieldValues['sWbs'];
        connection.zcommand.Params.ParamByName('Actividad').DataType:= ftString ;
        connection.zcommand.Params.ParamByName('Actividad').Value   := fotosPartidas.FieldValues['sNumeroActividad'];
        connection.zcommand.Params.ParamByName('Fase').DataType     := ftString ;
        connection.zcommand.Params.ParamByName('Fase').Value        := fotosPartidas.FieldValues['sFasePartida'];
        connection.zCommand.ExecSQL();

        // Actualizo Kardex del Sistema ....
        connection.zCommand.Active := False ;
        connection.zCommand.SQL.Clear ;
        connection.zcommand.SQL.Add ('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen)' +
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
        connection.zcommand.Params.ParamByName('Descripcion').Value := 'Agrega Fotografias a Reporte Actafotografica ' + comboActas.Text + ' del dia ' + DateToStr(Date) ;
        connection.zcommand.Params.ParamByName('Origen').DataType := ftString ;
        connection.zcommand.Params.ParamByName('Origen').Value := 'Reporte Fotografico' ;
        connection.zCommand.ExecSQL ;

        fotografico_Acta.Active := False;
        fotografico_acta.ParamByName('Contrato').AsString := global_contrato;
        fotografico_acta.ParamByName('Orden').AsString    := tsNumeroOrdenActa.Text;
        fotografico_acta.ParamByName('Acta').AsString     := ComboActas.Text;
        fotografico_acta.Open;

        fotografico_acta.Last;
    end;
end;

procedure TfrmReportePeriodo.FormShow(Sender: TObject);
var
    QryConfiguracion           : tzReadOnlyquery ;
begin
   sMenuP:=stMenu;
   QryConfiguracion := tzReadOnlyQuery.Create(Self) ;
   QryConfiguracion.Connection := connection.zConnection ;

   BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'opTiemposM');

   QryConfiguracion.Active := False;
   QryConfiguracion.SQL.Clear;
   QryConfiguracion.SQL.Add('select c.*, c2.* '+
                            'From contratos c2 INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
                            'Where c2.sContrato = :Contrato');
   QryConfiguracion.ParamByName('contrato').AsString := global_contrato;
   QryConfiguracion.Open;

   dsConfiguracion.FieldAliases.Clear;
   dsConfiguracion.DataSet  := QryConfiguracion ;
   dsConfiguracion.UserName := 'dsConfiguracion' ;

  If global_grupo = 'INTEL-CODE' Then
      rReporte.PreviewOptions.Buttons := [pbPrint,pbExport,pbZoom,pbFind,pbOutline,pbPageSetup,pbTools,pbEdit,pbExportQuick]
  Else
      rReporte.PreviewOptions.Buttons := [pbPrint,pbExport,pbZoom,pbFind,pbOutline,pbPageSetup,pbTools,pbExportQuick] ;

  OrdenesdeTrabajo.Active := False ;
  OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString ;
  OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := Global_Contrato ;
  //OrdenesdeTrabajo.Params.ParamByName('status').DataType := ftString ;
  //OrdenesdeTrabajo.Params.ParamByName('status').Value :=  connection.configuracion.FieldValues [ 'cStatusProceso' ];
  OrdenesdeTrabajo.Open ;

  Afectaciones.Active := False ;
  Afectaciones.Params.ParamByName('Contrato').DataType := ftString ;
  Afectaciones.Params.ParamByName('Contrato').Value    := Global_Contrato ;
  Afectaciones.Open ;

  tsOrdenesdeTrabajo.Clear ;

  If OrdenesdeTrabajo.RecordCount > 0 Then
  Begin
      tsNumeroOrden.KeyValue := OrdenesdeTrabajo.FieldValues['sNumeroOrden'] ;
      While NOT OrdenesdeTrabajo.Eof Do
      Begin
          tsOrdenesdeTrabajo.Items.Add(OrdenesdeTrabajo.FieldValues['sNumeroOrden'] ) ;
          OrdenesdeTrabajo.Next
      End
  End ;
  tdFechaInicial.Date := Date ;
  tdFechaFinal.Date := Date ;

  


  pgInformes.ActivePageIndex := 0 ;

  if not BotonPermiso.imprimir then
  begin
    btnImprimeDiarios.Enabled := false;
  end;

  ZQRContenido.Open;
  if ZQRContenido.RecordCount>0 then
    CargarContenidoLB(LsContenido);
end;

procedure TfrmReportePeriodo.CargarContenidoLB(Cmp:TCheckListBox);
var
  chkbx:TCheckBox;
  hints:Tcadena;
begin
  Cmp.Items.Clear;
  ZQRContenido.First;
  while not ZQRContenido.Eof  do
  begin
    hints := Tcadena.Create;
    hints.sValor := ZQRContenido.FieldByName('sdescripcion').AsString;
    Cmp.Items.AddObject(ZQRContenido.FieldByName('snombreportada').AsString,hints);
    ZQRContenido.Next;
  end;
end;

procedure TfrmReportePeriodo.CheckBox1Click(Sender: TObject);
var x:Integer;
begin
  for x := 0 to LsContenido.Items.Count-1 do
  begin
    if TCheckBox(sender).Checked then
      LsContenido.Checked[x] := True
    else
      LsContenido.Checked[x] := False;
  end;
end;

procedure TfrmReportePeriodo.LsContenidoMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
  index: integer;
begin
  index := LsContenido.ItemAtPos(point(X, Y), true);
  if index <> -1 then
    LsContenido.Hint := tcadena(LsContenido.Items.Objects[index]).sValor// 'el hint'+inttostr(index)//CheckListBoxHints[index]
  else
    LsContenido.Hint := '';
  if index <> prevIndex then
    Application.CancelHint;
  prevIndex := index;
end;

procedure TfrmReportePeriodo.fotografico_actaAfterInsert(DataSet: TDataSet);
begin
    fotografico_acta.FieldValues['sContrato']        := global_contrato;
    fotografico_acta.FieldValues['sNumeroOrden']     := tsNumeroOrdenActa.Text;
    fotografico_acta.FieldValues['sActaFotografica'] := comboActas.Text;
    fotografico_acta.FieldValues['dIdFecha']         := date;
    fotografico_acta.FieldValues['bImagen']          := connection.configuracion.FieldValues['bImagen'];
    fotografico_acta.FieldValues['lImprime']         := 'Si';
    fotografico_acta.FieldValues['sFasePartida']     := 'Ninguno';
end;

procedure TfrmReportePeriodo.fotografico_actaAfterScroll(DataSet: TDataSet);
begin
    btnPreview2.Click;
end;

procedure TfrmReportePeriodo.FotosPartidasAfterScroll(DataSet: TDataSet);
begin
    btnPreview.Click;
end;

procedure TfrmReportePeriodo.tdFechaInicialChange(Sender: TObject);
begin
  tdFechaFinal.Date:=tdFechainicial.Date;
end;

procedure TfrmReportePeriodo.tdFechaInicialEnter(Sender: TObject);
begin
    tdFechaInicial.Color := global_color_entrada
end;

procedure TfrmReportePeriodo.tdFechaInicialExit(Sender: TObject);
begin
    tdFechaInicial.Color := global_color_salida
end;
function IsDate(ADate: string): Boolean;
var
  Dummy: TDateTime;
begin
  IsDate := TryStrToDate(ADate, Dummy);
end;
procedure TfrmReportePeriodo.tdFechaInicialKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tdFechaFinal.SetFocus 
end;

procedure TfrmReportePeriodo.tdFechaFinalEnter(Sender: TObject);
begin
    tdFechaFinal.Color := global_color_entrada
end;

procedure TfrmReportePeriodo.tdFechaFinalExit(Sender: TObject);
begin
    tdFechaFinal.Color := global_color_salida
end;

procedure TfrmReportePeriodo.tdFechaFinalKeyPress(Sender: TObject;
  var Key: Char);
begin
   If Key = #13 Then
      tsNumeroOrden3.SetFocus  
end;

procedure TfrmReportePeriodo.ReporteGetValue(const VarName: String;
  var Value: Variant);
Var
  Cual: Integer;
begin
  If CompareText(VarName, 'SEMANA') = 0 then
     Value := WeekOfTheMonth(tdFechaFinal.Date) ;

  If CompareText(VarName, 'DIAS_SEMANA') = 0 then
     Value := (tdFechaFinal.Date - tdFechaInicial.Date) + 1 ;

  if chkMoneda.Checked then
  begin
      If CompareText(VarName, 'MONEDA') = 0 then
         Value := 'M.N.' ;
  end
  else
  begin
       If CompareText(VarName, 'MONEDA') = 0 then
         Value := 'DLL';
  end;

  If CompareText(VarName, 'DESCRIPCION') = 0 then
      Value := ordenesdetrabajo.FieldValues['mDescripcion'] ;

  If CompareText(VarName, 'ORDEN') = 0 then
      Value := ordenesdetrabajo.FieldValues['sNumeroOrden'] ;

  If CompareText(VarName, 'MONTO_MODIFICADO') = 0 then
      Value := dMontoModificado ;
  If CompareText(VarName, 'MONTO_CONTRATO') = 0 then
      Value := dMontoContrato ;

  If CompareText(VarName, 'POLIZA') = 0 then
      Value := sPoliza ;
  If CompareText(VarName, 'FIANZA') = 0 then
      Value := sFianza ;

  If CompareText(VarName, 'PUESTO') = 0 then
      Value := StringPuesto.Text ;

  If CompareText(VarName, 'DIRECTORIO') = 0 then
      Value := StringNombre.Text ;

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

  If CompareText(VarName, 'INICIO_ORIGINAL') = 0 then
      Value := sConvenioInicio ;

  If CompareText(VarName, 'TERMINO_ORIGINAL') = 0 then
      Value := sConvenioFinal ;

  If CompareText(VarName, 'ACTA') = 0 then
      Value := sActa ;

  If CompareText(VarName, 'NUEVO_INICIO') = 0 then
      Value := sNuevoInicio ;
  If CompareText(VarName, 'NUEVO_TERMINO') = 0 then
      Value := sNuevoFinal ;

  If CompareText(VarName, 'DURACION') = 0 then
  Begin
      If sNuevoFinal <> '' Then
          Value := StrToDate(sNuevoFinal) - StrToDate(sConvenioInicio) + 1
      Else
          Value := StrToDate(sConvenioFinal) - StrToDate(sConvenioInicio) + 1
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
  If CompareText(VarName, 'ORDEN') = 0 then
      Value := OrdenesdeTrabajo.FieldValues['sDescripcionCorta'] ;

  If CompareText(VarName, 'PERIODO') = 0 then
      Value := 'Del ' + FormatDateTime('dddd d "de" mmmm "del" yyyy' , tdFechaInicial.Date) + ' al ' + FormatDateTime('dddd d "de" mmmm "del" yyyy' , tdFechaFinal.Date);

  If CompareText(VarName, 'TIEMPO_MUERTO') = 0 then
      Value := sHoraResult ;
  If CompareText(VarName, 'REPORTES') = 0 then
      Value := iReportes ;

  If CompareText(VarName, 'ORDENES') = 0 then
      Value := sOrdenes ;

  // Pasar información para saber cuantas columnas se deben mostrar
  If CompareText(Copy(VarName,1,3), 'COL') = 0 then
  begin
    Cual := StrToInt(Copy(VarName,4,2));
    if Cual > 1 + TerminoReal - InicioReal then
      Value := ''
    else
    begin
      Cual := DayOf(InicioReal + Cual - 1);
      Value := RightStr('0' + IntToStr(Cual),2);
    end;
  end;

  If CompareText(VarName, 'MAXCOL') = 0 then
    Value := 1 + (TerminoReal - InicioReal);

  If CompareText(VarName, 'DIAS1') = 0 then
    Value := Dias1;

  If CompareText(VarName, 'DIAS2') = 0 then
    Value := Dias2;

  If CompareText(VarName, 'MES1') = 0 then
    Value := NomMes[MonthOf(InicioReal)];

  If CompareText(VarName, 'MES2') = 0 then
    Value := NomMes[MonthOf(TerminoReal)];

  If CompareText(VarName, 'PROGRAMADO') = 0 then
    Value := '0';
end;


procedure TfrmReportePeriodo.ReportesCheckClick(Sender: TObject);
var
  cdFiltro: TZQuery;
begin

  if (reportesCheck.Checked = true) and (length(trim(tsNumeroOrden3.Text)) > 0)  then
  begin
    try
      cdfiltro := TZQuery.create(nil);
      cdFiltro.Connection := Connection.zConnection;
      cdFiltro.SQL.clear;
      cdfiltro.sql.add('select min(dIdFecha) as inicio, max(didfecha) as final '+
                       'from bitacoradeactividades '+
                       'where sContrato = :contrato '+
                       'and sNumeroOrden = :folio ');
      cdFiltro.ParamByName('contrato').AsString := global_contrato;
      cdfiltro.ParamByName('folio').asString := tsNumeroOrden3.text;
      cdFiltro.Open;
    finally
      tdFechaInicial.Date := cdfiltro.fieldbyname('inicio').asDatetime;
      tdFechaFinal.date := cdfiltro.fieldbyname('final').asDatetime;
      cdfiltro.Free;
    end;
  end;
  if rbCronologias.Checked then
  begin
     try
      cdfiltro := TZQuery.create(nil);
      cdFiltro.Connection := Connection.zConnection;
      cdFiltro.SQL.clear;
      cdfiltro.sql.add('select min(dIdFecha) as inicio, max(didfecha) as final '+
                       'from bitacoradeactividades '+
                       'where sContrato = :contrato ');
      cdFiltro.ParamByName('contrato').AsString := global_contrato;
      cdFiltro.Open;
    finally
      tdFechaInicial.Date := cdfiltro.fieldbyname('inicio').asDatetime;
      tdFechaFinal.date := cdfiltro.fieldbyname('final').asDatetime;
      cdfiltro.Destroy;
    end;
  
  end;
end;

procedure TfrmReportePeriodo.frxFotograficoGetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'PERIODO') = 0 then
      Value := 'Del ' + FormatDateTime('dddd d "de" mmmm "del" yyyy' , tdFechaInicial.Date) + ' al ' + FormatDateTime('dddd d "de" mmmm "del" yyyy' , tdFechaFinal.Date);


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

procedure TfrmReportePeriodo.OnClick3D(Sender: TObject);
begin
  showmessage('rangel');
  {if CompareText(Sender.ClassName, 'TButton') = 0 then
    TButton(Sender).Checked := Not TButton(Sender).Checked;}
end;

procedure TfrmReportePeriodo.pgInformesChange(Sender: TObject);
begin
    if pgInformes.ActivePageIndex = 1 then
    begin
        fotografico_Acta.Active := False;
        fotografico_acta.ParamByName('Contrato').AsString := global_contrato;
        fotografico_acta.ParamByName('Orden').AsString    := tsNumeroOrdenActa.Text;
        fotografico_acta.ParamByName('Acta').AsString     := ComboActas.Text;
        fotografico_acta.Open;

        //Ahora el llenado del combobox de las actas fotograficas,
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('select sActaFotografica from reportefotografico_acta where sContrato =:Contrato and sNumeroOrden =:Orden group by sActaFotografica ');
        connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
        connection.QryBusca.ParamByName('Orden').AsString    := tsNumeroOrdenActa.Text;
        connection.QryBusca.Open;

        ComboActas.Items.Clear;
        while not connection.QryBusca.Eof do
        begin
            ComboActas.Items.Add(connection.QryBusca.FieldValues['sActaFotografica']);
            connection.QryBusca.Next;
        end;
    end;
end;

procedure TfrmReportePeriodo.btnPersonalProgramadoClick(Sender: TObject);
Var
    QryPersonal : TzReadOnlyQuery ;
    QryEquipos  : TzReadOnlyQuery ;
    dsPersonal  : TfrxDBDataSet ;
    dsEquipos  : TfrxDBDataSet ;
    dsQryGraficaTiemposMuertos : TfrxDBDataSet;
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdFechaFinal.Date<tdFechaInicial.Date then
   begin
   showmessage('la fecha final de impresión es menor a la fecha inicial de impresión');
   tdFechaFinal.SetFocus;
   exit;
   end;

  if not BotonPermiso.imprimir then
  begin
    showmessage('No tiene permisos de impresión');
    exit;
  end;

  try
    dsQryGraficaTiemposMuertos := TfrxDBDataSet.Create(Self) ;
    dsQryGraficaTiemposMuertos.FieldAliases.Clear ;
    dsQryGraficaTiemposMuertos.UserName := 'dsQryGraficaTiemposMuertos' ;

    QryPersonal := tzReadOnlyQuery.Create(Self) ;
    QryPersonal.Connection := connection.zConnection ;
    QryEquipos := tzReadOnlyQuery.Create(Self) ;
    QryEquipos.Connection := connection.zConnection ;

    dsPersonal := TfrxDBDataSet.Create(Self) ;
    dsPersonal.DataSet := QryPersonal ;
    dsPersonal.UserName := 'dsPersonal' ;

    dsEquipos := TfrxDBDataSet.Create(Self) ;
    dsEquipos.DataSet := QryEquipos ;
    dsEquipos.UserName := 'dsEquipos' ;


    If chkDetalle.Checked Then
    Begin
        QryPersonal.Active := False ;
        QryPersonal.SQL.Clear ;

        QryEquipos.Active := False ;
        QryEquipos.SQL.Clear ;

        QryGraficaTiemposMuertos.Active := False ;
        QryGraficaTiemposMuertos.SQL.Clear ;

        If MessageDlg('Desea agrupar la informacion por contrato?' , mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
            QryPersonal.SQL.Add('Select a.sContrato, a.iItemOrden, a.sIdPersonal, a.sDescripcion, a.sMedida, a.dVentaMN, a.dCostoMN, bp.dIdFecha, Sum(bp.dCantidad) as dCantidad from personal a ' +
                                'inner join bitacoradepersonal bp on (a.sContrato = bp.sContrato And a.sIdPersonal = bp.sIdPersonal and bp.dIdFecha >= :FechaInicio And bp.dIdFecha <= :FechaFinal) ' +
                                'inner join bitacoradeactividades b ON (b.sContrato = bp.sContrato and b.dIdFecha = bp.dIdFecha And b.iIdDiario = bp.iIdDiario) ' +
                                'inner join ordenesdetrabajo o ON (b.sContrato = o.sContrato and b.sNumeroOrden = o.sNumeroOrden And o.cIdStatus = :status) ' +
                                'INNER JOIN turnos t ON (b.sContrato = t.sContrato and b.sIdTurno = t.sIdTurno and t.sOrigenTierra = "No") ' +
                                'Where a.sContrato = :contrato Group By bp.sContrato, bp.sIdPersonal, bp.dIdFecha ' +
                                'Order By a.iItemOrden, bp.sIdPersonal, bp.dIdFecha') ;
            QryEquipos.SQL.Add('Select a.sContrato, a.iItemOrden, a.sIdEquipo, a.sDescripcion, a.sMedida, a.dVentaMN, a.dCostoMN, bp.dIdFecha, Sum(bp.dCantidad) as dCantidad from equipos a ' +
                               'inner join bitacoradeequipos bp on (a.sContrato = bp.sContrato And a.sIdEquipo = bp.sIdEquipo and bp.dIdFecha >= :FechaInicio And bp.dIdFecha <= :FechaFinal) ' +
                               'inner join bitacoradeactividades b ON (b.sContrato = bp.sContrato and b.dIdFecha = bp.dIdFecha And b.iIdDiario = bp.iIdDiario) ' +
                               'inner join ordenesdetrabajo o ON (b.sContrato = o.sContrato and b.sNumeroOrden = o.sNumeroOrden And o.cIdStatus = :status) ' +
                               'INNER JOIN turnos t ON (b.sContrato = t.sContrato and b.sIdTurno = t.sIdTurno and t.sOrigenTierra = "No") ' +
                               'Where a.sContrato = :contrato Group By bp.sContrato, bp.sIdEquipo, bp.dIdFecha ' +
                               'Order By a.iItemOrden, bp.sIdEquipo, bp.dIdFecha') ;
            QryGraficaTiemposMuertos.SQL.Add('Select r.sContrato, r.dIdFecha, ' +
                                             '(sum(substr(sTiempoMuertoReal, 1 , 2)) + sum(substr(sTiempoMuertoReal, 4 , 2)) div 60 + ' +
                                             '(sum(substr(sTiempoMuertoReal, 4 , 2)) % 60 ) / 100 ) as iTiempoMuertoReal from reportediario r ' +
                                             'INNER JOIN ordenesdetrabajo o ON (r.sContrato = o.sContrato and r.sNumeroOrden = o.sNumeroOrden And o.cIdStatus = :status) ' +
                                             'INNER JOIN turnos t ON (r.sContrato = t.sContrato and r.sIdTurno = t.sIdTurno and t.sOrigenTierra = "No") ' +
                                             'Where r.sContrato = :Contrato And r.dIdFecha >= :FechaI And r.dIdFecha <= :FechaF And r.lStatus = "Autorizado" ' +
                                             'Group By r.sNumeroOrden, r.dIdfecha Order By r.sContrato, r.dIdFecha')
        End
        Else
        begin
            QryPersonal.SQL.Add('Select b.sNumeroOrden as sContrato, a.iItemOrden, a.sIdPersonal, a.sDescripcion, a.sMedida, a.dVentaMN, a.dCostoMN, bp.dIdFecha, Sum(bp.dCantidad) as dCantidad from personal a ' +
                                'inner join bitacoradepersonal bp on (a.sContrato = bp.sContrato And a.sIdPersonal = bp.sIdPersonal and bp.dIdFecha >= :FechaInicio And bp.dIdFecha <= :FechaFinal) ' +
                                'inner join bitacoradeactividades b ON (b.sContrato = bp.sContrato and b.dIdFecha = bp.dIdFecha And b.iIdDiario = bp.iIdDiario) ' +
                                'inner join ordenesdetrabajo o ON (b.sContrato = o.sContrato and b.sNumeroOrden = o.sNumeroOrden And o.cIdStatus = :status) ' +
                                'inner join turnos t ON (b.sContrato = t.sContrato and b.sIdTurno = t.sIdTurno and t.sOrigenTierra = "No") ' +
                                'Where a.sContrato = :contrato Group By b.sNumeroOrden, bp.sIdPersonal, bp.dIdFecha ' +
                                'Order By a.iItemOrden, bp.sIdPersonal, bp.dIdFecha') ;
            QryEquipos.SQL.Add('Select b.sNumeroOrden as sContrato, a.iItemOrden, a.sIdEquipo, a.sDescripcion, a.sMedida, a.dVentaMN, a.dCostoMN, bp.dIdFecha, Sum(bp.dCantidad) as dCantidad from equipos a ' +
                               'inner join bitacoradeequipos bp on (a.sContrato = bp.sContrato And a.sIdEquipo = bp.sIdEquipo and bp.dIdFecha >= :FechaInicio And bp.dIdFecha <= :FechaFinal) ' +
                               'inner join bitacoradeactividades b ON (b.sContrato = bp.sContrato and b.dIdFecha = bp.dIdFecha And b.iIdDiario = bp.iIdDiario) ' +
                               'inner join ordenesdetrabajo o ON (b.sContrato = o.sContrato and b.sNumeroOrden = o.sNumeroOrden And o.cIdStatus = :status) ' +
                               'inner join turnos t ON (b.sContrato = t.sContrato and b.sIdTurno = t.sIdTurno and t.sOrigenTierra = "No") ' +
                               'Where a.sContrato = :contrato Group By b.sNumeroOrden, bp.sIdEquipo, bp.dIdFecha ' +
                               'Order By a.iItemOrden, bp.sIdEquipo, bp.dIdFecha') ;
            QryGraficaTiemposMuertos.SQL.Add('Select r.sNumeroOrden as sContrato, r.dIdFecha, ' +
                                             '(sum(substr(sTiempoMuertoReal, 1 , 2)) + sum(substr(sTiempoMuertoReal, 4 , 2)) div 60 + ' +
                                             '(sum(substr(sTiempoMuertoReal, 4 , 2)) % 60 ) / 100 ) as iTiempoMuertoReal from reportediario r ' +
                                             'INNER JOIN ordenesdetrabajo o ON (r.sContrato = o.sContrato and r.sNumeroOrden = o.sNumeroOrden And o.cIdStatus = :status) ' +
                                             'INNER JOIN turnos t ON (r.sContrato = t.sContrato and r.sIdTurno = t.sIdTurno and t.sOrigenTierra = "No") ' +
                                             'Where r.sContrato = :Contrato And r.dIdFecha >= :FechaI And r.dIdFecha <= :FechaF And r.lStatus = "Autorizado" ' +
                                             'Group By r.sNumeroOrden, r.dIdfecha Order By r.sNumeroOrden, r.dIdFecha') ;
        end ;

        dsPersonal.FieldAliases.Clear ;
        dsPersonal.DataSet := QryPersonal ;
        dsEquipos.FieldAliases.Clear ;
        dsEquipos.DataSet := QryEquipos ;

        QryPersonal.Params.ParamByName('Contrato').DataType := ftString ;
        QryPersonal.Params.ParamByName('Contrato').Value := global_Contrato ;
        QryPersonal.Params.ParamByName('FechaInicio').DataType := ftDate ;
        QryPersonal.Params.ParamByName('FechaInicio').Value := tdFechaInicial.Date ;
        QryPersonal.Params.ParamByName('FechaFinal').DataType := ftDate ;
        QryPersonal.Params.ParamByName('FechaFinal').Value := tdFechaFinal.Date ;
        QryPersonal.Params.ParamByName('status').DataType := ftString ;
        QryPersonal.Params.ParamByName('status').Value :=  connection.configuracion.FieldValues [ 'cStatusProceso' ];
        QryPersonal.Open ;

        QryEquipos.Active := False ;
        QryEquipos.Params.ParamByName('Contrato').DataType := ftString ;
        QryEquipos.Params.ParamByName('Contrato').Value := global_Contrato ;
        QryEquipos.Params.ParamByName('FechaInicio').DataType := ftDate ;
        QryEquipos.Params.ParamByName('FechaInicio').Value := tdFechaInicial.Date ;
        QryEquipos.Params.ParamByName('FechaFinal').DataType := ftDate ;
        QryEquipos.Params.ParamByName('FechaFinal').Value := tdFechaFinal.Date ;
        QryEquipos.Params.ParamByName('status').DataType := ftString ;
        QryEquipos.Params.ParamByName('status').Value :=  connection.configuracion.FieldValues [ 'cStatusProceso' ];
        QryEquipos.Open ;

        QryGraficaTiemposMuertos.Params.ParamByName('Contrato').DataType := ftString ;
        QryGraficaTiemposMuertos.Params.ParamByName('Contrato').Value := global_contrato ;
        QryGraficaTiemposMuertos.Params.ParamByName('FechaI').DataType := ftDate ;
        QryGraficaTiemposMuertos.Params.ParamByName('FechaI').Value := tdFechaInicial.Date ;
        QryGraficaTiemposMuertos.Params.ParamByName('FechaF').DataType := ftDate ;
        QryGraficaTiemposMuertos.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
        QryGraficaTiemposMuertos.Params.ParamByName('status').DataType := ftString ;
        QryGraficaTiemposMuertos.Params.ParamByName('status').Value := connection.configuracion.FieldValues [ 'cStatusProceso' ] ;
        QryGraficaTiemposMuertos.Open ;

        if (QryPernocta.RecordCount > 0) or
           (rxPlataforma.RecordCount > 0) or
           (rxPernocta.RecordCount > 0) or
           (QryPlataforma.RecordCount > 0) or
           (rxReporteFotografico.RecordCount > 0) then
        begin
            rReporte.PreviewOptions.MDIChild := False ;
            rReporte.PreviewOptions.Modal := True ;
            rReporte.PreviewOptions.Maximized := lCheckMaximized () ;
            rReporte.PreviewOptions.ShowCaptions := False ;
            rReporte.Previewoptions.ZoomMode := zmPageWidth ;
            rReporte.LoadFromFile (global_files + 'DetalledeRecursosPersonalyEquipo.fr3') ;
            rReporte.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP)) ;
            if not FileExists(global_files + 'DetalledeRecursosPersonalyEquipo.fr3') then
               showmessage('El archivo de reporte DetalledeRecursosPersonalyEquipo.fr3 no existe, notifique al administrador del sistema');
        end
        else begin
            showmessage('No hay datos para imprimir');
        end;
    End
    Else
    Begin
        sPernocta := '' ;
        sPlataforma := '' ;

        dsPersonalPernocta.OnFirst := ActualizaPernocta ;
        dsPersonalPlataforma.OnFirst := ActualizaPlataforma ;
        dsPersonalPernocta.OnNext := ActualizaPernocta ;
        dsPersonalPlataforma.OnNext := ActualizaPlataforma ;

        // Personal Programado
        Connection.QryBusca.Active := False ;
        Connection.QryBusca.SQL.Clear ;
        Connection.QryBusca.SQL.Add('Select AVG(dCantidad) as dProgramado From personalprogramado Where ' +
                                    'sContrato = :Contrato And dIdFecha >= :FechaI and dIdFecha <= :FechaF ' +
                                    'Group By sContrato') ;
        Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.QryBusca.Params.ParamByName('FechaI').DataType := ftDate ;
        Connection.QryBusca.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
        Connection.QryBusca.Params.ParamByName('FechaF').DataType := ftDate ;
        Connection.QryBusca.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
        Connection.QryBusca.Open ;
        dProgramado := 0 ;
        If Connection.QryBusca.RecordCount > 0 Then
             dProgramado := Connection.QryBusca.FieldValues['dProgramado'] ;

        // Personal promedio real ...
        Connection.QryBusca.Active := False ;
        Connection.QryBusca.SQL.Clear ;
        Connection.QryBusca.SQL.Add('Select Sum(b.dCantidad) as dReal ' +
                                    'From bitacoradepersonal b INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                    'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                    'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sContrato') ;
        Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.QryBusca.Params.ParamByName('FechaI').DataType := ftDate ;
        Connection.QryBusca.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
        Connection.QryBusca.Params.ParamByName('FechaF').DataType := ftDate ;
        Connection.QryBusca.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
        Connection.QryBusca.Open ;
        dPromedio := 0 ;
        If Connection.QryBusca.RecordCount > 0 Then
        Begin
            dPromedio := Connection.QryBusca.FieldValues['dReal'] ;
            Connection.QryBusca.Active := False ;
            Connection.QryBusca.SQL.Clear ;
            Connection.QryBusca.SQL.Add('Select distinct b.dIdFecha ' +
                                       'From bitacoradepersonal b INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                       'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                       'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF') ;
            Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
            Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
            Connection.QryBusca.Params.ParamByName('FechaI').DataType := ftDate ;
            Connection.QryBusca.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
            Connection.QryBusca.Params.ParamByName('FechaF').DataType := ftDate ;
            Connection.QryBusca.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
            Connection.QryBusca.Open ;
            If Connection.QryBusca.RecordCount > 0 Then
                 dPromedio := RoundTo(dPromedio / Connection.QryBusca.RecordCount,0) ;
        End ;

        // Qrys por Pernocta ....
        If chkPernocta.Checked Then
        Begin
            QryPersonalPernocta.Active := False ;
            QryPersonalPernocta.SQL.Clear ;
            QryPersonalPernocta.SQL.Add('Select b.sIdPernocta, b.dIdFecha, Sum(b.dCantidad) as dReal From bitacoradepersonal b ' +
                                        'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                        'INNER JOIN reportediario r ON (b2.sContrato = r.sContrato And b2.dIdFecha = r.dIdFecha And b2.sNumeroOrden = r.sNumeroOrden) ' +
                                        'INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                        'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sIdPernocta, b.dIdFecha ' +
                                        'Order By b.dIdFecha, b.sIdPernocta ') ;
            QryPersonalPernocta.Params.ParamByName('Contrato').DataType := ftString ;
            QryPersonalPernocta.Params.ParamByName('Contrato').Value := global_contrato ;
            QryPersonalPernocta.Params.ParamByName('FechaI').DataType := ftDate ;
            QryPersonalPernocta.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
            QryPersonalPernocta.Params.ParamByName('FechaF').DataType := ftDate ;
            QryPersonalPernocta.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
            dsPersonalPernocta.FieldAliases.Clear ;
            dsPersonalPernocta.DataSet := QryPersonalPernocta ;
            QryPersonalPernocta.Open ;

            QryPernocta.Active := False ;
            QryPernocta.SQL.Clear ;
            QryPernocta.SQL.Add('Select b.sIdPernocta, Sum(b.dCantidad) as dReal From bitacoradepersonal b ' +
                                'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                'INNER JOIN reportediario r ON (b2.sContrato = r.sContrato And b2.dIdFecha = r.dIdFecha And b2.sNumeroOrden = r.sNumeroOrden) ' +
                                'INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sContrato, b.sIdPernocta ') ;
            QryPernocta.Params.ParamByName('Contrato').DataType := ftString ;
            QryPernocta.Params.ParamByName('Contrato').Value := global_contrato ;
            QryPernocta.Params.ParamByName('FechaI').DataType := ftDate ;
            QryPernocta.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
            QryPernocta.Params.ParamByName('FechaF').DataType := ftDate ;
            QryPernocta.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
            dsPernocta.FieldAliases.Clear ;
            dsPernocta.DataSet := QryPernocta ;
            QryPernocta.Open ;
       End
       Else
       Begin
            QryPersonalPernocta.Active := False ;
            QryPersonalPernocta.SQL.Clear ;
            QryPersonalPernocta.SQL.Add('Select b.sIdPernocta, b.dIdFecha, Sum(b.dCantidad) as dReal From bitacoradepersonal b ' +
                                        'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                        'INNER JOIN reportediario r ON (b2.sContrato = r.sContrato And b2.dIdFecha = r.dIdFecha And b2.sNumeroOrden = r.sNumeroOrden) ' +
                                        'INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                        'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sIdPernocta, b.dIdFecha ' +
                                        'Order By b.dIdFecha, b.sIdPernocta ') ;
            QryPersonalPernocta.Params.ParamByName('Contrato').DataType := ftString ;
            QryPersonalPernocta.Params.ParamByName('Contrato').Value := '' ;
            QryPersonalPernocta.Params.ParamByName('FechaI').DataType := ftDate ;
            QryPersonalPernocta.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
            QryPersonalPernocta.Params.ParamByName('FechaF').DataType := ftDate ;
            QryPersonalPernocta.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
            dsPersonalPernocta.FieldAliases.Clear ;
            dsPersonalPernocta.DataSet := QryPersonalPernocta ;
            QryPersonalPernocta.Open ;

            QryPernocta.Active := False ;
            QryPernocta.SQL.Clear ;
            QryPernocta.SQL.Add('Select b.sIdPernocta, Sum(b.dCantidad) as dReal From bitacoradepersonal b ' +
                                'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                'INNER JOIN reportediario r ON (b2.sContrato = r.sContrato And b2.dIdFecha = r.dIdFecha And b2.sNumeroOrden = r.sNumeroOrden) ' +
                                'INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sContrato, b.sIdPernocta ') ;
            QryPernocta.Params.ParamByName('Contrato').DataType := ftString ;
            QryPernocta.Params.ParamByName('Contrato').Value := '' ;
            QryPernocta.Params.ParamByName('FechaI').DataType := ftDate ;
            QryPernocta.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
            QryPernocta.Params.ParamByName('FechaF').DataType := ftDate ;
            QryPernocta.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
            dsPernocta.FieldAliases.Clear ;
            dsPernocta.DataSet := QryPernocta ;
            QryPernocta.Open ;
       End ;

       If chkPlataforma.Checked Then
       Begin
            // Por plataforma
            QryPersonalPlataforma.Active := False ;
            QryPersonalPlataforma.SQL.Clear ;
            QryPersonalPlataforma.SQL.Add('Select b.sIdPlataforma, b.dIdFecha, Sum(b.dCantidad) as dReal From bitacoradepersonal b ' +
                                          'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                          'INNER JOIN reportediario r ON (b2.sContrato = r.sContrato And b2.dIdFecha = r.dIdFecha And b2.sNumeroOrden = r.sNumeroOrden) ' +
                                          'INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                          'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sIdPlataforma, b.dIdFecha ' +
                                          'Order By b.dIdFecha, b.sIdPlataforma') ;
            QryPersonalPlataforma.Params.ParamByName('Contrato').DataType := ftString ;
            QryPersonalPlataforma.Params.ParamByName('Contrato').Value := global_contrato ;
            QryPersonalPlataforma.Params.ParamByName('FechaI').DataType := ftDate ;
            QryPersonalPlataforma.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
            QryPersonalPlataforma.Params.ParamByName('FechaF').DataType := ftDate ;
            QryPersonalPlataforma.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
            dsPersonalPlataforma.FieldAliases.Clear ;
            dsPersonalPlataforma.DataSet := QryPersonalPlataforma ;
            QryPersonalPlataforma.Open ;

            // Grafica de Pernocta y Plataforma

            QryPlataforma.Active := False ;
            QryPlataforma.SQL.Clear ;
            QryPlataforma.SQL.Add('Select b.sIdPlataforma, Sum(b.dCantidad) as dReal From bitacoradepersonal b ' +
                                  'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                  'INNER JOIN reportediario r ON (b2.sContrato = r.sContrato And b2.dIdFecha = r.dIdFecha And b2.sNumeroOrden = r.sNumeroOrden) ' +
                                  'INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                  'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sContrato, b.sIdPlataforma ') ;
            QryPlataforma.Params.ParamByName('Contrato').DataType := ftString ;
            QryPlataforma.Params.ParamByName('Contrato').Value := global_contrato ;
            QryPlataforma.Params.ParamByName('FechaI').DataType := ftDate ;
            QryPlataforma.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
            QryPlataforma.Params.ParamByName('FechaF').DataType := ftDate ;
            QryPlataforma.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
            dsPlataforma.FieldAliases.Clear ;
            dsPlataforma.DataSet := QryPlataforma ;
            QryPlataforma.Open ;
       End
       Else
       Begin
            // Por plataforma
            QryPersonalPlataforma.Active := False ;
            QryPersonalPlataforma.SQL.Clear ;
            QryPersonalPlataforma.SQL.Add('Select b.sIdPlataforma, b.dIdFecha, Sum(b.dCantidad) as dReal From bitacoradepersonal b ' +
                                          'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                          'INNER JOIN reportediario r ON (b2.sContrato = r.sContrato And b2.dIdFecha = r.dIdFecha And b2.sNumeroOrden = r.sNumeroOrden) ' +
                                          'INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                          'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sIdPlataforma, b.dIdFecha ' +
                                          'Order By b.dIdFecha, b.sIdPlataforma') ;
            QryPersonalPlataforma.Params.ParamByName('Contrato').DataType := ftString ;
            QryPersonalPlataforma.Params.ParamByName('Contrato').Value := '' ;
            QryPersonalPlataforma.Params.ParamByName('FechaI').DataType := ftDate ;
            QryPersonalPlataforma.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
            QryPersonalPlataforma.Params.ParamByName('FechaF').DataType := ftDate ;
            QryPersonalPlataforma.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
            dsPersonalPlataforma.FieldAliases.Clear ;
            dsPersonalPlataforma.DataSet := QryPersonalPlataforma ;
            QryPersonalPlataforma.Open ;

            // Grafica de Pernocta y Plataforma

            QryPlataforma.Active := False ;
            QryPlataforma.SQL.Clear ;
            QryPlataforma.SQL.Add('Select b.sIdPlataforma, Sum(b.dCantidad) as dReal From bitacoradepersonal b ' +
                                  'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                  'INNER JOIN reportediario r ON (b2.sContrato = r.sContrato And b2.dIdFecha = r.dIdFecha And b2.sNumeroOrden = r.sNumeroOrden) ' +
                                  'INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                  'Where b.sContrato = :Contrato And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sContrato, b.sIdPlataforma ') ;
            QryPlataforma.Params.ParamByName('Contrato').DataType := ftString ;
            QryPlataforma.Params.ParamByName('Contrato').Value := '' ;
            QryPlataforma.Params.ParamByName('FechaI').DataType := ftDate ;
            QryPlataforma.Params.ParamByName('FechaI').Value := tdFechaInicial.Date  ;
            QryPlataforma.Params.ParamByName('FechaF').DataType := ftDate ;
            QryPlataforma.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
            dsPlataforma.FieldAliases.Clear ;
            dsPlataforma.DataSet := QryPlataforma ;
            QryPlataforma.Open ;
       End ;



           rReporte.PreviewOptions.MDIChild := False ;
           rReporte.PreviewOptions.Modal := True ;
           rReporte.PreviewOptions.Maximized := lCheckMaximized () ;
           rReporte.PreviewOptions.ShowCaptions := False ;
           rReporte.Previewoptions.ZoomMode := zmPageWidth ;
           rReporte.LoadFromFile (global_files + 'ResumendeCostosporInstalacion.fr3') ;
           rDiarioFirmas (global_contrato, '', 'A',tdFechaFinal.Date , frmReportePeriodo ) ;
           rReporte.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
           if not FileExists(global_files + 'ResumendeCostosporInstalacion.fr3') then
              showmessage('El archivo de reporte ResumendeCostosporInstalacion.fr3 no existe, notifique al administrador del sistema');

    End ;
    QryPersonal.Destroy ;
    QryEquipos.Destroy ;
    dsPersonal.Destroy ;
    dsEquipos.Destroy ;
    dsQryGraficaTiemposMuertos.Destroy ;
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Album Fotografico/Tiempos Muertos ...', 'Al imprimir Distribucion de Personal', 0);
    end;
  end;
end;

procedure TfrmReportePeriodo.PowerPointAlbum(roqDatos : TRxMemoryData; fileName : string);
var
  PowerPointApp: OLEVariant;
  Cuenta, Hojas: Integer;
  xLeft, xTop: Real;
  bS: TStream;
  Pic: TJpegImage;
  BlobField: tField;
  OldFecha: TDate;

  function Depura(Valor: String): String;
  Const
    NoIncluir = '\/:*?<>|"';
  var
    Resultado: String;
  begin
    Resultado := '';

    while Length(Resultado) < Length(Valor) do
      if Pos(Valor[Length(Resultado) + 1], NoIncluir) > 0 then
        Resultado := Resultado + '-'
      else
        Resultado := Resultado + Valor[Length(Resultado) + 1];

    Result := Resultado;
  end;

begin
  roqDatos.First;
  try
    PowerPointApp := CreateOleObject('PowerPoint.Application')
  except
    Exit;
  end;

  try
    ProgressBar1.Visible := True;
    try
      PowerPointApp.Visible := True;
      PowerPointApp.Presentations.Add(-1);
      PowerPointApp.WindowState := 2;

      Cuenta := 0;
      Hojas := 1;
      OldFecha := Trunc(roqDatos.FieldByName('dIdFecha').AsDateTime);
      ProgressBar1.Position := 0;
      ProgressBar1.Max := roqDatos.RecordCount;
      while Not roqDatos.Eof do
      begin
        ProgressBar1.Position := ProgressBar1.Position + 1;
        Self.Repaint;
        //ProgressBar1.Repaint;
        if Cuenta = 0 then
        begin
          PowerPointApp.ActiveWindow.View.GotoSlide(PowerPointApp.ActivePresentation.Slides.Add(Hojas, 1).SlideIndex);
          PowerPointApp.ActiveWindow.Selection.SlideRange.Layout := 12;     // Limpiar la hoja (Layout - 12 es hoja en blanco)

          // Poner los encabezados
          PowerPointApp.ActiveWindow.Selection.SlideRange.Shapes.AddTextbox(1, 0, 0, 314.625, 28.875).Select;
          PowerPointApp.ActiveWindow.Selection.ShapeRange.TextFrame.WordWrap := -1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.LineRuleWithin := -1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.SpaceWithin := 1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.LineRuleBefore := -1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.SpaceBefore := 0.5;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.LineRuleAfter := -1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.SpaceAfter := 0;
          PowerPointApp.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Characters(1, 0).Select;
          PowerPointApp.ActiveWindow.Selection.TextRange.Text := roqDatos.FieldByName('sContrato').AsString;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Name := 'Arial';
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Size := 10;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Bold := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Italic := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Underline := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Shadow := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Emboss := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.BaselineOffset := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.AutoRotateNumbers := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Color.SchemeColor := 2;

          PowerPointApp.ActiveWindow.Selection.SlideRange.Shapes.AddTextbox(1, 592.5, 0, 127.5, 28.875).Select;
          PowerPointApp.ActiveWindow.Selection.ShapeRange.TextFrame.WordWrap := -1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.LineRuleWithin := -1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.SpaceWithin := 1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.LineRuleBefore := -1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.SpaceBefore := 0.5;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.LineRuleAfter := -1;
          PowerPointApp.ActiveWindow.Selection.TextRange.ParagraphFormat.SpaceAfter := 0;
          PowerPointApp.ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Characters(1, 0).Select;
          PowerPointApp.ActiveWindow.Selection.TextRange.Text := 'FECHA: ' + roqDatos.FieldByName('dIdFecha').AsString;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Name := 'Arial';
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Size := 10;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Bold := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Italic := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Underline := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Shadow := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Emboss := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.BaselineOffset := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.AutoRotateNumbers := 0;
          PowerPointApp.ActiveWindow.Selection.TextRange.Font.Color.SchemeColor := 2;
        end;

        // Pasar la imagen a disco
        BlobField := roqDatos.FieldByName('bImagen');
        BS := roqDatos.CreateBlobStream ( BlobField , bmRead ) ;
        If bs.Size > 1 Then
        Begin
          try
            Pic:=TJpegImage.Create;
            try
              Pic.LoadFromStream(bS);
              bImagen.Picture.Graphic:=Pic;
              bImagen.Picture.SaveToFile('C:\MiImagen.jpg');
            finally
              Pic.Free;
            end;
          finally
            bS.Free
          End
        End
        Else
          bImagen.Picture := Nil;

        case Cuenta of
          0: begin xLeft := -175.12; xTop := -128.62; end;
          1: begin xLeft := 175.12; xTop := -128.62; end;
          2: begin xLeft := -175.12; xTop := 128.62; end;
          3: begin xLeft := 175.12; xTop := 128.62; end;
        end;

        PowerPointApp.ActiveWindow.Selection.SlideRange.Shapes.AddPicture('C:\MiImagen.jpg', 0, -1, 201, 155, 319, 230).Select;
        PowerPointApp.ActiveWindow.Selection.ShapeRange.IncrementLeft(xLeft);
        PowerPointApp.ActiveWindow.Selection.ShapeRange.IncrementTop(xTop);
        PowerPointApp.ActiveWindow.Selection.ShapeRange.ScaleWidth(0.89, 0, 0);
        PowerPointApp.ActiveWindow.Selection.ShapeRange.ScaleHeight(0.89, 0, 0);

        Inc(Cuenta);
        roqDatos.Next;

        if (Cuenta > 3) or (OldFecha <> Trunc(roqDatos.FieldByName('dIdFecha').AsDateTime)) then
        begin
          Inc(Hojas);
          Cuenta := 0;
          OldFecha := Trunc(roqDatos.FieldByName('dIdFecha').AsDateTime);
        end;
      end;
      FileName := FileName + ' Fotografías incluidas en reportes diarios del ' + DateToStr(tdFechaInicial.Date) + ' al ' + DateToStr(tdFechaFinal.Date) + '.ppt';
      FileName := Depura(FileName);
      PowerPointApp.ActivePresentation.SaveAs(FileName);
      MessageDlg('El libro de fotografías ha sido generado con el nombre: ' + #10 + FileName, mtInformation, [mbOk], 0);
      PowerPointApp.WindowState := 3;
    except
      on e:exception do
      begin
        PowerPointApp.Quit;
        messagedlg('Ha ocurrido un error al intentar generar el libro de fotografías solicitado.' + #10 + #10 +
                   'Informe de esto al administrador del sistema:' + #10 + e.Message, mtWarning, [mbOk], 0);
      end;
    end;
  finally
    ProgressBar1.Visible := False;
    PowerPointApp := Null;
    SysUtils.DeleteFile('C:\MiImagen.jpg');     // Borrar el archivo que pudo haber sido generado
  end;
end;

procedure TfrmReportePeriodo.frxGerencialGetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'MONTO_MODIFICADO') = 0 then
      Value := dMontoModificado ;
  If CompareText(VarName, 'MONTO_CONTRATO') = 0 then
      Value := dMontoContrato ;

  If CompareText(VarName, 'POLIZA') = 0 then
      Value := sPoliza ;
  If CompareText(VarName, 'FIANZA') = 0 then
      Value := sFianza ;

  If CompareText(VarName, 'PUESTO') = 0 then
      Value := StringPuesto.Text ;

  If CompareText(VarName, 'DIRECTORIO') = 0 then
      Value := StringNombre.Text ;

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

  If CompareText(VarName, 'INICIO_ORIGINAL') = 0 then
      Value := sConvenioInicio ;

  If CompareText(VarName, 'TERMINO_ORIGINAL') = 0 then
      Value := sConvenioFinal ;

  If CompareText(VarName, 'ACTA') = 0 then
      Value := sActa ;

  If CompareText(VarName, 'NUEVO_INICIO') = 0 then
      Value := sNuevoInicio ;
  If CompareText(VarName, 'NUEVO_TERMINO') = 0 then
      Value := sNuevoFinal ;

  If CompareText(VarName, 'DURACION') = 0 then
  Begin
      If sNuevoFinal <> '' Then
          Value := StrToDate(sNuevoFinal) - StrToDate(sConvenioInicio) + 1
      Else
          Value := StrToDate(sConvenioFinal) - StrToDate(sConvenioInicio) + 1
  End

end;

procedure TfrmReportePeriodo.ActualizaPlataforma(Sender: TObject);
begin
IF QryPersonalPlataforma.Active THEN BEGIN
  If QryPersonalPlataforma.RecordCount > 0 Then
  Begin
      If QryPersonalPlataforma.FieldValues['sIdPlataforma'] <> sPlataforma Then
      Begin
           Connection.QryBusca.Active := False ;
           Connection.QryBusca.SQL.Clear ;
           Connection.QryBusca.SQL.Add('Select Sum(b.dCantidad) as dReal ' +
                                        'From bitacoradepersonal b INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                        'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                        'Where b.sContrato = :Contrato And b.sIdPlataforma = :Plataforma And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sContrato, b.sIdPlataforma') ;
           Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
           Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
           Connection.QryBusca.Params.ParamByName('Plataforma').DataType := ftString ;
           Connection.QryBusca.Params.ParamByName('Plataforma').Value := QryPersonalPlataforma.FieldValues['sIdPlataforma'] ;
           Connection.QryBusca.Params.ParamByName('FechaI').DataType := ftDate ;
           Connection.QryBusca.Params.ParamByName('FechaI').Value := tdFechaInicial.Date ;
           Connection.QryBusca.Params.ParamByName('FechaF').DataType := ftDate ;
           Connection.QryBusca.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
           Connection.QryBusca.Open ;
           dReal := 0 ;
           If Connection.QryBusca.RecordCount > 0 Then
           Begin
               dReal := Connection.QryBusca.FieldValues['dReal'] ;
               Connection.QryBusca.Active := False ;
               Connection.QryBusca.SQL.Clear ;
               Connection.QryBusca.SQL.Add('Select distinct b.dIdFecha ' +
                                           'From bitacoradepersonal b INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                           'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                           'Where b.sContrato = :Contrato And b.sIdPlataforma = :Plataforma And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sContrato, b.sIdPlataforma, b.dIdFecha') ;
               Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
               Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
               Connection.QryBusca.Params.ParamByName('Plataforma').DataType := ftString ;
               Connection.QryBusca.Params.ParamByName('Plataforma').Value := QryPersonalPlataforma.FieldValues['sIdPlataforma'] ;
               Connection.QryBusca.Params.ParamByName('FechaI').DataType := ftDate ;
               Connection.QryBusca.Params.ParamByName('FechaI').Value := tdFechaInicial.Date ;
               Connection.QryBusca.Params.ParamByName('FechaF').DataType := ftDate ;
               Connection.QryBusca.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
               Connection.QryBusca.Open ;
               If Connection.QryBusca.RecordCount > 0 Then
                    dReal := dReal / Connection.QryBusca.RecordCount ;

               sPlataforma := QryPersonalPlataforma.FieldValues['sIdPlataforma']
          End
      End
    End

END;
end;

procedure TfrmReportePeriodo.btnImportarClick(Sender: TObject);
 Var
   size: Real ;
   indice, iItem  : integer;
   lTamanio : boolean;
   sArchivo : string;
begin
    if ComboActas.Text = '' then
    begin
        messageDLG('Seleccione un Acta Fotografica!', mtInformation, [mbOk], 0);
        comboActas.SetFocus;
    end
    else
    begin
        if tsNumeroActividad.Text = '' then
        begin
            messageDLG('Selecccione una Partida!', mtInformation, [mbOk], 0);
            tsNumeroActividad.SetFocus;
            exit;
        end;     
        OpenPicture.Title  := 'Inserta Imagen';
        sArchivo           := '' ;
        lTamanio           := True;

       If OpenPicture.Execute then
        begin
            indice := 0;
            OpenPicture.Files.Count;
            while indice < OpenPicture.Files.Count  do
            begin
                try
                    sArchivo := OpenPicture.Files.Strings[indice] ;
                    size := Tamanyo(sArchivo) ;
                    If size <= 1024 Then
                        bImagen2.Picture.LoadFromFile(OpenPicture.Files.Strings[indice])
                    Else
                    begin
                       MessageDlg('La imagen de demaciado grande, se adaptara a una mejor sin alterar su archivo original.', mtInformation, [mbOk], 0);
                       sArchivo := RedimensionarJPG(OpenPicture.Files[indice]);
                       size := Tamanyo(sArchivo);
                        lTamanio := True;
                    end;
                except
//                    bImagen.Picture.LoadFromFile('') ;
                end;
                inc(indice);
            end;
            if lTamanio = False then
               sArchivo := ''
            else
            begin
                If sArchivo <> '' Then
                Begin
                    btnPreview.Enabled := True ;
                    iItem  := 1 ;
                    indice := 0;
                    while indice < OpenPicture.Files.Count  do
                    begin
                        sArchivo := OpenPicture.Files.Strings[indice];

                        connection.zCommand.Active := False;
                        connection.zCommand.SQL.Clear;
                        connection.zCommand.SQL.Add('select max(iImagen) as iImagen from reportefotografico_acta '+
                                                    'where sContrato =:Contrato and sNumeroOrden =:Orden and sActaFotografica =:Acta group by sContrato');
                        connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                        connection.zcommand.Params.ParamByName('Contrato').Value    := Global_Contrato ;
                        connection.zcommand.Params.ParamByName('Orden').DataType    := ftString ;
                        connection.zcommand.Params.ParamByName('Orden').Value       := tsNumeroOrdenActa.Text;
                        connection.zcommand.Params.ParamByName('Acta').DataType     := ftString ;
                        connection.zcommand.Params.ParamByName('Acta').Value        := ComboActas.Text ;
                        connection.zCommand.Open;

                        if connection.zCommand.RecordCount > 0 then
                           iItem := connection.zCommand.FieldValues['iImagen'] + 1
                        else
                           iItem := 1;

                        connection.zCommand.Active := False ;
                        connection.zCommand.SQL.Clear ;
                        connection.zcommand.SQL.Add ( 'Insert Into reportefotografico_acta (sContrato, sNumeroOrden, sActaFotografica, dIdFecha, iImagen, bImagen, sWbs, sNumeroActividad) ' +
                                                      'Values (:Contrato, :Orden, :Acta, :Fecha, :Item, :Imagen, :Wbs, :Actividad)') ;
                        connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                        connection.zcommand.Params.ParamByName('Contrato').Value    := Global_Contrato ;
                        connection.zcommand.Params.ParamByName('Orden').DataType    := ftString ;
                        connection.zcommand.Params.ParamByName('Orden').Value       := tsNumeroOrdenActa.Text;
                        connection.zcommand.Params.ParamByName('Acta').DataType     := ftString ;
                        connection.zcommand.Params.ParamByName('Acta').Value        := ComboActas.Text ;
                        connection.zcommand.Params.ParamByName('Fecha').DataType    := ftDate ;
                        connection.zcommand.Params.ParamByName('Fecha').Value       := date ;
                        connection.zcommand.Params.ParamByName('Item').DataType     := ftInteger ;
                        connection.zcommand.Params.ParamByName('Item').Value        := iItem ;
                        connection.zcommand.Params.ParamByName('Imagen').LoadFromFile(sArchivo, ftGraphic) ;
                        connection.zcommand.Params.ParamByName('Wbs').DataType      := ftString ;
                        connection.zcommand.Params.ParamByName('Wbs').Value         := QryPartidasEfectivas.FieldValues['sWbs'];
                        connection.zcommand.Params.ParamByName('Actividad').DataType:= ftString ;
                        connection.zcommand.Params.ParamByName('Actividad').Value   := QryPartidasEfectivas.FieldValues['sNumeroActividad'];
                        connection.zCommand.ExecSQL();
                        inc(indice);
                    end;

                    // Actualizo Kardex del Sistema ....
                    connection.zCommand.Active := False ;
                    connection.zCommand.SQL.Clear ;
                    connection.zcommand.SQL.Add ('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen)' +
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
                    connection.zcommand.Params.ParamByName('Descripcion').Value := 'Agrega Fotografias a Reporte Actafotografica ' + comboActas.Text + ' del dia ' + DateToStr(Date) ;
                    connection.zcommand.Params.ParamByName('Origen').DataType := ftString ;
                    connection.zcommand.Params.ParamByName('Origen').Value := 'Reporte Fotografico' ;
                    connection.zCommand.ExecSQL ;

                    fotografico_Acta.Active := False;
                    fotografico_acta.ParamByName('Contrato').AsString := global_contrato;
                    fotografico_acta.ParamByName('Orden').AsString    := tsNumeroOrdenActa.Text;
                    fotografico_acta.ParamByName('Acta').AsString     := ComboActas.Text;
                    fotografico_acta.Open;

                    messageDLG('El Archivo fotografico se Guardó Correctamente!', mtInformation, [mbOk], 0);
                End;
            end;
        end;
    end;
end;

procedure TfrmReportePeriodo.btnAlbumClick(Sender: TObject);
Var
  Libro, Excel, Hoja, iHojas, XLShape: Variant;
  x, iFila, iCounter, iLoop, iColumnaProrrateo, ItemMayor, Longitud, Item: Integer;
  NombreDelExcel,
  SQLExtra,
  StrAux,
  StrOrden: String;

  SumaTotal, dDiferencia, SumaPorHorario, dMaximoValor: Real;

  ValoresProrrateados: Array Of Real;

  TmpName: String;
  TempPath: array [0..MAX_PATH-1] of Char;
  Fs: TStream;
  Pic : TJpegImage;
  imgAux: TImage;

  QueryOrden,
  QueryPartidas,
  QueryFases,
  QueryFotos: TZQuery;

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

begin
  Try
    QueryOrden:= TZQuery.Create(Self);
    QueryOrden.Connection := Connection.zConnection;

    QueryPartidas := TZQuery.Create(Self);
    QueryPartidas.Connection := Connection.zConnection;

    QueryFases := TZQuery.Create(Self);
    QueryFases.Connection := Connection.zConnection;

    QueryFotos := TZQuery.Create(Self);
    QueryFotos.Connection := Connection.zConnection;
    
    NombreDelExcel := 'ReporteFotografico.xls';
    Try
      Excel := CreateOleObject('Excel.Application');
    Except
      On E: Exception do begin
        FreeAndNil(Excel);
        ShowMessage(E.Message);
        Exit;
      end;
    End;
    Excel.Visible := True;
    Excel.DisplayAlerts:= False;
    Libro := Excel.Workbooks.Add;


    for Item := 0 to tsOrdenesdeTrabajo.Items.Count - 1 do begin
      if tsOrdenesDeTrabajo.Checked[Item] then begin
        QueryOrden.Active := False;
        QueryOrden.SQL.Text := 'SELECT * FROM ordenesdetrabajo WHERE sNumeroOrden = :Orden AND sContrato = :Contrato';
        QueryOrden.ParamByName('Contrato').AsString := Global_Contrato;
        QueryOrden.ParamByName('Orden').AsString := tsOrdenesDeTrabajo.Items.Strings[Item];
        QueryOrden.Open;

        StrOrden := 'FOLIO: ' + QueryOrden.FieldByName('sDescripcionCorta').AsString + #10#10 + QueryOrden.FieldByName('mDescripcion').AsString;

        QueryPartidas.Active := False;
        QueryPartidas.SQL.Text := 'SELECT rf.sNumeroOrden, ao.sNumeroActividad, ao.mDescripcion FROM reportefotografico AS rf INNER JOIN actividadesxorden AS ao ' +
                                  'ON (ao.sContrato = rf.sContrato AND ao.sNumeroOrden = rf.sNumeroOrden AND ao.sNumeroActividad = rf.sNumeroActividad) WHERE rf.sContrato = :Contrato AND rf.sNumeroOrden = :Folio GROUP BY rf.sNumeroActividad;';
        QueryPartidas.ParamByName('Contrato').AsString := Global_Contrato;
        QueryPartidas.ParamByName('Folio').AsString := QueryOrden.FieldByName('sNumeroOrden').AsString;
        QueryPartidas.Open;

        for iLoop := 1 to QueryPartidas.RecordCount do begin
          if Libro.Sheets.Count < iLoop then begin
            Libro.Sheets.Add;
          end;
        end;

        while Not QueryPartidas.Eof do begin
          if Libro.Sheets.Count < QueryPartidas.RecNo then begin
            Libro.Sheets.Add;
          end;

          Excel.WorkBooks[1].WorkSheets[QueryPartidas.RecNo].Name := 'PARTIDA ' + QueryPartidas.FieldByName('sNumeroActividad').AsString;
          Hoja := Excel.WorkBooks[1].WorkSheets[QueryPartidas.RecNo];
          Libro.Sheets[QueryPartidas.RecNo].Select;

          Excel.Columns[ColumnaNombre(1)+':'+ColumnaNombre(59)].ColumnWidth := 1;

          Excel.ActiveSheet.PageSetup.PrintTitleRows := '$1:$8';

          {$REGION 'IMAGENES DE CABECERA'}
          Try
            TmpName := '';
            imgAux := TImage.Create(nil);
            if TmpName='' then begin
              GetTempPath(SizeOf(TempPath), TempPath);
              TmpName:=TempPath +'imgtempSln1'+formatdatetime('dddddd hhnnss',now)+'.jpg';
              fs := Connection.contrato.CreateBlobStream(Connection.contrato.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
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
          Hoja.Shapes.AddPicture(TmpName, True, True, Excel.Cells[1, 2].Left, 2, 112, 52);
    
          Try
            TmpName := '';
            imgAux := TImage.Create(nil);
            if TmpName='' then begin
              GetTempPath(SizeOf(TempPath), TempPath);
              TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
          Hoja.Shapes.AddPicture(TmpName, True, True, (Excel.Cells[1, 49].Left - 2), 2, 79, 59);
          {$ENDREGION}

          {$REGION 'CABECERA - TEXTO'}
          XLShape := Hoja.Shapes.AddTextbox(1, Excel.Cells[1, 15].Left, 4, 290, 55);

          XLShape.Line.Visible := 0;
          XLShape.Fill.Visible := 0;

          Longitud := Length(Connection.contrato.FieldByName('sCliente').AsString);

          XLShape.TextFrame2.TextRange.Characters.Text := Connection.contrato.FieldByName('sCliente').AsString;
          XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Name := 'Arial';
          XLShape.TextFrame2.TextRange.Characters[1, Longitud].ParagraphFormat.Alignment := 2;
          XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Size := 8;
          XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Bold := True;
          {$ENDREGION}

          {$REGION 'DATOS DE FOLIO'}
          XLShape := Hoja.Shapes.AddTextbox(1, Excel.Cells[5, 2].Left, (Excel.Cells[5, 2].Top + 5), 500, 55);

          XLShape.Fill.Visible := msoTrue;
          XLShape.Fill.ForeColor.RGB := RGB(203, 203, 203);
          XLShape.Fill.Transparency := 0;

          XLShape.Line.Style := msoLineThickThin;
          XLShape.Line.Visible := msoTrue;
          XLShape.Line.ForeColor.RGB := RGB(0, 0, 255);
          XLShape.Line.Transparency := 0;
          XLShape.Line.Weight := 4.5;

          Longitud := Length(Trim(StrOrden));

          XLShape.TextFrame2.TextRange.Characters.Text := StrOrden;
          XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Name := 'Arial';
          XLShape.TextFrame2.TextRange.Characters[1, Longitud].ParagraphFormat.Alignment := 2;
          XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Size := 8;
          XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Bold := True;
          {$ENDREGION}

          QueryFases.Active := False;
          QueryFases.SQL.Text := 'SELECT sFasePartida FROM reportefotografico WHERE sContrato = :Contrato AND sNumeroOrden = :Orden AND sNumeroActividad = :Partida GROUP BY sFasePartida';
          QueryFases.ParamByName('Contrato').AsString := Global_Contrato;
          QueryFases.ParamByName('Orden').AsString := QueryOrden.FieldByName('sNumeroOrden').AsString;
          QueryFases.ParamByName('Partida').AsString := QueryPartidas.FieldByName('sNumeroActividad').AsString;
          QueryFases.Open;

          iFila := 10;

          while Not QueryFases.Eof do begin

            {$REGION 'DATOS DE PARTIDA'}
            XLShape := Hoja.Shapes.AddTextbox(1, Excel.Cells[iFila, 2].Left, (Excel.Cells[iFila, 2].Top), 500, 48);

            XLShape.Fill.Visible := msoTrue;
            XLShape.Fill.ForeColor.RGB := RGB(203, 203, 203);
            XLShape.Fill.Transparency := 0;

            XLShape.Line.Style := msoLineThickThin;
            XLShape.Line.Visible := msoTrue;
            XLShape.Line.ForeColor.RGB := RGB(0, 0, 255);
            XLShape.Line.Transparency := 0;
            XLShape.Line.Weight := 4.5;

            StrAux := #13 + 'CONTRATO: ' + Global_Contrato_Barco +#10#13+ 'INSTALACIÓN: PLATAFORMA ' + QueryOrden.FieldByName('sIdPlataforma').AsString;

            Longitud := Length(StrAux);

            XLShape.TextFrame2.TextRange.Characters.Text := StrAux;
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Name := 'Arial';
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].ParagraphFormat.Alignment := msoAlignLeft;
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Size := 8;
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Bold := True;

            {$REGION 'APARTADO DE PARTIDA'}
            XLShape := Hoja.Shapes.AddTextbox(1, Excel.Cells[iFila + 1, 42].Left, (Excel.Cells[iFila + 1, 42].Top + 3), 120, 20);

            XLShape.Line.Visible := 0;
            XLShape.Fill.Visible := 0;

            StrAux := 'PARTIDA '+QueryPartidas.FieldByName('sNumeroActividad').AsString+' ('+UpperCase(QueryFases.FieldByName('sFasePartida').AsString)+')';

            Longitud := Length(StrAux);

            XLShape.TextFrame2.TextRange.Characters.Text := StrAux;
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Name := 'Arial';
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].ParagraphFormat.Alignment := 2;
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Size := 8;
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Bold := True;
            {$ENDREGION}

            {$ENDREGION}

            {$REGION 'CUADRO DE FOTOS MAYOR'}
            Inc(iFila, 4);
            XLShape := Hoja.Shapes.AddTextbox(1, Excel.Cells[iFila, 2].Left, (Excel.Cells[iFila, 2].Top), 500, 430);

            XLShape.Fill.Visible := 0;

            XLShape.Line.Style := msoLineThickThin;
            XLShape.Line.Visible := msoTrue;
            XLShape.Line.ForeColor.RGB := RGB(0, 0, 255);
            XLShape.Line.Transparency := 0;
            XLShape.Line.Weight := 4.5;
            {$ENDREGION}

            {$REGION 'CUADRO DE FOTOS CENTRAL'}
            Inc(iFila, 3);
            XLShape := Hoja.Shapes.AddTextbox(1, Excel.Cells[iFila, 9].Left, (Excel.Cells[iFila, 9].Top), 380, 330);

            XLShape.Fill.Visible := 0;

            XLShape.Line.Style := msoLineThickThin;
            XLShape.Line.Visible := msoTrue;
            XLShape.Line.ForeColor.RGB := RGB(0, 0, 255);
            XLShape.Line.Transparency := 0;
            XLShape.Line.Weight := 4.5;
            {$ENDREGION}

            QueryFotos.Active := False;
            QueryFotos.SQL.Text := 'SELECT * FROM reportefotografico WHERE sContrato = :Contrato AND sNumeroOrden = :Orden AND sNumeroActividad = :Partida AND sFasePartida = :Fase';
            QueryFotos.ParamByName('Contrato').AsString := Global_Contrato;
            QueryFotos.ParamByName('Orden').AsString := QueryOrden.FieldByName('sNumeroOrden').AsString;
            QueryFotos.ParamByName('Partida').AsString := QueryPartidas.FieldByName('sNumeroActividad').AsString;
            QueryFotos.ParamByName('Fase').AsString := QueryFases.FieldByName('sFasePartida').AsString;
            QueryFotos.Open;

            if QueryFotos.RecordCount = 1 then begin

              {$REGION 'IMAGENES TIPO 1'}
              Try
                TmpName := '';
                imgAux := TImage.Create(nil);
                if TmpName='' then begin
                  GetTempPath(SizeOf(TempPath), TempPath);
                  TmpName:=TempPath +'imgtempSln'+IntToStr(QueryFotos.RecordCount)+formatdatetime('dddddd hhnnss',now)+'.jpg';
                  fs := QueryFotos.CreateBlobStream(QueryFotos.FieldByName('bImagen'), bmRead);
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
              Hoja.Shapes.AddPicture(TmpName, True, True, (Excel.Cells[iFila, 10].Left - 2), (Excel.Cells[iFila, 10].Top + 5), 360, 300);
              {$ENDREGION}

            end;

            if QueryFotos.RecordCount = 2 then begin
              while Not QueryFotos.Eof do begin

                {$REGION 'IMAGENES TIPO 2'}
                Try
                  TmpName := '';
                  imgAux := TImage.Create(nil);
                  if TmpName='' then begin
                    GetTempPath(SizeOf(TempPath), TempPath);
                    TmpName:=TempPath +'imgtempSln'+IntToStr(QueryFotos.RecordCount)+formatdatetime('dddddd hhnnss',now)+'.jpg';
                    fs := QueryFotos.CreateBlobStream(QueryFotos.FieldByName('bImagen'), bmRead);
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
                {$ENDREGION}

                Hoja.Shapes.AddPicture(TmpName, True, True, (Excel.Cells[iFila, 10].Left - 3), (Excel.Cells[iFila, 10].Top + 5), 360, 140);

                Inc(iFila, 11);
                QueryFotos.Next;
              end;
              iFila := iFila - 22;
            end;

            if QueryFotos.RecordCount > 2 then begin
              while Not QueryFotos.Eof do begin
                
                {$REGION 'IMAGENES TIPO 4'}
                Try
                  TmpName := '';
                  imgAux := TImage.Create(nil);
                  if TmpName='' then begin
                    GetTempPath(SizeOf(TempPath), TempPath);
                    TmpName:=TempPath +'imgtempSln'+IntToStr(QueryFotos.RecordCount)+formatdatetime('dddddd hhnnss',now)+'.jpg';
                    fs := QueryFotos.CreateBlobStream(QueryFotos.FieldByName('bImagen'), bmRead);
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

                {$ENDREGION}

                if (QueryFotos.RecNo mod 2) <> 0 then begin
                  Hoja.Shapes.AddPicture(TmpName, True, True, (Excel.Cells[iFila, 10].Left - 3), (Excel.Cells[iFila, 10].Top + 5), 180, 140);
                end else begin
                  Hoja.Shapes.AddPicture(TmpName, True, True, (Excel.Cells[iFila, 31].Left - 3), (Excel.Cells[iFila, 31].Top + 5), 180, 140);
                  Inc(iFila, 11);
                end;

                QueryFotos.Next;
              end;
              iFila := iFila - 22;
            end;

            {$REGION 'DESCRIPCIÓN DE LA PARTIDA'}
            Inc(iFila, 25);
            XLShape := Hoja.Shapes.AddTextbox(1, Excel.Cells[iFila, 2].Left, (Excel.Cells[iFila, 2].Top + 25), 500, 50);

            XLShape.Fill.Visible := msoTrue;
            XLShape.Fill.ForeColor.RGB := RGB(203, 203, 203);
            XLShape.Fill.Transparency := 0;

            XLShape.Line.Style := msoLineThickThin;
            XLShape.Line.Visible := msoTrue;
            XLShape.Line.ForeColor.RGB := RGB(0, 0, 255);
            XLShape.Line.Transparency := 0;
            XLShape.Line.Weight := 4.5;

            StrAux := 'DESCRIPCIÓN ' +#10#13+ QueryPartidas.FieldByName('mDescripcion').AsString;

            Longitud := Length(StrAux);

            XLShape.TextFrame2.TextRange.Characters.Text := StrAux;
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Name := 'Arial';
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].ParagraphFormat.Alignment := msoAlignCenter;
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Size := 8;
            XLShape.TextFrame2.TextRange.Characters[1, Longitud].Font.Bold := True;
            {$ENDREGION}

            Inc(iFila, 7);
            
            Hoja.Rows[iFila - 1].PageBreak := 1;
            
            QueryFases.Next;
          end;

          {$REGION 'PROPIEDADES DE HOJA'}
          Excel.ActiveWindow.View := xlPageBreakPreview;
          Excel.ActiveWindow.DisplayGridlines := False;
          Excel.ActiveWindow.Zoom := 100;

          Excel.ActiveSheet.PageSetup.TopMargin  := Excel.InchesToPoints(0.196850393700787);//Excel.InchesToPoints(0.2);
          Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(0.196850393700787);//Excel.InchesToPoints(1.521);
          Excel.ActiveSheet.PageSetup.LeftMargin  := Excel.InchesToPoints(0.196850393700787);
          Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0.196850393700787);
          Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0.196850393700787);
          Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0.196850393700787);
          Excel.ActiveSheet.PageSetup.FitToPagesWide := 1;
          Excel.ActiveSheet.PageSetup.FitToPagesTall := 1;
          {$ENDREGION}

          QueryPartidas.Next;
        end;
      end;
    end;

  Finally

  End;
end;

procedure TfrmReportePeriodo.btnDeleteClick(Sender: TObject);
begin
   If MessageDlg('Desea eliminar la Imagen '+ IntToStr(fotografico_acta.FieldValues['iImagen']) + ' del '+ comboActas.text, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
   Begin
       fotografico_acta.Delete;

       // Actualizo Kardex del Sistema ....
        connection.zCommand.Active := False ;
        connection.zCommand.SQL.Clear ;
        connection.zcommand.SQL.Add ('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen)' +
                                     'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
        connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
        connection.zcommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
        connection.zcommand.Params.ParamByName('Usuario').DataType := ftString ;
        connection.zcommand.Params.ParamByName('Usuario').Value := Global_Usuario ;
        connection.zcommand.Params.ParamByName('Fecha').DataType := ftDate ;
        connection.zcommand.Params.ParamByName('Fecha').Value          := Date ;
        connection.zcommand.Params.ParamByName('Hora').DataType        := ftString ;
        connection.zcommand.Params.ParamByName('Hora').value           := FormatDateTime('hh:mm:ss', Now) ;
        connection.zcommand.Params.ParamByName('Descripcion').DataType := ftString ;
        connection.zcommand.Params.ParamByName('Descripcion').Value    := 'Elimina Fotografia '+ IntToStr(fotografico_acta.FieldValues['iImagen']) +' a Reporte Actafotografica ' + comboActas.Text + ' del dia ' + DateToStr(Date) ;
        connection.zcommand.Params.ParamByName('Origen').DataType      := ftString ;
        connection.zcommand.Params.ParamByName('Origen').Value         := 'Reporte Fotografico' ;
        connection.zCommand.ExecSQL ;
   End;
end;

procedure TfrmReportePeriodo.btnDownClick(Sender: TObject);
begin
    OrdenarFotos('Abajo');
end;

procedure TfrmReportePeriodo.BtnImpLBClick(Sender: TObject);
var Ih:Integer;
  ExcelAp,Libro: Variant;


procedure ActualiaFactorGeneradorEQ(sParamEmbarcacion: string; sParamOrden: string; sParamFolio: string);
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

    zqFactoresEquipo.First;
    while not zqFactoresEquipo.Eof do
    begin
        TotalProgreso := TotalProgreso + Progreso;


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

procedure ActualiaFactorGeneradorPER(sParamEmbarcacion: string; sParamOrden: string; sParamFolio: string);
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

    zqFactoresPersonal.First;
    while not zqFactoresPersonal.Eof do
    begin
        TotalProgreso := TotalProgreso + Progreso;
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

begin
  if length(trim(dbFolio01.Text)) = 0then
  begin
   dbFolio01.SetFocus;
   raise Exception.Create('Seleccione un folio porfavor.');
  end;

  if not BotonPermiso.imprimir then
    raise Exception.Create('No tiene permisos de impresión');

  ZqrContrato.Close;
  ZqrContrato.ParamByName('contrato').AsString := global_contrato;
  ZqrContrato.Open;

  ZQRFolio.CLOSE;
  ZQRFolio.ParamByName('Folio').AsString := dbFolio01.Text;
  ZQRFolio.ParamByName('contrato').AsString := global_contrato;
  ZQRFolio.Open;

  if ZQRFolio.RecordCount = 0 then
    raise Exception.Create('No se encontró el folio especificado.');


  ZQRperiodo.Close;
  ZQRperiodo.ParamByName('Folio').AsString := dbFolio01.Text;
  ZQRperiodo.ParamByName('contrato').AsString := global_contrato;
  ZQRperiodo.Open;

  if ZQRperiodo.RecordCount = 0 then
    raise Exception.Create('No se encontró reportado para el folio seleccionado.');

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
  Libro := ExcelAp.Workbooks.Add;

  ActualiaFactorGeneradorPER(global_barco, global_contrato, ZQRFolio.FieldValues['sNumeroOrden']);
  ActualiaFactorGeneradorEQ(global_barco, global_contrato, ZQRFolio.FieldValues['sNumeroOrden']);



  for Ih := LsContenido.Count - 1 downto 0 do
  begin
    if LsContenido.Checked[Ih] then
    begin
      ImprimeHoja(LsContenido.items[Ih],Libro,ExcelAp);
    end;
  end;
end;

procedure TfrmReportePeriodo.ImprimeHoja(nombre:string;Lb,excel:Variant);
var Libro,Hoja:Variant;
  lAplicaLibro:Boolean;
  ContAux:Integer;
  IVarAux,IVarAux2:Integer;
  icolumna,ifila,ciclos:Integer;
  tmpname,TempPath,sVarAux,sVarAux2,svaraux3,svar:string;
  dtVarAux1,dtVarAux2,RgFecha1,RgFecha2:TDateTime;

  DVarAux1,DVarAux2:Real;

  Dia, Mes, Año,Dia2, Mes2, Año2: Word;

  imgaux:timage;
  fs:TStream;
  Pic:TJpegImage;

  QryConsulta1,QryConsulta2:TZReadOnlyQuery;//Consultas varias

  ImprimirTotal:Boolean;

  {$REGION 'PROCEDIMIENTOS VARIOS'}
  procedure EncabezadoTexto;
  begin
    if ZQRConfiguracion.State = dsBrowse then
    begin
      if not ZQRConfiguracion.Locate('scontrato',global_contrato,[]) then
      begin
        ZQRConfiguracion.close;
        ZQRConfiguracion.ParamByName('contrato').asstring:=global_contrato;
        ZQRConfiguracion.Open;
      end;
    end
    else
    begin
      ZQRConfiguracion.close;
      ZQRConfiguracion.ParamByName('contrato').asstring:=global_contrato;
      ZQRConfiguracion.Open;
    end;

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.Value := ZQRConfiguracion.FieldByName('sNombre').AsString;
    Inc (iFila);

    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 38, True, 8, clBlack, 'Arial');
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.Value := ZQRConfiguracion.FieldByName('sDireccion1').AsString + ' '+#10
    + ZQRConfiguracion.FieldByName('sDireccion2').AsString +' '+#10+
    ZQRConfiguracion.FieldByName('sCiudad').AsString;
    Excel.Selection.ReadingOrder := xlContext;
    Excel.Selection.WrapText := True;
    Hoja.PageSetup.PrintTitleRows := '$1:$5';
    Inc (iFila,2);
  end;

  procedure EncabezadoImagen(Izquierda,Derecha:boolean;Modo:integer = 1);
  begin
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
          Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 1, 150, 85);

        if Modo = 2 then
          Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 1, 135, 70);
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
        if Modo = 1 then
          Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 375, 1, 90, 85);
        if Modo = 2 then
          Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 490, 1, 75, 70);
      end;
    end;
  end;

  procedure FormatoColumnas(VarEx:Variant;Col:integer = 1;Modo:integer = 1);
  begin
    if Modo = 1 then
    begin
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 3.57;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 5.86;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 10.71;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 6.86;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 6;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 8.57;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 7.43;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 9.14;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 20.71;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 3.29;
      Inc(col);
    end;
    if Modo = 2 then
    begin
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 0.75;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 9;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 9;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 9;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 9;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 9;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 10;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 10;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 9;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 7;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 9;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 9;
      Inc(col);
      VarEx.Columns[ColumnaNombre(col)+':'+ColumnaNombre(col)].ColumnWidth := 0.75;
      Inc(col);
    end;
  end;

  procedure HojaCarta(VarEx:Variant;Tpo:Integer = 1);
  begin
    //Ajustamos la hoja para mejor presentacion
    VarEx.ActiveWindow.View := 2;
    VarEx.ActiveSheet.PageSetup.LeftMargin := 0.7;
    VarEx.ActiveSheet.PageSetup.RightMargin := 0.7;
    VarEx.ActiveSheet.PageSetup.TopMargin := 0.75;
    VarEx.ActiveSheet.PageSetup.BottomMargin := 0.75;
    VarEx.ActiveSheet.PageSetup.HeaderMargin := 0.3 ;
    VarEx.ActiveSheet.PageSetup.FooterMargin := 0.3;
    VarEx.ActiveSheet.PageSetup.PrintHeadings := False;
    VarEx.ActiveSheet.PageSetup.PrintGridlines := False;
    VarEx.ActiveSheet.PageSetup.PrintQuality := 600;
    VarEx.ActiveSheet.PageSetup.CenterHorizontally := False;
    VarEx.ActiveSheet.PageSetup.CenterVertically := False;
    VarEx.ActiveSheet.PageSetup.Draft := False;
    VarEx.ActiveSheet.PageSetup.PaperSize := 1;
    VarEx.ActiveSheet.PageSetup.BlackAndWhite := False;
    VarEx.ActiveSheet.PageSetup.Zoom := False;
    VarEx.ActiveSheet.PageSetup.FitToPagesWide := 1;
    if tpo = 1 then  //Todo en una pagina
      VarEx.ActiveSheet.PageSetup.FitToPagesTall := 1;
    if tpo = 2 then
      VarEx.ActiveSheet.PageSetup.FitToPagesTall := 2;
    VarEx.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
    VarEx.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := True;
    VarEx.ActiveWindow.Zoom := 100;
  end;
  {$ENDREGION}

  function cuentasabadosydomingos(fechaInicial:Tdatetime;fechafinal:Tdatetime):Integer;
  var
    Dias:Integer;
    e,c:integer;
    totaldominsaba:integer;
  begin
    Dias := Trunc (fechafinal) - Trunc (fechainicial);
    dias:=dias+1;
    totaldominsaba:=0;
    c:=(Dias-1);
    for e:=0 to c do
    begin
      if (DayOfTheWeek(fechaInicial)=6) or(DayOfTheWeek(fechaInicial)=7) then
      begin
        totaldominsaba:=totaldominsaba+1;
      end;
      fechaInicial:=IncDay(fechaInicial,1);
    end;
    cuentasabadosydomingos:=totaldominsaba;
  end;

  function CuentaDomingos(IniDate: TDateTime; EndDate: TDateTime): Integer;
  var
    Sundays: Integer;
    i,Dias : Integer;
  begin
    Sundays := 0;
    Dias := DaysBetween(IniDate,EndDate);
    for I := 0 to dias - 1 do begin
        If DayOfWeek(IniDate+i) = 1 then Sundays := Sundays + 1;
        If DayOfWeek(IniDate+i) = 7 then Sundays := Sundays + 1;
    end;
    Result := SunDays;
  end;
begin
  {
    dataset de folio precargado con folio ya zqrfolio
    dataset con contenido precargado zqrcontenido
    dataset de configuracion zqrconfiguracion, este debe abrirse y cerrarse
    de acuerdo a la configuracion de contrato que se necesite
  }
  try
    Hoja := Lb.Sheets.Add;
    Hoja.Name := nombre;

    if not zqrcontenido.locate('snombreportada',nombre,[]) then
      raise exception.Create('***');

    if zqrcontenido.fieldbyname('ltipo').asstring = 'PORTADA' then
    begin
      {$REGION 'PORTADA'}
      iColumna := 1;
      iFila := 1;
      FormatoColumnas(Excel,icolumna);
      EncabezadoImagen(True,True);

      Excel.Range['A1:J6'].Select;
      PFormatosExcel_H2(Excel, 0, True, 10);
      Excel.activeWindow.DisplayGridlines := false;

      {$REGION 'CONTENIDO'}

      iFila := 7;

      Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 65, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := Connection.contrato.FieldByName('mComentarios').AsString;
      Inc (iFila,3);

      Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'CONTRATO:' + Global_Contrato_Barco;
      Inc (iFila,3);

      Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 57, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      svarAux := connection.contrato.FieldByName('stitulo').AsString;
      svaraux := AnsireplaceText( svaraux, 'OBJETO DEL CONTRATO:', '' );
      Excel.Selection.Value := svaraux;

      Inc (iFila);

      Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 57, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlRight;
      Excel.Selection.Value := 'TRABAJOS REALIZADOS EN INSTALACIÓN:';

      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 57, True, 11, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'PLATAFORMA:' + zqrfolio.FieldByName('sIdPlataforma').AsString ;
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 15, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'FOLIO:' ;
      Inc (iFila);

      Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 15, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrfolio.FieldByName('sIdFolio').AsString;

      ZqrConfiguracion.Close;
      ZqrConfiguracion.parambyname('contrato').asstring := global_contrato_barco;
      ZqrConfiguracion.Open;

      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then
        begin
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtempPortada.jpg';
          fs := ZQRContenido.CreateBlobStream(ZQRContenido.FieldByName('bImagen1'), bmRead);
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
        Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 20, 390, 410, 245);

      Excel.Range[ColumnaNombre(1)+IntToStr(18)+':'+ColumnaNombre(10)+IntToStr(36)].Select;
      PFormatosExcel_H2(Excel, 12, True, 8, clBlack, 'Arial');

      Excel.Range[ColumnaNombre(2)+IntToStr(37)+':'+ColumnaNombre(9)+IntToStr(46)].Select;
      PFormatosExcel_H2(Excel, 12, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrfolio.FieldByName('mDescripcion').AsString;

      {$ENDREGION}

        //Ajustamos la hoja para mejor presentacion
        HojaCarta(Excel);

      {$ENDREGION}
    end;

    if zqrcontenido.fieldbyname('ltipo').asstring = 'PRESENTACION' then
    begin
      {$REGION 'PLANTILLA PRESENTACION'}
      iColumna := 1;
      iFila := 1;
      FormatoColumnas(Excel,icolumna);

      EncabezadoImagen(False,True);

      //Texto de Cabecera
      Excel.Range['A1:J1'].Select;
      PFormatosExcel_H2(Excel, 0, True, 10);

      Excel.activeWindow.DisplayGridlines := false;


      iFila := 2;
      EncabezadoTexto;

      inc(ifila,7);
      Excel.Rows[ifila].RowHeight := 40;

      inc(ifila,2);
      Excel.Rows[ifila].RowHeight := 40;

      inc(ifila,2);
      Excel.Rows[ifila].RowHeight := 40;
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;

      Excel.Selection.Value := ZQRContenido.fieldbyname('sdescripcion').asstring;

      Excel.Rows[ifila].RowHeight := 50;
      Excel.Selection.Borders.LineStyle := xlContinuous;
      excel.selection.interior.colorindex := 15;

      inc(ifila,1);
      Excel.Rows[ifila].RowHeight := 40;
      inc(ifila,2);
      Excel.Rows[ifila].RowHeight := 40;
      inc(ifila,8);
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := '';

      Excel.Rows[ifila].RowHeight := 10;

      HojaCarta(Excel);
      {$ENDREGION}
    end;

    if zqrcontenido.fieldbyname('ltipo').asstring = 'INDICE' then
    begin
      {$REGION 'CONTENIDO'}
      iColumna := 1;
      iFila := 1;
      FormatoColumnas(Excel,icolumna);
      EncabezadoImagen(False,True);

      Excel.Range['A1:J1'].Select;
      PFormatosExcel_H2(Excel, 0, True, 10);
      Excel.activeWindow.DisplayGridlines := false;

      ifila := 2;
      EncabezadoTexto;
      Inc (iFila,4);

      {$REGION 'CUERPO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;

      Excel.Selection.Value := ZQRContenido.fieldbyname('sdescripcion').asstring;
      try
        ivaraux := length(ZQRContenido.fieldbyname('sdescripcion').asstring);
        if (IVarAux mod 50) > 0 then
          IVarAux := IVarAux + 50;
        ivaraux := ivaraux  div 50;
        ivaraux := ivaraux * 18;
        Excel.Rows[ifila].RowHeight := ivaraux;
      Except
        Excel.Rows[ifila].RowHeight := 18;
      end;
      Excel.Selection.Borders.LineStyle := xlContinuous;
      excel.selection.interior.colorindex := 15;
      inc(ifila,1);

      ContAux := zqrContenido.RecNo;
      try
        ZQRContenido.First;
        while not ZQRContenido.Eof do
        begin
          if (ZQRContenido.FieldByName('lincluirindice').AsString = 'Si') then
          begin
            IVarAux := LsContenido.Items.IndexOf(ZQRContenido.FieldByName('snombreportada').asstring);
            if IVarAux > -1 then
              if LsContenido.Checked[ivaraux] then
              begin
                inc(ifila,1);
                Excel.Rows[ifila].RowHeight := 25;
                inc(ifila,1);
                Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
                PFormatosExcel_H2(Excel, 16, True, 11, clBlack, 'Arial');
                Excel.Selection.HorizontalAlignment := xlLeft;
                Excel.Selection.Value := ZQRContenido.fieldbyname('sdescripcion').asstring;
                try
                  ivaraux := length(ZQRContenido.fieldbyname('sdescripcion').asstring);
                  if (IVarAux mod 50) > 0 then
                    IVarAux := IVarAux + 50;
                  ivaraux := ivaraux  div 50;
                  ivaraux := ivaraux * 18;
                  Excel.Rows[ifila].RowHeight := ivaraux;
                Except
                  Excel.Rows[ifila].RowHeight := 18;
                end;

              end;
          end;
          ZQRContenido.Next;
        end;

      finally
        ZQRContenido.RecNo := ContAux;
      end;


      {$ENDREGION}

      HojaCarta(Excel);
      {$ENDREGION}
    end;

    if zqrcontenido.fieldbyname('ltipo').asstring = 'OFICIO-T1' then
    begin
      {$REGION 'OFICIO1'}
      iColumna := 1;
      iFila := 1;
      FormatoColumnas(Excel,icolumna);
      EncabezadoImagen(False,True);

      Excel.Range['A1:J1'].Select;
      PFormatosExcel_H2(Excel, 0, True, 10);
      Excel.activeWindow.DisplayGridlines := false;

      ifila := 2;
      EncabezadoTexto;

      {$REGION 'CONTENIDO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      //Excel.Selection.Value := zqrconfiguracion.FieldByName('sCiudad').AsString+' a '+FormatDateTime('dd'+'" de "'+'mmmm'+'" de "'+ 'yyyy',zqrfolio.FieldByName('dffprogramado').AsDateTime +15 );
      IVarAux := 15;
      if DayOfWeek(zqrperiodo.FieldByName('maximo').AsDateTime +15) = 1 then
        IVarAux := ivaraux+1;
      Excel.Selection.Value := zqrconfiguracion.FieldByName('sCiudad').AsString+' a '+FormatDateTime('dd'+'" de "'+'mmmm'+'" de "'+ 'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime +ivaraux );

      //PFormatosExcel_SoloBorde(Excel);
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'OFICIO: ' + zqrfolio.FieldByName('soficioautorizacion').asstring;
      //PFormatosExcel_SoloBorde(Excel);
      Inc (iFila,2);


      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila+3)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := Xlleft;
      Excel.Selection.formula :=  Connection.contrato.FieldByName('mComentarios').AsString;
      Inc (iFila,5);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := Xlleft;
      if length(trim(zqrcontenido.FieldByName('snombref1').asstring)) > 0 then
        Excel.Selection.Value := 'ATN '+uppercase(zqrcontenido.FieldByName('snombref1').asstring);
      //PFormatosExcel_SoloBorde(Excel);
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := Xlleft;
      Excel.Selection.Value := zqrcontenido.FieldByName('scargof1').AsString;
      //PFormatosExcel_SoloBorde(Excel);
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False,8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := XlRight;
      Excel.Selection.Value := 'Asunto: '+zqrcontenido.FieldByName('stexto1').AsString;
      //PFormatosExcel_SoloBorde(Excel);
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := zqrcontenido.FieldByName('mtexto2').AsString;
      //PFormatosExcel_SoloBorde(Excel);
      Excel.Rows[ifila].RowHeight := 50;
      Inc (iFila,3);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 10, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'FOLIO';
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 10, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrfolio.FieldByName('snumeroorden').asstring;
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'PROGRAMA:';
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 59, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrfolio.FieldByName('mdescripcion').asstring;
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'Instalación:';

      Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'Periodo de ejecución:';
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'PLATAFORMA: ' + zqrfolio.FieldByName('sidplataforma').asstring;

      Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 9, True, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;

      if FormatDateTime('yyyy',zqrperiodo.FieldByName('minimo').AsDateTime) = FormatDateTime('yyyy',zqrperiodo.FieldByName('maximo').AsDateTime)  then
      begin
        if FormatDateTime('mm',zqrperiodo.FieldByName('minimo').AsDateTime) = FormatDateTime('mm',zqrperiodo.FieldByName('maximo').AsDateTime)  then
          Excel.Selection.Value := 'DEL '+FormatDateTime('dd',zqrperiodo.FieldByName('minimo').AsDateTime)+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime))
        else
          Excel.Selection.Value := 'DEL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm',zqrperiodo.FieldByName('minimo').AsDateTime))+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime));
      end
      else
        Excel.Selection.Value := 'DEL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('minimo').AsDateTime))+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime));


      IVarAux2 := Length('PLATAFORMA: ' + zqrfolio.FieldByName('sidplataforma').asstring);
      if (IVarAux2 mod 23) > 0 then
        IVarAux2 := IVarAux2+23;
      IVarAux2 := IVarAux2 div 23;
      IVarAux2 := IVarAux2 * 11;
      if Excel.Rows[ifila].RowHeight < IVarAux2 then
        Excel.Rows[ifila].RowHeight := IVarAux2;

      Inc (iFila,3);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := Xlleft;
      Excel.Selection.Value := zqrcontenido.FieldByName('mtexto3').asstring;
      //PFormatosExcel_SoloBorde(Excel);
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := XlCenter;
      Excel.Selection.Value := 'Atentamente';
      //PFormatosExcel_SoloBorde(Excel);
      Inc (iFila,4);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False,8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrcontenido.FieldByName('snombref2').AsString;
      //PFormatosExcel_SoloBorde(Excel);
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrcontenido.FieldByName('scargof2').AsString;
      //PFormatosExcel_SoloBorde(Excel);
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'c.c.p. Archivo';
      Inc (iFila);

    {$ENDREGION}

      HojaCarta(Excel);
    {$ENDREGION}
    end;

    if zqrcontenido.fieldbyname('ltipo').asstring = 'OFICIO-T2' then
    begin
      {$REGION 'OFICIO2'}
      iColumna := 1;
      iFila := 1;
      FormatoColumnas(Excel,icolumna);
      EncabezadoImagen(False,True);

      Excel.Range['A1:J1'].Select;
      PFormatosExcel_H2(Excel, 0, True, 10);
      Excel.activeWindow.DisplayGridlines := false;

      {$REGION 'CABECERA'}
      iFila := 2;
      encabezadotexto;
      {$ENDREGION}

      {$REGION 'CONTENIDO'}
      zqrconfiguracion.Active := False;
      zqrconfiguracion.ParamByName('contrato').AsString := global_contrato;
      zqrconfiguracion.Open;

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 86, True, 10, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := Connection.contrato.FieldByName('mComentarios').AsString;
      Excel.Rows[iFila].RowHeight := 86.25;
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := '';
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'TRABAJOS REALIZADOS EN:';
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'PLATAFORMA:' + zqrfolio.fieldbyname('sidplataforma').asstring;
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := '';
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'FOLIO';
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrfolio.fieldbyname('snumeroorden').asstring;
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 31, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := '';
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'PROGRAMA';
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 82, True,9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrfolio.fieldbyname('mdescripcion').asstring;
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := '';
      Inc (iFila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'Periodo de Ejecución';
      Inc (iFila,3);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;

      if FormatDateTime('mm yyyy',zqrperiodo.FieldByName('minimo').AsDateTime) = FormatDateTime('mm yyyy',zqrperiodo.FieldByName('maximo').AsDateTime)  then
        Excel.Selection.Value := 'DEL '+FormatDateTime('dd',zqrperiodo.FieldByName('minimo').AsDateTime)+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime))
      else
        Excel.Selection.Value := 'DEL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('minimo').AsDateTime))+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime));

      Inc (iFila,5);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'CONTRATO:'+global_contrato_barco;
      Inc (iFila,4);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 10, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrcontenido.FieldByName('stexto1').asstring;
      Inc (iFila);
      {$ENDREGION}

      HojaCarta(Excel);
      {$ENDREGION}
    end;

    if zqrcontenido.fieldbyname('ltipo').asstring = 'OFICIO-T3' then
    begin
      {$REGION 'OFICIO3'}
      iColumna := 1;
      iFila := 1;
      FormatoColumnas(Excel,icolumna);
      for IVarAux := 2 to 9 do
        Excel.Columns[ColumnaNombre(IVarAux)+':'+ColumnaNombre(IVarAux)].ColumnWidth := 10;

      EncabezadoImagen(False,True);
      ifila := 2;
      EncabezadoTexto;
      Excel.activeWindow.DisplayGridlines := false;
      {$REGION 'CONTENIDO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := FormatDateTime('dd'+'" de "'+'mmmm'+'" de "'+ 'yyyy',ZQRperiodo.FieldByName('maximo').AsDateTime );
      Inc (iFila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila+3)].Select;
      //Excel.Selection.MergeCells := True;
      PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'ATN: '+zqrContenido.FieldByName('snombref1').AsString +#10+zqrContenido.FieldByName('sCargof1').AsString;
      Inc (iFila,4);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'Asunto: '+zqrcontenido.FieldByName('stexto1').AsString;
      Inc (iFila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlJustify;
      Excel.Selection.Value := zqrcontenido.FieldByName('mtexto2').AsString;
      IVarAux := Length(zqrcontenido.FieldByName('mtexto2').AsString);
      if (IVarAux mod 99) > 0 then
        IVarAux := IVarAux+99;
      IVarAux := IVarAux div 99;
      IVarAux := IVarAux * 15;
      if Excel.Rows[ifila].RowHeight < IVarAux then
        Excel.Rows[ifila].RowHeight := IVarAux;
      Inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
      Excel.Selection.Value := 'PARTIDA';
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.interior.colorindex := 28;
      PFormatosExcel_SoloBorde(Excel);


      Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Value  := 'ACTIVIDADES';
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.interior.colorindex := 28;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      Excel.Selection.Value  := 'AVANCE';
      Excel.Selection.interior.colorindex := 28;
      Excel.Selection.HorizontalAlignment := xlCenter;
      PFormatosExcel_SoloBorde(Excel);
      Inc(ifila,1);

      //partidas
      QryConsulta1 := TZReadOnlyQuery.Create(nil);
      try
        QryConsulta1.Connection := connection.zConnection;
        QryConsulta1.Active := False;
        QryConsulta1.SQL.Clear;

        QryConsulta2 := TZReadOnlyQuery.Create(nil);
        try
          QryConsulta2.Connection := connection.zConnection;
          QryConsulta2.Active := False;
          QryConsulta2.SQL.Clear;

          QryConsulta1.SQL.Text := 'SELECT stipoactividad,scontrato,snumeroorden,snumeroactividad,swbs,mdescripcion FROM actividadesxorden WHERE sContrato = :Contrato AND sNumeroOrden = :folio ORDER BY iItemOrden  ';
          QryConsulta1.ParamByName('contrato').AsString := global_contrato;
          QryConsulta1.ParamByName('folio').AsString := ZQRFolio.FieldByName('snumeroorden').AsString;
          QryConsulta1.Open;


          while not QryConsulta1.Eof do
          begin
            QryConsulta2.sql.clear;
            if QryConsulta1.FieldByName('stipoactividad').AsString = 'Actividad' then
            begin
              QryConsulta2.SQL.Text := 'SELECT IFNULL((SUM(dcantidad)),0) as avance FROM bitacoradeactividades WHERE scontrato = :contrato AND sNumeroOrden = :folio AND sIdTipoMovimiento = "ED" AND sWbs = :swbs  ';
              QryConsulta2.Active := False;
              QryConsulta2.ParamByName('contrato').AsString := QryConsulta1.FieldByName('scontrato').AsString;
              QryConsulta2.ParamByName('folio').AsString := QryConsulta1.FieldByName('snumeroorden').AsString;
              QryConsulta2.ParamByName('swbs').AsString := QryConsulta1.FieldByName('swbs').AsString;
              QryConsulta2.Open;
            end;
            if QryConsulta1.FieldByName('stipoactividad').AsString = 'Paquete' then
            begin
              QryConsulta2.SQL.Text :=  'select ifnull(sum(t1.cantidad),0) /100 as avance from (select ao.inivel,ao.sNumeroActividad,ao.dPonderado,ao.snumeroorden, '+
                                        'sum(ba.dcantidad*ao.dponderado) as cantidad '+
                                        'FROM actividadesxorden ao '+
                                        'inner join bitacoradeactividades ba on (ba.scontrato = ao.scontrato and ba.snumeroorden = ao.snumeroorden and ba.snumeroactividad = ao.snumeroactividad and ba.sidtipomovimiento = "ED")'+
                                        'where ao.sContrato = :contrato and ao.snumeroorden= :Folio and ao.swbs like :swbs '+
                                        'group by ao.sNumeroActividad)t1';
              QryConsulta2.Active := False;
              QryConsulta2.ParamByName('contrato').AsString := QryConsulta1.FieldByName('scontrato').AsString;
              QryConsulta2.ParamByName('folio').AsString := QryConsulta1.FieldByName('snumeroorden').AsString;
              QryConsulta2.ParamByName('swbs').AsString := QryConsulta1.FieldByName('swbs').AsString+'%';
              QryConsulta2.Open;
            end;

            //inserccion de datos
            Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
            excel.Selection.NumberFormat := '@';
            if Length(Trim(QryConsulta1.FieldByName('snumeroactividad').AsString)) = 1 then
              Excel.Selection.value := QryConsulta1.FieldByName('snumeroactividad').AsString+'.0'
            else
              Excel.Selection.value := QryConsulta1.FieldByName('snumeroactividad').AsString;
            if QryConsulta1.FieldByName('stipoactividad').AsString = 'Paquete' then
            begin
              PFormatosExcel_H2(Excel, 16, True, 7, clBlack, 'Arial');
              Excel.Selection.interior.colorindex := 28;
            end
            else
              PFormatosExcel_H2(Excel, 16, False, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlCenter;
            PFormatosExcel_SoloBorde(Excel);


            Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
            Excel.Selection.MergeCells := True;
            if QryConsulta1.FieldByName('stipoactividad').AsString = 'Paquete' then
            begin
              PFormatosExcel_H2(Excel, 16, True, 7, clBlack, 'Arial');
              Excel.Selection.interior.colorindex := 28;
            end
            else
              PFormatosExcel_H2(Excel, 16, False, 7, clBlack, 'Arial');
            Excel.Selection.WrapText := True;
            Excel.Selection.HorizontalAlignment := xlLeft;
            if QryConsulta1.FieldByName('stipoactividad').AsString = 'Paquete' then
               Excel.Selection.value := ZQRFolio.FieldByName('mdescripcion').AsString
            else
              Excel.Selection.value  := QryConsulta1.FieldByName('mdescripcion').AsString;
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
            if QryConsulta2.FieldByName('avance').AsFloat > 1 then
              Excel.Selection.Value := 1
            else
              Excel.Selection.Value  := QryConsulta2.FieldByName('avance').AsFloat;
            if QryConsulta1.FieldByName('stipoactividad').AsString = 'Paquete' then
            begin
              PFormatosExcel_H2(Excel, 16, True, 7, clBlack, 'Arial');
              Excel.Selection.interior.colorindex := 28;
            end
            else
              PFormatosExcel_H2(Excel, 16, False, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlCenter;
            excel.Selection.NumberFormat := '0.00%';
            PFormatosExcel_SoloBorde(Excel);

            IVarAux := Length(QryConsulta1.FieldByName('mdescripcion').AsString);
            if (IVarAux mod 77) > 0 then
              IVarAux := IVarAux+77;
            IVarAux := IVarAux div 77;
            IVarAux := IVarAux * 11;
            if Excel.Rows[ifila].RowHeight < IvarAux  then
              Excel.Rows[ifila].RowHeight := IVarAux;            
            Inc(ifila,1);
            QryConsulta1.Next;
          end;
        finally
          QryConsulta2.free;
        end;
      finally
        QryConsulta1.free;
      end;
      Inc(ifila,2);
      //notas
      QryConsulta1 := TZReadOnlyQuery.Create(nil);
      try
        QryConsulta1.Connection := connection.zConnection;
        QryConsulta1.Active := False;
        QryConsulta1.SQL.Clear;
        QryConsulta1.SQL.Text := 'SELECT * FROM bitacoradeactividades WHERE scontrato = :contrato AND sIdTipoMovimiento = "ng" AND snumeroorden = :folio ORDER BY didfecha limit 0';
        QryConsulta1.ParamByName('contrato').AsString := ZQRFolio.FieldByName('scontrato').AsString;
        QryConsulta1.ParamByName('folio').AsString := ZQRFolio.FieldByName('snumeroorden').AsString;
        QryConsulta1.Open;
        laplicalibro := QryConsulta1.FieldDefs.IndexOf('laplicalibro') > -1;
        if laplicalibro then
        begin
          QryConsulta1.Active := False;
          QryConsulta1.SQL.Clear;
          QryConsulta1.SQL.Text := 'SELECT * FROM notas_generales limit 0';
          QryConsulta1.Open;
          laplicalibro := QryConsulta1.FieldDefs.IndexOf('laplicalibro') > -1;
        end;
        QryConsulta1.Active := False;
        QryConsulta1.SQL.Clear;
        if laplicalibro then
        begin
          QryConsulta1.SQL.Text := 'select * from (SELECT mdescripcion,didfecha FROM bitacoradeactividades WHERE scontrato = :contrato AND '+
                                   'snumeroorden = :folio and lAplicaLibro = "Si" ORDER BY didfecha) t1 '+
                                  'union '+
                                  'select * from(SELECT snotageneral as mdescripcion,didfecha from notas_generales where sContrato = :contrato '+
                                  'and lAplicaLibro = "Si" order by iOrden,didfecha) t2 '+
                                  'ORDER BY didfecha ';

        end
        else
          QryConsulta1.SQL.Text := 'select * from (SELECT mdescripcion,didfecha FROM bitacoradeactividades WHERE scontrato = :contrato AND '+
                                   'snumeroorden = :folio ORDER BY didfecha) t1 '+
                                  'union '+
                                  'select * from(SELECT snotageneral as mdescripcion,didfecha from notas_generales where sContrato = :contrato '+
                                  ' order by iOrden,didfecha) t2 '+
                                  'ORDER BY didfecha ';
        QryConsulta1.ParamByName('contrato').AsString := ZQRFolio.FieldByName('scontrato').AsString;
        QryConsulta1.ParamByName('folio').AsString := ZQRFolio.FieldByName('snumeroorden').AsString;
        QryConsulta1.Open;

        QryConsulta1.First;
        while not QryConsulta1.Eof do
        begin
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          Excel.Selection.MergeCells := True;
          PFormatosExcel_H2(Excel, 16, False, 7, clBlack, 'Arial');
          Excel.Selection.Value  := QryConsulta1.FieldByName('mdescripcion').AsString;
          Excel.Selection.HorizontalAlignment := xlJustify;

          IVarAux := Length(QryConsulta1.FieldByName('mdescripcion').AsString);
          if (IVarAux mod 99) > 0 then
            IVarAux := IVarAux+99;
          IVarAux := IVarAux div 99;
          IVarAux := IVarAux * 15;
          if Excel.Rows[ifila].RowHeight < IVarAux then
            Excel.Rows[ifila].RowHeight := IVarAux;
          inc(ifila,1);
          QryConsulta1.Next;
        end;
      finally
        QryConsulta1.Free;
      end;
      //nota final
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      Excel.Selection.MergeCells := True;
      PFormatosExcel_H2(Excel, 16, False, 7, clBlack, 'Arial');
      Excel.Selection.Value  := zqrcontenido.FieldByName('mtexto3').AsString;
      Excel.Selection.HorizontalAlignment := xlJustify;

      IVarAux := Length(zqrcontenido.FieldByName('mtexto3').AsString);
      if (IVarAux mod 99) > 0 then
        IVarAux := IVarAux+99;
      IVarAux := IVarAux div 99;
      IVarAux := IVarAux * 15;
      if Excel.Rows[ifila].RowHeight < IVarAux then
        Excel.Rows[ifila].RowHeight := IVarAux;
      inc(ifila,1);


      Inc(ifila,3);
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'ATENTAMENTE';

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := 'RECIBÍ EN CONFORMIDAD';
      Inc(ifila,4);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrcontenido.FieldByName('snombref2').AsString;
      IVarAux := Length(zqrcontenido.FieldByName('snombref2').AsString);
      if (IVarAux mod 50) > 0 then
        IVarAux := IVarAux+50;
      IVarAux := IVarAux div 50;
      IVarAux := IVarAux * 11;
      if Excel.Rows[ifila].RowHeight < IVarAux then
        Excel.Rows[ifila].RowHeight := IVarAux;

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrcontenido.FieldByName('snombref3').asstring;
      IVarAux := Length(zqrcontenido.FieldByName('snombref3').AsString);
      if (IVarAux mod 33) > 0 then
        IVarAux := IVarAux+33;
      IVarAux := IVarAux div 33;
      IVarAux := IVarAux * 11;
      if Excel.Rows[ifila].RowHeight < IVarAux then
        Excel.Rows[ifila].RowHeight := IVarAux;
      inc(ifila);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrcontenido.FieldByName('scargof2').AsString;
      IVarAux := Length(zqrcontenido.FieldByName('scargof2').AsString);
      if (IVarAux mod 50) > 0 then
        IVarAux := IVarAux+50;
      IVarAux := IVarAux div 50;
      IVarAux := IVarAux * 11;
      if Excel.Rows[ifila].RowHeight < IVarAux then
        Excel.Rows[ifila].RowHeight := IVarAux;

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.Value := zqrcontenido.FieldByName('scargof3').asstring;
      IVarAux := Length(zqrcontenido.FieldByName('scargof3').AsString);
      if (IVarAux mod 50) > 0 then
        IVarAux := IVarAux+50;
      IVarAux := IVarAux div 50;
      IVarAux := IVarAux * 11;
      if Excel.Rows[ifila].RowHeight < IVarAux then
        Excel.Rows[ifila].RowHeight := IVarAux;
      {$ENDREGION}
      HojaCarta(Excel,2);
      {$ENDREGION}
    end;

    if zqrcontenido.fieldbyname('ltipo').asstring = 'OFICIO-T4' then
    begin  //oficio acta
      {$REGION 'OFICIO4'}
      iColumna := 1;
      iFila := 1;
      FormatoColumnas(Excel,icolumna);
      Excel.Columns['C:C'].ColumnWidth := 22.43;

      EncabezadoImagen(False,True);
      ifila := 2;
      EncabezadoTexto;
      Excel.activeWindow.DisplayGridlines := false;

      {$REGION 'CONTENIDO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := FormatDateTime('dd'+'" de "'+'mmmm'+'" de "'+ 'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime );
      Inc (iFila,6);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'Asunto: '+zqrcontenido.FieldByName('stexto1').AsString;
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'RESUMEN DE OBRA:';
      inc(ifila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'PERIODO: ';
      PFormatosExcel_SoloBorde(Excel);


      Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Active := False;
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select * from convenios where sContrato = :barco';
        qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
        qryconsulta1.Open;

        if qryconsulta1.RecordCount > 0 then
          Excel.Selection.Value := 'DEL '+uppercase(FormatDateTime('dd'+'" de "'+'mmmm'+'" de "'+ 'yyyy',qryconsulta1.FieldByName('dfechainicio').AsDateTime ))+' AL '+ uppercase(FormatDateTime('dd'+'" de "'+'mmmm'+'" de "'+ 'yyyy',qryconsulta1.FieldByName('dfechafinal').AsDateTime ));

      finally
        qryconsulta1.Free;
      end;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'CONTRATO: ';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := global_contrato_barco;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'DESCRIPCIÓN: ';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      svarAux := connection.contrato.FieldByName('stitulo').AsString;
      svaraux := AnsireplaceText( svaraux, 'OBJETO DEL CONTRATO:', '' );
      Excel.Selection.Value := svaraux;
      Excel.Rows[ifila].RowHeight := 49.5;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'FOLIO: ';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := zqrfolio.FieldByName('snumeroorden').asstring;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'PERIODO DE EJECUCION: ';
      inc(ifila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      if FormatDateTime('yyyy',zqrperiodo.FieldByName('minimo').AsDateTime) = FormatDateTime('yyyy',zqrperiodo.FieldByName('maximo').AsDateTime)  then
      begin
        if FormatDateTime('mm',zqrperiodo.FieldByName('minimo').AsDateTime) = FormatDateTime('mm',zqrperiodo.FieldByName('maximo').AsDateTime)  then
          Excel.Selection.Value := 'DEL '+FormatDateTime('dd',zqrperiodo.FieldByName('minimo').AsDateTime)+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime))
        else
          Excel.Selection.Value := 'DEL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm',zqrperiodo.FieldByName('minimo').AsDateTime))+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime));
      end
      else
        Excel.Selection.Value := 'DEL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('minimo').AsDateTime))+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime));

      //Excel.Selection.Value :=  'DEL '+uppercase(FormatDateTime('dd'+'" de "'+'mmmm'+'" de "'+ 'yyyy',zqrperiodo.FieldByName('minimo').AsDateTime ))+' AL '+ uppercase(FormatDateTime('dd'+'" de "'+'mmmm'+'" de "'+ 'yyyy',zqrperiodo.FieldByName('maximo').AsDateTime ));
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,2);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'RESUMEN DE COSTOS ';
      inc(ifila,2);

      {$REGION 'ENCABEZADO CUADRO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PARTIDA ANEXO C';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'DESCRIPCIÓN';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'IMPORTE TOTAL';
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := '';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := '';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'M.N';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'U.S.D';
      PFormatosExcel_SoloBorde(Excel);
      {$ENDREGION}

      {$REGION 'MOVIMIENTO DE BARCO'}
      inc(ifila,1);
      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.Active := False;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text :=  'select tm.sidtipomovimiento,tm.sdescripcion,tm.stipo,ifnull(tf.factor,0),ifnull(tm.dventamn,0),ifnull(tm.dventadll,0),  round(ifnull((tm.dVentaMN*tf.factor),0),2) as smn,'+
                                  'round(ifnull((tm.dVentadll*tf.factor),0),2) as sdll from tiposdemovimiento tm left join '+
                                  '(select mb.sClasificacion,sum(mf.sfactor) as factor from movimientosdeembarcacion mb inner join '+
                                  'movimientosxfolios mf on (mf.iiddiario = mb.iiddiario and mf.scontrato = :barco and sNumeroOrden = :contrato and mf.sfolio = :folio) '+
                                  'where mb.sContrato = :barco group by mb.sClasificacion) tf '+
                                  'on (tf.sclasificacion = sIdTipoMovimiento) where tm.sClasificacion = "Movimiento de Barco" and tm.sIdTipoMovimiento <> "s/n"';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
        qryconsulta1.Open;
        ivarAux := ifila;
        while not qryconsulta1.eof do
        begin
          if (qryconsulta1.FieldByName('smn').AsFloat > 0) or (qryconsulta1.FieldByName('smn').AsFloat > 0) or (AnsiMatchStr( qryconsulta1.FieldByName('sidtipomovimiento').AsString, ['1.1','1.2','1.3'])) then
          begin
            Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := qryconsulta1.FieldByName('sidtipomovimiento').AsString;
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            svar := qryconsulta1.FieldByName('sdescripcion').AsString;
            svar := AnsireplaceText( svar, 'en operacion ', '' );
            svar:= AnsireplaceText( svar, 'EN OPERACIÓN ', '' );
            Excel.Selection.Value := svar;

            PFormatosExcel_SoloBorde(Excel);

            //movimientosxfolio.factor * tiposdemovimeinto.dventamn  , tiposdemovimiento.dventadll
            Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('smn').AsFloat;
            PFormatosExcel_SoloBorde(Excel);
            Excel.Selection.NumberFormat := '#,##0.00';

            Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('sdll').AsFloat;
            PFormatosExcel_SoloBorde(Excel);
            Excel.Selection.NumberFormat := '#,##0.00';

            IVarAux2 := Length(qryconsulta1.FieldByName('sdescripcion').AsString);
            if (IVarAux2 mod 34) > 0 then
              IVarAux2 := IVarAux2+34;
            IVarAux2 := IVarAux2 div 34;
            IVarAux2 := IVarAux2 * 11;
            if Excel.Rows[ifila].RowHeight < IVarAux2 then
              Excel.Rows[ifila].RowHeight := IVarAux2;

            inc(ifila,1);
          end;

          qryconsulta1.Next;
        end;
      finally
        qryconsulta1.Free;
      end;
      {$ENDREGION}

      {$REGION 'PERSONAL'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value :='ANEXO C-1 "PERSONAL OPT."';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PERSONAL';
      PFormatosExcel_SoloBorde(Excel);

      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.Active := False;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select sum(t1.smn) as smn, sum(t1.sdll) as sdll from  (select (round(ifnull(round(sum(bp.dCantHHgenerador),2),0)*ifnull(p.dVentaMN,0),2)) as smn, (round(ifnull(round(sum(bp.dCantHHgenerador),2),0)*ifnull(p.dVentadll,0),2)) as sdll '+
                                 'from personal p inner join bitacoradepersonal bp on (p.sidpersonal = bp.sIdPersonal and bp.scontrato = :Contrato  '+
                                 'and bp.snumeroorden = :Folio ) where p.sContrato = :barco group by p.sIdPersonal order by p.iItemOrden) t1';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
        qryconsulta1.Open;

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.NumberFormat := '#,##0.00';
        if qryconsulta1.RecordCount > 0 then
          Excel.selection.value := qryconsulta1.FieldByName('smn').AsFloat
        else
          Excel.selection.value := 0;
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.NumberFormat := '#,##0.00';
        if qryconsulta1.RecordCount > 0 then
          Excel.selection.value := qryconsulta1.FieldByName('sdll').AsFloat
        else
          Excel.selection.value := 0;
        PFormatosExcel_SoloBorde(Excel);
      finally
        qryconsulta1.Free;
      end;
      inc(ifila,1);
      {$ENDREGION}

      {$REGION 'EQUIPO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value :='ANEXO C-2 "EQUIPO OPT."';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'EQUIPO';
      PFormatosExcel_SoloBorde(Excel);

      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.Active := False;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select sum(t1.smn) as smn, sum(t1.sdll) as sdll from  (select (round(ifnull(round(sum(be.dCantHHgenerador),2),0)*ifnull(e.dVentaMN,0),2)) as smn, (round(ifnull(round(sum(be.dCantHHgenerador),2),0)*ifnull(e.dVentadll,0),2)) as sdll  '+
                                 'from equipos e inner join bitacoradeequipos be on (e.sidequipo = be.sIdequipo and be.scontrato = :Contrato '+
                                 'and be.snumeroorden = :Folio ) where e.sContrato = :barco group by e.sIdequipo order by e.iItemOrden) t1';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
        qryconsulta1.Open;

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.NumberFormat := '#,##0.00';
        if qryconsulta1.RecordCount > 0 then
          Excel.selection.value := qryconsulta1.FieldByName('smn').AsFloat
        else
          Excel.selection.value := 0;
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.NumberFormat := '#,##0.00';
        if qryconsulta1.RecordCount > 0 then
          Excel.selection.value := qryconsulta1.FieldByName('sdll').AsFloat
        else
          Excel.selection.value := 0;
        PFormatosExcel_SoloBorde(Excel);
      finally
        qryconsulta1.Free;
      end;

      inc(ifila,1);
      {$ENDREGION}

      {$REGION 'PERNOCTAS'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value :='ANEXO C-3 "SERV. DE HOTEL."';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'SERVICIOS DE HOTELERIA';
      PFormatosExcel_SoloBorde(Excel);


      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.Active := False;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select round(ifnull(sum(tcalc.tmn),0),2) as smn, round(ifnull(sum(tcalc.tdll),0),2) as sdll from '+
                                 '(select cu.sidcuenta, cu.dventamn,cu.dventadll,tt.tpern,(round(ifnull(tt.tpern,0),2)*ifnull(cu.dventamn,0)) as tmn,(round(ifnull(tt.tpern,0),2)*ifnull(cu.dventadll,0)) as tdll from cuentas cu inner join '+
                                 '(select prn.sidcuenta,(ifnull(prn.cp,0) + ifnull(prd.cd,0))+ifnull(pre.cpe,0) as tpern from (select c.sidcuenta, sum(p.dCantHHgenerador) as cp  '+
                                 'from cuentas c inner join bitacoradepersonal p on (p.stipopernocta = c.sidcuenta and p.snumeroorden = :folio and p.scontrato = :Contrato ) group by c.sIdCuenta) prn '+
                                 'left join '+
                                 '(select c2.sidcuenta, sum(bp.dcantidad) as cd from cuentas c2 inner join bitacoradepernocta bp on (c2.sidcuenta = if(length(bp.sIdCuenta) '+
                                 '= 0,"4.1",bp.sIdCuenta) and bp.sContrato = :Contrato2 and bp.sNumeroOrden = :folio2) group by c2.sIdCuenta) prd on (prn.sidcuenta = prd.sidcuenta) '+
                                 '  left join '+
                                 '(select c3.sidcuenta, sum(pe.dCantidad) as cpe '+
                                 'from cuentas c3 inner join bitacoradepersonal_cuadre pe on (pe.stipopernocta = c3.sidcuenta and pe.snumeroorden = :folio and pe.scontrato = :Contrato ) group by c3.sIdCuenta) pre '+
                                 'on (prn.sidcuenta = pre.sidcuenta) '+
                                 ') tt '+
                                 'on (tt.sidcuenta = cu.sidcuenta)) tcalc';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.ParamByName('contrato2').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio2').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.Open;

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.NumberFormat := '#,##0.00';
        if qryconsulta1.RecordCount > 0 then
          Excel.selection.value := qryconsulta1.FieldByName('smn').AsFloat
        else
          Excel.selection.value := 0;
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.NumberFormat := '#,##0.00';
        if qryconsulta1.RecordCount > 0 then
          Excel.selection.value := qryconsulta1.FieldByName('sdll').AsFloat
        else
          Excel.selection.value := 0;
        PFormatosExcel_SoloBorde(Excel);
      finally
        qryconsulta1.Free;
      end;
      inc(ifila,1);
      {$ENDREGION}

      {$REGION 'TOTALES'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := '';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'TOTALES';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.NumberFormat := '#,##0.00';
      for ivaraux2 := ivaraux to ifila-1 do
      begin
        if length(trim(Excel.Selection.formula)) = 0 then
          Excel.Selection.formula := '='+ColumnaNombre(8)+inttostr(ivaraux2)
        else
          Excel.Selection.formula := Excel.Selection.formula+'+'+ColumnaNombre(8)+inttostr(ivaraux2);
      end;

      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.NumberFormat := '#,##0.00';
      for ivaraux2 := ivaraux to ifila-1 do
      begin
        if length(trim(Excel.Selection.formula)) = 0 then
          Excel.Selection.formula := '='+ColumnaNombre(9)+inttostr(ivaraux2)
        else
          Excel.Selection.formula := Excel.Selection.formula+'+'+ColumnaNombre(9)+inttostr(ivaraux2);
      end;

      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,3);
      {$ENDREGION}

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'ATENTAMENTE';


      Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'RECIBÍ EN CONFORMIDAD';

      inc(ifila,1);
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila+2)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := zqrcontenido.FieldByName('snombref1').AsString+#10+zqrcontenido.FieldByName('scargof1').AsString;


      Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila+2)].Select;
      PFormatosExcel_H2(Excel, 16, true, 8, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := zqrcontenido.FieldByName('snombref2').AsString+#10+zqrcontenido.FieldByName('scargof2').AsString;
      {$ENDREGION}

      HojaCarta(Excel);
      {$ENDREGION}
    end;

    if zqrcontenido.fieldbyname('ltipo').asstring = 'OFICIO-T5' then
    begin   //obra
      {$REGION 'OFICIO5'}
      icolumna := 1;
      ifila := 1;
      FormatoColumnas(Excel,icolumna,2);
      EncabezadoImagen(True,True,2);
      ifila := 6;

      sVarAux := '= SUM(';
      sVarAux2 := '= SUM(';

      {$REGION 'CUADRO SUPERIOR'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'CONTRATO';
      PFormatosExcel_SoloBorde(Excel);
      Excel.Rows[ifila].RowHeight := 15;

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := global_contrato_barco;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'FOLIO';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := zqrfolio.FieldByName('snumeroorden').AsString;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);


      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila+1)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'DESCRIPCIÓN:';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila+1)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      svar := Connection.contrato.FieldByName('sTitulo').AsString;
      svar := AnsireplaceText( svar, 'objeto del contrato:', '' );
      svar := AnsireplaceText( svar, 'OBJETO DEL CONTRATO:', '' );
      Excel.Selection.Value := svar;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'OBRA';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := zqrfolio.FieldByName('mdescripcion').AsString;
      PFormatosExcel_SoloBorde(Excel);
      Excel.Rows[ifila].RowHeight := 54.75;
      inc(ifila,1);
      Excel.Rows[ifila].RowHeight := 15;
      
      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'LOCALIZACION';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'PLATAFORMA: '+zqrfolio.FieldByName('sidplataforma').AsString;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,2);
      {$ENDREGION}

      {$REGION 'ENCABEZADO CUADRO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'MOVIMIENTO DE EMBARCACION';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PARTIDA';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'DESCRIPCIÓN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'CLAS.';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'CANTIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PU MN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PU USD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'IMP MN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'IMP USD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);
      {$ENDREGION}

      {$REGION 'MOVIMIENTOS DE BARCO'}
      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.Active := False;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text :=  'select tm.sidtipomovimiento,tm.sdescripcion,tm.stipo,ifnull(tf.factor,0) as factor,ifnull(tm.dventamn,0) as dventamn,ifnull(tm.dventadll,0) as dventadll,  ifnull((tm.dVentaMN*tf.factor),0) as smn,'+
                                  'ifnull((tm.dVentadll*tf.factor),0) as sdll from tiposdemovimiento tm left join '+
                                  '(select mb.sClasificacion,sum(mf.sfactor) as factor from movimientosdeembarcacion mb inner join '+
                                  'movimientosxfolios mf on (mf.iiddiario = mb.iiddiario and mf.scontrato = :barco and sNumeroOrden = :contrato and mf.sfolio = :folio) '+
                                  'where mb.sContrato = :barco group by mb.sClasificacion) tf '+
                                  'on (tf.sclasificacion = sIdTipoMovimiento)  where tm.sClasificacion = "Movimiento de barco"';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
        qryconsulta1.Open;
        ivarAux := ifila;
        while not qryconsulta1.eof do
        begin
          if (qryconsulta1.FieldByName('smn').asfloat > 0) or (qryconsulta1.FieldByName('sdll').asfloat > 0) or (AnsiMatchStr( qryconsulta1.FieldByName('sidtipomovimiento').AsString, ['1.1','1.2','1.3'])) then
          begin

            Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := qryconsulta1.FieldByName('sidtipomovimiento').AsString;
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlleft;

            if AnsiMatchStr( qryconsulta1.FieldByName('sidtipomovimiento').AsString, ['1.1','1.2']) then
            begin
              svar := qryconsulta1.FieldByName('sdescripcion').AsString;
              svar := AnsireplaceText( svar, 'en operacion ', '' );
              svar := AnsireplaceText( svar, 'EN OPERACIÓN ', '' );
              Excel.Selection.Value := svar;
            end
            else
            begin
              Excel.Selection.Value := qryconsulta1.FieldByName('sdescripcion').AsString;
            end;

           { if AnsiMatchStr( qryconsulta1.FieldByName('sidtipomovimiento').AsString, ['1.1','1.2']) then
              Excel.Selection.Value := 'BARCO '+ qryconsulta1.FieldByName('sdescripcion').AsString
            else
              Excel.Selection.Value := qryconsulta1.FieldByName('sdescripcion').AsString;
                 }
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := qryconsulta1.FieldByName('stipo').AsString;
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('factor').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00000000';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('dventamn').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('dventadll').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.formula := '=ROUND((RC[-2]*RC[-3]),2)';
            //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;  valor real de consulta
            Excel.Selection.NumberFormat :=  '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
            //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            IVarAux2 := Length(qryconsulta1.FieldByName('sdescripcion').AsString);
            if (IVarAux2 mod 34) > 0 then
              IVarAux2 := IVarAux2+34;
            IVarAux2 := IVarAux2 div 34;
            IVarAux2 := IVarAux2 * 11;
            if Excel.Rows[ifila].RowHeight < IVarAux2 then
              Excel.Rows[ifila].RowHeight := IVarAux2;

            inc(ifila,1);
          end;
          qryconsulta1.Next;
        end;
      finally
        qryconsulta1.Free;
      end;
      {$ENDREGION}

      {$REGION 'TOTALES'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'IMPORTE BARCO:';

      Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula := '= SUM('+ColumnaNombre(11)+inttostr(ivaraux)+':'+ColumnaNombre(11)+inttostr(ifila-1)+')';
      Excel.Selection.NumberFormat :=  '$#,##0.00';
      sVarAux := sVarAux + ColumnaNombre(11)+inttostr(ifila);

      Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula :=  '= SUM('+ColumnaNombre(12)+inttostr(ivaraux)+':'+ColumnaNombre(12)+inttostr(ifila-1)+')';
      Excel.Selection.NumberFormat := '$#,##0.00';
      Excel.Rows[ifila].RowHeight := 24.75;
      sVarAux2 := sVarAux2 + ColumnaNombre(12)+inttostr(ifila);
      {$ENDREGION}

      inc(ifila,2);

      {$REGION 'ENCABEZADO CUADRO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PERSONAL';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PARTIDA';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'DESCRIPCIÓN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'UNIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'CANTIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PU MN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PU USD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'IMP MN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'IMP USD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);
      {$ENDREGION}

      {$REGION 'PERSONAL'}
      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.Active := False;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select p.sidpersonal,p.sdescripcion,p.smedida,ifnull(sum(bp.dCantHHgenerador),0)as cantidad,'+
                                 'ifnull(p.dVentaMN,0) as pumn,ifnull(p.dVentadll,0) as pudll,ifnull(sum(bp.dCantHHgenerador),0)*ifnull(p.dVentaMN,0) as smn, ifnull(sum(bp.dCantHHgenerador),0)*ifnull(p.dVentadll,0) as sdll '+
                                 'from personal p inner join bitacoradepersonal bp on (p.sidpersonal = bp.sIdPersonal and bp.scontrato = :Contrato '+
                                 'and bp.snumeroorden = :Folio ) where p.sContrato = :barco group by p.sIdPersonal order by p.iItemOrden';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
        qryconsulta1.Open;
        ivarAux := ifila;
        while not qryconsulta1.eof do
        begin
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('sidpersonal').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlleft;
          Excel.Selection.Value := qryconsulta1.FieldByName('sdescripcion').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('smedida').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.formula := '=ROUND(('+qryconsulta1.FieldByName('cantidad').asstring+'),2)';
          //Excel.Selection.Value := qryconsulta1.FieldByName('cantidad').asstring;
          Excel.Selection.NumberFormat :=  '#,##0.00000000';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.Value := qryconsulta1.FieldByName('pumn').asfloat;
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.Value := qryconsulta1.FieldByName('pudll').asfloat;
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.formula := '=ROUND((RC[-2]*RC[-3]),2)';
          //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
          Excel.Selection.NumberFormat :=  '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
          Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          IVarAux2 := Length(qryconsulta1.FieldByName('sdescripcion').AsString);
          if (IVarAux2 mod 34) > 0 then
            IVarAux2 := IVarAux2+34;
          IVarAux2 := IVarAux2 div 34;
          IVarAux2 := IVarAux2 * 11;
          if Excel.Rows[ifila].RowHeight < IVarAux2 then
            Excel.Rows[ifila].RowHeight := IVarAux2;

          inc(ifila,1);
          qryconsulta1.Next;
        end;
      finally
        qryconsulta1.Free;
      end;
      {$ENDREGION}

      {$REGION 'TOTALES'}
      Excel.Rows[ifila].RowHeight := 24.75;
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'IMPORTE PERSONAL:';

      Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula := '= SUM('+ColumnaNombre(11)+inttostr(ivaraux)+':'+ColumnaNombre(11)+inttostr(ifila-1)+')';
      Excel.Selection.NumberFormat :=  '$#,##0.00';
      sVarAux := sVarAux +'+'+ ColumnaNombre(11)+inttostr(ifila);

      Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula :=  '= SUM('+ColumnaNombre(12)+inttostr(ivaraux)+':'+ColumnaNombre(12)+inttostr(ifila-1)+')';
      Excel.Selection.NumberFormat := '$#,##0.00';
      sVarAux2 := sVarAux2 +'+'+ ColumnaNombre(12)+inttostr(ifila);
      Excel.Rows[ifila].RowHeight := 24.75;
      {$ENDREGION}

      inc(ifila,2);

      {$REGION 'ENCABEZADO CUADRO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'EQUIPO';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PARTIDA';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'DESCRIPCIÓN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'UNIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'CANTIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PU MN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PU USD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'IMP MN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'IMP USD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);
      {$ENDREGION}

      {$REGION 'EQUIPO'}
      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.Active := False;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select e.sIdEquipo,e.sDescripcion,e.sMedida,ifnull(sum(be.dcantHHgenerador),0) as cantidad,'+
                                 'ifnull(e.dventamn,0) as pumn,ifnull(e.dventadll,0) as pudll , (ifnull(e.dventamn,0)*ifnull(sum(be.dcantHHgenerador),0)) '+
                                 'as smn,(ifnull(e.dventadll,0)*ifnull(sum(be.dcantHHgenerador),0)) as sdll '+
                                 'from equipos e  inner join bitacoradeequipos be on '+
                                 '(be.sIdEquipo = e.sIdEquipo and be.sContrato = :Contrato '+
                                 'and be.sNumeroOrden = :folio) where e.sContrato = :barco group by e.sIdEquipo ORDER BY e.iItemOrden';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
        qryconsulta1.Open;
        ivarAux := ifila;
        while not qryconsulta1.eof do
        begin
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('sidequipo').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlleft;
          Excel.Selection.Value := qryconsulta1.FieldByName('sdescripcion').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('smedida').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.formula := '=ROUND(('+qryconsulta1.FieldByName('cantidad').asstring+'),2)';
          //Excel.Selection.Value := qryconsulta1.FieldByName('cantidad').asstring;
          Excel.Selection.NumberFormat :=  '#,##0.00000000';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.Value := qryconsulta1.FieldByName('pumn').asfloat;
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.Value := qryconsulta1.FieldByName('pudll').asfloat;
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.formula := '=ROUND((RC[-2]*RC[-3]),2)';
          //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
          Excel.Selection.NumberFormat :=  '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
          Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          IVarAux2 := Length(qryconsulta1.FieldByName('sdescripcion').AsString);
          if (IVarAux2 mod 38) > 0 then
            IVarAux2 := IVarAux2+38;
          IVarAux2 := IVarAux2 div 38;
          IVarAux2 := IVarAux2 * 11;
          if Excel.Rows[ifila].RowHeight < IVarAux2 then
            Excel.Rows[ifila].RowHeight := IVarAux2;

          inc(ifila,1);
          qryconsulta1.Next;
        end;
      finally
        qryconsulta1.Free;
      end;
      {$ENDREGION}

      {$REGION 'TOTALES'}
      Excel.Rows[ifila].RowHeight := 24.75;
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'IMPORTE EQUIPOS:';

      Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula := '= SUM('+ColumnaNombre(11)+inttostr(ivaraux)+':'+ColumnaNombre(11)+inttostr(ifila-1)+')';
      Excel.Selection.NumberFormat :=  '$#,##0.00';
      sVarAux := sVarAux +'+'+ ColumnaNombre(11)+inttostr(ifila);

      Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula :=  '= SUM('+ColumnaNombre(12)+inttostr(ivaraux)+':'+ColumnaNombre(12)+inttostr(ifila-1)+')';
      Excel.Selection.NumberFormat := '$#,##0.00';
      sVarAux2 := sVarAux2 +'+'+ ColumnaNombre(12)+inttostr(ifila);

      Excel.Rows[ifila].RowHeight := 24.75;
      {$ENDREGION}

      inc(ifila,2);

      {$REGION 'ENCABEZADO CUADRO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PERNOCTAS';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PARTIDA';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'DESCRIPCIÓN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'UNIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'CANTIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PU MN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'PU USD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'IMP MN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'IMP USD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);
      {$ENDREGION}

      {$REGION 'PERNOCTAS'}
      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.Active := False;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select cu.sidcuenta,cu.sDescripcion,cu.sMedida ,ifnull(tt.tpern,0) as cantidad, cu.dventamn as pumn,cu.dventadll as pudll, '+
                                 '(ifnull(tt.tpern,0)*ifnull(cu.dventamn,0)) as smn,(ifnull(tt.tpern,0)*ifnull(cu.dventadll,0)) as sdll from cuentas cu left join  '+
                                 '(select prn.sidcuenta,(ifnull(prn.cp,0) + ifnull(prd.cd,0)+ifnull(pre.cpe,0)) as tpern from (select c.sidcuenta, sum(p.dCantHHgenerador) as cp  '+
                                 'from cuentas c inner join bitacoradepersonal p on (p.stipopernocta = c.sidcuenta and p.snumeroorden = :folio and p.scontrato = :Contrato ) group by c.sIdCuenta) prn '+
                                 'left join '+
                                 '(select c2.sidcuenta, sum(bp.dcantidad) as cd from cuentas c2 inner join bitacoradepernocta bp on (c2.sidcuenta = if(length(bp.sIdCuenta) '+
                                 '= 0,"4.1",bp.sIdCuenta) and bp.sContrato = :Contrato2 and bp.sNumeroOrden = :folio2) group by c2.sIdCuenta) prd  '+
                                 'on (prn.sidcuenta = prd.sidcuenta)  '+
                                 '  left join '+
                                 '(select c3.sidcuenta, sum(pe.dCantidad) as cpe '+
                                 'from cuentas c3 inner join bitacoradepersonal_cuadre pe on (pe.stipopernocta = c3.sidcuenta and pe.snumeroorden = :folio and pe.scontrato = :Contrato ) group by c3.sIdCuenta) pre '+
                                 'on (prn.sidcuenta = pre.sidcuenta) '+
                                 ') tt '+
                                 'on (tt.sidcuenta = cu.sidcuenta)';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.ParamByName('contrato2').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio2').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.Open;
        ivarAux := ifila;
        while not qryconsulta1.eof do
        begin
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('sidcuenta').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlleft;
          Excel.Selection.Value := qryconsulta1.FieldByName('sdescripcion').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('smedida').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          //Excel.Selection.Value := qryconsulta1.FieldByName('cantidad').asstring;
          Excel.Selection.formula := '=ROUND(('+qryconsulta1.FieldByName('cantidad').asstring+'),2)';
          Excel.Selection.NumberFormat :=  '#,##0.00000000';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.Value := qryconsulta1.FieldByName('pumn').asfloat;
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.Value := qryconsulta1.FieldByName('pudll').asfloat;
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
          Excel.Selection.formula := '=ROUND((RC[-2]*RC[-3]),2)';
          Excel.Selection.NumberFormat :=  '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
          Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          IVarAux2 := Length(qryconsulta1.FieldByName('sdescripcion').AsString);
          if (IVarAux2 mod 34) > 0 then
            IVarAux2 := IVarAux2+34;
          IVarAux2 := IVarAux2 div 34;
          IVarAux2 := IVarAux2 * 11;
          if Excel.Rows[ifila].RowHeight < IVarAux2 then
            Excel.Rows[ifila].RowHeight := IVarAux2;

          inc(ifila,1);
          qryconsulta1.Next;
        end;
      finally
        qryconsulta1.Free;
      end;
      {$ENDREGION}

      {$REGION 'TOTALES'}
      Excel.Rows[ifila].RowHeight := 24.75;
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'IMPORTE PERNOCTAS:';

      Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula := '= SUM('+ColumnaNombre(11)+inttostr(ivaraux)+':'+ColumnaNombre(11)+inttostr(ifila-1)+')';
      Excel.Selection.NumberFormat :=  '$#,##0.00';
      sVarAux := sVarAux +'+'+ ColumnaNombre(11)+inttostr(ifila);

      Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula :=  '= SUM('+ColumnaNombre(12)+inttostr(ivaraux)+':'+ColumnaNombre(12)+inttostr(ifila-1)+')';
      Excel.Selection.NumberFormat := '$#,##0.00';
      sVarAux2 := sVarAux2 +'+'+ ColumnaNombre(12)+inttostr(ifila);

      Excel.Rows[ifila].RowHeight := 24.75;
      {$ENDREGION}

      inc(ifila,2);
      svaraux := svaraux+')';
      svaraux2 := svaraux2+')';

      {$REGION 'ENCABEZADO CUADRO'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'MATERIAL';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'TRAZABILIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'DESCRIPCIÓN';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'UNIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'CANTIDAD';
      Excel.selection.interior.colorindex := 16;
      PFormatosExcel_SoloBorde(Excel);

      inc(ifila,1);
      {$ENDREGION}

      {$REGION 'MATERIAL'}
      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Connection := connection.zConnection;
        qryconsulta1.Active := False;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select bs.sTrazabilidad,i.mDescripcion,i.sMedida,ifnull(sum(bs.dCantidad),0) as cantidad from bitacoradesalida bs '+
                                 'inner join insumos i  '+
                                 'on (  bs.sIdInsumo = i.sIdInsumo and  bs.sTrazabilidad = i.strazabilidad and i.sContrato = :Barco ) '+
                                 'where bs.sContrato = :contrato and bs.sNumeroOrden = :Folio group by bs.sIdInsumo, bs.sTrazabilidad';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
        qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
        qryconsulta1.Open;
        ivarAux := ifila;
        while not qryconsulta1.eof do
        begin
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('strazabilidad').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('mdescripcion').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('smedida').AsString;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := qryconsulta1.FieldByName('cantidad').asstring;
          Excel.Selection.NumberFormat := '#,##0.00';
          PFormatosExcel_SoloBorde(Excel);

          IVarAux2 := Length(qryconsulta1.FieldByName('mdescripcion').AsString);
          if (IVarAux2 mod 34) > 0 then
            IVarAux2 := IVarAux2+34;
          IVarAux2 := IVarAux2 div 34;
          IVarAux2 := IVarAux2 * 11;
          if Excel.Rows[ifila].RowHeight < IVarAux2 then
            Excel.Rows[ifila].RowHeight := IVarAux2;

          inc(ifila,1);
          qryconsulta1.Next;
        end;
      finally
        qryconsulta1.Free;
      end;
      {$ENDREGION}

      {$REGION 'TOTAL'}
      inc(ifila,3);

      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := 'COSTO TOTAL DE LA ACTIVIDAD:';

      Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula := svaraux;
      Excel.Selection.NumberFormat := '#,##0.00';


      Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.formula := svaraux2;
      Excel.Selection.NumberFormat := '#,##0.00';
      {$ENDREGION}

      hojacarta(excel,2);
      {$ENDREGION}
    end;

    if zqrcontenido.FieldByName('ltipo').AsString = 'OFICIO-T6' then
    begin  //nota de campo
      icolumna := 1;
      ifila := 1;
      NotaCampoExcel(excel,hoja,libro);
    end;

    if zqrcontenido.FieldByName('ltipo').AsString = 'OFICIO-T7' then
    begin   //desgloce de costos
      {$REGION 'OFICIO 7'}
      icolumna := 1;
      ifila := 1;
      FormatoColumnas(Excel,icolumna,2);
      EncabezadoImagen(True,True,2);
      ifila := 6;

      sVarAux := '= SUM(';
      sVarAux2 := '= SUM(';

      {$REGION 'CUADRO SUPERIOR'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlright;
      Excel.Selection.Value := 'CONTRATO';
      PFormatosExcel_SoloBorde(Excel);
      Excel.Rows[ifila].RowHeight := 15;

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlcenter;
      Excel.Selection.Value := global_contrato_barco;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'FOLIO';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := zqrfolio.FieldByName('snumeroorden').AsString;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,1);


      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila+1)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'DESCRIPCIÓN:';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila+1)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;

      svarAux := connection.contrato.FieldByName('stitulo').AsString;
      svaraux := AnsireplaceText( svaraux, 'OBJETO DEL CONTRATO:', '' );
      svaraux := AnsireplaceText( svaraux, 'objeto del contrato:', '' );
      Excel.Selection.Value := svaraux;
      //Excel.Selection.Value := Connection.contrato.FieldByName('sTitulo').AsString;
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'OBRA';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := zqrfolio.FieldByName('mdescripcion').AsString;
      PFormatosExcel_SoloBorde(Excel);
      Excel.Rows[ifila].RowHeight := 54.75;
      inc(ifila,1);
      Excel.Rows[ifila].RowHeight := 15;
      
      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'LOCALIZACION';
      PFormatosExcel_SoloBorde(Excel);

      Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
      Excel.Selection.HorizontalAlignment := xlleft;
      Excel.Selection.Value := 'PLATAFORMA: '+zqrfolio.FieldByName('sidplataforma').AsString;
      PFormatosExcel_SoloBorde(Excel);
      inc(ifila,2);
      {$ENDREGION}

      {$REGION 'OBTENER FECHAS MINIMAS Y MAXIMAS'}
      dtvaraux1 := now+500;
      dtvaraux2 := strtodatetime('01/01/1800');

      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Active:=False;
        qryconsulta1.connection := connection.zConnection;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select min(didfecha) as minimo, max(didfecha) as maximo from bitacoradeactividades where sContrato = :Contrato and sNumeroOrden = :Folio ';
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').AsString;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').AsString;
        qryconsulta1.Open;

        if qryconsulta1.RecordCount = 1 then
        begin
          if qryconsulta1.FieldByName('minimo').AsDateTime < dtvaraux1  then
            dtvaraux1 := qryconsulta1.FieldByName('minimo').AsDateTime;
          if qryconsulta1.FieldByName('maximo').AsDateTime > dtvaraux2  then
            dtvaraux2 := qryconsulta1.FieldByName('maximo').AsDateTime;
        end;
      finally
        qryconsulta1.Free;
      end;

      qryconsulta1 := tzreadonlyquery.Create(nil);
      try
        qryconsulta1.Active:=False;
        qryconsulta1.connection := connection.zConnection;
        qryconsulta1.SQL.Clear;
        qryconsulta1.SQL.Text := 'select min(mb.didfecha) as minimo,max(mb.didfecha) as maximo from movimientosdeembarcacion mb inner join ' +
                                 'movimientosxfolios mf on (mf.iiddiario = mb.iiddiario and mf.scontrato = :barco and sNumeroOrden = :contrato and mf.sfolio = :folio) ' +
                                 'where mb.sContrato = :barco ';
        qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
        qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').AsString;
        qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').AsString;
        qryconsulta1.Open;

        if qryconsulta1.RecordCount = 1 then
        begin{
          if qryconsulta1.FieldByName('minimo').AsDateTime < dtvaraux1  then
            showmessage('Revise los movimientos de embarcación ya que hay folios en fechas inferiores a los reportados.');
          if qryconsulta1.FieldByName('maximo').AsDateTime > dtvaraux2  then
            showmessage('Revise los movimientos de embarcación ya que hay folios en fechas posteriores a los reportados.');
             }
        end;
      finally
        qryconsulta1.Free;
      end;

      if dtvaraux1 > dtvaraux2 then
        dtvaraux1 := dtvaraux2;

      if dtvaraux2 = strtodatetime('01/01/1800') then
        if dtvaraux2 = dtvaraux1 then
        begin
          dtvaraux1 := now;
          dtvaraux2 := now;
        end
        else
          dtvaraux2 := dtVarAux1;
      {$ENDREGION}

      DecodeDate(dtvaraux1, Año, Mes, Dia);
      Dia := 1;
      dtvaraux1 := EncodeDate(año,mes,dia);
      DecodeDate(dtvaraux2, Año2, Mes2, Dia2);
      Dia2 := DaysInMonth( dtVarAux2 );
      dtvaraux2 := EncodeDate(año2,mes2,dia2);
     { if (mes <> mes2) or (año <> año2) then
        ImprimirTotal := True
      else
        Imprimirtotal := False;}

      Ciclos := 0;
      while dtvaraux1 < dtvaraux2 do
      begin
        rgFecha1 := IncDay( dtvaraux1, strtoint(vartostr(DaysInMonth( dtVarAux1 )))-1 );
        sVarAux := '= ';
        sVarAux2 := '= ';
        try
          {$REGION 'Imprimir cuadro?'}
          qryconsulta1 := tzreadonlyquery.Create(nil);
          try
            dVaraux1 := 0;
            dvaraux2 := 0;
            qryconsulta1.Connection := connection.zConnection;
            qryconsulta1.Active := False;
            qryconsulta1.SQL.Clear;
            qryconsulta1.SQL.Text :=  'select tm.sidtipomovimiento,tm.sdescripcion,tm.smedida,tm.stipo,ifnull(tf.factor,0) as factor,'+
                                      'ifnull(tm.dventamn,0) as dventamn,ifnull(tm.dventadll,0) as dventadll,  ifnull((tm.dVentaMN*tf.factor),0) as smn,  '+
                                      'ifnull((tm.dVentadll*tf.factor),0) as sdll from tiposdemovimiento tm left join  '+
                                      '(select mb.sClasificacion,sum(mf.sfactor) as factor from movimientosdeembarcacion mb inner join '+
                                      'movimientosxfolios mf on (mf.iiddiario = mb.iiddiario and mf.scontrato = :barco and sNumeroOrden = :contrato and mf.sfolio = :folio) '+
                                      'where mb.sContrato = :barco and mb.dIdFecha >= :Fechai and mb.didfecha <= :Fechaf group by mb.sClasificacion  ) tf '+
                                      'on (tf.sclasificacion = sIdTipoMovimiento) where tm.sClasificacion = "Movimiento de Barco" and tm.sIdTipoMovimiento <> "s/n"';
            qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
            qryconsulta1.ParamByName('fechai').AsDateTime := dtvaraux1;
            qryconsulta1.ParamByName('fechaf').AsDateTime := rgFecha1;
            qryconsulta1.Open;
            qryconsulta1.First;
            while not qryconsulta1.eof do
            begin
              dVaraux1 := dvaraux1 + qryconsulta1.FieldByName('smn').AsFloat;
              dVaraux2 := dvaraux2 + qryconsulta1.FieldByName('sdll').AsFloat;
              qryconsulta1.Next;
            end;

            qryconsulta1.Active := False;
            qryconsulta1.SQL.Clear;
            qryconsulta1.SQL.Text :=  'select sum(t1.smn) as smn, sum(t1.sdll) as sdll from  (select (round(ifnull(round(sum(bp.dCantHHgenerador),2),0)*ifnull(p.dVentaMN,0),2)) as smn, (round(ifnull(round(sum(bp.dCantHHgenerador),2),0)*ifnull(p.dVentadll,0),2)) as sdll '+
                                      'from personal p inner join bitacoradepersonal bp on (p.sidpersonal = bp.sIdPersonal and bp.scontrato = :Contrato '+
                                      'and bp.snumeroorden = :Folio ) where p.sContrato = :barco and  bp.didfecha >= :Fechai and bp.dIdFecha  <= :Fechaf group by p.sIdPersonal  order by p.iItemOrden) t1';
            qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
            qryconsulta1.ParamByName('fechai').AsDateTime := dtvaraux1;
            qryconsulta1.ParamByName('fechaf').AsDateTime := rgFecha1;
            qryconsulta1.Open;
            qryconsulta1.First;
            while not qryconsulta1.eof do
            begin
              dVaraux1 := dvaraux1 + qryconsulta1.FieldByName('smn').AsFloat;
              dVaraux2 := dvaraux2 + qryconsulta1.FieldByName('sdll').AsFloat;
              qryconsulta1.Next;
            end;

            qryconsulta1.Active := False;
            qryconsulta1.SQL.Clear;
            qryconsulta1.SQL.Text :=  'select sum(t1.smn) as smn, sum(t1.sdll) as sdll from  (select (round(ifnull(round(sum(be.dCantHHgenerador),2),0)*ifnull(e.dVentaMN,0),2)) as smn, (round(ifnull(round(sum(be.dCantHHgenerador),2),0)*ifnull(e.dVentadll,0),2)) as sdll '+
                                      'from equipos e inner join bitacoradeequipos be on (e.sidequipo = be.sIdequipo and be.scontrato = :Contrato '+
                                      'and be.snumeroorden = :Folio ) where e.sContrato = :barco and  be.didfecha >= :Fechai and be.dIdFecha  <= :Fechaf group by e.sIdequipo order by e.iItemOrden) t1';
            qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
            qryconsulta1.ParamByName('fechai').AsDateTime := dtvaraux1;
            qryconsulta1.ParamByName('fechaf').AsDateTime := rgFecha1;
            qryconsulta1.Open;
            qryconsulta1.First;
            while not qryconsulta1.eof do
            begin
              dVaraux1 := dvaraux1 + qryconsulta1.FieldByName('smn').AsFloat;
              dVaraux2 := dvaraux2 + qryconsulta1.FieldByName('sdll').AsFloat;
              qryconsulta1.Next;
            end;

            qryconsulta1.Active := False;
            qryconsulta1.SQL.Clear;
            qryconsulta1.SQL.Text :=  'select sum(tcalc.tmn) as smn, sum(tcalc.tdll) as sdll from  '+
                                       '(select cu.sidcuenta, cu.dventamn,cu.dventadll,tt.tpern, '+
                                      '(round(ifnull(tt.tpern,0),2)*ifnull(cu.dventamn,0)) as tmn,(round(ifnull(tt.tpern,0),2)*ifnull(cu.dventadll,0)) as tdll '+
                                       'from cuentas cu inner join '+
                                         '(select prn.sidcuenta,(ifnull(prn.cp,0) + ifnull(prd.cd,0)+ifnull(pre.cpe,0)) as tpern from (select c.sidcuenta, sum(p.dCantHHgenerador) as cp  '+
                                        'from cuentas c inner join bitacoradepersonal p on (p.stipopernocta = c.sidcuenta and p.snumeroorden = :folio and p.scontrato = :Contrato )  '+
                                        'where p.dIdFecha >= :Fechai and p.dIdFecha <= :Fechaf '+
                                        'group by c.sIdCuenta) prn '+
                                      'left join '+
                                        '(select c2.sidcuenta, sum(bp.dcantidad) as cd from cuentas c2 '+
                                        'inner join bitacoradepernocta bp on (c2.sidcuenta = if(length(bp.sIdCuenta) = 0,"4.1",bp.sIdCuenta) and bp.sContrato = :Contrato2 and bp.sNumeroOrden = :folio2) '+
                                        'where bp.dIdFecha >= :Fechai and bp.dIdFecha <= :Fechaf '+
                                        'group by c2.sIdCuenta) prd '+
                                        'on (prn.sidcuenta = prd.sidcuenta) '+
                                     'left join '+
                                   '(select c3.sidcuenta, sum(pe.dCantidad) as cpe '+
                                   'from cuentas c3 inner join bitacoradepersonal_cuadre pe on (pe.stipopernocta = c3.sidcuenta and pe.snumeroorden = :folio and pe.scontrato = :Contrato ) where pe.dIdFecha >= :Fechai and pe.dIdFecha <= :Fechaf  group by c3.sIdCuenta) pre '+
                                   'on (prn.sidcuenta = pre.sidcuenta) '+
                                       ' ) tt '+
                                      'on (tt.sidcuenta = cu.sidcuenta) '+
                                      ')tcalc ';
            qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('contrato2').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio2').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('fechai').AsDateTime := dtvaraux1;
            qryconsulta1.ParamByName('fechaf').AsDateTime := rgFecha1;
            qryconsulta1.Open;
            qryconsulta1.First;
            while not qryconsulta1.eof do
            begin
              dVaraux1 := dvaraux1 + qryconsulta1.FieldByName('smn').AsFloat;
              dVaraux2 := dvaraux2 + qryconsulta1.FieldByName('sdll').AsFloat;
              qryconsulta1.Next;
            end;

          finally
            qryconsulta1.Free;
          end;
          
          if dvaraux1+dvaraux2 = 0 then
            raise exception.Create('siguiente');
          {$ENDREGION}

          {$REGION 'FECHA'}
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.Value := 'PERIODO';
          PFormatosExcel_SoloBorde(Excel);
          Excel.selection.interior.colorindex := 16;

          Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value :=  'DEL '+FormatDateTime('dd',dtvaraux1)+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',rgfecha1)) ;
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlleft;
          Excel.Selection.Value := 'ESTIMACION No.';
          PFormatosExcel_SoloBorde(Excel);
          Excel.selection.interior.colorindex := 16;

          zqrestimacion.Active := False;
          zqrestimacion.ParamByName('barco').AsString := global_contrato_barco;
          zqrestimacion.ParamByName('fechai').AsDatetime := dtvaraux1;
          zqrestimacion.ParamByName('fechaf').AsDateTime := rgFecha1;
          zqrestimacion.Open;

          Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlleft;
          if zqrestimacion.RecordCount > 0 then
            Excel.Selection.Value := zqrestimacion.FieldByName('inumeroestimacion').AsString;

          PFormatosExcel_SoloBorde(Excel);

          inc(ifila,2);

          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := 'NIVEL';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := 'CATEGORIA / CONCEPTO';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := 'CANTIDAD';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := 'UNIDAD';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := 'PRECIO M.N.';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := 'PRECIO DLS.';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := 'IMPORTE M.N.';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlcenter;
          Excel.Selection.Value := 'IMPORTE DLS.';
          PFormatosExcel_SoloBorde(Excel);
          Excel.Rows[ifila].RowHeight := 21;

          inc(ifila,1);
          {$ENDREGION}

          {$REGION 'MOVIMIENTOS DE BARCO'}
          qryconsulta1 := tzreadonlyquery.Create(nil);
          try
            qryconsulta1.Connection := connection.zConnection;
            qryconsulta1.Active := False;
            qryconsulta1.SQL.Clear;
            qryconsulta1.SQL.Text :=  'select tm.sidtipomovimiento,tm.sdescripcion,tm.smedida,tm.stipo,ifnull(tf.factor,0) as factor,'+
                                      'ifnull(tm.dventamn,0) as dventamn,ifnull(tm.dventadll,0) as dventadll,  ifnull((tm.dVentaMN*tf.factor),0) as smn,  '+
                                      'ifnull((tm.dVentadll*tf.factor),0) as sdll from tiposdemovimiento tm left join  '+
                                      '(select mb.sClasificacion,sum(mf.sfactor) as factor from movimientosdeembarcacion mb inner join '+
                                      'movimientosxfolios mf on (mf.iiddiario = mb.iiddiario and mf.scontrato = :barco and sNumeroOrden = :contrato and mf.sfolio = :folio) '+
                                      'where mb.sContrato = :barco and mb.dIdFecha >= :Fechai and mb.didfecha <= :Fechaf group by mb.sClasificacion  ) tf '+
                                      'on (tf.sclasificacion = sIdTipoMovimiento) where tm.sClasificacion = "Movimiento de Barco" and tm.sIdTipoMovimiento <> "s/n"';
            qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
            qryconsulta1.ParamByName('fechai').AsDateTime := dtvaraux1;
            qryconsulta1.ParamByName('fechaf').AsDateTime := rgFecha1;
            qryconsulta1.Open;
            ivarAux := ifila;
            while not qryconsulta1.eof do
            begin
              if (qryconsulta1.FieldByName('smn').asfloat > 0) or (qryconsulta1.FieldByName('sdll').asfloat > 0) or (AnsiMatchStr( qryconsulta1.FieldByName('sidtipomovimiento').AsString, ['1.1','1.2','1.3'] )) then
              begin
                Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
                PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
                Excel.Selection.HorizontalAlignment := xlcenter;
                Excel.Selection.Value := qryconsulta1.FieldByName('sidtipomovimiento').AsString;
                PFormatosExcel_SoloBorde(Excel);

                Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
                PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
                Excel.Selection.HorizontalAlignment := xlleft;
                svarAux3 := qryconsulta1.FieldByName('sdescripcion').AsString;
                svaraux3 := AnsireplaceText( svaraux3, 'en operacion ', '' );
                svaraux3 := AnsireplaceText( svaraux3, 'EN OPERACIÓN ', '' );
                Excel.Selection.Value := svaraux3;
                //Excel.Selection.Value := qryconsulta1.FieldByName('sdescripcion').AsString;
                PFormatosExcel_SoloBorde(Excel);

                Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
                PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
                Excel.Selection.HorizontalAlignment := xlcenter;
                Excel.Selection.Value := qryconsulta1.FieldByName('factor').asfloat;
                Excel.Selection.NumberFormat := '#,##0.00000000';
                PFormatosExcel_SoloBorde(Excel);

                Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
                PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
                Excel.Selection.HorizontalAlignment := xlcenter;
                Excel.Selection.Value := qryconsulta1.FieldByName('smedida').AsString;
                Excel.Selection.NumberFormat := '@';
                PFormatosExcel_SoloBorde(Excel);

                Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
                PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
                Excel.Selection.HorizontalAlignment := xlright;
                Excel.Selection.Value := qryconsulta1.FieldByName('dventamn').asfloat;
                Excel.Selection.NumberFormat := '#,##0.00';
                PFormatosExcel_SoloBorde(Excel);

                Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
                PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
                Excel.Selection.HorizontalAlignment := xlright;
                Excel.Selection.Value := qryconsulta1.FieldByName('dventadll').asfloat;
                Excel.Selection.NumberFormat := '#,##0.00';
                PFormatosExcel_SoloBorde(Excel);

                Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
                PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
                Excel.Selection.HorizontalAlignment := xlright;
                //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
                Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
                Excel.Selection.NumberFormat :=  '#,##0.00';
                PFormatosExcel_SoloBorde(Excel);
                svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

                Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
                PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
                Excel.Selection.HorizontalAlignment := xlright;
                //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
                Excel.Selection.formula := '=ROUND((RC[-2]*RC[-5]),2)';
                Excel.Selection.NumberFormat := '#,##0.00';
                PFormatosExcel_SoloBorde(Excel);
                svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);

                IVarAux2 := Length(qryconsulta1.FieldByName('sdescripcion').AsString);
                if (IVarAux2 mod 34) > 0 then
                  IVarAux2 := IVarAux2+34;
                IVarAux2 := IVarAux2 div 34;
                IVarAux2 := IVarAux2 * 11;
                if Excel.Rows[ifila].RowHeight < IVarAux2 then
                  Excel.Rows[ifila].RowHeight := IVarAux2;

                inc(ifila,1);
              end;
              qryconsulta1.Next;
            end;
          finally
            qryconsulta1.Free;
          end;
        {$ENDREGION}

          {$REGION 'FILA BLANCA'}
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);
          Excel.Rows[ifila].RowHeight := 11.25;
          inc(ifila,1);
          {$ENDREGION}

          {$REGION 'PERSONAL'}
          qryconsulta1 := tzreadonlyquery.Create(nil);
          try
            qryconsulta1.Connection := connection.zConnection;
            qryconsulta1.Active := False;
            qryconsulta1.SQL.Clear;
            qryconsulta1.SQL.Text :=  'select sum(t1.smn) as smn, sum(t1.sdll) as sdll from  (select (round(ifnull(round(sum(bp.dCantHHgenerador),2),0)*ifnull(p.dVentaMN,0),2)) as smn, (round(ifnull(round(sum(bp.dCantHHgenerador),2),0)*ifnull(p.dVentadll,0),2)) as sdll '+
                                      'from personal p inner join bitacoradepersonal bp on (p.sidpersonal = bp.sIdPersonal and bp.scontrato = :Contrato '+
                                      'and bp.snumeroorden = :Folio ) where p.sContrato = :barco and  bp.didfecha >= :Fechai and bp.dIdFecha  <= :Fechaf group by p.sIdPersonal  order by p.iItemOrden) t1';
            qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
            qryconsulta1.ParamByName('fechai').AsDateTime := dtvaraux1;
            qryconsulta1.ParamByName('fechaf').AsDateTime := rgFecha1;
            qryconsulta1.Open;
            ivarAux := ifila;
            while not qryconsulta1.eof do
            begin
              Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := '2';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlleft;
              Excel.Selection.Value := 'PERSONAL DE CONSTRUCCIÓN ANEXO C-1';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := 1;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := 'LOTE';
              Excel.Selection.NumberFormat := '@';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
              Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
              Excel.Selection.NumberFormat :=  '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);
              svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

              Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
              Excel.Selection.formula := '=ROUND((RC[-2]*RC[-5]),2)';
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);
              svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);

              Excel.Rows[ifila].RowHeight := 16;

              inc(ifila,1);
              qryconsulta1.Next;
            end;
          finally
            qryconsulta1.Free;
          end;
        {$ENDREGION}

          {$REGION 'EQUIPO'}
          qryconsulta1 := tzreadonlyquery.Create(nil);
          try
            qryconsulta1.Connection := connection.zConnection;
            qryconsulta1.Active := False;
            qryconsulta1.SQL.Clear;
            qryconsulta1.SQL.Text :=  'select sum(t1.smn) as smn, sum(t1.sdll) as sdll from  (select (round(ifnull(round(sum(be.dCantHHgenerador),2),0)*ifnull(e.dVentaMN,0),2)) as smn, (round(ifnull(round(sum(be.dCantHHgenerador),2),0)*ifnull(e.dVentadll,0),2)) as sdll '+
                                      'from equipos e inner join bitacoradeequipos be on (e.sidequipo = be.sIdequipo and be.scontrato = :Contrato '+
                                      'and be.snumeroorden = :Folio ) where e.sContrato = :barco and  be.didfecha >= :Fechai and be.dIdFecha  <= :Fechaf group by e.sIdequipo order by e.iItemOrden) t1';
            qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
            qryconsulta1.ParamByName('fechai').AsDateTime := dtvaraux1;
            qryconsulta1.ParamByName('fechaf').AsDateTime := rgFecha1;
            qryconsulta1.Open;
            ivarAux := ifila;
            while not qryconsulta1.eof do
            begin
              Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := '3';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlleft;
              Excel.Selection.Value := 'EQUIPOS OPTATIVOS PARA TRABAJOS DE REHABILITACION, MANTENIMIENTO, BUCEO Y SEGURIDAD ANEXO C-2';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := 1;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := 'LOTE';
              Excel.Selection.NumberFormat := '@';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
              Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
              Excel.Selection.NumberFormat :=  '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);
              svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

              Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
              Excel.Selection.formula := '=ROUND((RC[-2]*RC[-5]),2)';
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);
              svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);

              Excel.Rows[ifila].RowHeight := 32;

              inc(ifila,1);
              qryconsulta1.Next;
            end;
          finally
            qryconsulta1.Free;
          end;
        {$ENDREGION}

          {$REGION 'PERNOCTA'}
          qryconsulta1 := tzreadonlyquery.Create(nil);
          try
            qryconsulta1.Connection := connection.zConnection;
            qryconsulta1.Active := False;
            qryconsulta1.SQL.Clear;
            qryconsulta1.SQL.Text :=  'select sum(tcalc.tmn) as smn, sum(tcalc.tdll) as sdll from  '+
                                       '(select cu.sidcuenta, cu.dventamn,cu.dventadll,tt.tpern, '+
                                      '(round(ifnull(tt.tpern,0),2)*ifnull(cu.dventamn,0)) as tmn,(round(ifnull(tt.tpern,0),2)*ifnull(cu.dventadll,0)) as tdll '+
                                       'from cuentas cu inner join '+
                                         '(select prn.sidcuenta,(ifnull(prn.cp,0) + ifnull(prd.cd,0)+ifnull(pre.cpe,0)) as tpern from (select c.sidcuenta, sum(p.dCantHHgenerador) as cp  '+
                                        'from cuentas c inner join bitacoradepersonal p on (p.stipopernocta = c.sidcuenta and p.snumeroorden = :folio and p.scontrato = :Contrato )  '+
                                        'where p.dIdFecha >= :Fechai and p.dIdFecha <= :Fechaf '+
                                        'group by c.sIdCuenta) prn '+
                                      'left join '+
                                        '(select c2.sidcuenta, sum(bp.dcantidad) as cd from cuentas c2 '+
                                        'inner join bitacoradepernocta bp on (c2.sidcuenta = if(length(bp.sIdCuenta) = 0,"4.1",bp.sIdCuenta) and bp.sContrato = :Contrato2 and bp.sNumeroOrden = :folio2) '+
                                        'where bp.dIdFecha >= :Fechai and bp.dIdFecha <= :Fechaf '+
                                        'group by c2.sIdCuenta) prd '+
                                        'on (prn.sidcuenta = prd.sidcuenta) '+
                                     'left join '+
                                   '(select c3.sidcuenta, sum(pe.dCantidad) as cpe '+
                                   'from cuentas c3 inner join bitacoradepersonal_cuadre pe on (pe.stipopernocta = c3.sidcuenta and pe.snumeroorden = :folio and pe.scontrato = :Contrato ) where pe.dIdFecha >= :Fechai and pe.dIdFecha <= :Fechaf  group by c3.sIdCuenta) pre '+
                                   'on (prn.sidcuenta = pre.sidcuenta) '+
                                       ' ) tt '+
                                      'on (tt.sidcuenta = cu.sidcuenta) '+
                                      ')tcalc ';
            qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('contrato2').AsString := zqrfolio.FieldByName('scontrato').asstring;
            qryconsulta1.ParamByName('folio2').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
            qryconsulta1.ParamByName('fechai').AsDateTime := dtvaraux1;
            qryconsulta1.ParamByName('fechaf').AsDateTime := rgFecha1;
            qryconsulta1.Open;
            ivarAux := ifila;
            while not qryconsulta1.eof do
            begin
              Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := '4';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlleft;
              Excel.Selection.Value := 'SERVICIOS DE ALIMENTACIÓN Y HOSPEDAJE ANEXO C-3';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := 1;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := 'LOTE';
              Excel.Selection.NumberFormat := '@';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
              Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
              Excel.Selection.NumberFormat :=  '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);
              svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

              Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
              Excel.Selection.formula := '=ROUND((RC[-2]*RC[-5]),2)';
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);
              svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);

              Excel.Rows[ifila].RowHeight := 32;

              inc(ifila,1);
              qryconsulta1.Next;
            end;
          finally
            qryconsulta1.Free;
          end;
        {$ENDREGION}

          {$REGION 'FILA BLANCA'}
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.Value := '';
          PFormatosExcel_SoloBorde(Excel);
          Excel.Rows[ifila].RowHeight := 11.25;
          inc(ifila,1);
          {$ENDREGION}

          {$REGION 'FILA PRECIOS OPTATIVOS'}
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.Value := '6';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlleft;
          Excel.Selection.Value := 'PRECIOS UNITARIOS OPTATIVOS (C-6).';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.Value := '1';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.Value := 'LOTE';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.Value := '0';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.Value := '0';
          PFormatosExcel_SoloBorde(Excel);

          Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.Value := '0';
          PFormatosExcel_SoloBorde(Excel);
          svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

          Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
          Excel.Selection.Value := '0';
          PFormatosExcel_SoloBorde(Excel);
          Excel.Rows[ifila].RowHeight := 16;
          svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);
          inc(ifila,1);
          {$ENDREGION}

          {$REGION 'TOTAL'}
          inc(ifila,1);
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, FALSE, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.Value := 'IMPORTE TOTAL:';

          Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.NumberFormat := '#,##0.00';
          Excel.Selection.FORMULA := svaraux;

          Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlright;
          Excel.Selection.formula:= svaraux2;
          Excel.Selection.NumberFormat := '#,##0.00';
          {$ENDREGION}

          inc(ciclos);
          inc(ifila,3);
        except
          on e:exception do
            if e.message <> 'siguiente' then
              raise; 
        end;

          //Excel.Rows[ifila].PageBreak: = 1;
          //hojacarta(excel,2);
        dtVarAux1 := IncMonth(dtVarAux1);

      end;

      {$REGION 'ACUMULADO ACTUAL DE ESTIMACIONES'}
      if ciclos > 1 then
      begin
        rgFecha1 := IncDay( dtvaraux1, strtoint(vartostr(DaysInMonth( dtVarAux1 )))-1 );
        sVarAux := '= ';
        sVarAux2 := '= ';

        {$REGION 'TITULO-ENCABEZADO'}
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := XLCENTER;
        Excel.Selection.Value := 'ACUMULADO ACTUAL DE ESTIMACIONES';
        PFormatosExcel_SoloBorde(Excel);
        Excel.selection.interior.colorindex := 16;
        Excel.Rows[ifila].RowHeight := 21;

        inc(ifila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlcenter;
        Excel.Selection.Value := 'NIVEL';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlcenter;
        Excel.Selection.Value := 'CATEGORIA / CONCEPTO';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlcenter;
        Excel.Selection.Value := 'CANTIDAD';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlcenter;
        Excel.Selection.Value := 'UNIDAD';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlcenter;
        Excel.Selection.Value := 'PRECIO M.N.';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlcenter;
        Excel.Selection.Value := 'PRECIO DLS.';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlcenter;
        Excel.Selection.Value := 'IMPORTE M.N.';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlcenter;
        Excel.Selection.Value := 'IMPORTE DLS.';
        PFormatosExcel_SoloBorde(Excel);
        Excel.Rows[ifila].RowHeight := 21;

        inc(ifila,1);
        {$ENDREGION}

        {$REGION 'MOVIMIENTOS DE BARCO'}
        qryconsulta1 := tzreadonlyquery.Create(nil);
        try
          qryconsulta1.Connection := connection.zConnection;
          qryconsulta1.Active := False;
          qryconsulta1.SQL.Clear;
          qryconsulta1.SQL.Text :=  'select tm.sidtipomovimiento,tm.sdescripcion,tm.smedida,tm.stipo,ifnull(tf.factor,0) as factor,'+
                                    'ifnull(tm.dventamn,0) as dventamn,ifnull(tm.dventadll,0) as dventadll,  ifnull((tm.dVentaMN*tf.factor),0) as smn,  '+
                                    'ifnull((tm.dVentadll*tf.factor),0) as sdll from tiposdemovimiento tm left join  '+
                                    '(select mb.sClasificacion,sum(mf.sfactor) as factor from movimientosdeembarcacion mb inner join '+
                                    'movimientosxfolios mf on (mf.iiddiario = mb.iiddiario and mf.scontrato = :barco and sNumeroOrden = :contrato and mf.sfolio = :folio) '+
                                    'where mb.sContrato = :barco  group by mb.sClasificacion  ) tf '+
                                    'on (tf.sclasificacion = sIdTipoMovimiento) where tm.sClasificacion = "Movimiento de Barco" and tm.sIdTipoMovimiento <> "s/n"';
          qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
          qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
          qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
          qryconsulta1.Open;
          ivarAux := ifila;
          while not qryconsulta1.eof do
          begin
            if (qryconsulta1.FieldByName('smn').asfloat > 0) or (qryconsulta1.FieldByName('sdll').asfloat > 0) or (AnsiMatchStr( qryconsulta1.FieldByName('sidtipomovimiento').AsString, ['1.1','1.2','1.3'] )) then
            begin

              Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := qryconsulta1.FieldByName('sidtipomovimiento').AsString;
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlleft;
              svarAux3 := qryconsulta1.FieldByName('sdescripcion').AsString;
              svaraux3 := AnsireplaceText( svaraux3, 'en operacion ', '' );
              svaraux3 := AnsireplaceText( svaraux3, 'EN OPERACIÓN ', '' );
              Excel.Selection.Value := svaraux3;
              //Excel.Selection.Value := qryconsulta1.FieldByName('sdescripcion').AsString;
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := qryconsulta1.FieldByName('factor').asfloat;
              Excel.Selection.NumberFormat := '#,##0.00000000';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlcenter;
              Excel.Selection.Value := qryconsulta1.FieldByName('smedida').AsString;
              Excel.Selection.NumberFormat := '@';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              Excel.Selection.Value := qryconsulta1.FieldByName('dventamn').asfloat;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              Excel.Selection.Value := qryconsulta1.FieldByName('dventadll').asfloat;
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);

              Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
              Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
              Excel.Selection.NumberFormat :=  '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);
              svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

              Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
              PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
              Excel.Selection.HorizontalAlignment := xlright;
              //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
              Excel.Selection.formula := '=ROUND((RC[-2]*RC[-5]),2)';
              Excel.Selection.NumberFormat := '#,##0.00';
              PFormatosExcel_SoloBorde(Excel);
              svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);

              IVarAux2 := Length(qryconsulta1.FieldByName('sdescripcion').AsString);
              if (IVarAux2 mod 34) > 0 then
                IVarAux2 := IVarAux2+34;
              IVarAux2 := IVarAux2 div 34;
              IVarAux2 := IVarAux2 * 11;
              if Excel.Rows[ifila].RowHeight < IVarAux2 then
                Excel.Rows[ifila].RowHeight := IVarAux2;

              inc(ifila,1);
            end;
            qryconsulta1.Next;
          end;
        finally
          qryconsulta1.Free;
        end;
      {$ENDREGION}

        {$REGION 'FILA BLANCA'}
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);
        Excel.Rows[ifila].RowHeight := 11.25;
        inc(ifila,1);
        {$ENDREGION}

        {$REGION 'PERSONAL'}
        qryconsulta1 := tzreadonlyquery.Create(nil);
        try
          qryconsulta1.Connection := connection.zConnection;
          qryconsulta1.Active := False;
          qryconsulta1.SQL.Clear;
          qryconsulta1.SQL.Text :=  'select sum(t1.smn) as smn, sum(t1.sdll) as sdll from  (select (round(ifnull(round(sum(bp.dCantHHgenerador),2),0)*ifnull(p.dVentaMN,0),2)) as smn, (round(ifnull(round(sum(bp.dCantHHgenerador),2),0)*ifnull(p.dVentadll,0),2)) as sdll '+
                                    'from personal p inner join bitacoradepersonal bp on (p.sidpersonal = bp.sIdPersonal and bp.scontrato = :Contrato '+
                                    'and bp.snumeroorden = :Folio ) where p.sContrato = :barco group by p.sIdPersonal order by p.iItemOrden) t1';
          qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
          qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
          qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
          qryconsulta1.Open;
          ivarAux := ifila;
          while not qryconsulta1.eof do
          begin
            Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := '2';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlleft;
            Excel.Selection.Value := 'PERSONAL DE CONSTRUCCIÓN ANEXO C-1';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := 1;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := 'LOTE';
            Excel.Selection.NumberFormat := '@';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
            Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
            Excel.Selection.NumberFormat :=  '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);
            svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

            Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
            Excel.Selection.formula := '=ROUND((RC[-2]*RC[-5]),2)';
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);
            svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);

            Excel.Rows[ifila].RowHeight := 16;

            inc(ifila,1);
            qryconsulta1.Next;
          end;
        finally
          qryconsulta1.Free;
        end;
      {$ENDREGION}

        {$REGION 'EQUIPO'}
        qryconsulta1 := tzreadonlyquery.Create(nil);
        try
          qryconsulta1.Connection := connection.zConnection;
          qryconsulta1.Active := False;
          qryconsulta1.SQL.Clear;
          qryconsulta1.SQL.Text :=  'select sum(t1.smn) as smn, sum(t1.sdll) as sdll from  (select (round(ifnull(round(sum(be.dCantHHgenerador),2),0)*ifnull(e.dVentaMN,0),2)) as smn, (round(ifnull(round(sum(be.dCantHHgenerador),2),0)*ifnull(e.dVentadll,0),2)) as sdll  '+
                                    'from equipos e inner join bitacoradeequipos be on (e.sidequipo = be.sIdequipo and be.scontrato = :Contrato '+
                                    'and be.snumeroorden = :Folio ) where e.sContrato = :barco group by e.sIdequipo order by e.iItemOrden) t1';
          qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
          qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
          qryconsulta1.ParamByName('barco').AsString := global_contrato_barco;
          qryconsulta1.Open;
          ivarAux := ifila;
          while not qryconsulta1.eof do
          begin
            Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := '3';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlleft;
            Excel.Selection.Value := 'EQUIPOS OPTATIVOS PARA TRABAJOS DE REHABILITACION, MANTENIMIENTO, BUCEO Y SEGURIDAD ANEXO C-2';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := 1;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := 'LOTE';
            Excel.Selection.NumberFormat := '@';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
            Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
            Excel.Selection.NumberFormat :=  '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);
            svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

            Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
            Excel.Selection.formula := '=ROUND((RC[-2]*RC[-5]),2)';
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);
            svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);

            Excel.Rows[ifila].RowHeight := 32;

            inc(ifila,1);
            qryconsulta1.Next;
          end;
        finally
          qryconsulta1.Free;
        end;
      {$ENDREGION}

        {$REGION 'PERNOCTA'}
        qryconsulta1 := tzreadonlyquery.Create(nil);
        try
          qryconsulta1.Connection := connection.zConnection;
          qryconsulta1.Active := False;
          qryconsulta1.SQL.Clear;
          qryconsulta1.SQL.Text :=  'select ifnull(sum(tcalc.tmn),0) as smn, ifnull(sum(tcalc.tdll),0) as sdll from '+
                                 '(select cu.sidcuenta, cu.dventamn,cu.dventadll,tt.tpern,(round(ifnull(tt.tpern,0),2)*ifnull(cu.dventamn,0)) as tmn,(round(ifnull(tt.tpern,0),2)*ifnull(cu.dventadll,0)) as tdll from cuentas cu inner join  '+
                                 '(select prn.sidcuenta,(ifnull(prn.cp,0) + ifnull(prd.cd,0))+ifnull(pre.cpe,0) as tpern from (select c.sidcuenta, sum(p.dCantHHgenerador) as cp '+
                                 'from cuentas c inner join bitacoradepersonal p on (p.stipopernocta = c.sidcuenta and p.snumeroorden = :folio and p.scontrato = :Contrato ) group by c.sIdCuenta) prn '+
                                 'left join  '+
                                 '(select c2.sidcuenta, sum(bp.dcantidad) as cd from cuentas c2 inner join bitacoradepernocta bp on (c2.sidcuenta = if(length(bp.sIdCuenta) '+
                                 '= 0,"4.1",bp.sIdCuenta) and bp.sContrato = :Contrato2 and bp.sNumeroOrden = :folio2) group by c2.sIdCuenta) prd on (prn.sidcuenta = prd.sidcuenta) '+
                                 '  left join '+
                                 '(select c3.sidcuenta, sum(pe.dCantidad) as cpe '+
                                 'from cuentas c3 inner join bitacoradepersonal_cuadre pe on (pe.stipopernocta = c3.sidcuenta and pe.snumeroorden = :folio and pe.scontrato = :Contrato ) group by c3.sIdCuenta) pre '+
                                 'on (prn.sidcuenta = pre.sidcuenta) '+
                                 ') tt '+
                                 'on (tt.sidcuenta = cu.sidcuenta)) tcalc';
          qryconsulta1.ParamByName('contrato').AsString := zqrfolio.FieldByName('scontrato').asstring;
          qryconsulta1.ParamByName('folio').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
          qryconsulta1.ParamByName('contrato2').AsString := zqrfolio.FieldByName('scontrato').asstring;
          qryconsulta1.ParamByName('folio2').AsString := zqrfolio.FieldByName('snumeroorden').asstring;
          qryconsulta1.Open;
          ivarAux := ifila;
          while not qryconsulta1.eof do
          begin
            Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := '4';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlleft;
            Excel.Selection.Value := 'SERVICIOS DE ALIMENTACIÓN Y HOSPEDAJE ANEXO C-3';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := 1;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlcenter;
            Excel.Selection.Value := 'LOTE';
            Excel.Selection.NumberFormat := '@';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);

            Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            //Excel.Selection.Value := qryconsulta1.FieldByName('smn').asfloat;
            Excel.Selection.formula := '=ROUND((RC[-2]*RC[-4]),2)';
            Excel.Selection.NumberFormat :=  '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);
            svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

            Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
            Excel.Selection.HorizontalAlignment := xlright;
            //Excel.Selection.Value := qryconsulta1.FieldByName('sdll').asfloat;
            Excel.Selection.formula := '=ROUND((RC[-2]*RC[-5]),2)';
            Excel.Selection.NumberFormat := '#,##0.00';
            PFormatosExcel_SoloBorde(Excel);
            svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);

            Excel.Rows[ifila].RowHeight := 32;

            inc(ifila,1);
            qryconsulta1.Next;
          end;
        finally
          qryconsulta1.Free;
        end;
      {$ENDREGION}

        {$REGION 'FILA BLANCA'}
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.Value := '';
        PFormatosExcel_SoloBorde(Excel);
        Excel.Rows[ifila].RowHeight := 11.25;
        inc(ifila,1);
        {$ENDREGION}

        {$REGION 'FILA PRECIOS OPTATIVOS'}
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
        Excel.Selection.Value := '6';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlleft;
        Excel.Selection.Value := 'PRECIOS UNITARIOS OPTATIVOS (C-6).';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
        Excel.Selection.Value := '1';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
        Excel.Selection.Value := 'LOTE';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
        Excel.Selection.Value := '0';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
        Excel.Selection.Value := '0';
        PFormatosExcel_SoloBorde(Excel);

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
        Excel.Selection.Value := '0';
        PFormatosExcel_SoloBorde(Excel);
        svaraux := svaraux + '+'+ColumnaNombre(11)+IntToStr(iFila);

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, false, 7, clBlack, 'Arial');
        Excel.Selection.Value := '0';
        PFormatosExcel_SoloBorde(Excel);
        Excel.Rows[ifila].RowHeight := 16;
        svaraux2 := svaraux2 + '+'+ColumnaNombre(12)+IntToStr(iFila);
        inc(ifila,1);
        {$ENDREGION}

        {$REGION 'TOTAL'}
        inc(ifila,1);
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, FALSE, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.Value := 'IMPORTE TOTAL:';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.NumberFormat := '#,##0.00';
        Excel.Selection.FORMULA := svaraux;

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, true, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.formula:= svaraux2;
        Excel.Selection.NumberFormat := '#,##0.00';
        {$ENDREGION}

      end;
      {$ENDREGION}
      hojacarta(excel,2);
      {$ENDREGION}
    end;

  except
    on e:Exception do
      if e.message <> '***' then
        ShowMessage(e.message);
  end;
end;

procedure TfrmReportePeriodo.ActualizaPernocta(Sender: TObject);
begin
IF QryPersonalPernocta.Active THEN BEGIN
  If QryPersonalPernocta.RecordCount > 0 Then
  Begin
      If QryPersonalPernocta.FieldValues['sIdPernocta'] <> sPernocta Then
      Begin
           Connection.QryBusca.Active := False ;
           Connection.QryBusca.SQL.Clear ;
           Connection.QryBusca.SQL.Add('Select Sum(b.dCantidad) as dReal ' +
                                       'From bitacoradepersonal b INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                       'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                       'Where b.sContrato = :Contrato And b.sIdPernocta = :Pernocta And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sContrato, b.sIdPernocta') ;
           Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
           Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
           Connection.QryBusca.Params.ParamByName('Pernocta').DataType := ftString ;
           Connection.QryBusca.Params.ParamByName('Pernocta').Value := QryPersonalPernocta.FieldValues['sIdPernocta'] ;
           Connection.QryBusca.Params.ParamByName('FechaI').DataType := ftDate ;
           Connection.QryBusca.Params.ParamByName('FechaI').Value := tdFechaInicial.Date ;
           Connection.QryBusca.Params.ParamByName('FechaF').DataType := ftDate ;
           Connection.QryBusca.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
           Connection.QryBusca.Open ;
           dReal := 0 ;
           If Connection.QryBusca.RecordCount > 0 Then
           Begin
               dReal := Connection.QryBusca.FieldValues['dReal'] ;
               Connection.QryBusca.Active := False ;
               Connection.QryBusca.SQL.Clear ;
               Connection.QryBusca.SQL.Add('Select distinct b.dIdFecha ' +
                                           'From bitacoradepersonal b INNER JOIN personal p2 ON (b.sContrato = p2.sContrato And b.sIdPersonal = p2.sIdPersonal) ' +
                                           'INNER JOIN bitacoradeactividades b2 ON (b.sContrato = b2.sContrato And b.dIdFecha = b2.dIdFecha And b.iIdDiario = b2.iIdDiario) ' +
                                           'Where b.sContrato = :Contrato And b.sIdPernocta = :Pernocta And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF Group By b.sContrato, b.sIdPernocta, b.dIdFecha') ;
               Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
               Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
               Connection.QryBusca.Params.ParamByName('Pernocta').DataType := ftString ;
               Connection.QryBusca.Params.ParamByName('Pernocta').Value := QryPersonalPernocta.FieldValues['sIdPernocta'] ;
               Connection.QryBusca.Params.ParamByName('FechaI').DataType := ftDate ;
               Connection.QryBusca.Params.ParamByName('FechaI').Value := tdFechaInicial.Date ;
               Connection.QryBusca.Params.ParamByName('FechaF').DataType := ftDate ;
               Connection.QryBusca.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
               Connection.QryBusca.Open ;
               If Connection.QryBusca.RecordCount > 0 Then
                    dReal := dReal / Connection.QryBusca.RecordCount ;

               sPernocta := QryPersonalPernocta.FieldValues['sIdPernocta']
          End
      End
    End
END;
end;

procedure TfrmReportePeriodo.frxProgramacionGetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'PROGRAMADO') = 0 then
      Value := floattostr(RoundTo(dProgramado,-2)) ;
  If CompareText(VarName, 'REAL') = 0 then
      Value := floattostr(RoundTo(dReal,-2)) ;
  If CompareText(VarName, 'PROMEDIO') = 0 then
      Value := floattostr(RoundTo(dPromedio,0)) ;

  If CompareText(VarName, 'PERIODO') = 0 then
      Value := DateToStr (tdFechaInicial.Date ) + ' AL ' +  DateToStr (tdFechaFinal.Date ) ;

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


procedure TfrmReportePeriodo.btnImprimeDiariosClick(Sender: TObject);
var
  OldWbs, NomCampo,
  sTipoMoneda, sCadena, sDia, MiWbs : String;
  iPos, NumPaq, i  : Integer;
  Q_Paquetes       : tzReadOnlyquery ;
  ArrayPaquetes : array [1..10, 1..2] of String;
begin
//Verifica que la fecha final no sea menor que la fecha inicio      //
{ If reportesCheck.Checked Then
    Begin


    End;    }
   if tdFechaFinal.Date<tdFechaInicial.Date then
   begin
   showmessage('la fecha final de impresión es menor a la fecha inicial de impresión' );
   tdFechaFinal.SetFocus;
   exit;
   end;   

    if not BotonPermiso.imprimir then
    begin
      showmessage('No tiene permisos de impresión');
      exit;
    end;
    Q_Paquetes := tzReadOnlyQuery.Create(Self) ;
    Q_Paquetes.Connection := connection.zConnection ;

    if rbCronologias.Checked then begin
      ImprimeExcel_Cronologias;
    end;                           

    if rbCronologiasFiltradas.Checked then begin
      ImprimeExcel_Cronologias;
    end;

    if rbNotaCampo.Checked then begin
      ImprimeExcel_NotaDeCampo;
    end;

end;

procedure TfrmReportePeriodo.btnimprimirClick(Sender: TObject);
var
  Libro, Excel, Hoja, iHojas: Variant;
  OldWbs, NomCampo,
  sTipoMoneda, sCadena, NombreDelExcel, sDia, MiWbs : String;
  iPos, NumPaq  : Integer;
  Q_Paquetes       : tzReadOnlyquery ;
  iCounter, iLoop, iColumnaProrrateo, ItemMayor: Integer;
begin
   if tdFechaFinal.Date<tdFechaInicial.Date then
   begin
     showmessage('la fecha final de impresión es menor a la fecha inicial de impresión' );
     tdFechaFinal.SetFocus;
     exit;
   end;

  if not BotonPermiso.imprimir then
  begin
    showmessage('No tiene permisos de impresión');
    exit;
  end;
  Q_Paquetes := tzReadOnlyQuery.Create(Self) ;
  Q_Paquetes.Connection := connection.zConnection ;

  if tdFechaFinal.Date<tdFechaInicial.Date then
  begin
    showmessage('la fecha final de impresión es menor a la fecha inicial de impresión' );
    tdFechaFinal.SetFocus;
  exit;
  end;

  if not BotonPermiso.imprimir then
  begin
    showmessage('No tiene permisos de impresión');
    exit;
  end;
  Q_Paquetes := tzReadOnlyQuery.Create(Self) ;
  Q_Paquetes.Connection := connection.zConnection ;


  imprimirLibro;


end;

procedure TfrmReportePeriodo.rbAnalisisFinancieroEnter(Sender: TObject);
begin
     chkFases.Enabled := False;
     chkFases.Checked := False;
end;

procedure TfrmReportePeriodo.rbComentariosEnter(Sender: TObject);
begin
     chkFases.Enabled := False;
     chkFases.Checked := False;
end;

procedure TfrmReportePeriodo.rbDetalleAvancesEnter(Sender: TObject);
begin
     chkFases.Enabled := False;
     chkFases.Checked := False;
end;

procedure TfrmReportePeriodo.rbDetalleInstalacionEnter(Sender: TObject);
begin
     chkFases.Enabled := True;
end;

procedure TfrmReportePeriodo.rbProgramadoEnter(Sender: TObject);
begin
     chkFases.Enabled := False;
     chkFases.Checked := False;
end;

procedure TfrmReportePeriodo.rbResumenGralEnter(Sender: TObject);
begin
     chkFases.Enabled := False;
     chkFases.Checked := False;
end;

procedure TfrmReportePeriodo.rbResumenInstalacionEnter(Sender: TObject);
begin
     chkFases.Enabled := False;
     chkFases.Checked := False;
end;

procedure TfrmReportePeriodo.rbVolumenGeneralEnter(Sender: TObject);
begin
    chkFases.Enabled := True;
end;

procedure TfrmReportePeriodo.rDiarioGetValue(const VarName: String;
  var Value: Variant);
begin
  if chkMoneda.Checked then
  begin
      If CompareText(VarName, 'MONEDA') = 0 then
         Value := 'M.N.' ;
  end
  else
  begin
       If CompareText(VarName, 'MONEDA') = 0 then
         Value := 'DLL';
  end;

  If CompareText(VarName, 'ORDEN') = 0 then
      Value := 'DEL CONTRATO' ;

  If CompareText(VarName, 'FECHA_INICIO') = 0 then
      Value := tdFechaInicial.Date  ;

  If CompareText(VarName, 'FECHA_FINAL') = 0 then
      Value := tdFechaFinal.Date ;

  If CompareText(VarName, 'DESCRIPCION_CORTA') = 0 then
      Value := sDiarioDescripcionCorta ;

  If CompareText(VarName, 'IMPRIME_AVANCES') = 0 then
      Value := sDiarioComentario ;

  If CompareText(VarName, 'sNewTexto') = 0 then
      Value := sDiarioTitulo ;

  If CompareText(VarName, 'PERIODO') = 0 then
      Value := sDiarioPeriodo ;


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

procedure TfrmReportePeriodo.chkDetalleClick(Sender: TObject);
begin
    if chkDetalle.checked then
    begin
        chkPernocta.enabled := false ;
        chkPlataforma.enabled := false ;
    end
    else
    begin
        chkPernocta.enabled := true ;
        chkPlataforma.enabled := true ;
    end
end;         

procedure TfrmReportePeriodo.chkDLLClick(Sender: TObject);
begin
    if chkMoneda.Checked then
       chkMoneda.Checked := False;
end;

procedure TfrmReportePeriodo.chkMonedaClick(Sender: TObject);
begin
     if chkDLL.Checked then
        chkDLL.Checked := False;
end;

procedure TfrmReportePeriodo.cmdEliminarClick(Sender: TObject);
var
   indice : integer;
begin
    if ComboActas.Items.Count  > 0 then
    begin
        indice := ComboActas.ItemIndex;
        ComboActas.Items.Delete(indice);
    end;
end;

procedure TfrmReportePeriodo.cmdImprimeActaClick(Sender: TObject);
var
   QryrFotografico   : tzReadOnlyQuery ;
   limiteFotos : Integer;
begin
  if chk4img.Checked then
    limiteFotos := 4;
  if chk2img.Checked then
    limiteFotos := 2;
  if numeroDeImagenes(limiteFotos) then
  begin
    QryrFotografico := tzReadOnlyquery.Create(Self) ;
    QryrFotografico.Connection := connection.zConnection ;

    QryrFotografico.Active := False ;
    QryrFotografico.SQL.Clear ;
    QryrFotografico.SQL.Add('Select f.*, o.mDescripcion as DescripcionOrden, o.sDescripcionCorta, o.sIdFolio, o.sIdPlataforma, a.mDescripcion, '+
                            '(select mDescripcion from contratos where sContrato =:codigo) as ContratoPrincipal From reportefotografico_acta f '+
                            'inner join ordenesdetrabajo o ON (f.sContrato = o.sContrato And f.sNumeroOrden = o.sNumeroOrden) '+
                            'inner join actividadesxorden a ON (a.sContrato = f.sContrato and a.sIdConvenio =:Convenio And a.sNumeroOrden = f.sNumeroOrden and a.sWbs = f.sWbs and a.sNumeroActividad = f.sNumeroActividad and a.sTipoActividad = "Actividad") '+
                            'Where f.sContrato =:Contrato and f.sNumeroOrden =:Orden and f.sActaFotografica =:Acta And f.lImprime = "Si" '+
                            'Order By f.sNumeroOrden, f.iImagen') ;
    QryrFotografico.Params.ParamByName('contrato').DataType := ftString ;
    QryrFotografico.Params.ParamByName('contrato').Value    := global_contrato ;
    QryrFotografico.Params.ParamByName('convenio').DataType := ftString ;
    QryrFotografico.Params.ParamByName('convenio').Value    := global_convenio ;
    QryrFotografico.Params.ParamByName('codigo').DataType   := ftString ;
    QryrFotografico.Params.ParamByName('codigo').Value      := connection.contrato.FieldValues['sCodigo'];
    QryrFotografico.Params.ParamByName('orden').DataType    := ftString ;
    QryrFotografico.Params.ParamByName('orden').Value       := tsNumeroOrdenActa.Text;
    QryrFotografico.Params.ParamByName('Acta').DataType     := ftString ;
    QryrFotografico.Params.ParamByName('Acta').Value        := ComboActas.Text;
    QryrFotografico.Open ;

    rxReporteFotografico.EmptyTable;
    While NOT QryrFotografico.Eof Do
    Begin
        rxReporteFotografico.Append ;
        rxReporteFotografico.FieldValues['sContrato']         := QryrFotografico.FieldValues['sContrato'] ;
        rxReporteFotografico.FieldValues['iImagen']           := QryrFotografico.FieldValues['iImagen'] ;
        rxReporteFotografico.FieldValues['bImagen']           := QryrFotografico.FieldValues['bImagen'] ;
        rxReporteFotografico.FieldValues['sDescripcion']      := QryrFotografico.FieldValues['DescripcionOrden'] ;
        rxReporteFotografico.FieldValues['sDescripcionCorta'] := QryrFotografico.FieldValues['sDescripcionCorta'] ;
        rxReporteFotografico.FieldValues['mDescripcion']      := QryrFotografico.FieldValues['mDescripcion'] ;
        rxReporteFotografico.FieldValues['sTituloOrden']      := QryrFotografico.FieldValues['sIdFolio'] ;
        rxReporteFotografico.FieldValues['ContratoPrincipal'] := QryrFotografico.FieldValues['ContratoPrincipal'] ;
        rxReporteFotografico.FieldValues['sNumeroOrden']      := QryrFotografico.FieldValues['sNumeroOrden'] ;
        rxReporteFotografico.FieldValues['dIdFecha']          := QryrFotografico.FieldValues['dIdFecha'] ;
        rxReporteFotografico.FieldValues['sNumeroActividad']  := QryrFotografico.FieldValues['sNumeroActividad'] ;
        rxReporteFotografico.FieldValues['sFasePartida']      := QryrFotografico.FieldValues['sFasePartida'] ;
        rxReporteFotografico.FieldValues['sIdFolio']          := QryrFotografico.FieldValues['sIdFolio'] ;
        rxReporteFotografico.FieldValues['sIdPlataforma']     := QryrFotografico.FieldValues['sIdPlataforma'] ;
        rxReporteFotografico.Post ;
        QryrFotografico.Next
    End;

    rReporte.PreviewOptions.MDIChild := True ;
    rReporte.PreviewOptions.Modal := False ;
    rReporte.PreviewOptions.Maximized := lCheckMaximized () ;
    rReporte.PreviewOptions.ShowCaptions := False ;
    rReporte.Previewoptions.ZoomMode := zmPageWidth ;
    if chk4img.Checked then
    begin
      rReporte.LoadFromFile(global_files + 'ActaFotografica_4.fr3');
    end
    else
    begin
      rReporte.LoadFromFile(global_files + 'ActaFotografica_2.fr3');
    end;

    rReporte.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));

    if chk4.Checked then
    begin
      if not FileExists(global_files + 'ActaFotografica_4.fr3') then
        showmessage('El archivo de reporte ActaFotografica.fr3 no existe, notifique al administrador del sistema');
    end
    else
    begin
      if not FileExists(global_files + 'ActaFotografica_2.fr3') then
        showmessage('El archivo de reporte ActaFotografica2.fr3 no existe, notifique al administrador del sistema');
    end;
  end;    
end;

procedure TfrmReportePeriodo.cmdNuevoClick(Sender: TObject);
begin
    ComboActas.SetFocus;
    ComboActas.Text := '';
    cmdNuevo.Enabled := False;
    cmdNuevo.Caption := ' Enter';
end;

procedure TfrmReportePeriodo.ComboActasEnter(Sender: TObject);
begin
    comboActas.color := global_color_entrada;
    if tsNumeroOrdenActa.Text = '' then
    begin
        messageDLG('Seleccione un frente de Trabajo!', mtInformation, [mbOk], 0);
        tsNumeroOrdenActa.SetFocus;
    end;
end;

procedure TfrmReportePeriodo.ComboActasExit(Sender: TObject);
begin
    comboActas.color := global_color_salida;
    fotografico_Acta.Active := False;
    fotografico_acta.ParamByName('Contrato').AsString := global_contrato;
    fotografico_acta.ParamByName('Orden').AsString    := tsNumeroOrdenActa.Text;
    fotografico_acta.ParamByName('Acta').AsString     := ComboActas.Text;
    fotografico_acta.Open;
end;

procedure TfrmReportePeriodo.ComboActasKeyPress(Sender: TObject; var Key: Char);
var
   i : integer;
begin
    if cmdNuevo.Enabled = False then
    begin
        if key = #13  then
        begin
            //Primero Verificamos si existe el Acta fotografica..
            connection.QryBusca.Active := False;
            connection.QryBusca.SQL.Clear;
            connection.QryBusca.SQL.Add('select sActaFotografica from reportefotografico_acta where sContrato =:Contrato and sNumeroOrden =:Orden and sActaFotografica =:Acta');
            connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
            connection.QryBusca.ParamByName('Orden').AsString    := tsNumeroOrdenActa.Text;
            connection.QryBusca.ParamByName('Acta').AsString     := ComboActas.Text;
            connection.QryBusca.Open;

            if connection.QryBusca.RecordCount > 0 then
            begin
                messageDLG('El Nombre de Acta Fotografica "'+ComboActas.text+'" ya Existe! ', mtInformation, [mbOk], 0);
                exit;
            end
            else
               ComboActas.Items.Add(ComboActas.Text);
            if comboActas.Text = '' then
            begin
                messageDLG('Escriba un Nombre para la Acta Fotografica!', mtInformation,[mbOk],0);
                comboActas.SetFocus;
            end;
            cmdNuevo.Enabled := True;
            cmdNuevo.Caption := 'Nuevo';
        end;
    end;
end;

procedure TfrmReportePeriodo.cxCheckBox1Click(Sender: TObject);
begin
  if cxCheckBox1.Checked then begin
    chkCaratula.Checked := True;
    chkOficio.Checked := True;
    chk1.Checked := True;
    chk2.Checked := True;
    chk4.Checked := True;
    chk5.Checked := True;
    chk6.Checked := True;
    chk7.Checked := True;
    chk8.Checked := True;
  end
  else begin
    chkCaratula.Checked := False;
    chkOficio.Checked := False;
    chk1.Checked := False;
    chk2.Checked := False;
    chk4.Checked := False;
    chk5.Checked := False;
    chk6.Checked := False;
    chk7.Checked := False;
    chk8.Checked := False;
  end;

end;

procedure TfrmReportePeriodo.dbFiltroEnter(Sender: TObject);
begin
    dbFiltro.Color := global_color_entrada
end;

procedure TfrmReportePeriodo.dbFiltroExit(Sender: TObject);
begin
      dbFiltro.Color := global_color_salida
end;

procedure TfrmReportePeriodo.dbFiltroKeyPress(Sender: TObject;
  var Key: Char);
begin
      if key = #8 then
         dbFiltro.KeyValue := null;
end;

procedure TfrmReportePeriodo.ActividadesxOrdenCalcFields(
  DataSet: TDataSet);
var
    sTipoMoneda,
    sCalculoMoneda : string;
begin
     if chkMoneda.Checked then
        sTipoMoneda := 'b.dCantidad * a.dVentaMN'
     else
         sTipoMoneda := 'b.dCantidad * a.dVentaDLL';

     If ActividadesxOrden.FieldValues['sWbs'] <> Null Then
         ActividadesxOrdensWbsSpace.Text := espaces (ActividadesxOrden.FieldValues['iNivel']) + ActividadesxOrden.FieldValues['sWbs'] ;

     If ActividadesxOrden.FieldValues['sTipoActividad'] = 'Actividad' Then
     Begin
          Connection.qryBusca2.Active := False ;
          Connection.qryBusca2.SQL.Clear ;
          Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad) as Instalado From bitacoradeactividades b ' +
                                        'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.dIdFecha < :Fecha And b.sWbs = :Wbs And b.sNumeroActividad = :Actividad ' +
                                        'Group By b.sWbs, b.sNumeroActividad') ;
          Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato ;
          Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden3.Text ;
          Connection.QryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
          Connection.QryBusca2.Params.ParamByName('Fecha').Value := tdFechaInicial.Date ;
          Connection.QryBusca2.Params.ParamByName('Wbs').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Wbs').Value := ActividadesxOrden.FieldValues['sWbs'] ;
          Connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad'] ;
          Connection.qryBusca2.Open ;
          If Connection.qryBusca2.RecordCount > 0 then
               ActividadesxOrdendAcumuladoAnterior.Value := Connection.qryBusca2.FieldValues['Instalado']
          Else
               ActividadesxOrdendAcumuladoAnterior.Value := 0 ;

          Connection.qryBusca2.Active := False ;
          Connection.qryBusca2.SQL.Clear ;
          Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad) as Instalado From bitacoradeactividades b ' +
                                        'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF And b.sWbs = :Wbs And b.sNumeroActividad = :Actividad ' +
                                        'Group By b.sWbs, b.sNumeroActividad') ;
          Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato ;
          Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden3.Text ;
          Connection.QryBusca2.Params.ParamByName('FechaI').DataType := ftDate ;
          Connection.QryBusca2.Params.ParamByName('FechaI').Value := tdFechaInicial.Date ;
          Connection.QryBusca2.Params.ParamByName('FechaF').DataType := ftDate ;
          Connection.QryBusca2.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
          Connection.QryBusca2.Params.ParamByName('Wbs').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Wbs').Value := ActividadesxOrden.FieldValues['sWbs'] ;
          Connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad'] ;
          Connection.qryBusca2.Open ;
          If Connection.qryBusca2.RecordCount > 0 then
               ActividadesxOrdendCantidadPeriodo.Value := Connection.qryBusca2.FieldValues['Instalado']
          Else
               ActividadesxOrdendCantidadPeriodo.Value := 0 ;

         ActividadesxOrdendAcumulado.Value := ActividadesxOrdendCantidadPeriodo.Value + ActividadesxOrdendAcumuladoAnterior.Value ;
         if chkMoneda.Checked then
         begin
              ActividadesxOrdendTotal.Value := ActividadesxOrdendCantidadPeriodo.Value * ActividadesxOrdendVentaMN.Value ;
              ActividadesxOrdendTotalAcumulado.Value := ActividadesxOrdendAcumulado.Value * ActividadesxOrdendVentaMN.Value ;
         end
         else
         begin
              ActividadesxOrdendTotal.Value := ActividadesxOrdendCantidadPeriodo.Value * ActividadesxOrdendVentaDLL.Value ;
              ActividadesxOrdendTotalAcumulado.Value := ActividadesxOrdendAcumulado.Value * ActividadesxOrdendVentaDLL.Value ;
         end;
     End
     Else
     Begin
         // Es Paquete ...

          Connection.qryBusca2.Active := False ;
          Connection.qryBusca2.SQL.Clear ;
          Connection.qryBusca2.SQL.Add('Select Sum('+sTipoMoneda+') as dTotal From bitacoradeactividades b ' +
                                       'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sNumeroOrden = b.sNumeroOrden and a.sWbs = b.sWbs And a.sNumeroActividad = b.sNumeroActividad and a.sIdConvenio = :Convenio) ' +
                                       'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.dIdFecha >= :FechaI and b.dIdFecha <= :FechaF And b.sWbs Like :Wbs ' +
                                       'Group By b.sContrato') ;
          Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato ;
          Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden3.Text ;
          Connection.QryBusca2.Params.ParamByName('convenio').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('convenio').Value := global_convenio ;
          Connection.QryBusca2.Params.ParamByName('FechaI').DataType := ftDate ;
          Connection.QryBusca2.Params.ParamByName('FechaI').Value := tdFechaInicial.Date ;
          Connection.QryBusca2.Params.ParamByName('FechaF').DataType := ftDate ;
          Connection.QryBusca2.Params.ParamByName('FechaF').Value := tdFechaFinal.Date ;
          Connection.QryBusca2.Params.ParamByName('Wbs').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Wbs').Value := Trim(ActividadesxOrden.FieldValues['sWbs']) + '.%' ;
          Connection.qryBusca2.Open ;
          If Connection.qryBusca2.RecordCount > 0 then
               ActividadesxOrdendTotal.Value := connection.QryBusca2.FieldValues['dTotal']
          Else
               ActividadesxOrdendTotal.Value := 0 ;

          Connection.qryBusca2.Active := False ;
          Connection.qryBusca2.SQL.Clear ;
          Connection.qryBusca2.SQL.Add('Select Sum('+sTipoMoneda+') as dTotal From bitacoradeactividades b ' +
                                       'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sNumeroOrden = b.sNumeroOrden and a.sWbs = b.sWbs And a.sNumeroActividad = b.sNumeroActividad and a.sIdConvenio = :Convenio) ' +
                                       'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.dIdFecha <= :Fecha And b.sWbs Like :Wbs ' +
                                       'Group By b.sContrato') ;
          Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato ;
          Connection.QryBusca2.Params.ParamByName('Orden').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden3.Text ;
          Connection.QryBusca2.Params.ParamByName('convenio').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('convenio').Value := global_convenio ;
          Connection.QryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
          Connection.QryBusca2.Params.ParamByName('Fecha').Value := tdFechaFinal.Date ;
          Connection.QryBusca2.Params.ParamByName('Wbs').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Wbs').Value := Trim(ActividadesxOrden.FieldValues['sWbs']) + '.%' ;
          Connection.qryBusca2.Open ;
          If Connection.qryBusca2.RecordCount > 0 then
               ActividadesxOrdendTotalAcumulado.Value := connection.QryBusca2.FieldValues['dTotal']
          Else
               ActividadesxOrdendTotalAcumulado.Value := 0 ;

     End ;

end;

procedure TfrmReportePeriodo.OrdenarFotos(sParamOrden: string);
var
   idAuxiliar, idAuxiliar2 : integer;
   SavePlace   : TBookmark;
begin
    if fotografico_acta.RecordCount > 0 then
    begin
        if sParamOrden = 'Arriba' then
        begin
            idAuxiliar2 := fotografico_acta.FieldValues['iImagen'];
            fotografico_acta.Prior;

            idAuxiliar  := fotografico_acta.FieldValues['iImagen'];
            fotografico_acta.Next;
        end;

        if sParamOrden = 'Abajo' then
        begin
            idAuxiliar2 := fotografico_acta.FieldValues['iImagen'];
            fotografico_acta.Next;

            idAuxiliar  := fotografico_acta.FieldValues['iImagen'];
            fotografico_acta.Prior;
        end;
        //Colocamos un id mayor para evitar duplicidad..
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE reportefotografico_acta SET iImagen = :DiarioNuevo ' +
                                    'where sContrato = :contrato and sNumeroOrden =:Orden And iImagen = :diario ');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := Global_Contrato;
        Connection.zCommand.Params.ParamByName('Orden').DataType    := ftString;
        Connection.zCommand.Params.ParamByName('Orden').value       := tsNumeroOrdenActa.Text;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := idAuxiliar2;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').DataType := ftInteger;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').value    := idAuxiliar + 500;
        connection.zCommand.ExecSQL;

        //Ahora actualizamos el item mayor
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE reportefotografico_acta SET iImagen = :DiarioNuevo ' +
                                    'where sContrato = :contrato And sNumeroOrden =:Orden And iImagen = :diario ');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := Global_Contrato;
        Connection.zCommand.Params.ParamByName('Orden').DataType    := ftString;
        Connection.zCommand.Params.ParamByName('Orden').value       := tsNumeroOrdenActa.Text;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := idAuxiliar;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').DataType := ftInteger;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').value    := idAuxiliar2;
        connection.zCommand.ExecSQL;

         //Ahora actualizamos el item alterado
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE reportefotografico_acta SET iImagen = :DiarioNuevo ' +
                                    'where sContrato = :contrato and sNumeroOrden =:Orden And iImagen = :diario ');
        Connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('contrato').value    := Global_Contrato;
        Connection.zCommand.Params.ParamByName('Orden').DataType    := ftString;
        Connection.zCommand.Params.ParamByName('Orden').value       := tsNumeroOrdenActa.Text;
        Connection.zCommand.Params.ParamByName('diario').DataType   := ftInteger;
        Connection.zCommand.Params.ParamByName('diario').value      := idAuxiliar + 500;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').DataType := ftInteger;
        Connection.zCommand.Params.ParamByName('DiarioNuevo').value    := idAuxiliar;
        connection.zCommand.ExecSQL;

        if sParamOrden = 'Arriba' then
           fotografico_acta.Prior
        else
           fotografico_acta.Next;

        SavePlace := fotografico_acta.GetBookmark;
        fotografico_acta.Refresh;
        fotografico_acta.GotoBookmark(SavePlace);
        fotografico_acta.FreeBookmark(SavePlace);

    end;
end;

Procedure TfrmReportePeriodo.FormatoNormal(var Excel: Variant;Cadena:string; Align: Integer;Negrita,Ajustar:Boolean);
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

Procedure TfrmReportePeriodo.FormatoNormal(var Excel: Variant;Cadena:Variant; Align: Integer;Negrita,Ajustar:Boolean;Formato:String;Column:Integer=0);
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

procedure TfrmReportePeriodo.AjustarTexto(var rangoE: Variant;TotalR:Integer);
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

procedure TfrmReportePeriodo.imprimirLibro;
Var
  Libro, Excel, Hoja: Variant;
  sFolios: String;
  iFila, iColumna, iHojas, iCounter, iLoop: Integer;
  NombreDelExcel, SQLExtra: String;
  TempPath: String;
  Fs: TStream;
  Pic : TJpegImage;
  imgAux: TImage;
  dContrato_Inicio, dContrato_Final: TDateTime;
  TmpName: String;
  QryImagen : TZQuery;

  sColumna : string;
  sFila : string;

  sumaMN : Double;
  sumaUSD : Double;
  i, iAnt : integer;
  finPartidas : integer;
  texto : string;
  sInicio : string;

  ficha : Integer;

  buscaContenido : TZQuery;
  haymasportadas : Boolean;

  zqFolio: tzreadonlyquery;

  {$REGION 'PROCEDIMIENTOS UTILES'}
  procedure EstablecerContornos;
  begin
    Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeLeft].Weight := xlThin;
    Excel.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeTop].Weight := xlThin;
    Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;
    Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeRight].Weight := xlThin;
  end;

  procedure DarFormato(combinar : Boolean ; alinear : string ; Negritas : Boolean ; contornos : Boolean);
  begin
    if combinar then
    begin
      Excel.Selection.MergeCells := True;
    end;
    
    if alinear = 'centro' then
    begin
      Excel.Selection.HorizontalAlignment := xlCenter;
    end

    else if alinear = 'der' then
    begin
      Excel.Selection.HorizontalAlignment := xlRight;
    end

    else if alinear = 'izq' then
    begin
      Excel.Selection.HorizontalAlignment  := xlLeft;
    end;

    if Negritas then
    begin
      Excel.Selection.Font.Bold := True;
    end;

    if contornos then
    begin
      EstablecerContornos;
    end;

    Excel.Selection.VerticalAlignment := xlCenter;
    Excel.Selection.Font.Size := 7;
  end;

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

  procedure imprimirPortadaContenido (tituloportada : string; nombre : string);
  var
    x : integer;
  begin
      Libro.Sheets.Add;
      if nombre <> '' then begin
        Excel.WorkBooks[1].WorkSheets[1].Name := nombre;
      end
      else begin
        Excel.WorkBooks[1].WorkSheets[1].Name := 'no asignado';
      end;

      iColumna := 1;
      Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 0.75;
      for x := 1 to 8 - 1 do
      begin
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 11;
      end;
      Inc(iColumna);
      Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 0.75;
    
      //Imagen Derecha
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then begin
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
        Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 365, 1, 80, 62);


      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select mDescripcion  from actividadesxorden '+
                                  'where sContrato = :contrato '+
                                  'and sIdConvenio = :convenio '+
                                  'and sNumeroOrden = :folio '+
                                  'and sTipoActividad = "Actividad" '+
                                  'order by dFechaInicio');

      connection.QryBusca.Params.ParamByName('contrato').DataType := ftString;
      connection.QryBusca.Params.ParamByName('convenio').DataType := ftString;
      connection.QryBusca.Params.ParamByName('folio').DataType    := ftString;
      connection.QryBusca.Params.ParamByName('contrato').Value    := global_contrato;
      connection.QryBusca.Params.ParamByName('convenio').Value    := global_convenio;
      connection.QryBusca.Params.ParamByName('folio').Value       := dbFolio01.Text;
      connection.QryBusca.Open;

      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Clear;
      connection.QryBusca2.SQL.Add('select sNombre, sDireccion1, mCuerpoFolio from configuracion where sContrato = :contrato');
      connection.QryBusca2.Params.ParamByName('contrato').DataType := ftString;
      connection.QryBusca2.Params.ParamByName('contrato').Value := global_contrato;
      connection.QryBusca2.Open;

      if connection.QryBusca2.RecordCount > 0 then begin
        Excel.activeWindow.DisplayGridlines := false;
    
        Excel.Range['D2:F2'].Select;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Font.Size := 12;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Value := connection.QryBusca2.FieldByName('sNombre').AsString;

        Excel.Rows[3].RowHeight := 37.5;
        Excel.Range['D3:F3'].Select;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Font.Size := 8;
        Excel.Selection.Value := connection.QryBusca2.FieldByName('sDireccion1').AsString;
        Excel.Selection.WrapText := True;
      end;
      if connection.QryBusca.RecordCount > 0 then begin
        for x := 4 to 11 do
        begin
          Excel.Rows[x].RowHeight := 15;
        end;
        Excel.Rows[12].RowHeight := 50;
        Excel.Rows[13].RowHeight := 15;
        Excel.Rows[14].RowHeight := 50;
        Excel.Rows[15].RowHeight := 15;
        Excel.Rows[16].RowHeight := 50;
        Excel.Rows[17].RowHeight := 60;

        Excel.Range['B17:H17'].Select;
        Excel.Selection.MergeCells := True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment := xlCenter;
        Excel.Selection.Interior.ColorIndex := 16;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Font.Size := 16;
        Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeLeft].Weight := xlThin;
        Excel.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeTop].Weight := xlThin;
        Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;
        Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlEdgeRight].Weight := xlThin;
        Excel.Selection.Value := tituloPortada;
      end;
  end;

  procedure HojaCarta(VarEx:Variant);
  begin
      //Ajustamos la hoja para mejor presentacion
      VarEx.ActiveWindow.View := 2;
      VarEx.ActiveSheet.PageSetup.LeftMargin := 0.7;
      VarEx.ActiveSheet.PageSetup.RightMargin := 0.7;
      VarEx.ActiveSheet.PageSetup.TopMargin := 0.75;
      VarEx.ActiveSheet.PageSetup.BottomMargin := 0.75;
      VarEx.ActiveSheet.PageSetup.HeaderMargin := 0.3 ;
      VarEx.ActiveSheet.PageSetup.FooterMargin := 0.3;
      VarEx.ActiveSheet.PageSetup.PrintHeadings := False;
      VarEx.ActiveSheet.PageSetup.PrintGridlines := False;
      VarEx.ActiveSheet.PageSetup.PrintQuality := 600;
      VarEx.ActiveSheet.PageSetup.CenterHorizontally := False;
      VarEx.ActiveSheet.PageSetup.CenterVertically := False;
      VarEx.ActiveSheet.PageSetup.Draft := False;
      VarEx.ActiveSheet.PageSetup.PaperSize := 1;
      VarEx.ActiveSheet.PageSetup.BlackAndWhite := False;
      VarEx.ActiveSheet.PageSetup.Zoom := False;
      VarEx.ActiveSheet.PageSetup.FitToPagesWide := 1;
      VarEx.ActiveSheet.PageSetup.FitToPagesTall := 1;
      VarEx.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
      VarEx.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := True;
  end;
  {$ENDREGION}
begin
  if dbFolio01.Text <> '' then
  begin
    buscaContenido := TZQuery.Create(nil);


    zqFolio := tzreadonlyquery.create(nil);
    try
      {$REGION 'Consultas'}
      //Contenido del libro
      buscaContenido.Connection := connection.zConnection;
      buscaContenido.Active := False;
      buscaContenido.SQL.Clear;
      buscaContenido.SQL.Add('select * from contenidonotacampo order by iOrden');
      buscaContenido.Open;

      //datos del folio
      zqfolio.connection :=  connection.zConnection;
      zqfolio.sql.clear;
      zqfolio.sql.text := 'SELECT * FROM Ordenesdetrabajo WHERE ' +
                                      'sContrato = :Contrato AND sNumeroOrden = :Folio ';
      zqfolio.Params.ParamByName('Contrato').AsString := Global_Contrato;
      zqfolio.Params.ParamByName('Folio').AsString := dbfolio01.KeyValue;
      zqfolio.Open;
      {$ENDREGION}

      if zqfolio.recordcount < 1 then
        raise exception.create('no se pudo localizar el folio posiblemente necesite recargar los datos de la ventana.');

      if buscaContenido.RecordCount > 0 then
      begin
        buscaContenido.First;
        haymasportadas := True;
      end
      else
      begin
        haymasportadas := False;
      end;

      {$REGION 'ACCEDER A EXCEL'}
      NombreDelExcel := PGetTempDir + 'TEMP~' + PNombreAleatorio(3) + 'NotaDeCampo.xls';
      Try
        Excel := CreateOleObject('Excel.Application');
      Except
        On E: Exception do begin
          FreeAndNil(Excel);
          ShowMessage(E.Message);
          Exit;
        end;
      End;

      Excel.Visible := True;
      Excel.DisplayAlerts:= False;
      Libro := Excel.Workbooks.Add;


    //IMPRIME LOS FOLIOS
    {$ENDREGION}

      ficha := 1;

      if chkCaratula.Checked then
      begin
      {$REGION 'OFICIO'}
        Libro.Sheets.Add;
        Excel.WorkBooks[1].WorkSheets[1].Name := 'PORTADA';

        iColumna := 1;
        iFila := 1;

        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 3.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 5.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 10.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 8.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7.43;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9.14;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 20.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 3.29;
        Inc(iColumna);

        {$REGION 'IMAGENES DE CABECERA'}
        //Imagen Izquierda
        Try
          TmpName := '';
          imgAux := TImage.Create(nil);
          if TmpName='' then begin
      //      GetTempPath(SizeOf(TempPath), TempPath);
            TempPath := ExtractFilePath(Application.Exename);
            TmpName:=TempPath +'imgtempSln1'+formatdatetime('dddddd hhnnss',now)+'.jpg';
            fs := Connection.contrato.CreateBlobStream(Connection.contrato.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
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
          Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 1, 150, 85);
        //Imagen Derecha
        Try
          TmpName := '';
          imgAux := TImage.Create(nil);
          if TmpName='' then begin
      //      GetTempPath(SizeOf(TempPath), TempPath);
            TempPath := ExtractFilePath(Application.Exename);
            TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
        Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 375, 1, 90, 85);
        //Texto de Cabecera
        Excel.Range['A1:J6'].Select;
        PFormatosExcel_H2(Excel, 0, True, 10);

        Excel.activeWindow.DisplayGridlines := false;

        {$ENDREGION}

        {$REGION 'CABECERA'}
           {
          Connection.QryBusca.SQL.Clear;
          Connection.QryBusca.SQL.Text := 'SELECT * ' +
                                          '		FROM ' +
                                          'Ordenesdetrabajo ' +
                                          '		WHERE ' +
                                          'sContrato = :Contrato AND sNumeroOrden = :Folio ';
          Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
          Connection.QryBusca.Params.ParamByName('Folio').AsString := dbfolio01.KeyValue;
          Connection.QryBusca.Open;
            }
          iFila := 7;

          Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 65, True, 12, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.Value := Connection.contrato.FieldByName('mComentarios').AsString;
          Inc (iFila,3);

          Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.Value := 'CONTRATO:' + Global_Contrato_Barco;
          Inc (iFila,3);

          Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 57, True, 12, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.Value := Connection.contrato.FieldByName('sTitulo').AsString;
          Inc (iFila);

          Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 57, True, 12, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.Value := 'TRABAJOS REALIZADOS EN INSTALACIÓN:';

          Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 57, True, 11, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.Value := 'PLATAFORMA:' + zqfolio.FieldByName('sIdPlataforma').AsString ;
          Inc (iFila,2);

          Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 15, True, 12, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.Value := 'FOLIO:' ;
          Inc (iFila);

          Excel.Range[ColumnaNombre(1)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 15, True, 12, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.Value := zqfolio.FieldByName('sIdFolio').AsString;

          QryImagen:=TZQuery.Create (Self);
          QryImagen.connection:= connection.zConnection;
          QryImagen.SQL.Clear;
          QryImagen.SQL.Text:= 'SELECT * ' +
                                          '		FROM ' +
                                          'Configuracion ' +
                                          '		WHERE ' +
                                          'sContrato = :Contrato';
          QryImagen.ParamByName('Contrato').AsString := Global_Contrato_Barco;
          QryImagen.Open;

          Try
          TmpName := '';
          imgAux := TImage.Create(nil);
          if TmpName='' then begin
            TempPath := ExtractFilePath(Application.Exename);
            TmpName:=TempPath +'imgtempSln1'+formatdatetime('dddddd hhnnss',now)+'.jpg';
            fs := QryImagen.CreateBlobStream(QryImagen.FieldByName('bImagenAux1'), bmRead);
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
          Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 20, 390, 410, 245);

        Excel.Range[ColumnaNombre(1)+IntToStr(18)+':'+ColumnaNombre(10)+IntToStr(36)].Select;
        PFormatosExcel_H2(Excel, 12, True, 8, clBlack, 'Arial');

        Excel.Range[ColumnaNombre(2)+IntToStr(37)+':'+ColumnaNombre(9)+IntToStr(46)].Select;
        PFormatosExcel_H2(Excel, 12, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := zqfolio.FieldByName('mDescripcion').AsString;

        {$ENDREGION}

        //Ajustamos la hoja para mejor presentacion
        HojaCarta(Excel);

       {$ENDREGION}
      end;

      if chkOficio.Checked then
      begin
      {$REGION 'OFICIO'}
        if haymasportadas then begin
          if not buscaContenido.Eof then begin
            imprimirPortadaContenido(buscaContenido.FieldByName('sDescripcion').AsString , buscaContenido.FieldByName('sNombrePortada').AsString);
            buscaContenido.Next;
            Inc(ficha);
          end;
        end;


        Libro.Sheets.Add;
        Excel.WorkBooks[1].WorkSheets[1].Name := 'OFICIO';
        Inc(ficha);


        iColumna := 1;
        iFila := 1;

        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 3.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 5.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 10.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 8.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7.43;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9.14;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 20.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 3.29;
        Inc(iColumna);

      {$REGION 'IMAGENES DE CABECERA'}
      //Imagen Derecha
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then begin
    //      GetTempPath(SizeOf(TempPath), TempPath);
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
        Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 375, 1, 85, 80);
      //Texto de Cabecera
      Excel.Range['A1:J1'].Select;
      PFormatosExcel_H2(Excel, 0, True, 10);

      Excel.activeWindow.DisplayGridlines := false;

      {$ENDREGION}

      {$REGION 'CABECERA'}

        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Text := 'SELECT * ' +
                                        '		FROM ' +
                                        'Configuracion' +
                                        '		WHERE ' +
                                        'sContrato = :Contrato ';
        Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
      //  Connection.QryBusca.Params.ParamByName('Folio').AsString := dbfolio01.KeyValue;
        Connection.QryBusca.Open;


        iFila := 2;
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('sNombre').AsString;
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 38, True, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('sDireccion1').AsString + ' '+#10
        + Connection.QryBusca.FieldByName('sDireccion2').AsString +' '+#10+
        Connection.QryBusca.FieldByName('sCiudad').AsString; {+ Connection.QryBusca.FieldByName('sDireccion3').AsString +}
        Excel.Selection.ReadingOrder := xlContext;
        Excel.Selection.WrapText := True;
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('sCiudad').AsString+' a '+FormatDateTime('dd'+'" de "'+'mmmm'+'" de "'+ 'yyyy',zqfolio.FieldByName('dffprogramado').AsDateTime +15 );
        PFormatosExcel_SoloBorde(Excel);
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlright;
        Excel.Selection.Value := 'OFICIO: ' + zqfolio.FieldByName('soficioautorizacion').asstring;
        PFormatosExcel_SoloBorde(Excel);
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := Xlleft;
        Excel.Selection.Value := 'PEMEX EXPLORACION PRODUCCION';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila+3)].Select;
        PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := Xlleft;
        Excel.Selection.formula :=  Connection.contrato.FieldByName('mComentarios').AsString;
        Inc (iFila,5);
        {
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := Xlleft;
        Excel.Selection.Value := 'GERENCIA DE MANTENIMIENTO INTEGRAL MARINO';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := Xlleft;
        Excel.Selection.Value := 'COORDINACION DE MANTENIMIENTO INTEGRAL, LITORAL, GTDH, GOLFO NORTE Y SOPORTE A';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := Xlleft;
        Excel.Selection.Value := 'PERFORACION.';
        Inc (iFila,2);
          }
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := Xlleft;
        if length(trim(connection.QryBusca.FieldByName('sadministradorcontrato').asstring)) > 0 then
          Excel.Selection.Value := 'ATN '+uppercase(connection.QryBusca.FieldByName('sadministradorcontrato').asstring);
        PFormatosExcel_SoloBorde(Excel);
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := Xlleft;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('spuestoadministradorcontrato').AsString+ ' No. '+global_contrato_barco;
        PFormatosExcel_SoloBorde(Excel);
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False,8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := XlRight;
        Excel.Selection.Value := 'Asunto: Entrega de Reporte Final';
        PFormatosExcel_SoloBorde(Excel);
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlleft;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('mcuerpofolio').AsString;
        PFormatosExcel_SoloBorde(Excel);
        Excel.Rows[ifila].RowHeight := 50;
        Inc (iFila,3);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 10, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'FOLIO';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 10, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := zqfolio.FieldByName('snumeroorden').asstring;
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'PROGRAMA:';
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 59, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := zqfolio.FieldByName('mdescripcion').asstring;
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'Instalación:';

        Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'Periodo de ejecución:';
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'PLATAFORMA: ' + zqfolio.FieldByName('sidplataforma').asstring;

        Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 9, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;

        if FormatDateTime('mm',zqfolio.FieldByName('dfiprogramado').AsDateTime) = FormatDateTime('mm',zqfolio.FieldByName('dffprogramado').AsDateTime)  then
          Excel.Selection.Value := 'DEL '+FormatDateTime('dd',zqfolio.FieldByName('dfiprogramado').AsDateTime)+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqfolio.FieldByName('dffprogramado').AsDateTime))
        else
          Excel.Selection.Value := 'DEL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqfolio.FieldByName('dfiprogramado').AsDateTime))+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqfolio.FieldByName('dffprogramado').AsDateTime));

        Inc (iFila,3);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := Xlleft;
        Excel.Selection.Value := 'Sin otro particular al respecto, quedo a su disposición para cualquier aclaración.';
        PFormatosExcel_SoloBorde(Excel);
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := XlCenter;
        Excel.Selection.Value := 'Atentamente';
        PFormatosExcel_SoloBorde(Excel);
        Inc (iFila,4);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False,8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('srepresentante').AsString;
        PFormatosExcel_SoloBorde(Excel);
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('srepresentanteobra').AsString;
        PFormatosExcel_SoloBorde(Excel);
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, False, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlleft;
        Excel.Selection.Value := 'c.c.p. Archivo';
        Inc (iFila);

    {$ENDREGION}

       {$ENDREGION}
       HojaCarta(Excel);
      end;

      if chk1.Checked then
      begin
      {$REGION 'OFICIO'}
        if haymasportadas then begin
          if not buscaContenido.Eof then begin
            imprimirPortadaContenido(buscaContenido.FieldByName('sDescripcion').AsString , buscaContenido.FieldByName('sNombrePortada').AsString);
            buscaContenido.Next;
            Inc(ficha);
          end;
        end;

        Libro.Sheets.Add;
        Excel.WorkBooks[1].WorkSheets[1].Name := 'OFICIO2';
        Inc(ficha);

        iColumna := 1;
        iFila := 1;

        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 3.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 5.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 10.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 8.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7.43;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9.14;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 20.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 3.29;
        Inc(iColumna);

    {$REGION 'IMAGENES DE CABECERA'}
      //Imagen Derecha
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then begin
    //      GetTempPath(SizeOf(TempPath), TempPath);
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
        Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 375, 1, 90, 85);
      //Texto de Cabecera
      Excel.Range['A1:J1'].Select;
      PFormatosExcel_H2(Excel, 0, True, 10);

      Excel.activeWindow.DisplayGridlines := false;

      {$ENDREGION}

       {$REGION 'CABECERA'}

        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Text := 'SELECT * ' +
                                        '		FROM ' +
                                        'Configuracion' +
                                        '		WHERE ' +
                                        'sContrato = :Contrato ';
        Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
      //  Connection.QryBusca.Params.ParamByName('Folio').AsString := dbfolio01.KeyValue;
        Connection.QryBusca.Open;



        iFila := 2;
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('sNombre').AsString;
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 38, True, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('sDireccion1').AsString + '                                                                                                                                    '
        + Connection.QryBusca.FieldByName('sDireccion2').AsString + '                                                                                                                                    '
        + Connection.QryBusca.FieldByName('sDireccion3').AsString + Connection.QryBusca.FieldByName('sCiudad').AsString;
        Excel.Selection.ReadingOrder := xlContext;
        Excel.Selection.WrapText := True;
        Inc (iFila,3);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 86, True, 10, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := Connection.contrato.FieldByName('mComentarios').AsString;
        Excel.Rows[iFila].RowHeight := 86.25;
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'TRABAJOS REALIZADOS EN:';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'PLATAFORMA:' + zqfolio.fieldbyname('sidplataforma').asstring;
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'FOLIO';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := zqfolio.fieldbyname('snumeroorden').asstring;
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 31, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'PROGRAMA';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 82, True,9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := zqfolio.fieldbyname('mdescripcion').asstring;
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'Periodo de Ejecución';
        Inc (iFila,3);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;

        if FormatDateTime('mm',zqfolio.FieldByName('dfiprogramado').AsDateTime) = FormatDateTime('mm',zqfolio.FieldByName('dffprogramado').AsDateTime)  then
          Excel.Selection.Value := 'DEL '+FormatDateTime('dd',zqfolio.FieldByName('dfiprogramado').AsDateTime)+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqfolio.FieldByName('dffprogramado').AsDateTime))
        else
          Excel.Selection.Value := 'DEL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqfolio.FieldByName('dfiprogramado').AsDateTime))+' AL '+uppercase(FormatDateTime('dd'+'" DE "'+'mmmm'+'" DE "'+'yyyy',zqfolio.FieldByName('dffprogramado').AsDateTime));

        Inc (iFila,5);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 9, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'CONTRATO:'+global_contrato_barco;
        Inc (iFila,4);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 10, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;


        //Excel.Selection.Value := zqrfolio.FieldByName('sbarco').AsString;
        Excel.Selection.Value := '"B.P.D. ISLAND PIONEER"';
        Inc (iFila);

        {$ENDREGION}

        HojaCarta(Excel);

       {$ENDREGION}
      end;

      if chk2.Checked then
      begin
      {$REGION 'OFICIO'}
        if haymasportadas then begin
          if not buscaContenido.Eof then begin
            imprimirPortadaContenido(buscaContenido.FieldByName('sDescripcion').AsString , buscaContenido.FieldByName('sNombrePortada').AsString);
            buscaContenido.Next;
            Inc(ficha);
          end;
        end;

        Libro.Sheets.Add;
        Excel.WorkBooks[1].WorkSheets[1].Name := 'CONTENIDO (0)';
        Inc(ficha);

        iColumna := 1;
        iFila := 1;

        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 3.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 5.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 10.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 8.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7.43;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9.14;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 20.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 3.29;
        Inc(iColumna);

      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then begin
    //      GetTempPath(SizeOf(TempPath), TempPath);
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
        Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 375, 1, 90, 85);
      //Texto de Cabecera
      Excel.Range['A1:J1'].Select;
      PFormatosExcel_H2(Excel, 0, True, 10);

      Excel.activeWindow.DisplayGridlines := false;

        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Text := 'SELECT * ' +
                                        '		FROM ' +
                                        'Configuracion' +
                                        '		WHERE ' +
                                        'sContrato = :Contrato ';
        Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
      //  Connection.QryBusca.Params.ParamByName('Folio').AsString := dbfolio01.KeyValue;
        Connection.QryBusca.Open;

        iFila := 2;
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('sNombre').AsString;
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 38, True, 8, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := Connection.QryBusca.FieldByName('sDireccion1').AsString + '                                                                                                                                    '
        + Connection.QryBusca.FieldByName('sDireccion2').AsString + '                                                                                                                                    '
        +  Connection.QryBusca.FieldByName('sCiudad').AsString;
        Excel.Selection.ReadingOrder := xlContext;
        Excel.Selection.WrapText := True;
        Inc (iFila,5);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 86, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'CONTENIDO';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '1.- OFICIO DE SOLICITUD DE TRABAJOS DE "P.E.P" A EL "CONTRATISTA"';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '2.-';
        Inc (iFila);

           Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '3.-';
        Inc (iFila);

           Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 41, True, 12, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '4.-';
        Inc (iFila);
    
       {$ENDREGION}
      end;

      if chk4.Checked then
      begin
      {$REGION 'OFICIO'}
        if haymasportadas then begin
          if not buscaContenido.Eof then begin
            imprimirPortadaContenido(buscaContenido.FieldByName('sDescripcion').AsString , buscaContenido.FieldByName('sNombrePortada').AsString);
            buscaContenido.Next;
            Inc(ficha);
          end;
        end;

        Libro.Sheets.Add;
        Excel.WorkBooks[1].WorkSheets[1].Name := 'OFICIO ACTA(4)';
        Inc(ficha);

        sInicio := '';
        iFila := 1;
        iColumna := 1;

    
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 0.75;
        for i := 0 to 9 - 1 do
        begin
          Inc(iColumna);
          Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        end;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 2.86;
        //Imagen Derecha
        Try
          TmpName := '';
          imgAux := TImage.Create(nil);
          if TmpName='' then begin
      //      GetTempPath(SizeOf(TempPath), TempPath);
            TempPath := ExtractFilePath(Application.Exename);
            TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
          Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 365, 1, 80, 62);

        Excel.activeWindow.DisplayGridlines := false;
          //Texto de Cabecera


          Excel.Range['B2:J2'].Select;
          DarFormato(True, 'centro', True, false);
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Value := 'OCEANOGRAFIA, S.A. DE C.V.';

          Excel.Rows[3].RowHeight := 37.5;
          Excel.Range['E3:G3'].Select;
          DarFormato(True, 'centro', False, False);
          Excel.Selection.Font.Size := 8;
          Excel.Selection.Value := 'AV. 4 ORIENTE MANZANA D LOTE 3 '+
                                   'PUERTO IND. PESQUERO LAGUNA AZUL '+
                                   '24140 CIUDAD DEL CARMEN, CAMPECHE';
          Excel.Selection.WrapText := True;

          Excel.Range['B5:J5'].Select;
          DarFormato(True, 'der', False, False);
          Excel.Selection.Value := DateToStr(Now());

          for i := 6 to 9 do
          begin
            Excel.Rows[i].RowHeight:= 11.25;
          end;
          Excel.Range['B11:J11'].Select;
          Excel.Selection.Font.Size := 8;
          DarFormato(True, 'der', False, False);
          Excel.Selection.Value := 'Asunto: Acta de Entrega';

          Excel.Range['B13:J13'].Select;
          DarFormato(True, 'izq', True, False);
          Excel.Selection.Value := 'RESUMEN DE OBRA:';

          Excel.Range['B15:D15'].Select;
          DarFormato(True, 'izq', True, True);
          Excel.Selection.Value := 'PERIODO:';

          Excel.Range['B16:D16'].Select;
          DarFormato(True, 'izq', True, True);
          Excel.Selection.Value := 'CONTRATO:';

          Excel.Range['B17:D17'].Select;
          Excel.Rows[17].RowHeight := 50;
          DarFormato(True, 'izq', True, True);
          Excel.Selection.Value := 'DESCRIPCION:';

          Excel.Range['B18:D18'].Select;
          DarFormato(True, 'izq', True, True);
          Excel.Selection.Value := 'DESCRIPCION:';


          Excel.Range['E15:J15'].Select;
          DarFormato(True, 'izq', True, True);
          Excel.Selection.Value := 'DEL 03 DE DICIEMBRE DEL 2012 AL 02 DE DICIEMBRE DEL 2014:';

          Excel.Range['E16:J16'].Select;
          DarFormato(True, 'izq', True, True);
          Excel.Selection.Value := '428232833:';

          Excel.Range['E17:J17'].Select;
          Excel.Rows[17].RowHeight := 50;
          DarFormato(True, 'izq', True, True);
          Excel.Selection.Value := 'MANTENIMIENTO Y APOYO AL PERSONAL...:';

          Excel.Range['E18:J18'].Select;
          DarFormato(True, 'izq', True, True);
          Excel.Selection.Value := 'ISPSB2O323_J2H30003:';

          for i := 19 to 21 do
          begin
            Excel.Rows[i].RowHeight := 9;
          end;

          Excel.Range['B22:C22'].Select;
          DarFormato(True, 'izq', True, False);
          Excel.Rows[24].RowHeight := 11.25;
          Excel.Selection.Font.Size := 8;
          Excel.Selection.Value := 'PERIODO DE EJECUCION:';

          Excel.Range['B24:J24'].Select;
          Excel.Rows[24].RowHeight := 18.75;
          Excel.Selection.Font.Size := 8;
          DarFormato(True, 'izq', True, True);
          Excel.Selection.Value := 'DEL 18 AL 21 DE OCTUBRE';

          Excel.Range['B26:C26'].Select;
          DarFormato(True, 'izq', True, False);
          Excel.Rows[24].RowHeight := 11.25;
          Excel.Selection.Font.Size := 8;
          Excel.Selection.Value := 'RESUMEN DE COSTOS:';

          Excel.Range['B28:D28'].Select;
          DarFormato(True, 'centro', True, True);
          Excel.Selection.Font.Size := 9;
          Excel.Selection.Value := 'PARTIDA ANEXO C';

          Excel.Range['E28:H28'].Select;
          DarFormato(True, 'centro', True, True);
          Excel.Selection.Font.Size := 9;
          Excel.Selection.Value := 'DESCRIPCIÓN';

          Excel.Range['I28:J28'].Select;
          DarFormato(True, 'centro', True, True);
          Excel.Selection.Font.Size := 9;
          Excel.Selection.Value := 'IMPORTE TOTAL';
          Excel.Range['I29:I29'].Select;
          DarFormato(True, 'centro', True, True);
          Excel.Selection.Font.Size := 9;
          Excel.Selection.Value := 'M.N.';
          Excel.Range['J29:J29'].Select;
          DarFormato(True, 'centro', True, True);
          Excel.Selection.Font.Size := 9;
          Excel.Selection.Value := 'U.S.D.';

          sumaMN := 0;
          sumaUSD := 0;

          for i := 29 to 35 do
          begin
            Excel.Range['B' + IntToStr(i) + ':D' + IntToStr(i)].Select;
            DarFormato(True, 'centro', False, True);
            Excel.Selection.Value := IntToStr(i);
            Excel.Selection.Font.Size := 8;

            Excel.Range['E' + IntToStr(i) + ':H' + IntToStr(i)].Select;
            DarFormato(True, 'centro', False, True);
            Excel.Selection.Value := 'DESCRIPCION NUMERO: ' + IntToStr(i);
            Excel.Selection.Font.Size := 8;

            Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
            DarFormato(True, 'centro', True, True);
            Excel.Selection.Value := IntToStr(i);
            Excel.Selection.Font.Size := 8;
            sumaMN := sumaMN + i;


            Excel.Range['J' + IntToStr(i) + ':J' + IntToStr(i)].Select;
            DarFormato(True, 'centro', True, True);
            Excel.Selection.Value := IntToStr(i);
            Excel.Selection.Font.Size := 8;
            sumaUSD := sumaUSD + i;
          end;

          Excel.Range['B' + IntToStr(i) + ':D' + IntToStr(i)].Select;
          DarFormato(True, 'centro', False, True);
          Excel.Selection.Value := '';
          Excel.Selection.Font.Size := 8;

          Excel.Range['E' + IntToStr(i) + ':H' + IntToStr(i)].Select;
          DarFormato(True, 'centro', True, True);
          Excel.Selection.Value := 'TOTALES';
          Excel.Selection.Font.Size := 8;

          Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
          DarFormato(True, 'centro', True, True);
          Excel.Selection.Value := FloatToStr(sumaMN);
          Excel.Selection.Font.Size := 8;

          Excel.Range['J' + IntToStr(i) + ':J' + IntToStr(i)].Select;
          DarFormato(True, 'centro', True, True);
          Excel.Selection.Value := FloatToStr(sumaUSD);
          Excel.Selection.Font.Size := 8;

          i := i + 3;
          Excel.Range['B' + IntToStr(i) + ':E' + IntToStr(i)].Select;
          DarFormato(True, 'centro', True, False);
          Excel.Selection.Font.Size := 8;
          Excel.Selection.Value := 'ATENTAMENTE';

          Excel.Range['G' + IntToStr(i) + ':J' + IntToStr(i)].Select;
          DarFormato(True, 'centro', True, False);
          Excel.Selection.Font.Size := 8;
          Excel.Selection.Value := 'RECIBI EN CONFORMIDAD';

          Inc(i);
          Excel.Rows[i].RowHeight := 28.5;
          Excel.Range['B' + IntToStr(i + 1) + ':E' + IntToStr(i + 3)].Select;
          DarFormato(True, 'centro', True, False);
          Excel.Selection.Font.Size := 8;
          Excel.Selection.Value := 'SHJJFHCGSEJHCFSCFGYSC'+
                                   'HAGDGAFDHAGFDUYThhgsjhasv'+
                                   'gsdhf87w46ewjbheruye';
          Excel.Selection.WrapText := True;

          Excel.Range['G' + IntToStr(i + 1) + ':J' + IntToStr(i + 3)].Select;
          DarFormato(True, 'centro', True, False);
          Excel.Selection.Font.Size := 8;
          Excel.Selection.Value := 'SHJJFHCGSEJHCFSCFGYSC'+
                                   'HAGDGAFDHAGFDUYThhgsjhasv'+
                                   'gsdhf87w46ewjbheruye';
          Excel.Selection.WrapText := True;
       {$ENDREGION}
      end;

      if chk5.Checked then
      begin
      {$REGION 'OFICIO'}
        if haymasportadas then begin
          if not buscaContenido.Eof then begin
            imprimirPortadaContenido(buscaContenido.FieldByName('sDescripcion').AsString , buscaContenido.FieldByName('sNombrePortada').AsString);
            buscaContenido.Next;
            Inc(ficha);
          end;
        end;

        Libro.Sheets.Add;
        Excel.WorkBooks[1].WorkSheets[1].Name := 'OBRA (6)';
        Inc(ficha);

        sInicio := '';
        iFila := 1;
        iColumna := 1;

    
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 0.75;
        for i := 1 to 12 - 1 do
        begin
          Inc(iColumna);
          Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 8.438;
        end;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 0.75;
        //Imagen Derecha
        Try
          TmpName := '';
          imgAux := TImage.Create(nil);
          if TmpName='' then begin
            TempPath := ExtractFilePath(Application.Exename);
            TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
          Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 425, 1, 80, 62);
        Excel.activeWindow.DisplayGridlines := false;

        Excel.Range['B2:J2'].Select;
        DarFormato(True, 'centro', True, false);
        Excel.Selection.Font.Size := 12;
        Excel.Selection.Value := 'OCEANOGRAFIA, S.A. DE C.V.';

        Excel.Rows[3].RowHeight := 37.5;
        Excel.Range['E3:G3'].Select;
        DarFormato(True, 'centro', False, False);
        Excel.Selection.Font.Size := 8;
        Excel.Selection.Value := 'AV. 4 ORIENTE MANZANA D LOTE 3 '+
                                 'PUERTO IND. PESQUERO LAGUNA AZUL '+
                                 '24140 CIUDAD DEL CARMEN, CAMPECHE';
        Excel.Selection.WrapText := True;

        Excel.Range['B6:C6'].Select;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := 'CONTRATO:';

        Excel.Range['D6:F6'].Select;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := '';

        Excel.Range['G6:H6'].Select;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := 'FOLIO';

        Excel.Range['I6:L6'].Select;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := '';

        Excel.Range['B7:C8'].Select;
        Excel.Rows[7].RowHeight := 45.75;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := 'DESCRIPCION:';

        Excel.Range['D7:F8'].Select;
        DarFormato(True, 'izq', False, True);
        Excel.Selection.Value := 'COMENTARIOS;';
        Excel.Selection.WrapText := True;

        Excel.Range['G7:H7'].Select;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := 'OBRA:';

        Excel.Range['G8:H8'].Select;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := 'LOCALIZACIÓN';

        Excel.Range['I6:L6'].Select;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := 'SJHDSF3U4273V_K34R';

        Excel.Range['I7:L7'].Select;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := 'SDHGFWHEFDHXJGWYVETWY26T';
        Excel.Selection.WrapText := True;

        Excel.Range['I8:L8'].Select;
        DarFormato(True, 'izq', True, True);
        Excel.Selection.Value := 'PLATAFORMA: ' + '';



        ///////////////////////////////////////////////////////////////////////////
        //Tabla 1
        Excel.Range['B10:L11'].Select;
        Excel.Selection.Interior.ColorIndex := 16;

        Excel.Range['B10:L10'].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'MOVIMIENTOS DE EMBARCACIÓN';

        Excel.Range['B11:B11'].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range['C11:F11'].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'DESCRIPCIÓN';

        Excel.Range['G11:G11'].Select;
        DarFormato(True, 'centro', False, True);
        Excel.Selection.Value := 'CLAS.';

        Excel.Range['H11:H11'].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'CANTIDAD';

        Excel.Range['I11:I11'].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'PU MN';

        Excel.Range['J11:J11'].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'PU USD';

        Excel.Range['K11:K11'].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'IMP MN';

        Excel.Range['L11:L11'].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'IMP USD';

        sumaMN := 0.0;
        sumaUSD := 0.0;
        for i := 12 to 20 do
        begin
          //partida
          Excel.Range['B' + IntToStr(i) + ':B' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);

          //descripcion
          Excel.Range['C' + IntToStr(i) + ':F' + IntToStr(i)].Select;
          DarFormato(True, 'centro', False, True);
          Excel.Selection.Value := 'DESCRIASDFPCIONDJHFHGDFQ';
          Excel.Selection.WrapText := True;


          //clasificacion
          Excel.Range['G' + IntToStr(i) + ':G' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := 'FDH';

          //cantidad
          Excel.Range['H' + IntToStr(i) + ':H' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);

          //pu mn
          Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := FloatToStr(i * 3.25);

          //pu usd
          Excel.Range['J' + IntToStr(i) + ':J' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := FloatToStr(i * 5.75);

          //imp mn
          Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);
          sumaMN := sumaMN + i;

          //imp usd
          Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);
          sumaUSD := sumaUSD + i;
        end;

        Inc(i);

        Excel.Rows[i].RowHeight := 25;

        Excel.Range['B' + IntToStr(i) + ':F' + IntToStr(i)].Select;
        DarFormato(True, 'izq', False, False);
        Excel.Selection.Value := 'IMPORTE BARCO:';

        Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
        DarFormato(False, 'izq', True, False);
        Excel.Selection.Value := '$' + FloatToStr(sumaMN);

        Excel.Range['L' + IntToStr(i) + ':K' + IntToStr(i)].Select;
        DarFormato(False, 'izq', True, False);
        Excel.Selection.Value := '$' + FloatToStr(sumaUSD);

        ////////////////////////////////////////////////////////////////////////////
        //Tabla 2

        i := i + 2;
        Excel.Range['B' + IntToStr(i) + ':L' + IntToStr(i + 1)].Select;
        Excel.Selection.Interior.ColorIndex := 16;

        Excel.Range['B' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'PERSONAL';

        Inc(i);;

        Excel.Range['B' + IntToStr(i) + ':B' + IntToStr(i)].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range['C' + IntToStr(i) + ':F' + IntToStr(i)].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'DESCRIPCIÓN';

        Excel.Range['G' + IntToStr(i) + ':G' + IntToStr(i)].Select;
        DarFormato(True, 'centro', False, True);
        Excel.Selection.Value := 'UNIDAD';

        Excel.Range['H' + IntToStr(i) + ':H' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'CANTIDAD';

          Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'PU MN';

        Excel.Range['J' + IntToStr(i) + ':J' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'PU USD';

        Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'IMP MN';

        Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'IMP USD';

        iAnt := i + 1;
        sumaMN := 0.0;
        sumaUSD := 0.0;
        for i := iAnt to iAnt + 23 do
        begin
          //partida
          Excel.Range['B' + IntToStr(i) + ':B' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);

          //descripcion
          Excel.Range['C' + IntToStr(i) + ':F' + IntToStr(i)].Select;
          DarFormato(True, 'centro', False, True);
          Excel.Selection.Value := 'psdbSDCHJSGD';
          Excel.Selection.WrapText := True;


          //clasificacion
          Excel.Range['G' + IntToStr(i) + ':G' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := 'JOR';

          //cantidad
          Excel.Range['H' + IntToStr(i) + ':H' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);

          //pu mn
          Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := FloatToStr(i * 3.25);

          //pu usd
          Excel.Range['J' + IntToStr(i) + ':J' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := FloatToStr(i * 5.75);

          //imp mn
          Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);
          sumaMN := sumaMN + i;

          //imp usd
          Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);
          sumaUSD := sumaUSD + i;
        end;

        Inc(i);

        Excel.Rows[i].RowHeight := 25;

        Excel.Range['B' + IntToStr(i) + ':F' + IntToStr(i)].Select;
        DarFormato(True, 'izq', False, False);
        Excel.Selection.Value := 'IMPORTE PERSONAL:';

        Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
        DarFormato(False, 'izq', True, False);
        Excel.Selection.Value := '$' + FloatToStr(sumaMN);

        Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(False, 'izq', True, False);
        Excel.Selection.Value := '$' + FloatToStr(sumaUSD);

        ////////////////////////////////////////////////////////////////////////////
        //Tabla 3
        i := i + 2;
        Excel.Range['B' + IntToStr(i) + ':L' + IntToStr(i + 1)].Select;
        Excel.Selection.Interior.ColorIndex := 16;

        Excel.Range['B' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'EQUIPO';

        Inc(i);

        Excel.Range['B' + IntToStr(i) + ':B' + IntToStr(i)].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range['C' + IntToStr(i) + ':F' + IntToStr(i)].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'DESCRIPCIÓN';

        Excel.Range['G' + IntToStr(i) + ':G' + IntToStr(i)].Select;
        DarFormato(True, 'centro', False, True);
        Excel.Selection.Value := 'CLAS.';

        Excel.Range['H' + IntToStr(i) + ':H' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'UNIDAD';

          Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'PU MN';

        Excel.Range['J' + IntToStr(i) + ':J' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'PU USD';

        Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'IMP MN';

        Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'IMP USD';

        iAnt := i + 1;
        sumaMN := 0.0;
        sumaUSD := 0.0;
        for i := iAnt to iAnt + 23 do
        begin
          //partida
          Excel.Range['B' + IntToStr(i) + ':B' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);

          //descripcion
          Excel.Range['C' + IntToStr(i) + ':F' + IntToStr(i)].Select;
          DarFormato(True, 'centro', False, True);
          Excel.Selection.Value := 'psdbSDCHJSGD';
          Excel.Selection.WrapText := True;


          //clasificacion
          Excel.Range['G' + IntToStr(i) + ':G' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := 'JOR';

          //cantidad
          Excel.Range['H' + IntToStr(i) + ':H' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);

          //pu mn
          Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := FloatToStr(i * 3.25);

          //pu usd
          Excel.Range['J' + IntToStr(i) + ':J' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := FloatToStr(i * 5.75);

          //imp mn
          Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);
          sumaMN := sumaMN + i;

          //imp usd
          Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);
          sumaUSD := sumaUSD + i;
        end;

        Inc(i);

        Excel.Rows[i].RowHeight := 25;

        Excel.Range['B' + IntToStr(i) + ':F' + IntToStr(i)].Select;
        DarFormato(True, 'izq', False, False);
        Excel.Selection.Value := 'IMPORTE EQUIPO:';

        Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
        DarFormato(False, 'izq', True, False);
        Excel.Selection.Value := '$' + FloatToStr(sumaMN);

        Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(False, 'izq', True, False);
        Excel.Selection.Value := '$' + FloatToStr(sumaUSD);


        ////////////////////////////////////////////////////////////////////////////
        //Tabla 4 Pernotas
        i := i + 2;
        Excel.Range['B' + IntToStr(i) + ':L' + IntToStr(i + 1)].Select;
        Excel.Selection.Interior.ColorIndex := 16;

        Excel.Range['B' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'PERNOCTAS';

        Inc(i);

        Excel.Range['B' + IntToStr(i) + ':B' + IntToStr(i)].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range['C' + IntToStr(i) + ':F' + IntToStr(i)].Select;
        DarFormato(True, 'centro', True, True);
        Excel.Selection.Value := 'DESCRIPCIÓN';

        Excel.Range['G' + IntToStr(i) + ':G' + IntToStr(i)].Select;
        DarFormato(True, 'centro', False, True);
        Excel.Selection.Value := 'CLAS.';

        Excel.Range['H' + IntToStr(i) + ':H' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'UNIDAD';

          Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'PU MN';

        Excel.Range['J' + IntToStr(i) + ':J' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'PU USD';

        Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'IMP MN';

        Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(False, 'centro', True, True);
        Excel.Selection.Value := 'IMP USD';

        iAnt := i + 1;
        sumaMN := 0.0;
        sumaUSD := 0.0;
        for i := iAnt to iAnt + 3 do
        begin
          //partida
          Excel.Range['B' + IntToStr(i) + ':B' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);

          //descripcion
          Excel.Range['C' + IntToStr(i) + ':F' + IntToStr(i)].Select;
          DarFormato(True, 'centro', False, True);
          Excel.Selection.Value := 'psdbSDCHJSGD';
          Excel.Selection.WrapText := True;


          //clasificacion
          Excel.Range['G' + IntToStr(i) + ':G' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := 'JOR';

          //cantidad
          Excel.Range['H' + IntToStr(i) + ':H' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);

          //pu mn
          Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := FloatToStr(i * 3.25);

          //pu usd
          Excel.Range['J' + IntToStr(i) + ':J' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := FloatToStr(i * 5.75);

          //imp mn
          Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);
          sumaMN := sumaMN + i;

          //imp usd
          Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
          DarFormato(False, 'centro', False, True);
          Excel.Selection.Value := IntToStr(i);
          sumaUSD := sumaUSD + i;
        end;

        Inc(i);

        Excel.Rows[i].RowHeight := 25;

        Excel.Range['B' + IntToStr(i) + ':F' + IntToStr(i)].Select;
        DarFormato(True, 'izq', False, False);
        Excel.Selection.Value := 'IMPORTE PERNOCTAS:';

        Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
        DarFormato(True, 'izq', True, False);
        Excel.Selection.Value := '$' + FloatToStr(sumaMN);

        Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(False, 'izq', True, False);
        Excel.Selection.Value := '$' + FloatToStr(sumaUSD);

        ////////////////////////////////////////////////////////////////////////////
        //Tabla 5 MATERIAL
        i := i + 2;
        Excel.Range['B' + IntToStr(i) + ':I' + IntToStr(i + 1)].Select;
        Excel.Selection.Interior.ColorIndex := 16;

        Excel.Range['B' + IntToStr(i) + ':I' + IntToStr(i)].Select;
        DarFormato(True,'centro', True, True);
        Excel.Selection.Value := 'MATERIALES';

        Inc(i);

        Excel.Range['B' + IntToStr(i) + ':C' + IntToStr(i)].Select;
        DarFormato(True,'centro', True, True);
        Excel.Selection.Value := 'TRAZABILIDAD';

        Excel.Range['D' + IntToStr(i) + ':G' + IntToStr(i)].Select;
        DarFormato(True,'centro', True, True);
        Excel.Selection.Value := 'DESCRIPCIÓN';

        Excel.Range['H' + IntToStr(i) + ':H' + IntToStr(i)].Select;
        DarFormato(False,'centro', True, True);
        Excel.Selection.Value := 'UNIDAD';

        Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
        DarFormato(False,'centro', True, True);
        Excel.Selection.Value := 'CANTIDAD';

        iAnt := i + 1;
        sumaMN := 0.0;
        sumaUSD := 0.0;
        for i := iAnt to iAnt + 12 do
        begin
          Excel.Rows[i].RowHeight := 20;
          Excel.Range['B' + IntToStr(i) + ':C' + IntToStr(i)].Select;
          DarFormato(True,'centro', True, True);
          Excel.Selection.Value := 'S/T';

          Excel.Range['D' + IntToStr(i) + ':G' + IntToStr(i)].Select;
          DarFormato(True,'centro', True, True);
          Excel.Selection.Value := 'KDGFWTR23YY';

          Excel.Range['H' + IntToStr(i) + ':H' + IntToStr(i)].Select;
          DarFormato(False,'centro', True, True);
          Excel.Selection.Value := IntToStr(i);
          sumaMN := sumaMN + i;

          Excel.Range['I' + IntToStr(i) + ':I' + IntToStr(i)].Select;
          DarFormato(False,'centro', True, True);
          Excel.Selection.Value := IntToStr(i);
          sumaUSD := sumaUSD + i;
        end;
      
        Inc(i);

        Excel.Range['D' + IntToStr(i) + ':F' + IntToStr(i)].Select;
        DarFormato(True, 'der', True, False);
        Excel.Selection.Value := 'COSTO TOTAL DE LA ACTIVIDAD:';

        Excel.Range['K' + IntToStr(i) + ':K' + IntToStr(i)].Select;
        DarFormato(True, 'der', True, False);
        Excel.Selection.Value := FloatToStr(sumaMN);

        Excel.Range['L' + IntToStr(i) + ':L' + IntToStr(i)].Select;
        DarFormato(True, 'der', True, False);
        Excel.Selection.Value := FloatToStr(sumaUSD);
        {$ENDREGION}
      end;

      if chk6.Checked then
      begin
      {$REGION 'OFICIO'}
        if haymasportadas then begin
          if not buscaContenido.Eof then begin
            imprimirPortadaContenido(buscaContenido.FieldByName('sDescripcion').AsString , buscaContenido.FieldByName('sNombrePortada').AsString);
            buscaContenido.Next;
            Inc(ficha);
          end;
        end;

        Libro.Sheets.Add;
        Excel.WorkBooks[1].WorkSheets[1].Name := 'OF NOTA CAMPO(7)';
        Inc(ficha);

        iColumna := 1;
        iFila := 1;

        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 0.75;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);

    {$REGION 'IMAGENES DE CABECERA'}
      //Imagen Izquierda
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then begin
    //      GetTempPath(SizeOf(TempPath), TempPath);
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtempSln1'+formatdatetime('dddddd hhnnss',now)+'.jpg';
          fs := Connection.contrato.CreateBlobStream(Connection.contrato.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
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
      Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 20, 70, 35);
      //Imagen Derecha
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then begin
    //      GetTempPath(SizeOf(TempPath), TempPath);
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
      Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 498, 20, 70, 40);
      //Texto de Cabecera
      Excel.Range['A1:L4'].Select;
      //PFormatosExcel_Bordes(Excel, False, True, False, False);

      Excel.Range['A2:L4'].Select;
      PFormatosExcel_H2(Excel, 0, True, 10);
      Excel.Selection.Value := Connection.configuracion.FieldByName('sNombre').AsString;



      Excel.Range['A4:L5'].Select;
      PFormatosExcel_H2(Excel, 0, True, 8);
     // PFormatosExcel_Bordes(Excel, False, True, False, False, -4119);
      Excel.Selection.Value := 'NOTA DE CAMPO';
      {$ENDREGION}

    {$REGION 'CABECERA'}
        iFila := 6;
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'CONTRATO:';

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Connection.contrato.FieldByName('sContrato').AsString;

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'FOLIO:';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Ordenesdetrabajo.FieldByName('sNumeroOrden').AsString;
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(7)+':'+ColumnaNombre(3)+IntToStr(8)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'DESCRIPCIÓN:';

        Excel.Range[ColumnaNombre(4)+IntToStr(7)+':'+ColumnaNombre(6)+IntToStr(8)].Select;
        PFormatosExcel_H2(Excel, 11, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Connection.contrato.FieldByName('mDescripcion').AsString;

        Excel.Range[ColumnaNombre(7)+IntToStr(7)+':'+ColumnaNombre(8)+IntToStr(7)].Select;
        PFormatosExcel_H2(Excel, 47, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'OBRA:';

        Excel.Range[ColumnaNombre(9)+IntToStr(7)+':'+ColumnaNombre(12)+IntToStr(7)].Select;
        PFormatosExcel_H2(Excel, 47, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Connection.contrato.FieldByName('sTitulo').AsString;
        Inc (iFila);
     
        Excel.Range[ColumnaNombre(7)+IntToStr(8)+':'+ColumnaNombre(8)+IntToStr(8)].Select;
        PFormatosExcel_H2(Excel, 21, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'LOCALIZACIÓN:';

        Excel.Range[ColumnaNombre(9)+IntToStr(8)+':'+ColumnaNombre(12)+IntToStr(8)].Select;
        PFormatosExcel_H2(Excel, 21, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Connection.contrato.FieldByName('sUbicacion').AsString;

    {$ENDREGION}

    {$REGION 'IMPRESION ACTIVIDADES'}
        iFila := 12;

        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add(' ' +
                                  'SELECT ' +
                                  '	sNumeroOrden, ' +
                                  '	mDescripcion, ' +
                                  '	sNumeroActividad ' +
                                  'FROM actividadesxorden ' +
                                  'WHERE sNumeroOrden = :Folio');
        Connection.QryBusca.ParamByName('Folio').AsString := Ordenesdetrabajo.FieldByName('sNumeroOrden').AsString;
        Connection.QryBusca.Open;

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'ACTIVIDAD';
        Inc(iFila);

        if Connection.QryBusca.RecordCount > 0 then begin
          while Not Connection.QryBusca.Eof do begin

            Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
            PFormatosExcel_Bordes(Excel);
            Excel.Selection.Value := Connection.QryBusca.FieldByName('sNumeroActividad').AsString;

            Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
            PFormatosExcel_Bordes(Excel);
            Excel.Selection.Value := Connection.QryBusca.FieldByName('mDescripcion').AsString;

            Inc (iFila);
            Connection.QryBusca.Next;
          end;
        end
        else
        begin
            Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
            PFormatosExcel_Bordes(Excel);
            Excel.Selection.Value := '';

            Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
            PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
            PFormatosExcel_Bordes(Excel);
            Excel.Selection.Value := '';
        end;

        Inc (iFila,3);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, False, 7, clBlack, 'Arial');
        Excel.Selection.Value := 'PERIODOS DE EJECUCION DE LA ACTIVIDAD';

        Inc(iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'FECHA';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'INICIO';

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'TERMINO';

        Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'AFECTACION';

        Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'INTERVALO TIEMPO';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'AVANCE ANTERIOR';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'AVANCE ACTUAL';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'AVANCE ACUMULADO';

        Inc(iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(4)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Inc(iFila);

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.Value := 'DURACION TIEMPO EFECTIVO (HRS):' ;

        Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '01:56' ;

        Inc (iFila);

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.Value := 'DURACION TIEMPO AFECTACIONES (HRS):' ;

        Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '00:00' ;

        Inc (iFila);

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(5)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, False, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.Value := 'TIEMPO TOTAL (HRS):' ;

        Excel.Range[ColumnaNombre(6)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 15, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '01:56' ;

        Inc (iFila,2);

    {$ENDREGION}

    {$REGION 'IMPRESION ACTIVIDADES BARCO'}
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'ACTIVIDAD';
        Inc(iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 38, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';
        Inc(iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'MOVIMIENTOS DE EMBARCACION';
        Inc(iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'DESCRIPCION';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'UNIDAD';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'CANTIDAD';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PU MN';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PU USD';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'IMP MN';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'IMP USD';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.Value := 'IMPORTE BARCO';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';
        Inc (iFila,2);
    {$ENDREGION}

    {$REGION 'IMPRESION PERSONAL'}
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PERSONAL';
        Inc(iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'DESCRIPCION';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'UNIDAD';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'CANTIDAD';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PU MN';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PU USD';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'IMP MN';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'IMP USD';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.Value := 'IMPORTE PERSONAL';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';
        Inc (iFila,2);
    {$ENDREGION}

    {$REGION 'IMPRESION EQUIPOS'}
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'EQUIPO';
        Inc(iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'DESCRIPCION';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'UNIDAD';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'CANTIDAD';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PU MN';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PU USD';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'IMP MN';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'IMP USD';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.Value := 'IMPORTE EQUIPO';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';
        Inc (iFila,2);
    {$ENDREGION}

    {$REGION 'IMPRESION PERCNOTAS'}
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 13, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PERNOCTAS';
        Inc(iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PARTIDA';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'DESCRIPCION';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'UNIDAD';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'CANTIDAD';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PU MN';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'PU USD';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'IMP MN';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'IMP USD';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';
        Inc (iFila);

        Excel.Range[ColumnaNombre(5)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.Value := 'IMPORTE PERCNOTAS';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'MATERIAL';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'TRAZABILIDAD';

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'DESCRIPCION';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'UNIDAD';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        PFormatosExcel_Rellenar(Excel, $00BBBBBB);
        Excel.Selection.Value := 'CANTIDAD';
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.Value := '';
        Inc (iFila,2);

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.Value := 'COSTO TOTAL DE LA ACTIVIDAD';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 18, True, 7, clBlack, 'Arial');
        Excel.Selection.Value := '$0';
       {$ENDREGION}
      {$ENDREGION}
      end;

      if chk7.Checked then
      begin
      {$REGION 'OFICIO'}
        if haymasportadas then begin
          if not buscaContenido.Eof then begin
            imprimirPortadaContenido(buscaContenido.FieldByName('sDescripcion').AsString , buscaContenido.FieldByName('sNombrePortada').AsString);
            buscaContenido.Next;
            Inc(ficha);
          end;
        end;

        Libro.Sheets.Add;
        Excel.WorkBooks[1].WorkSheets[1].Name := 'CAMPO(7)';
        Inc(ficha);

        iColumna := 1;
        iFila := 1;

        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 0.75;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9;
        Inc(iColumna);

    {$REGION 'IMAGENES DE CABECERA'}
      //Imagen Izquierda
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then begin
    //      GetTempPath(SizeOf(TempPath), TempPath);
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtempSln1'+formatdatetime('dddddd hhnnss',now)+'.jpg';
          fs := Connection.contrato.CreateBlobStream(Connection.contrato.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
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
      Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 20, 70, 35);
      //Imagen Derecha
      Try
        TmpName := '';
        imgAux := TImage.Create(nil);
        if TmpName='' then begin
    //      GetTempPath(SizeOf(TempPath), TempPath);
          TempPath := ExtractFilePath(Application.Exename);
          TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
      Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 498, 10, 70, 50);
      //Texto de Cabecera
      Excel.Range['A1:L5'].Select;
      //PFormatosExcel_Bordes(Excel, False, True, False, False);

     // Excel.Range['A2:L4'].Select;
     // PFormatosExcel_H2(Excel, 0, True, 10);
     // Excel.Selection.Value := Connection.configuracion.FieldByName('sNombre').AsString;



      Excel.Range['A1:L5'].Select;
      PFormatosExcel_H2(Excel, 0, True, 8);
     // PFormatosExcel_Bordes(Excel, False, True, False, False, -4119);
      Excel.Selection.Value := 'DESGLOSE DE COSTOS';
      Inc (iFila,4);
      {$ENDREGION}

    {$REGION 'CABECERA'}
        iFila := 6;
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(3)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'CONTRATO:';

        Excel.Range[ColumnaNombre(4)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Connection.contrato.FieldByName('sContrato').AsString;

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'FOLIO:';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Ordenesdetrabajo.FieldByName('sNumeroOrden').AsString;
        Inc (iFila);

        Excel.Range[ColumnaNombre(2)+IntToStr(7)+':'+ColumnaNombre(3)+IntToStr(8)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'DESCRIPCIÓN:';

        Excel.Range[ColumnaNombre(4)+IntToStr(7)+':'+ColumnaNombre(6)+IntToStr(8)].Select;
        PFormatosExcel_H2(Excel, 11, False, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Connection.contrato.FieldByName('mDescripcion').AsString;

        Excel.Range[ColumnaNombre(7)+IntToStr(7)+':'+ColumnaNombre(8)+IntToStr(7)].Select;
        PFormatosExcel_H2(Excel, 47, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'OBRA:';

        Excel.Range[ColumnaNombre(9)+IntToStr(7)+':'+ColumnaNombre(12)+IntToStr(7)].Select;
        PFormatosExcel_H2(Excel, 47, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Connection.contrato.FieldByName('sTitulo').AsString;
        Inc (iFila);

        Excel.Range[ColumnaNombre(7)+IntToStr(8)+':'+ColumnaNombre(8)+IntToStr(8)].Select;
        PFormatosExcel_H2(Excel, 21, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'LOCALIZACIÓN:';

        Excel.Range[ColumnaNombre(9)+IntToStr(8)+':'+ColumnaNombre(12)+IntToStr(8)].Select;
        PFormatosExcel_H2(Excel, 21, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := Connection.contrato.FieldByName('sUbicacion').AsString;

    {$ENDREGION}

        iFila:=10;

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'NIVEL';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := 'CATEGORIA/CONCEPTO';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'CANTIDAD';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'UNIDAD';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'PRECIO M.N';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'PRECIO DLS.';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'IMPORTE M.N';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'IMPORTE DLS.';
        Inc (ifila);

        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(2)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 11, True, 7, clBlack, 'Arial');
        PFormatosExcel_Bordes(Excel);
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(7)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(9)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := 'PRECIO M.N';

        Excel.Range[ColumnaNombre(10)+IntToStr(iFila)+':'+ColumnaNombre(10)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(11)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(12)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 28, True, 7, clBlack, 'Arial');
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.Value := '';
       {$ENDREGION}
      end;

      if chk8.Checked then
      begin
      {$REGION 'OFICIO'}
        if haymasportadas then begin
          if not buscaContenido.Eof then begin
            imprimirPortadaContenido(buscaContenido.FieldByName('sDescripcion').AsString , buscaContenido.FieldByName('sNombrePortada').AsString);
            buscaContenido.Next;
            Inc(ficha);
          end;
        end;

        Inc(ficha);
        Libro.Sheets.Add;
        Excel.WorkBooks[1].WorkSheets[1].Name := 'DESGLOSE COSTO(8)';

        sInicio := '';
        sColumna := '';
        sFila := '';
        iColumna := 1;
        iFila := 1;
        finPartidas := 13;
    
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 3.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 5.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6.86;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 6;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 8.57;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7.43;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 9.14;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 15.71;
        Inc(iColumna);
        Excel.Columns[ColumnaNombre(iColumna)+':'+ColumnaNombre(iColumna)].ColumnWidth := 7.29;
        Inc(iColumna);
        Try
          TmpName := '';
          imgAux := TImage.Create(nil);
          if TmpName='' then begin
            TempPath := ExtractFilePath(Application.Exename);
            TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
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
          Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 365, 1, 80, 62);
          //Texto de Cabecera
          Excel.Range['A1:J1'].Select;
          PFormatosExcel_H2(Excel, 0, True, 10);

          Excel.activeWindow.DisplayGridlines := false;

          Connection.QryBusca.SQL.Clear;
          Connection.QryBusca.SQL.Text := 'SELECT * ' +
                                          '		FROM ' +
                                          'Configuracion' +
                                          '		WHERE ' +
                                          'sContrato = :Contrato ';
          Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
          Connection.QryBusca.Open;

          iFila := 2;
          Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(9)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 16, True, 12, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.Value := Connection.QryBusca.FieldByName('sNombre').AsString;
          Inc (iFila);

          Excel.Range[ColumnaNombre(3)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
          PFormatosExcel_H2(Excel, 38, True, 8, clBlack, 'Arial');
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.Value := Connection.QryBusca.FieldByName('sDireccion1').AsString + '                                                                                                                                    '
          + Connection.QryBusca.FieldByName('sDireccion2').AsString + '                                                                                                                                    '
          + Connection.QryBusca.FieldByName('sDireccion3').AsString + Connection.QryBusca.FieldByName('sCiudad').AsString;
          Excel.Selection.ReadingOrder := xlContext;
          Excel.Selection.WrapText := True;

          //Inicio inserccion de datos...
          iFila := 6;
          iColumna := 1;

          Excel.Range['B5:J5'].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.Value := DateToStr(Now());

          Excel.Range['B12:J12'].Select;
          Excel.Selection.Interior.ColorIndex := 42;
          EstablecerContornos;

          Excel.Range['B6:J6'].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Value := 'AT"N: ING. ANTONIO CARABES GARCIA';//Insertar aqui el nombre... ejemplo : AT'N: ING. ANTONIO CARABES GARCIA    

          Excel.Range['B7:J7'].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Value := ' F=338155';//Insertar fila siguiente... ejemplo : F=338155

          Excel.Range['B8:J8'].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Value := 'SUPERVISOR DE P.E.P.';//Insertar cargo... ejemplo : SUPERVISOR DE P.E.P.

          Excel.Range['B9:J9'].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Value := 'A BORDO DE B.P.D. ISLAND PIONEER';//Insertar estado... ejemplo : A BORDO DE B.P.D. ISLAND PIONEER

          Excel.Range['B10:J10'].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Value := 'Asunto: Acta de Entrega';//Insertar el tipo de asunto..

          Excel.Range['B11:J11'].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Size := 7;
          Excel.Selection.Value := 'POR MEDIO DE LA PRESENTE SE HACE ENTREGA DE LOS TRABAJOS REALIZADOS POR EL PERSONAL DE CONSTRUCCIÓN CON '+
          'APOYO DE UNA EMBARCACION CON POSICIONAMIENTO DINAMICO, CON CARGO A LA ORDEN DE TRABAJO No. 020  ' +
          '"REHABILITACIÓN Y CORRECCIÓN DE ANOMALIAS EN LA PLATAFORMA MAY-DL-1 Y TRABAJOS DE AMARRE DE '+
          'POZOS". Y CON NUMERO DE FOLIO ISP020_0008_APLT, CORRESPONDIENTE A LAS ACTIVIDADES DE  "REHABILITACIÓN ' +
          'ESTRUCTURAL EN SEGUNDO NIVEL DE PLATAFORMA MAY-DL1 (CSU: 268-13-0527)". DE ACUERDO A OFICIO No. PEP- ' +
          'SMIL-CMIL-TDHNP-GMIM-ISP-193/2013.,  AMPARADOS BAJO EL CONTRATO 428232833, EJECUTÁNDOSE LAS SIGUIENTES '+
          'ACTIVIDADES.';//Insertar... la explicacion del asunto...
          Excel.Selection.WrapText := True;
          Excel.Rows[11].RowHeight := 93;

          Excel.Range['B12:B12'].Select;
          Excel.Selection.HorizontalAlignment := xlLeft;
          EstablecerContornos;
          Excel.Selection.Value := 'PARTIDA';

          Excel.Range['C12:I12'].Select;
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.MergeCells := True;
          EstablecerContornos;
          Excel.Selection.Value := 'ACTIVIDAD';

          Excel.Range['J12:J12'].Select;
          Excel.Selection.HorizontalAlignment := xlLeft;
          EstablecerContornos;
          Excel.Selection.Value := 'AVANCE';

          iFila := 13;


          for i := 0 to 7 do
          begin
            Excel.Rows[iFila].RowHeight := 40;
            sInicio := 'B' + IntToStr(iFila) + ':' + 'B' + IntToStr(iFila);
            Excel.Range[sInicio].Select;
            EstablecerContornos;
            Excel.Selection.Value := IntToStr(i);
            Excel.Selection.VerticalAlignment    := xlCenter;
            Excel.Selection.HorizontalAlignment := xlCenter;

            sInicio := 'C' + IntToStr(iFila) + ':' + 'I' + IntToStr(iFila);
            Excel.Range[sInicio].Select;
            Excel.Selection.MergeCells := True;
            EstablecerContornos;
            Excel.Selection.Value := 'Lorem ipsum dolor sit amet, consectetur adipisicing elit, '+
                                     'sed do eiusmod tempor incididunt ut labore et dolore '+
                                     'magna aliqua. Ut enim ad minim veniam.';
            Excel.Selection.VerticalAlignment    := xlCenter;
            Excel.Selection.HorizontalAlignment := xlLeft;
            Excel.Selection.WrapText := True;

                               
            sInicio := 'J' + IntToStr(iFila) + ':' + 'J' + IntToStr(iFila);
            Excel.Range[sInicio].Select;
            Excel.Selection.MergeCells := True;
            EstablecerContornos;
            Excel.Selection.Value := '100%';
            Excel.Selection.VerticalAlignment    := xlCenter;
            Excel.Selection.HorizontalAlignment := xlCenter;

            Inc(finPartidas);
            Inc(iFila);
          end;

          Inc(finPartidas);
    
          for i := 1 to 2 do
          begin
            Inc(finPartidas);
            Excel.Rows[finPartidas].RowHeight := 45;
            sInicio := 'B' + IntToStr(finPartidas) + ':J' + IntToStr(finPartidas);
            Excel.Range[sInicio].Select;
            Excel.Selection.MergeCells := True;
            Excel.Selection.HorizontalAlignment := xlLeft;
            Excel.Selection.VerticalAlignment := xlCenter;
            Excel.Selection.Value := 'NOTA 01:    CON ESTA FECHA 18 DE OCTUBRE DE 2013, EN HORARIO DE 18:00 HRS., '+
            'INICIAN LAS ACTIVIDADES DE ESTE FOLIO ISP020_0008_APLT: "REHABILITACIÓN ESTRUCTURAL EN SEGUNDO NIVEL '+
            'DE PLATAFORMA MAY-DL1 (CSU: 268-13-0527)".   DE ACUERDO A OFICIO: PEP-SMIL-CMIL-TDHNP-GMIM-ISP-193/2013.';
            Excel.Selection.WrapText := True;

            Inc(finPartidas);
          end;

          Inc(finPartidas);
          Excel.Range['B' + IntToStr(finPartidas) + ':E' + IntToStr(finPartidas)].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment:= xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Value := 'ATENTAMENTE';
    
          Excel.Range['G' + IntToStr(finPartidas) + ':J' + IntToStr(finPartidas)].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment:= xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Value := 'RECIBI EN CONFORMIDAD';

          Inc(finPartidas);
          Excel.Rows[finPartidas].RowHeight := 30;
          Inc(finPartidas);

          Excel.Rows[finPartidas + 2].RowHeight := 25;
          sInicio := 'B' + IntToStr(finPartidas) + ':E' + IntToStr(finPartidas + 2);
          Excel.Range[sInicio].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment:= xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 9;
          Excel.Selection.Value := 'ING. ARISTEO ANTONIO ESTRADA CHAVEZ '+
                                   'REPRESENTANTE DE OCEANOGRAFIA '+
                                   'A BORDO DEL B.P.D. ISLAND PIONEER';
          Excel.Selection.WrapText := True;

          sInicio := 'G' + IntToStr(finPartidas) + ':J' + IntToStr(finPartidas + 2);
          Excel.Range[sInicio].Select;
          Excel.Selection.MergeCells := True;
          Excel.Selection.HorizontalAlignment:= xlCenter;
          Excel.Selection.VerticalAlignment := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 9;
          Excel.Selection.Value := 'ING. ARISTEO ANTONIO ESTRADA CHAVEZ '+
                                   'REPRESENTANTE DE OCEANOGRAFIA '+
                                   'A BORDO DEL B.P.D. ISLAND PIONEER';
          Excel.Selection.WrapText := True;
    
      {$ENDREGION}
      end;

    finally
      if assigned(zqfolio) then
        freeandnil(zqfolio);

      if assigned(buscaContenido) then
        freeandnil(buscacontenido) 
    end;
    if ficha > 1 then
      Libro.Sheets[ficha - 1].Select;

  end
  else begin
    ShowMessage('Seleccione un folio.');
  end;
end;


Procedure TfrmReportePeriodo.NotaCampoExcel(excel:Variant;Hoja:Variant;Libro:Variant);
var

  sFileName:string;
  sDescFrente:string;
  pidl: PItemIDList;
  InFolder: array[0..MAX_PATH] of Char;
  QrConfiguracion,QrAux : TZReadOnlyQuery;
  tmpNombre,tmpNombreC:string;
  Embarcacion : string;



procedure ConfigurarHoja(var excel: Variant; var Hoja: Variant);
var
  pfHoja: Byte;
  SubCad,CadError: String;
  Difer, AcumDifer: Extended;
  sFirmante1,sFirmante2,sPuesto1,sPuesto2:string;
  QryBuscarFirmas: tzReadOnlyQuery;
  fs:TStream;
  imgAux:TImage;
  Pic : TJpegImage;
  TempPath: array [0..MAX_PATH-1] of Char;
  FNombre1,FNombre2:TFileName;
  sCadT:string;
begin

  // Seleccionar el periodo de firmantes
  application.ProcessMessages;
  imgAux:=TImage.Create(nil);
  QryBuscarFirmas:=TZReadOnlyQuery.Create(nil);
   Application.ProcessMessages;

    // Poner las firmas en todas las hojas del libro generado
    try


      begin

        Excel.ActiveSheet.PageSetup.PaperSize := xlPaperLetter;
        Excel.ActiveWindow.View :=xlPageLayoutView;

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

        Excel.ActiveSheet.PageSetup.CenterFooter :='&"Arial,Normal"&'+inttostr(TamFont)+'&P de &#';//'&Z&G&P de &#&D&G'; //'&P de &N';

        Excel.ActiveSheet.PageSetup.LeftMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.TopMargin := Excel.InchesToPoints(0.196850393700787);
        Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(1.65354330708661);
        Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0);
        Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0.393700787401575);
        Excel.ActiveSheet.PageSetup.PrintHeadings := False;
        Excel.ActiveSheet.PageSetup.PrintGridlines := False;

        Excel.ActiveSheet.PageSetup.PrintQuality := 600;
        Excel.ActiveSheet.PageSetup.CenterHorizontally := False;
        Excel.ActiveSheet.PageSetup.CenterVertically := False;

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

Procedure FormatoEncabezado(var Excel: Variant;Cadena:string; Align: Integer;Negrita:Boolean);
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

Procedure GenerarMarco(var Excel:Variant);
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



procedure GenerarNotaPdas(var Excel,Hoja:Variant;var QrDatos:TZReadOnlyQuery;var QrPdas:TZReadOnlyQuery);
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
                                'where scontrato=:Contrato )';
  connection.QryBusca.ParamByName('contrato').AsString := global_Contrato_Barco;

  connection.QryBusca.Open;

  if connection.QryBusca.RecordCount = 1 then
     global_barco := connection.QryBusca.FieldByName('sIdEmbarcacion').AsString;
  {$ENDREGION}

  bPasoAjuste := False;
  dAjustePernocta := 0;
  while not QrPdas.Eof do
  begin

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


      Hoja.Range['B'+IntToStr(ren)+':C'+IntToStr(ren)].Select;
      FormatoNormal(Excel,QrPdas.FieldByName('sNumeroActividad').AsString,xlCenter,True,false,'@');

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
                      dFactorMovimiento := dFactorMovimiento +  connection.QryBusca2.FieldValues['sFactor'];
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
              FormatoNormal(Excel, xRound(QryMovimientos.FieldByName('dVentaMN').Value * dFactorMovimiento, 2),xlCenter,false,False,'#,##0.00');
              dCostoMn:=dCostoMn + xRound((QryMovimientos.FieldByName('dVentaMN').AsFloat * dFactorMovimiento), 2) ;

              Hoja.Range['L'+IntToStr(ren)+':L'+IntToStr(ren)].Select;
              FormatoNormal(Excel, xRound(QryMovimientos.FieldByName('dVentaDLL').Value * dFactorMovimiento, 2) ,xlCenter,false,False,'#,##0.00');
              dCostoDll:=dCostoDll + xRound((QryMovimientos.FieldByName('dVentaDLL').AsFloat * dFactorMovimiento), 2);
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
end;




Procedure PonerEncabezado(var Excel: Variant; var Hoja: Variant; var QrDatos: TZReadOnlyQuery);
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


  begin

      QrConfiguracion := TZReadOnlyQuery.Create(Nil);
      try
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
        QrConfiguracion.ParamByName('convenio').AsString:= global_convenio;
        QrConfiguracion.ParamByName('Orden').AsString:= zqrfolio.fieldbyname('snumeroorden').asstring;
        QrConfiguracion.ParamByName('ContratoBarco').AsString:= global_Contrato_barco;
        QrConfiguracion.Open;
        Application.ProcessMessages;
        TamFont:=7;

        Application.ProcessMessages;
        PonerEncabezado(Excel,Hoja,QrConfiguracion);

        QrAux:=TZReadOnlyQuery.Create(nil);
        QrAux.Connection:=connection.zConnection;
        QrAux.Active := False;
        QrAux.SQL.Clear;
        QrAux.SQL.Add('select * from actividadesxorden where scontrato=:Contrato  and sNumeroOrden=:Orden  '+
                      ' order by iItemOrden ');
        QrAux.ParamByName('Contrato').AsString := ZQRFolio.FieldByName('scontrato').AsString;
        QrAux.ParamByName('Orden').AsString    :=  ZQRFolio.FieldByName('snumeroorden').AsString;
        QrAux.Open;

        //Primero el Id de la Embarcacion principal... OSA 2013 ivan,,
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('select sIdEmbarcacion from embarcacion_vigencia '+
                           'where sContrato =:Contrato order by dFechaInicio');
        connection.QryBusca.ParamByName('Contrato').AsString := global_contrato_barco;
        connection.QryBusca.Open;

        if connection.QryBusca.RecordCount > 0 then
            global_barco := connection.QryBusca.FieldValues['sIdEmbarcacion']
        else
           messageDLG('No existe una Vigencia de Embarcacion Principal', mtInformation, [mbOk], 0);

        //ActualiaFactorGeneradorPER(global_barco, global_contrato, ordenesdetrabajo.FieldValues['sNumeroOrden']);
        //ActualiaFactorGeneradorEQ(global_barco, global_contrato, ordenesdetrabajo.FieldValues['sNumeroOrden']);

        GenerarNotaPdas(Excel,Hoja,QrConfiguracion,QrAux);
        ConfigurarHoja(Excel,Hoja);
      finally
        QrConfiguracion.Free;
      end;


  end;

end;


function TfrmReportePeriodo.RedimensionarJPG(sFilePath: string): string;
var
  bmp: TBitmap;
  jpg: TJpegImage;
  scale: Double;
  sTemp : String;
begin
  if FileExists(sFilePath) then
  begin
    jpg := TJpegImage.Create;
    try
      jpg.Loadfromfile(sFilePath);

      if jpg.Height > jpg.Width then
        scale := 550 / jpg.Height
      else
        scale := 550 / jpg.Width;

      bmp := TBitmap.Create;
      try
        {Create thumbnail bitmap, keep pictures aspect ratio}
        bmp.Width := Round(jpg.Width * scale);
        bmp.Height := Round(jpg.Height * scale);
        bmp.Canvas.StretchDraw(bmp.Canvas.Cliprect, jpg);
        {Draw thumbnail as control}
        //Self.Canvas.Draw(100, 10, bmp);
        {Convert back to JPEG and save to file}
        jpg.Assign(bmp);

        sTemp := Copy(sFilePath,1, Length(sFilePath) - 4);

        jpg.SaveToFile(sTemp + '_MODIF_INTELIGENT.jpg');
        result := sFilePath;
      finally
        bmp.free;
      end;
    finally
      jpg.free;
    end;
  end
  else
  begin
    result := '';
  end;
end;

function TfrmReportePeriodo.numeroDeImagenes(limite : Integer) : boolean;
var
  iInicio, iDesarrollo, iTermino, iNinguno : Integer;
  sFase, sImprime : string;
begin
  iInicio     := 0;
  iDesarrollo := 0;
  iTermino    := 0;
  iNinguno    := 0;                      

  if fotografico_acta.RecordCount > 0 then
  begin
    fotografico_acta.First;
    while not fotografico_acta.Eof do
    begin
      sFase := UpperCase(fotografico_acta.FieldByName('sFasePartida').AsString);
      sImprime := fotografico_acta.FieldByName('lImprime').AsString;

      if (sFase = 'INICIO') and (sImprime = 'Si') then
      begin
        Inc(iInicio);
      end;

      if (sFase = 'DESARROLLO') and (sImprime = 'Si') then
      begin
        Inc(iDesarrollo);
      end;

      if (sFase = 'CONCLUSION') and (sImprime = 'Si') then
      begin
        Inc(iTermino);
      end;

      if (sFase = 'NINGUNO') and (sImprime = 'Si') then
      begin
        Inc(iNinguno);
      end;

      fotografico_acta.Next;
          //Fin evaluaciones para imagenes
    end;//Fin While not

    //Se evaluan las fotografias encontradas por cada fase y si se van a imprimir.
    if iNinguno = 0 then
      begin
      if (iInicio > limite) or (iDesarrollo > limite) or (iTermino > limite) then
      begin
        ShowMessage('Solo se permiten ' + IntToStr(limite) + ' fotografias por fase. '+
                    'INICIO: ' + IntToStr(iInicio) + ' imagenes, '+
                    'DESARROLLO: ' + IntToStr(iDesarrollo) + ' imagenes, '+
                    'CONCLUSION: ' + IntToStr(iTermino) + ' imagenes, '+
                    'NINGUNO: ' + IntToStr(iNinguno) + ' imagenes.');

        Result := False;
      end
      else
      begin
        Result := True;
      end;
    end
    else
    begin
      ShowMessage('El Acta de Entrega no admite fotografias con fase NINGUNO');
      Result := False;
    end;
  end;//Fin if recordcount
end;

end.
