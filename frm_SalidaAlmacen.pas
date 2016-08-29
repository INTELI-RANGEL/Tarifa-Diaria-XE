unit frm_SalidaAlmacen;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, DB, DBCtrls, global, 
  Mask, OleCtrls, Grids, DBGrids, frm_barra, ExtCtrls, Utilerias,
  Menus, frxClass, frxDBSet, RXDBCtrl,  DateUtils,
  RXCtrls, CheckLst, ZAbstractRODataset, ZDataset,
  rxCurrEdit, rxToolEdit, AdvGlowButton, UnitValidacion,
   udbgrid, unitexcepciones, unittbotonespermisos, UFunctionsGHH,
  ZAbstractDataset, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, cxLabel, ExcelXP, OleServer, Excel2000,
  AdvEdit, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel,
  dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven,
  dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus, dxSkinSilver,
  dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008, dxSkinTheAsphaltWorld,
  dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue;

type
  TfrmSalidaAlmacen = class(TForm)
    ds_anexo_suministro: TDataSource;
    frxDBEntrada: TfrxDBDataset;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    ds_pedido: TDataSource;
    N4: TMenuItem;
    frxEntrada: TfrxReport;
    Pedido: TZReadOnlyQuery;
    ds_pSalidas: TDataSource;
    pSalidas: TZReadOnlyQuery;
    Reporte: TZReadOnlyQuery;
    GBSuperior: TGroupBox;
    frmBarra2: TfrmBarra;
    PgControl: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    tsPlataforma: TLabel;
    imgNotas: TImage;
    frmBarra1: TfrmBarra;
    Label7: TLabel;
    Label3: TLabel;
    tdIdFecha: TDateTimePicker;
    Label17: TLabel;
    tsNumeroOrden: TComboBox;
    tmComentarios: TMemo;
    Label6: TLabel;
    tsInsumo: TEdit;
    tdCantidad: TRxCalcEdit;
    GridPartidas: TRxDBGrid;
    Grid_Pedido: TRxDBGrid;
    lblEncabezado: TStaticText;
    Agregar: TAdvGlowButton;
    Editar: TAdvGlowButton;
    Salvar: TAdvGlowButton;
    Cancelar: TAdvGlowButton;
    Eliminar: TAdvGlowButton;
    Label16: TLabel;
    txtNombre: TEdit;
    pSalidassContrato: TStringField;
    pSalidasiFolioSalida: TIntegerField;
    pSalidasdFechaSalida: TDateField;
    pSalidassIdInsumo: TStringField;
    pSalidassIdAlmacen: TStringField;
    pSalidasdCantidad: TFloatField;
    pSalidassIdUsuario: TStringField;
    pSalidassNumeroOrden: TStringField;
    pSalidasmDescripcion: TStringField;
    iFolio: TCurrencyEdit;
    pSalidassNumeroActividad: TStringField;
    Label1: TLabel;
    ds_tipomovimiento: TDataSource;
    zq_tipomovimiento: TZReadOnlyQuery;
    tsTipomovimiento: TDBLookupComboBox;
    anexo_suministro: TZQuery;
    tsPartida: TDBLookupComboBox;
    lbl1: TLabel;
    lbl2: TLabel;
    ds_plataformas: TDataSource;
    qryPlataforma: TZReadOnlyQuery;
    tsPlataformaRef: TDBLookupComboBox;
    ds_partidas: TDataSource;
    qryPartidasRef: TZReadOnlyQuery;
    lbl3: TLabel;
    lbl4: TLabel;
    tsTrazabilidad: TEdit;
    lbl5: TLabel;
    tsMedida: TEdit;
    lbl6: TLabel;
    lbl7: TLabel;
    lbl8: TLabel;
    tmDescripcion: TMemo;
    pSalidassIdPlataforma: TStringField;
    ReportesContrato: TStringField;
    dtfldReportedFechaSalida: TDateField;
    ReportesIdTipo: TStringField;
    ReportesNombre: TStringField;
    ReportesNumeroOrden: TStringField;
    ReportesNumeroActividad: TStringField;
    ReportesIdPlataforma: TStringField;
    ReportesIdUsuario: TStringField;
    ReportemComentarios: TStringField;
    ReportesContrato_1: TStringField;
    intgrfldReporteiFolioSalida_1: TIntegerField;
    dtfldReportedFechaSalida_1: TDateField;
    ReportesIdInsumo: TStringField;
    ReportesWbs: TStringField;
    ReportesNumeroActividad_1: TStringField;
    ReportesIdAlmacen: TStringField;
    fltfldReportedCantidad: TFloatField;
    ReportesIdUsuario_1: TStringField;
    ReportesNumeroOrden_1: TStringField;
    ReportesIdPlataforma_1: TStringField;
    ReportesTrazabilidad: TStringField;
    fltfldReportedExistencia: TFloatField;
    fltfldReportedCostoMN: TFloatField;
    ReportemDescripcion: TMemoField;
    ReportesMedida: TStringField;
    ReporteTipomovimiento: TStringField;
    pSalidassTrazabilidad: TStringField;
    SpeedButton1: TSpeedButton;
    pSalidasiItem: TIntegerField;
    GbInferior: TGroupBox;
    Splitter1: TSplitter;
    ReporteiItem: TIntegerField;
    lblEstadoFolio: TcxLabel;
    QryBuscaFolio: TZQuery;
    Grid_Entradas: TRxDBGrid;
    anexo_suministroEstadoFolio: TStringField;
    anexo_suministroiFolioSalida: TIntegerField;
    anexo_suministrodFechaSalida: TDateField;
    anexo_suministrosNumeroOrden: TStringField;
    anexo_suministrosNombre: TStringField;
    anexo_suministromComentarios: TStringField;
    anexo_suministrosIdPlataforma: TStringField;
    anexo_suministrosNumeroActividad: TStringField;
    pSalidasiFolioAnio: TIntegerField;
    anexo_suministroiFolioAnio: TIntegerField;
    lbl9: TLabel;
    tsArchivo: TEdit;
    btnFiles: TBitBtn;
    OpenXLS: TOpenDialog;
    ExcelApplication1: TExcelApplication;
    ExcelWorkbook1: TExcelWorkbook;
    Excel: TExcelWorksheet;
    btnImportarVales: TBitBtn;
    zq_tipomovimientosIdTipo: TStringField;
    zq_tipomovimientosDescripcion: TStringField;
    zq_tipomovimientosClasificacion: TStringField;
    anexo_suministrosIdTipo: TStringField;
    EdtExt: TAdvEdit;
    anexo_suministrosext: TStringField;
    anexo_suministrosFolio: TStringField;
    ReporteiFolioSalida: TIntegerField;
    ReportesFolioSalida: TStringField;
    pSalidasiFolioAviso: TIntegerField;
    PedidoiFolio: TIntegerField;
    PedidodCantidad: TFloatField;
    PedidodCantidadRestante: TFloatField;
    PedidosIdInsumo: TStringField;
    PedidosTrazabilidad: TStringField;
    PedidosMedida: TStringField;
    PedidosDescripcion: TStringField;
    PedidomDescripcion: TMemoField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BtnExitClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tsIsometricoReferenciaKeyPress(Sender: TObject;
      var Key: Char);
    procedure frxReport50GetValue(const VarName: String;
      var Value: Variant);
    procedure frmBarra2btnAddClick(Sender: TObject);
    procedure frmBarra2btnEditClick(Sender: TObject);
    procedure frmBarra2btnPostClick(Sender: TObject);
    procedure frmBarra2btnDeleteClick(Sender: TObject);
    procedure frmBarra2btnRefreshClick(Sender: TObject);
    procedure frmBarra2btnCancelClick(Sender: TObject);
    procedure frmBarra2btnExitClick(Sender: TObject);
    procedure tdIdFechaEnter(Sender: TObject);
    procedure tdIdFechaExit(Sender: TObject);
    procedure tsOrigenKeyPress(Sender: TObject; var Key: Char);
    procedure tmComentariosEnter(Sender: TObject);
    procedure tmComentariosExit(Sender: TObject);
    procedure tdCantidadEnter(Sender: TObject);
    procedure tdCantidadExit(Sender: TObject);
    procedure tsNumeroActividadKeyPress(Sender: TObject; var Key: Char);
    procedure frxEntradaGetValue(const VarName: String;
      var Value: Variant);
    procedure ComentariosAdicionalesClick(Sender: TObject);
    procedure tdFechaAvisoKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure ActivaBotones(Sender :boolean);
    procedure ActivaBotones2(Sender :boolean);
    procedure mComentariosKeyPress(Sender: TObject; var Key: Char);
    procedure tsFamiliaKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tmComentariosKeyPress(Sender: TObject; var Key: Char);
    procedure Grid_PedidoKeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra2btnPrinterClick(Sender: TObject);
    procedure tsAlmacenExit(Sender: TObject);
    procedure AgregarClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure SalvarClick(Sender: TObject);
    procedure CancelarClick(Sender: TObject);
    procedure EliminarClick(Sender: TObject);
    procedure txtNombreEnter(Sender: TObject);
    procedure txtNombreContextPopup(Sender: TObject; MousePos: TPoint;
      var Handled: Boolean);
    procedure txtNombreExit(Sender: TObject);
    procedure tsInsumoExit(Sender: TObject);
    procedure tsInsumoKeyPress(Sender: TObject; var Key: Char);
    procedure tdCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure tsInsumoEnter(Sender: TObject);
    procedure txtNombreKeyPress(Sender: TObject; var Key: Char);
    procedure iFolioKeyPress(Sender: TObject; var Key: Char);
    procedure tdIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure iFolioEnter(Sender: TObject);
    procedure iFolioExit(Sender: TObject);
    procedure tsAlmacenKeyPress(Sender: TObject; var Key: Char);
    procedure dbPartidasKeyPress(Sender: TObject; var Key: Char);
    procedure mDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure Grid_EntradasTitleClick(Column: TColumn);
    procedure Grid_PedidoTitleClick(Column: TColumn);
    procedure GridPartidasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure tdCantidadChange(Sender: TObject);
    procedure iFolioChange(Sender: TObject);
    procedure tsInsumoKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure tsTipomovimientoEnter(Sender: TObject);
    procedure tsTipomovimientoxit(Sender: TObject);
    procedure tsTipomovimientoKeyPress(Sender: TObject; var Key: Char);
    procedure tsPartidaKeyPress(Sender: TObject; var Key: Char);
    procedure tsPartidaEnter(Sender: TObject);
    procedure tsPartidaExit(Sender: TObject);
    procedure tsPlataformaRefKeyPress(Sender: TObject; var Key: Char);
    procedure PedidoAfterScroll(DataSet: TDataSet);
    procedure anexo_suministroAfterScroll(DataSet: TDataSet);
    procedure pSalidasAfterScroll(DataSet: TDataSet);
    procedure Grid_PedidoDblClick(Sender: TObject);
    procedure PgControlChanging(Sender: TObject; var AllowChange: Boolean);
    procedure SpeedButton1Click(Sender: TObject);
    procedure tsNumeroOrdenChange(Sender: TObject);
    procedure buscaEstadoFolio;
    procedure Grid_EntradasDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure anexo_suministroCalcFields(DataSet: TDataSet);
    procedure btnFilesClick(Sender: TObject);
    procedure tsArchivoChange(Sender: TObject);
    procedure btnImportarValesClick(Sender: TObject);
    procedure tsTrazabilidadKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure tsTrazabilidadEnter(Sender: TObject);
    procedure tsTrazabilidadExit(Sender: TObject);
  private
  ValorPrev:Real;
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSalidaAlmacen: TfrmSalidaAlmacen;
  sDescripcion  : String ;
  txtAux        : String ;
  lNuevo        : Boolean ;
  OpcButton1    : String ;
  FechaEAnt     : String ;
  SavePlace,
  SavePlace2    : TBookmark;
  sBackup,
  IdInsumo      : String ;
  filtro, buscar  : string;
  Cantidad      : Double;
  TipoExplosion : string;
  utgrid:ticdbgrid;
  utgrid2:ticdbgrid;
  utgrid3:ticdbgrid;
  botonpermiso  : tbotonespermisos;
  numeroOrden   : string;
  folioSalida   : Integer;
  lContinuar, lContinuaReg  : Boolean;
  OldFolio:Integer;
  OldExt:string;
  FolioAnt      : double;

implementation

uses frm_connection , frm_comentariosxanexo, frm_entradaanex;

{$R *.dfm}

procedure TfrmSalidaAlmacen.btnFilesClick(Sender: TObject);
begin
  OpenXLS.Filter := 'Hoja de calculo Excel 03-13|*.xlsx';
  OpenXLS.FilterIndex := 1;
  if OpenXLS.Execute then
  begin
    tsArchivo.Text := OpenXLS.FileName;
  end;
end;

procedure TfrmSalidaAlmacen.btnImportarValesClick(Sender: TObject);
var
  //Variables para la coneccion a la hoja de
  Mensajes: string;
  flcid : Integer;
  //TZQuery's para la inserccion
  qrBusca, qrInformacion, qrMateriales : TZQuery;
  //Filas
  iFila, iFilaMax, iFilaCount   : integer;
  //Informacion
  sValue : string;
  sFolio, sPartida, sPlataforma, sNumero, sIdTipo, sPuesto, sMedida, sProveedor, sAlmacen : string;
  //Materiales
  sIdInsumo, sTrazabilidad : string;
  dCantidad : Double;
  iFolioAnio, iItem, iFolioAviso : Integer;

  idNumero  : string;
  sfecha    : string;
  sConcepto : string;
  sCantidad : string;

  procedure vacio(sParamColumnaDato, sParamColumnaMsg, sParamTexto : string; iColor :integer);
  begin
    Excel.Range[sParamColumnaMsg, sParamColumnaMsg].Value := sParamTexto;
    Excel.Range[sParamColumnaMsg, sParamColumnaMsg].Font.Size  := 12;
    Excel.Range[sParamColumnaMsg, sParamColumnaMsg].font.Bold  := True;
    Excel.Range[sParamColumnaMsg, sParamColumnaMsg].Font.Color := clRed;
    Excel.Range[sParamColumnaDato, sParamColumnaDato].Interior.ColorIndex := iColor;
    lContinuar := False;
  end;

begin

{$REGION 'CONECTAR A EXCEL'}
  if trim(tsArchivo.Text) = '' then
  begin
      messageDLg('Debe seleccionar un archivo de Excel!', mtInformation, [mbOk], 0);
      exit;
  end;
  flcid := GetUserDefaultLCID;
  ExcelApplication1.Connect;
  ExcelApplication1.Visible[flcid] := true;
  ExcelApplication1.UserControl := true;

  ExcelWorkbook1.ConnectTo(ExcelApplication1.Workbooks.Open(tsArchivo.Text,
        emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam,
        emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, flcid));

  Excel.ConnectTo(ExcelWorkbook1.Sheets.Item[1] as ExcelWorkSheet);
{$ENDREGION}

  qrBusca       := TZQuery.Create(nil);
  qrInformacion := TZQuery.Create(nil);
  qrMateriales  := TZQuery.Create(nil);
  qrBusca.Connection       := connection.zConnection;
  qrInformacion.Connection := connection.zConnection;
  qrMateriales.Connection  := connection.zConnection;

{$REGION 'VALIDACION CARATULA SALIDA DE EXCEL'}
    idNumero      := Excel.Range['J4','J4'].Value2;
    sFecha        := Excel.Range['J8','J8'].Value2;
    sFolio        := Excel.Range['K11','K11'].Value2;
    sPartida      := Excel.Range['K12','K12'].Value2;
    sPlataforma   := Excel.Range['K13','K13'].Value2;

    {Limpiamos los colores..}
    Excel.Range['J4', 'J4' ].Interior.ColorIndex := 2;
    Excel.Range['J8', 'J8' ].Interior.ColorIndex := 2;
    Excel.Range['K11','K11'].Interior.ColorIndex := 2;
    Excel.Range['K12','K12'].Interior.ColorIndex := 2;
    Excel.Range['K13','K13'].Interior.ColorIndex := 2;

    {Eliminamos los textos..}
    Excel.Range['M4', 'M4' ].Value := '';
    Excel.Range['M8', 'M8' ].Value := '';
    Excel.Range['M11','M11'].Value := '';
    Excel.Range['M12','M12'].Value := '';
    Excel.Range['M13','M13'].Value := '';

    lContinuar := True;
    {Validacion de datos vacios}
    if idNumero = '' then
       vacio('J4', 'M4', 'ID VACIO', 6);
    if sFecha = '' then
       vacio('J8', 'M8', 'FECHA VACIA', 6);
    if sFolio = '' then
       vacio('K11', 'M11', 'FOLO VACIO', 6);
    if sPartida = '' then
       vacio('K12', 'M12', 'PARTIDA VACIA', 6);
    if sPlataforma = '' then
       vacio('K13', 'M13', 'PLATAFORMA VACIA', 6);

    {Ahora Validamos datos contenidos en la base de datos}
    if lContinuar then
    begin
        {folio o No. de salida}
        qrBusca.Active := False;
        qrBusca.SQL.Clear;
        qrBusca.SQL.Add('select * from almacen_salida '+
                        'where sContrato = :contrato and iFolioSalida = :folioS and iFolioAnio = :anio');
        qrBusca.Params.ParamByName('contrato').DataType := ftString;
        qrBusca.Params.ParamByName('contrato').Value    := param_global_contrato;
        qrBusca.Params.ParamByName('folioS').DataType   := ftInteger;
        qrBusca.Params.ParamByName('folioS').Value      := StrToInt(idNumero);
        qrBusca.Params.ParamByName('anio').DataType     := ftInteger;
        qrBusca.Params.ParamByName('anio').Value        := YearOf(StrToInt(sFecha));
        qrBusca.Open;

        if qrBusca.RecordCount > 0 then
        begin
            {Ahora agregamos la condicion donde decimos si se está subiendo mismo folio y partida}
            qrBusca.Active := False;
            qrBusca.SQL.Clear;
            qrBusca.SQL.Add('select * from almacen_salida '+
                            'where sContrato = :contrato and iFolioSalida = :folioS and iFolioAnio = :anio and sNumeroOrden =:Folio and sNumeroActividad =:Actividad ');
            qrBusca.Params.ParamByName('contrato').DataType  := ftString;
            qrBusca.Params.ParamByName('contrato').Value     := param_global_contrato;
            qrBusca.Params.ParamByName('folioS').DataType    := ftInteger;
            qrBusca.Params.ParamByName('folioS').Value       := StrToInt(idNumero);
            qrBusca.Params.ParamByName('anio').DataType      := ftInteger;
            qrBusca.Params.ParamByName('anio').Value         := YearOf(StrToInt(sFecha));
            qrBusca.Params.ParamByName('Folio').DataType     := ftString;
            qrBusca.Params.ParamByName('Folio').Value        := sFolio;
            qrBusca.Params.ParamByName('Actividad').DataType := ftString;
            qrBusca.Params.ParamByName('Actividad').Value    := sPartida;
            qrBusca.Open;

            {Aqui mandamos el mensaje para el usuario..}
            if qrBusca.RecordCount > 0 then
            begin
                vacio('J4', 'M4', 'EL NO. DE SALIDA YA EXISTE, CON EL MISMO FOLIO Y PARTIDA', 6);
                lContinuar := False;
            end;
        end;

        {Plataforma si no existe..}
        qrBusca.Active := False;
        qrBusca.SQL.Clear;
        qrBusca.SQL.Add('select sIdPlataforma from plataformas where sIdPlataforma = :plataforma');
        qrBusca.Params.ParamByName('plataforma').DataType := ftString;
        qrBusca.Params.ParamByName('plataforma').Value    := sPlataforma;
        qrBusca.Open;

        if qrBusca.RecordCount = 0 then
        begin
            vacio('K13', 'M13', 'LA PLATAFORMA NO EXISTE', 6);
            lContinuar := False;
        end;

        {Folio u orden de trabajo sino existe..}
        qrBusca.Active := False;
        qrBusca.SQL.Clear;
        qrBusca.SQL.Add('select sIdFolio from ordenesdetrabajo '+
                        'where sContrato = :contrato and sNumeroOrden = :folio');
        qrBusca.Params.ParamByName('contrato').DataType := ftString;
        qrBusca.Params.ParamByName('contrato').Value    := param_global_contrato;
        qrBusca.Params.ParamByName('folio').DataType    := ftString;
        qrBusca.Params.ParamByName('folio').Value       := sFolio;
        qrBusca.Open;

        if qrBusca.RecordCount = 0 then
        begin
            vacio('K11', 'M11', 'EL FOLIO NO EXISTE', 6);
            lContinuar := False;
        end;

        {Partida de anexo si existe o no existe..}
        qrBusca.Active := False;
        qrBusca.SQL.Clear;
        qrBusca.SQL.Add('select sNumeroActividad from actividadesxorden '+
                        'where sContrato = :contrato and sIdconvenio =:Convenio and sNumeroOrden =:folio and sNumeroActividad =:partida and sTipoActividad = "Actividad"');
        qrBusca.Params.ParamByName('contrato').DataType := ftString;
        qrBusca.Params.ParamByName('contrato').Value    := param_global_contrato;
        qrBusca.Params.ParamByName('convenio').DataType := ftString;
        qrBusca.Params.ParamByName('convenio').Value    := global_convenio;
        qrBusca.Params.ParamByName('folio').DataType    := ftString;
        qrBusca.Params.ParamByName('folio').Value       := sFolio;
        qrBusca.Params.ParamByName('partida').DataType  := ftString;
        qrBusca.Params.ParamByName('partida').Value     := sPartida;
        qrBusca.Open;

        if qrBusca.RecordCount = 0 then
        begin
            vacio('K12', 'M12', 'LA PARTIDA DEL FOLIO NO EXISTE', 6);
            lContinuar := False;
        end;
    end;
{$ENDREGION}

{$REGION 'VALIDACION E INSERCION DE MATERIALES EXCEL'}
  iFila := 1;
  iFilaMax := 0;
  {buscamos la fila donde se encuentra codigo}
  sValue := Excel.Range['B' + IntToStr(iFila),'B' + IntToStr(iFila)].Value2;
  while iFila < 100 do
  begin
      sValue := Excel.Range['B' + IntToStr(iFila),'B' + IntToStr(iFila)].Value2;
      if sValue = 'CODIGO' then
      begin
         iFilaMax := iFila;
         inc(iFila, 100);
      end;
      inc(iFila);
  end;

  if iFila = 0 then
  begin
      messageDLg('No se encontro la columna con el titulo CODIGO, no se podran Importar los Materiales.', mtWarning, [mbOk], 0);
      exit;
  end
  else
  begin
      iFila     := iFilaMax + 1;
      iFilaCount  := 0;
      sIdInsumo := Excel.Range['B' + IntToStr(iFila),'B' + IntToStr(iFila)].Value2;
      {Se condiciona que maximo puede haber 3 espacios en blanco..}
      while (sIdInsumo <> '') or (iFilaCount < 3) do
      begin
          sIdInsumo     := Excel.Range['B' + IntToStr(iFila),'B' + IntToStr(iFila)].Value2;
          sConcepto     := Excel.Range['C' + IntToStr(iFila),'C' + IntToStr(iFila)].Value2;
          sTrazabilidad := Excel.Range['H' + IntToStr(iFila),'H' + IntToStr(iFila)].Value2;
          sCantidad     := Excel.Range['J' + IntToStr(iFila),'J' + IntToStr(iFila)].Value2;
          sMedida       := Excel.Range['K' + IntToStr(iFila),'K' + IntToStr(iFila)].Value2;

          {Limpiamos los colores..}
          Excel.Range['C' + IntToStr(iFila),'C' + IntToStr(iFila)].Interior.ColorIndex := 2;
          Excel.Range['H' + IntToStr(iFila),'H' + IntToStr(iFila)].Interior.ColorIndex := 2;
          Excel.Range['J' + IntToStr(iFila),'J' + IntToStr(iFila)].Interior.ColorIndex := 2;

          {Eliminamos los textos..}
          Excel.Range['M' + IntToStr(iFila),'M' + IntToStr(iFila)].Value := '';

          if sIdInsumo <> '' then
          begin
              if trim(sConcepto) = '' then
                 vacio('C'+IntToStr(iFila), 'M'+IntToStr(iFila), 'CONCEPTO VACIO', 6);
              if trim(sCantidad) = '' then
                 vacio('J'+IntToStr(iFila), 'M'+IntToStr(iFila), 'CANTIDAD VACIA', 6);
              if trim(sMedida) = '' then
                 vacio('K'+IntToStr(iFila), 'M'+IntToStr(iFila), 'UNIDAD DE MEIDA VACIA', 6);
              iFilaCount := 0;
          end
          else
             inc(iFilaCount);

          Inc(iFila);
      end;
  end;
 {$ENDREGION}

 {$REGION 'INSERCION DE MATERIALES NO EXISTENTES AL CATALOGO DE MATERIALES'}
 if lContinuar then
 begin
     {buscamos el primer aviso de embarque..}
     connection.QryBusca.Active := False;
     connection.QryBusca.SQL.Clear;
     connection.QryBusca.SQl.Add('select max(iFolio) as Folio from anexo_suministro where sContrato = :contrato ');
     connection.QryBusca.Params.ParamByName('contrato').DataType     := ftString;
     connection.QryBusca.Params.ParamByName('contrato').Value        := param_global_contrato;
     connection.QryBusca.Open;

     if connection.QryBusca.RecordCount > 0 then
        iFolioAviso := connection.QryBusca.FieldValues['Folio']
     else
     begin
         messageDLG('No existe un Aviso de Embarque para la orden Actual, debe crearlo para poder continuar!', mtInformation, [mbOk], 0);
         exit;
     end;

     {Buscamos el proveedor PEP en el catalogod de proveedores}
     connection.QryBusca2.Active := False;
     connection.QryBusca2.SQL.Clear;
     connection.QryBusca2.SQL.Add('select sIdProveedor from proveedores where sIdProveedor like "%PEP%" limit 1');
     connection.QryBusca2.Open;

     if connection.QryBusca2.RecordCount > 0 then
        sProveedor := connection.QryBusca2.FieldValues['sIdProveedor'];

     {Buscamos el ALAMCEN}
     connection.QryBusca2.Active := False;
     connection.QryBusca2.SQL.Clear;
     connection.QryBusca2.SQL.Add('select sIdAlmacen from almacenes');
     connection.QryBusca2.Open;

     if connection.QryBusca2.RecordCount > 0 then
        sAlmacen := connection.QryBusca2.FieldValues['sIdAlmacen'];

     iFila      := iFilaMax + 1;
     iFilaCount := 0;
     sIdInsumo  := Excel.Range['B' + IntToStr(iFila),'B' + IntToStr(iFila)].Value2;
     {Se condiciona que maximo puede haber 3 espacios en blanco..}
     while (sIdInsumo <> '') or (iFilaCount < 3) do
     begin
         sIdInsumo     := Excel.Range['B' + IntToStr(iFila),'B' + IntToStr(iFila)].Value2;
         sConcepto     := Excel.Range['C' + IntToStr(iFila),'C' + IntToStr(iFila)].Value2;
         sTrazabilidad := Excel.Range['H' + IntToStr(iFila),'H' + IntToStr(iFila)].Value2;
         sCantidad     := Excel.Range['J' + IntToStr(iFila),'J' + IntToStr(iFila)].Value2;
         sMedida       := Excel.Range['K' + IntToStr(iFila),'K' + IntToStr(iFila)].Value2;

         if sIdInsumo <> '' then
         begin
             lContinuaReg := True;
             {Consultamos si existe el insumo por Descripcion}
             connection.QryBusca.Active := False;
             connection.QryBusca.SQL.Clear;
             connection.QryBusca.SQl.Add('select sIdInsumo from insumos '+
                                         'where sContrato = :contrato and mDescripcion = :descripcion and sTrazabilidad = :trazabilidad ');
             connection.QryBusca.Params.ParamByName('contrato').DataType     := ftString;
             connection.QryBusca.Params.ParamByName('contrato').Value        := global_contrato_barco;
             connection.QryBusca.Params.ParamByName('descripcion').DataType  := ftString;
             connection.QryBusca.Params.ParamByName('descripcion').Value     := sConcepto;
             connection.QryBusca.Params.ParamByName('trazabilidad').DataType := ftString;
             connection.QryBusca.Params.ParamByName('trazabilidad').Value    := sTrazabilidad;
             connection.QryBusca.Open;

             if connection.QryBusca.RecordCount = 0 then
                lContinuaReg := False;

             if lContinuaReg then
             begin
                 {Consultamos si existe el insumo por Id y trazabilidad}
                 connection.QryBusca.Active := False;
                 connection.QryBusca.SQL.Clear;
                 connection.QryBusca.SQl.Add('select sIdInsumo from insumos '+
                                             'where sContrato = :contrato and sIdInsumo = :insumo and sTrazabilidad = :trazabilidad');
                 connection.QryBusca.Params.ParamByName('contrato').DataType     := ftString;
                 connection.QryBusca.Params.ParamByName('contrato').Value        := global_contrato_barco;
                 connection.QryBusca.Params.ParamByName('insumo').DataType       := ftString;
                 connection.QryBusca.Params.ParamByName('insumo').Value          := sIdInsumo;
                 connection.QryBusca.Params.ParamByName('trazabilidad').DataType := ftString;
                 connection.QryBusca.Params.ParamByName('trazabilidad').Value    := sTrazabilidad;
                 connection.QryBusca.Open;

                 if connection.QryBusca.RecordCount = 0 then
                    lContinuaReg := False;
             end;

             if lContinuaReg = False then
             begin
                //Ahora insertamos el material..
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('INSERT INTO insumos ( sContrato, sIdInsumo, sIdProveedor, sIdAlmacen, sTipoActividad, mDescripcion, dFechaInicio, dCostoMN, dCostoDLL, dVentaMN, dVentaDLL, sMedida, '+
                ' dCantidad, dInstalado, sIdGrupo, dNuevoPrecio, sIdFase, sTrazabilidad, sLabelIdMaterial, sColumnaAux ) '+
                ' VALUES (:contrato, :insumo, :proveedor, :almacen, :tipoactividad, :Descripcion, :fechai, 0, 0, 0, 0, :medida, 0, 0, null, 0, null, :trazabilidad, :labelmaterial, :insumo)');
                connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
                connection.zCommand.Params.ParamByName('contrato').value    := global_contrato_barco;
                connection.zCommand.Params.ParamByName('insumo').DataType   := ftString;
                connection.zCommand.Params.ParamByName('insumo').value      := sIdInsumo;
                connection.zCommand.Params.ParamByName('almacen').DataType  := ftString;
                connection.zCommand.Params.ParamByName('almacen').value     := sAlmacen;
                connection.zCommand.Params.ParamByName('tipoactividad').DataType := ftString;
                connection.zCommand.Params.ParamByName('tipoactividad').value  := 'Material';
                connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString;
                connection.zCommand.Params.ParamByName('Descripcion').value    := sConcepto;
                connection.zCommand.Params.ParamByName('fechai').DataType      := ftDate;
                connection.zCommand.Params.ParamByName('fechai').value         := date;
                connection.zCommand.Params.ParamByName('medida').DataType      := ftString;
                connection.zCommand.Params.ParamByName('medida').value         := sMedida;
                connection.zCommand.Params.ParamByName('trazabilidad').DataType:= ftString;
                connection.zCommand.Params.ParamByName('trazabilidad').value   := sTrazabilidad;
                connection.zCommand.Params.ParamByName('labelmaterial').DataType := ftString;
                connection.zCommand.Params.ParamByName('labelmaterial').value    := sIdInsumo;
                connection.zCommand.Params.ParamByName('proveedor').DataType  := ftString;
                if sProveedor <> '' then
                   connection.zCommand.Params.ParamByName('proveedor').value  := sProveedor
                else
                   connection.zCommand.Params.ParamByName('proveedor').value  := Null;
                connection.zCommand.ExecSQL;
             end;

             {Ahora insertamos el material sino existe en el Aviso de embarque..}
             connection.QryBusca.Active := False;
             connection.QryBusca.SQL.Clear;
             connection.QryBusca.SQl.Add('select sIdInsumo from anexo_psuministro '+
                                         'where sContrato = :contrato and iFolio = :Folio and sIdInsumo = :insumo and sTrazabilidad =:trazabilidad ');
             connection.QryBusca.Params.ParamByName('contrato').DataType := ftString;
             connection.QryBusca.Params.ParamByName('contrato').Value    := param_global_contrato;
             connection.QryBusca.Params.ParamByName('Folio').DataType    := ftInteger;
             connection.QryBusca.Params.ParamByName('Folio').Value       := iFolioAviso;
             connection.QryBusca.Params.ParamByName('Insumo').DataType   := ftString;
             connection.QryBusca.Params.ParamByName('Insumo').Value      := sIdInsumo;
             connection.QryBusca.Params.ParamByName('trazabilidad').DataType := ftString;
             connection.QryBusca.Params.ParamByName('trazabilidad').Value    := sTrazabilidad;
             connection.QryBusca.Open;

             if connection.QryBusca.RecordCount = 0 then
             begin
                 {Insertamos el material que no existe en el aviso de embarque..}
                 connection.zCommand.Active := False;
                 connection.zCommand.SQL.Clear;
                 connection.zCommand.SQL.Add('INSERT INTO anexo_psuministro ( sContrato, iFolio ,swbs, sNumeroActividad, dCantidad, dCantidadRestante, sIdInsumo, sTrazabilidad ) ' +
                            'VALUES (:Contrato, :Folio,:swbs, :Actividad, :Cantidad, :Cantidad, :Insumo, :trazabilidad )');
                 connection.zCommand.Params.ParamByName('Contrato').DataType  := ftString;
                 connection.zCommand.Params.ParamByName('Contrato').value     := param_Global_Contrato;
                 connection.zCommand.Params.ParamByName('Folio').DataType     := ftInteger;
                 connection.zCommand.Params.ParamByName('Folio').value        := iFolioAviso;
                 connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
                 connection.zCommand.Params.ParamByName('Actividad').value    := '';
                 connection.zCommand.Params.ParamByName('Cantidad').DataType  := ftFloat;
                 connection.zCommand.Params.ParamByName('Cantidad').value     := StrToFloat(sCantidad);
                 connection.zCommand.Params.ParamByName('swbs').AsString      := '';
                 connection.zCommand.Params.ParamByName('Insumo').DataType    := ftString;
                 connection.zCommand.Params.ParamByName('Insumo').value       := sIdInsumo;
                 connection.zCommand.Params.ParamByName('Trazabilidad').DataType := ftString;
                 connection.zCommand.Params.ParamByName('Trazabilidad').value    := sTrazabilidad;
                 connection.zCommand.ExecSQL;
             end;
             iFilaCount := 0;
         end
         else
            inc(iFilaCount);

         Inc(iFila);
     end;
 end;
 {$ENDREGION}

{$REGION 'IMPORTACION'}
  if lcontinuar then
  begin
      try
        connection.zConnection.StartTransaction;
      {$REGION 'INFORMACION'}
        qrBusca.Active := False;
        qrBusca.SQl.Clear;
        qrBusca.SQL.Add('select a.sIdTipo, b.sPuesto from movimientosdealmacen a, usuarios b '+
                        'where a.sClasificacion = "Salida" and b.sIdUsuario = :usuario');
        qrBusca.Params.ParamByName('usuario').DataType := ftString;
        qrBusca.Params.ParamByName('usuario').Value    := global_usuario;
        qrBusca.Open;

        if qrBusca.RecordCount > 0 then
        begin
            sIdTipo := qrBusca.FieldByName('sIdTipo').AsString;
            sPuesto := qrBusca.FieldByName('sPuesto').AsString;
        end;

        qrBusca.Active := False;
        qrBusca.SQL.Clear;
        qrBusca.SQL.Add('select * from almacen_salida '+
                        'where sContrato = :contrato and iFolioSalida = :folioS and iFolioAnio = :anio');
        qrBusca.Params.ParamByName('contrato').DataType := ftString;
        qrBusca.Params.ParamByName('contrato').Value    := param_global_contrato;
        qrBusca.Params.ParamByName('folioS').DataType   := ftInteger;
        qrBusca.Params.ParamByName('folioS').Value      := StrToInt(idNumero);
        qrBusca.Params.ParamByName('anio').DataType     := ftInteger;
        qrBusca.Params.ParamByName('anio').Value        := YearOf(StrToInt(sFecha));
        qrBusca.Open;

        {Si la caratula no existe se inserta..}
        if qrBusca.RecordCount = 0 then
        begin
            {Insertamos la caratula de la salida de materiales}
            qrInformacion.Active := False;
            qrInformacion.SQL.Clear;
            qrInformacion.SQL.Add('insert into almacen_salida (sContrato, iFolioSalida, iFolioAnio, dFechaSalida, sIdTipo, '+
                                                              'sNombre, sNumeroOrden, sNumeroActividad, sIdPlataforma, '+
                                                              'sIdUsuario, mComentarios) '+
                                  'values (:contrato, :numero, :anio, :fecha, :tipo, :puesto, :folio, :actividad, :plataforma, :usuario, "*")');
            qrInformacion.Params.ParamByName('contrato').DataType    := ftString;
            qrInformacion.Params.ParamByName('contrato').Value       := param_global_contrato;
            qrInformacion.Params.ParamByName('numero').DataType      := ftInteger;
            qrInformacion.Params.ParamByName('numero').Value         := idNumero;
            qrInformacion.Params.ParamByName('anio').DataType        := ftInteger;
            qrInformacion.Params.ParamByName('anio').Value           := YearOf(StrToInt(sFecha));
            qrInformacion.Params.ParamByName('fecha').DataType       := ftDatetime;
            qrInformacion.Params.ParamByName('fecha').Value          := sFecha;
            qrInformacion.Params.ParamByName('tipo').DataType        := ftString;
            qrInformacion.Params.ParamByName('tipo').Value           := sIdTipo;
            qrInformacion.Params.ParamByName('puesto').DataType      := ftString;
            qrInformacion.Params.ParamByName('puesto').Value         := sPuesto;
            qrInformacion.Params.ParamByName('folio').DataType       := ftString;
            qrInformacion.Params.ParamByName('folio').Value          := sFolio;
            qrInformacion.Params.ParamByName('actividad').DataType   := ftString;
            qrInformacion.Params.ParamByName('actividad').Value      := sPartida;
            qrInformacion.Params.ParamByName('plataforma').DataType  := ftString;
            qrInformacion.Params.ParamByName('plataforma').Value     := sPlataforma;
            qrInformacion.Params.ParamByName('usuario').DataType     := ftString;
            qrInformacion.Params.ParamByName('usuario').Value        := global_usuario;
            qrInformacion.ExecSQL;
        end;
      {$ENDREGION}

      {$REGION 'MATERIALES'}
        iFila     := iFilaMax + 1;
        iFilaCount  := 0;
        sIdInsumo := Excel.Range['B' + IntToStr(iFila),'B' + IntToStr(iFila)].Value2;
        {Se condiciona que maximo puede haber 3 espacios en blanco..}
        while (sIdInsumo <> '') or (iFilaCount < 3) do
        begin
            sIdInsumo     := Excel.Range['B' + IntToStr(iFila),'B' + IntToStr(iFila)].Value2;
            sConcepto     := Excel.Range['C' + IntToStr(iFila),'C' + IntToStr(iFila)].Value2;
            sTrazabilidad := Excel.Range['H' + IntToStr(iFila),'H' + IntToStr(iFila)].Value2;
            sCantidad     := Excel.Range['J' + IntToStr(iFila),'J' + IntToStr(iFila)].Value2;
            sMedida       := Excel.Range['K' + IntToStr(iFila),'K' + IntToStr(iFila)].Value2;

            if sIdInsumo <> '' then
            begin
                qrBusca.Active := False ;
                qrBusca.SQL.Clear ;
                qrBusca.SQL.Add ('select max(iItem) as maximo from bitacoradesalida '+
                                 'where sContrato =:Contrato and iFolioSalida =:Folio and sIdInsumo =:Insumo and sIdAlmacen =:Almacen '+
                                 'and sNumeroActividad =:Actividad and sTrazabilidad =:trazabilidad') ;
                qrBusca.Params.ParamByName('Contrato').DataType  := ftString ;
                qrBusca.Params.ParamByName('Contrato').value     := param_global_contrato;
                qrBusca.Params.ParamByName('Folio').DataType     := ftInteger ;
                qrBusca.Params.ParamByName('Folio').value        := idNumero;
                qrBusca.Params.ParamByName('Insumo').DataType    := ftString ;
                qrBusca.Params.ParamByName('Insumo').value       := sIdInsumo;
                qrBusca.Params.ParamByName('Almacen').DataType   := ftString;
                qrBusca.Params.ParamByName('Almacen').Value      := sAlmacen;
                qrBusca.Params.ParamByName('Actividad').DataType := ftString;
                qrBusca.Params.ParamByName('Actividad').Value    := sPartida;
                qrBusca.Params.ParamByName('trazabilidad').DataType := ftString;
                qrBusca.ParamByName('trazabilidad').Value           := sTrazabilidad;
                qrBusca.Open ;
                if Length(trim(qrBusca.FieldByName('maximo').AsString)) > 0 then
                   iItem :=  qrBusca.FieldByName('maximo').AsInteger + 1
                else
                   iItem := 0;

                try
                  {$REGION 'INSERTAR'}
                  qrMateriales.Active := False ;
                  qrMateriales.SQL.Clear ;
                  qrMateriales.SQL.Add('insert into bitacoradesalida (sContrato, iFolioSalida, iFolioAnio, dFechaSalida, '+
                                               'swbs, sIdInsumo, dCantidad, sIdUsuario, sIdAlmacen, '+
                                               'sNumeroOrden, sNumeroActividad, sIdPlataforma, sTrazabilidad, iItem) '+
                                        'values (:contrato, :folio, :folioanio, :fechaS, "" ,:insumo, :cantidad, '+
                                        ':usuario, :almacen, :orden, :actividad, :plataforma, :trazabilidad, :item)');
                  qrMateriales.Params.ParamByName('contrato').DataType     := ftString;
                  qrMateriales.Params.ParamByName('contrato').value        := param_global_contrato ;
                  qrMateriales.Params.ParamByName('folio').DataType        := ftInteger;
                  qrMateriales.Params.ParamByName('folio').value           := idNumero;
                  qrMateriales.Params.ParamByName('folioanio').DataType    := ftInteger;
                  qrMateriales.Params.ParamByName('folioanio').value       := YearOf(StrToInt(sFecha));;
                  qrMateriales.Params.ParamByName('fechaS').DataType       := ftDate;
                  qrMateriales.Params.ParamByName('fechaS').value          := sFecha;
                  qrMateriales.Params.ParamByName('insumo').DataType       := ftString;
                  qrMateriales.Params.ParamByName('insumo').value          := sIdInsumo;
                  qrMateriales.Params.ParamByName('cantidad').DataType     := ftFloat;
                  qrMateriales.Params.ParamByName('cantidad').value        := StrToFloat(sCantidad);
                  qrMateriales.Params.ParamByName('usuario').DataType      := ftString;
                  qrMateriales.Params.ParamByName('usuario').value         := global_usuario;
                  qrMateriales.Params.ParamByName('almacen').DataType      := ftString;
                  qrMateriales.Params.ParamByName('almacen').Value         := sAlmacen ;
                  qrMateriales.Params.ParamByName('orden').DataType        := ftString;
                  qrMateriales.Params.ParamByName('orden').value           := sFolio;
                  qrMateriales.Params.ParamByName('actividad').DataType    := ftString;
                  qrMateriales.Params.ParamByName('actividad').value       := sPartida;
                  qrMateriales.Params.ParamByName('plataforma').DataType   := ftString;
                  qrMateriales.Params.ParamByName('plataforma').value      := sPlataforma;
                  qrMateriales.Params.ParamByName('trazabilidad').DataType := ftString;
                  qrMateriales.Params.ParamByName('trazabilidad').value    := sTrazabilidad;
                  qrMateriales.Params.ParamByName('item').DataType         := ftInteger;
                  qrMateriales.Params.ParamByName('item').value            := iItem;
                  qrMateriales.ExecSQL;
                  {$ENDREGION}
                except
                end;

            end
            else
                inc(iFilaCount);

            inc(iFila)

        end;
        {$ENDREGION}
        connection.zConnection.Commit;
        anexo_suministro.Refresh;
        pSalidas.Refresh;
        messageDLG('La información se cargó Correctamente!', mtInformation, [mbOk], 0);
      except
      on e:Exception do
      begin
        if connection.zConnection.InTransaction then
        begin
          connection.zConnection.Rollback
        end;
          ShowMessage(e.Message);
        end;
  end;

  qrBusca.Destroy;
  qrInformacion.Destroy;
  qrMateriales.Destroy;
  end;

end;
//Fin Evento

procedure TfrmSalidaAlmacen.buscaEstadoFolio;
var status : string;
begin
  QryBuscaFolio.Active := False;
  QryBuscaFolio.Params.ParamByName('contrato').DataType := ftString ;
  QryBuscaFolio.Params.ParamByName('contrato').Value := param_global_contrato ;
  QryBuscaFolio.Params.ParamByName('orden').DataType := ftString;
  QryBuscaFolio.Params.ParamByName('orden').Value := tsNumeroOrden.Text;
  QryBuscaFolio.Open;

  status := QryBuscaFolio.FieldByName('cIdStatus').AsString;
  //Si se pasa True como valor del parametro "afectaA" se afectara el ComboBox
  begin
    if status = 'T' then
    begin
      lblEstadoFolio.Visible := True;
      tsNumeroOrden.Font.Color:= $00011421;
      tsNumeroOrden.Font.Style := [fsBold];
      tsNumeroOrden.Color :=  $001FA3FA;
      lblEstadoFolio.Caption := 'Folio Terminado';
    end
    else if status = 'S' then
    begin
      lblEstadoFolio.Visible := True;
      tsNumeroOrden.Font.Color:= $00002222;
      tsNumeroOrden.Font.Style := [fsBold];
      tsNumeroOrden.Color :=  $0000D9D9;
      lblEstadoFolio.Caption := 'Folio Suspendido';
    end
    else if status = 'P' then
    begin
      lblEstadoFolio.Visible := True;
      tsNumeroOrden.Font.Color:= $00481B02;
      tsNumeroOrden.Font.Style := [fsBold];
      tsNumeroOrden.Color :=  $00FA9C29;
      lblEstadoFolio.Caption := 'Folio en Proceso';
    end
    else
    begin
      lblEstadoFolio.visible := False;
    end;
  end;
end;

procedure filtra;
begin
    if filtro = 'Trazabilidad' then
    begin
        if length(trim(frmSalidaAlmacen.tsTrazabilidad.Text)) > 0  then
        begin
            buscar := frmSalidaAlmacen.tsTrazabilidad.Text;
            buscar := buscar + '*';
            filtro := ' sTrazabilidad like ' + QuotedStr(buscar)
        end ;

    end
    else
    begin
        if length(trim(frmSalidaAlmacen.tsInsumo.Text)) > 0  then
        begin
            buscar := frmSalidaAlmacen.tsInsumo.Text;
            buscar :=  '*'+ buscar + '*';
            filtro := ' mDescripcion like ' + QuotedStr(buscar);
        end;
    end;

  if filtro <> '' then
  begin
      frmSalidaAlmacen.pedido.Filtered := False;
      frmSalidaAlmacen.pedido.Filter   := filtro;
      frmSalidaAlmacen.pedido.Filtered := True;
  end
  else
     frmSalidaAlmacen.pedido.Filtered := False;

  filtro := '';
end;

procedure TfrmSalidaAlmacen.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  botonpermiso.Free;
  action := cafree ;
  utgrid.Destroy;
  utgrid2.destroy;
  utgrid3.Destroy;
end;


procedure TfrmSalidaAlmacen.FormShow(Sender: TObject);
var
    SavePlace : TBookmark;
    status : string;
begin
    sMenuP:=stMenu;
    BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'MnuSalAlmacen', PopupPrincipal);
    UtGrid:=TicdbGrid.create(grid_entradas);
    UtGrid2:=TicdbGrid.create(grid_pedido);
    UtGrid3:=TicdbGrid.create(gridpartidas);
    ActivaBotones(False);

    Pedido.Active := False;
    Pedido.Params.ParamByName('Contrato').DataType := ftString;
    Pedido.Params.ParamByName('Contrato').Value := global_Contrato_Barco;
    Pedido.Params.ParamByName('Orden').DataType := ftString;
    Pedido.Params.ParamByName('Orden').Value    := param_global_contrato;
    Pedido.Open;

    tsNumeroOrden.Items.Clear ;
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('select sNumeroOrden from ordenesdetrabajo where sContrato = :Contrato ' +
                                'order by sNumeroOrden') ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := param_global_contrato ;
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 Then
    begin
        While NOT Connection.qryBusca.Eof Do
        Begin
            tsNumeroOrden.Items.Add(Connection.qryBusca.FieldValues['sNumeroOrden']) ;
            Connection.qryBusca.Next;
        End;
    end;

    qryPlataforma.Active := False;
    qryPlataforma.Open;

    zq_tipomovimiento.Active := False;
    zq_tipomovimiento.Open;
        
    anexo_suministro.Active := False ;
    anexo_suministro.Params.ParamByName('Contrato').DataType := ftString ;
    anexo_suministro.Params.ParamByName('Contrato').Value    := param_global_contrato;
    anexo_suministro.Open;
    
    numeroOrden := anexo_suministro.FieldByName('sNumeroOrden').AsString;

    if anexo_suministro.RecordCount > 0 then
    begin
        pSalidas.Active := False;
        pSalidas.ParamByName('Contrato').DataType := ftString;
        pSalidas.ParamByName('Contrato').Value    := global_contrato_barco;
        pSalidas.ParamByName('Orden').DataType    := ftString;
        pSalidas.ParamByName('Orden').Value       := param_global_contrato;
        pSalidas.ParamByName('Folio').DataType    := ftInteger;
        pSalidas.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioSalida'];
        pSalidas.Open;
    end;

    ActivaBotones2(False);

    if connection.configuracion.FieldValues['sExplosion'] = 'Recursos por Concepto/Partida' then
       TipoExplosion := 'recursosanexo'
    else
       TipoExplosion := 'recursosanexosnuevos';
    BotonPermiso.permisosBotones(frmBarra1);
    BotonPermiso.permisosBotones(frmBarra2);

    PgControl.ActivePageIndex := 1;
    //Grid_Pedido.Columns[1].Width := 64;
end;

procedure TfrmSalidaAlmacen.BtnExitClick(Sender: TObject);
begin
    Close ;
end;

procedure TfrmSalidaAlmacen.frmBarra1btnExitClick(Sender: TObject);
begin
  Insertar1.Enabled := True ;
  Editar1.Enabled := True ;
  Registrar1.Enabled := False ;
  Can1.Enabled := False ;
  Eliminar1.Enabled := True ;
  Refresh1.Enabled := True ;
  frmBarra1.btnExitClick(Sender);
end;

procedure TfrmSalidaAlmacen.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure TfrmSalidaAlmacen.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click
end;

procedure TfrmSalidaAlmacen.EditarClick(Sender: TObject);
begin
     If pSalidas.RecordCount > 0 Then
     Begin
         OpcButton         := 'Edit';
         Agregar.Enabled   := False ;
         Editar.Enabled    := False ;
         Salvar.Enabled    := True ;
         Cancelar.Enabled  := True ;
         Eliminar.Enabled  := False ;
         ActivaBotones2(true);
         tdCantidad.ReadOnly     := False;
         tsTrazabilidad.ReadOnly := False;
         tdCantidad.SetFocus;
         if Length(trim(tdCantidad.Text)) > 0 then
         begin
             ValorPrev := pSalidas.FieldByName('dcantidad').AsFloat;
             tdCantidad.Value := ValorPrev;
         end
         else
           ValorPrev := 0;
         tdCantidad.Enabled := True;
         PgControl.Pages[0].TabVisible := False;
         GBSuperior.Enabled := False;
     End;
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmSalidaAlmacen.Registrar1Click(Sender: TObject);
begin
    frmBarra2.btnPost.Click 
end;

procedure TfrmSalidaAlmacen.Can1Click(Sender: TObject);
begin
    frmBarra2.btnCancel.Click 
end;

procedure TfrmSalidaAlmacen.CancelarClick(Sender: TObject);
begin
    Agregar.Enabled  := True ;
    Editar.Enabled   := True;
    Salvar.Enabled   := False ;
    Cancelar.Enabled := False ;
    Eliminar.Enabled := True ;

    ActivaBotones2(False);
    GridPartidas.Enabled := True;
    BotonPermiso.permisosBotones(frmBarra1);
    tdCantidad.Enabled := False;
    tsTrazabilidad.ReadOnly := True;
    tdCantidad.ReadOnly     := True;
    PgControl.Pages[0].TabVisible := True;
    GBSuperior.Enabled := True;
end;

procedure TfrmSalidaAlmacen.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TfrmSalidaAlmacen.EliminarClick(Sender: TObject);
var
   sAlmacen : string;
begin
     If pSalidas.RecordCount > 0 Then
     Begin
         {Buscamos el ALAMCEN}
         connection.QryBusca2.Active := False;
         connection.QryBusca2.SQL.Clear;
         connection.QryBusca2.SQL.Add('select sIdAlmacen from almacenes');
         connection.QryBusca2.Open;

         if connection.QryBusca2.RecordCount > 0 then
            sAlmacen := connection.QryBusca2.FieldValues['sIdAlmacen'];

          connection.zCommand.Connection.StartTransaction;
          try
             connection.zCommand.Active := False ;
             connection.zCommand.SQL.Clear ;
             connection.zCommand.SQL.Add ('Delete from bitacoradesalida where sContrato = :Contrato ' +
                                          'and iFolioSalida =:Folio and iFolioAnio =:FolioA and sIdInsumo =:Insumo and sNumeroActividad =:Actividad and sIdAlmacen =:Almacen and dFechaSalida =:Fecha and iitem = :item ') ;
             connection.zcommand.Params.ParamByName('Contrato').DataType  := ftString ;
             connection.zcommand.Params.ParamByName('Contrato').value     := param_global_contrato ;
             connection.zcommand.Params.ParamByName('Folio').DataType     := ftInteger ;
             connection.zcommand.Params.ParamByName('Folio').value        := anexo_suministro.FieldValues['iFolioSalida'] ;
             connection.zcommand.Params.ParamByName('FolioA').DataType    := ftInteger ;
             connection.zcommand.Params.ParamByName('FolioA').value       := anexo_suministro.FieldValues['iFolioAnio'] ;
             connection.zcommand.Params.ParamByName('Insumo').DataType    := ftString ;
             connection.zcommand.Params.ParamByName('Insumo').value       := IdInsumo ;
             connection.zcommand.Params.ParamByName('Almacen').DataType   := ftString;
             connection.zcommand.Params.ParamByName('Almacen').Value      := sAlmacen;
             connection.zcommand.Params.ParamByName('Fecha').DataType     := ftDate ;
             connection.zcommand.Params.ParamByName('Fecha').value        := anexo_suministro.FieldValues['dFechaSalida'] ;
             connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
             connection.zCommand.Params.ParamByName('Actividad').Value    := pSalidas.FieldValues['sNumeroActividad'];
             connection.zCommand.Params.ParamByName('item').DataType := ftInteger;
             connection.zCommand.Params.ParamByName('item').Value    := pSalidas.FieldValues['iitem'];
             connection.zCommand.ExecSQL;

              //Actualiza consulta de las existencias...
              connection.zCommand.Active := False;
              connection.zCommand.SQL.Clear;
              connection.zCommand.SQL.Add('UPDATE anexo_psuministro SET dCantidadRestante = dCantidadRestante + :Cantidad ' +
              'WHERE sContrato = :Contrato And iFolio = :Folio And sIdInsumo = :Insumo and sTrazabilidad =:Trazabilidad ');
              connection.zCommand.Params.ParamByName('Contrato').DataType     := ftString;
              connection.zCommand.Params.ParamByName('Contrato').value        := param_Global_Contrato;
              connection.zCommand.Params.ParamByName('Folio').DataType        := ftInteger;
              connection.zCommand.Params.ParamByName('Folio').value           := pSalidas.FieldValues['iFolioAviso'];
              connection.zCommand.Params.ParamByName('Insumo').DataType       := ftString;
              connection.zCommand.Params.ParamByName('Insumo').value          := pSalidas.FieldValues['sIdInsumo'];
              connection.zCommand.Params.ParamByName('Cantidad').DataType     := ftFloat;
              connection.zCommand.Params.ParamByName('Cantidad').value        := pSalidas.FieldValues['dCantidad'];
              connection.zCommand.Params.ParamByName('Trazabilidad').DataType := ftString;
              connection.zCommand.Params.ParamByName('Trazabilidad').value    := pSalidas.FieldValues['sTrazabilidad'];
              connection.zCommand.ExecSQL;

              Pedido.Refresh;

              SavePlace := pSalidas.GetBookmark ;
              pSalidas.Refresh ;
              connection.zCommand.Connection.Commit;
              Try
                 pSalidas.GotoBookmark(SavePlace);
              except
                on e:exception do
                  pSalidas.FreeBookmark(SavePlace);
              End;

          Except
               connection.zCommand.Connection.Rollback;
               MessageDlg('Ocurrio un error al eliminar el registro.', mtInformation, [mbOk], 0);
          End
     End
     else
       tdCantidad.Text := '0';
end;

procedure TfrmSalidaAlmacen.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmSalidaAlmacen.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmSalidaAlmacen.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

procedure TfrmSalidaAlmacen.SalvarClick(Sender: TObject);
var
  iItem  : integer;
  ValAct : Real;
  sAlmacen : string;
begin
   {Buscamos el ALAMCEN}
   connection.QryBusca2.Active := False;
   connection.QryBusca2.SQL.Clear;
   connection.QryBusca2.SQL.Add('select sIdAlmacen from almacenes');
   connection.QryBusca2.Open;

   if connection.QryBusca2.RecordCount > 0 then
      sAlmacen := connection.QryBusca2.FieldValues['sIdAlmacen'];
   connection.zCommand.Connection.StartTransaction;
  try

     //Salida de materiales..
    If OpcButton = 'New' then
    Begin
     ValorPrev := 0;
      // Consulta de Insumo antes de insertar.

      if pedido.FieldValues['dCantidadRestante'] < tdCantidad.Value then
      begin
        showmessage('No se puede proporcionar la Cantidad Solicitada, verificar Existencias !');
        exit;
      end;

      connection.QryBusca2.Active := False ;
      connection.QryBusca2.SQL.Clear ;
      connection.QryBusca2.SQL.Add ('select sIdInsumo, dCantidad from bitacoradesalida where sContrato =:Contrato and iFolioSalida =:Folio and sIdInsumo =:Insumo '+
                                    'and sIdAlmacen =:Almacen and sNumeroActividad =:Actividad and sTrazabilidad =:trazabilidad and sExt = :Ext ') ;
      connection.QryBusca2.Params.ParamByName('Contrato').DataType  := ftString ;
      connection.QryBusca2.Params.ParamByName('Contrato').value     := param_global_contrato ;
      connection.QryBusca2.Params.ParamByName('Folio').DataType     := ftString ;
      connection.QryBusca2.Params.ParamByName('Folio').value        := anexo_suministro.FieldValues['iFolioSalida'] ;
      connection.QryBusca2.Params.ParamByName('Insumo').DataType    := ftString ;
      connection.QryBusca2.Params.ParamByName('Insumo').value       := tsInsumo.Text;
      connection.QryBusca2.Params.ParamByName('Almacen').DataType   := ftString;
      connection.QryBusca2.Params.ParamByName('Almacen').Value      := sAlmacen ;
      connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString;
      connection.QryBusca2.Params.ParamByName('Actividad').Value    := pSalidas.FieldValues['sNumeroActividad'];
      connection.QryBusca2.Params.ParamByName('trazabilidad').DataType := ftString;
      connection.QryBusca2.Params.ParamByName('trazabilidad').Value    := tsTrazabilidad.Text;
      connection.QryBusca2.Params.ParamByName('ext').AsString := anexo_suministro.FieldByName('sext').AsString;
      connection.QryBusca2.Open ;

      try
        connection.QryBusca.Active := False ;
        connection.QryBusca.SQL.Clear ;
        connection.QryBusca.SQL.Add ('select max(iItem) as maximo from bitacoradesalida where sContrato =:Contrato and iFolioSalida =:Folio and sIdInsumo =:Insumo '+
                                    'and sIdAlmacen =:Almacen and sNumeroActividad =:Actividad and sTrazabilidad =:trazabilidad ') ;
        connection.QryBusca.Params.ParamByName('Contrato').DataType  := ftString ;
        connection.QryBusca.Params.ParamByName('Contrato').value     := param_global_contrato ;
        connection.QryBusca.Params.ParamByName('Folio').DataType     := ftString ;
        connection.QryBusca.Params.ParamByName('Folio').value        := anexo_suministro.FieldValues['iFolioSalida'] ;
        connection.QryBusca.Params.ParamByName('Insumo').DataType    := ftString ;
        connection.QryBusca.Params.ParamByName('Insumo').value       := tsInsumo.Text;
        connection.QryBusca.Params.ParamByName('Almacen').DataType   := ftString;
        connection.QryBusca.Params.ParamByName('Almacen').Value      := sAlmacen ;
        connection.QryBusca.Params.ParamByName('Actividad').DataType := ftString;
        connection.QryBusca.Params.ParamByName('Actividad').Value    := pSalidas.FieldValues['sNumeroActividad'];
        connection.QryBusca.Params.ParamByName('trazabilidad').DataType := ftString;
        connection.QryBusca.Params.ParamByName('trazabilidad').Value    := tsTrazabilidad.Text;
        connection.QryBusca.Open ;
        if Length(trim(connection.QryBusca.FieldByName('maximo').AsString)) > 0 then
           iitem :=  connection.QryBusca.FieldByName('maximo').AsInteger + 1
        else
           iitem := 0;

      except
        on e:Exception do
          iitem := 0;
      end;

      //if connection.QryBusca2.RecordCount = 0 then
      begin
        // soad -> Inbsercion de los datos en la bitacora de Entrada....
        //****************************************************************
        connection.zCommand.Active := False ;
        connection.zCommand.SQL.Clear ;
        connection.zCommand.SQL.Add ( 'INSERT INTO bitacoradesalida ( sContrato, iFolioSalida,sExt, iFolioAnio, dFechaSalida, swbs, sIdInsumo, dCantidad, sIdUsuario, sIdAlmacen, sNumeroOrden, sNumeroActividad, sIdPlataforma, sTrazabilidad, iitem, iFolioAviso ) ' +
                                      'VALUES (:Contrato, :Folio,:Ext, :FolioAnio, :FechaS, "" ,:Insumo, :Cantidad, :Usuario, :Almacen, :Orden, :actividad, :plataforma, :trazabilidad, :item, :FolioAviso)') ;
        connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
        connection.zCommand.Params.ParamByName('Contrato').value    := param_global_contrato ;
        connection.zCommand.Params.ParamByName('Folio').DataType    := ftString ;
        connection.zCommand.Params.ParamByName('Folio').value       := anexo_suministro.FieldValues['iFolioSalida'] ;
        connection.zCommand.Params.ParamByName('FechaS').DataType   := ftDate ;
        connection.zCommand.Params.ParamByName('FechaS').value      := anexo_suministro.FieldValues['dFechaSalida'];
        connection.zCommand.Params.ParamByName('Insumo').DataType   := ftString ;
        connection.zCommand.Params.ParamByName('Insumo').value      := pedido.FieldValues['sIdInsumo'];
        connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat ;
        connection.zCommand.Params.ParamByName('Cantidad').value    := tdCantidad.Value ;
        connection.zCommand.Params.ParamByName('Usuario').DataType  := ftString ;
        connection.zCommand.Params.ParamByName('Usuario').value     := global_usuario;
        connection.zCommand.Params.ParamByName('Almacen').DataType  := ftString;
        connection.zCommand.Params.ParamByName('Almacen').Value     := sAlmacen ;
        connection.zCommand.Params.ParamByName('Orden').DataType      := ftString ;
        connection.zCommand.Params.ParamByName('Orden').value         := anexo_suministro.FieldValues['sNumeroOrden'];
        connection.zCommand.Params.ParamByName('Actividad').DataType  := ftString ;
        connection.zCommand.Params.ParamByName('Actividad').value     := anexo_suministro.FieldValues['sNumeroActividad'];
        connection.zCommand.Params.ParamByName('Plataforma').DataType := ftString ;
        connection.zCommand.Params.ParamByName('Plataforma').value    := anexo_suministro.FieldValues['sIdPlataforma'];
        connection.zCommand.Params.ParamByName('trazabilidad').DataType := ftString ;
        connection.zCommand.Params.ParamByName('trazabilidad').value    := tsTrazabilidad.Text;
        connection.zCommand.Params.ParamByName('item').DataType         := ftInteger;
        connection.zCommand.Params.ParamByName('item').value            := iitem;
        connection.zCommand.Params.ParamByName('FolioAnio').DataType    := ftInteger;
        connection.zCommand.Params.ParamByName('FolioAnio').value       := anexo_suministro.FieldValues['iFolioAnio'];
        connection.zCommand.Params.ParamByName('FolioAviso').DataType   := ftInteger;
        connection.zCommand.Params.ParamByName('FolioAviso').value      := pedido.FieldValues['iFolio'];
        connection.zCommand.Params.ParamByName('Ext').AsString          := anexo_suministro.FieldByName('sExt').AsString;
        connection.zCommand.ExecSQL;
      end;
    End;

    If OpcButton = 'Edit' then
    Begin
      // Consulta de Insumo antes de insertar.     
      if pedido.FieldValues['dCantidadRestante'] < tdCantidad.Value then
      begin
        showmessage('No se puede proporcionar la Cantidad Solicitada, verificar Existencias !');
        exit;
      end;
      try
        // soad -> Edita de los datos en la bitacora de Entrada....
        //****************************************************************
        connection.zCommand.Active := False ;
        connection.zCommand.SQL.Clear ;
        connection.zCommand.SQL.Add ('UPDATE bitacoradesalida SET dCantidad =:Cantidad where sContrato =:Contrato and iFolioSalida =:Folio and iFolioAnio =:FolioAnio '+
                                     'and sIdInsumo =:Insumo and sExt= :Ext and sIdAlmacen =:Almacen and sNumeroActividad =:Actividad and sTrazabilidad = :Trazabilidad and iitem = :Item') ;
        connection.zCommand.Params.ParamByName('Contrato').DataType  := ftString ;
        connection.zCommand.Params.ParamByName('Contrato').value     := param_global_contrato ;
        connection.zCommand.Params.ParamByName('Folio').DataType     := ftString ;
        connection.zCommand.Params.ParamByName('Folio').value        := anexo_suministro.FieldValues['iFolioSalida'] ;
        connection.zCommand.Params.ParamByName('Ext').AsString       := anexo_suministro.FieldByName('sext').AsString;
        connection.zCommand.Params.ParamByName('Insumo').DataType    := ftString ;
        connection.zCommand.Params.ParamByName('Insumo').value       := tsInsumo.Text;
        connection.zCommand.Params.ParamByName('Cantidad').DataType  := ftFloat ;
        connection.zCommand.Params.ParamByName('Cantidad').value     := tdCantidad.Value ;
        connection.zCommand.Params.ParamByName('Almacen').DataType   := ftString;
        connection.zCommand.Params.ParamByName('Almacen').Value      := sAlmacen ;
        connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
        connection.zCommand.Params.ParamByName('Actividad').Value    := pSalidas.FieldValues['sNumeroActividad'];
        connection.zCommand.Params.ParamByName('trazabilidad').DataType := ftString ;
        connection.zCommand.Params.ParamByName('trazabilidad').value    := pSalidas.FieldValues['sTrazabilidad'];
        connection.zCommand.Params.ParamByName('Item').DataType      := ftInteger ;
        connection.zCommand.Params.ParamByName('Item').value         := pSalidas.FieldValues['iItem'];
        connection.zCommand.Params.ParamByName('FolioAnio').DataType := ftInteger;
        connection.zCommand.Params.ParamByName('FolioAnio').value    := pSalidas.FieldValues['iFolioAnio'];
        connection.zCommand.ExecSQL ;
      except
        begin
          MessageDlg('Ocurrio un error al actualizar el registro', mtWarning, [mbOk], 0);
          exit;
        end;
      End;

      Agregar.Enabled  := True ;
      Editar.Enabled   := True ;
      Salvar.Enabled   := False ;
      Cancelar.Enabled := False ;
      Eliminar.Enabled := True ;
      tdCantidad.Enabled := False;
      GBSuperior.Enabled := True;
      PgControl.Pages[0].TabVisible := True;
    End;

    ValAct := tdCantidad.Value;

    If OpcButton = 'New' then
       ValAct := pedido.FieldByName('dCantidadRestante').AsFloat - ValAct
    else
    begin
       {Buscamos el Aviso d emebarque de donde se extrae la Cantidad..}
       connection.QryBusca.Active := False;
       connection.QryBusca.SQL.Clear;
       connection.QryBusca.SQL.Add('select dCantidadRestante from anexo_psuministro ' +
                 'WHERE sContrato = :Contrato And iFolio = :Folio And sIdInsumo = :Insumo and sTrazabilidad =:Trazabilidad ');
       connection.QryBusca.Params.ParamByName('Contrato').DataType     := ftString;
       connection.QryBusca.Params.ParamByName('Contrato').value        := param_Global_Contrato;
       connection.QryBusca.Params.ParamByName('Folio').DataType        := ftInteger;
       connection.QryBusca.Params.ParamByName('Folio').value           := pSalidas.FieldValues['iFolioAviso'];
       connection.QryBusca.Params.ParamByName('Insumo').DataType       := ftString;
       connection.QryBusca.Params.ParamByName('Insumo').value          := pSalidas.FieldValues['sIdInsumo'];
       connection.QryBusca.Params.ParamByName('Trazabilidad').DataType := ftString;
       connection.QryBusca.Params.ParamByName('Trazabilidad').value    := pSalidas.FieldValues['sTrazabilidad'];
       connection.QryBusca.Open;

       ValAct := (connection.QryBusca.FieldByName('dCantidadRestante').AsFloat + ValorPrev) - ValAct;
    end;

    //Actualiza consulta de las existencias...
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('UPDATE anexo_psuministro SET dCantidadRestante = :Cantidad ' +
      'WHERE sContrato = :Contrato And iFolio = :Folio And sIdInsumo = :Insumo and sTrazabilidad =:Trazabilidad ');
    connection.zCommand.Params.ParamByName('Contrato').DataType     := ftString;
    connection.zCommand.Params.ParamByName('Contrato').value        := param_Global_Contrato;
    connection.zCommand.Params.ParamByName('Folio').DataType        := ftInteger;
    If OpcButton = 'New' then
       connection.zCommand.Params.ParamByName('Folio').value        := pedido.FieldValues['iFolio']
    else
       connection.zCommand.Params.ParamByName('Folio').value        := pSalidas.FieldValues['iFolioAviso'];
    connection.zCommand.Params.ParamByName('Insumo').DataType       := ftString;
    If OpcButton = 'New' then
       connection.zCommand.Params.ParamByName('Insumo').value       := Pedido.FieldValues['sIdInsumo']
    else
       connection.zCommand.Params.ParamByName('Insumo').value       := pSalidas.FieldValues['sIdInsumo'];
    connection.zCommand.Params.ParamByName('Cantidad').DataType     := ftFloat;
    connection.zCommand.Params.ParamByName('Cantidad').value        := ValAct;
    connection.zCommand.Params.ParamByName('Trazabilidad').DataType := ftString;
    If OpcButton = 'New' then
       connection.zCommand.Params.ParamByName('Trazabilidad').value := trim(tsTrazabilidad.Text)
    else
       connection.zCommand.Params.ParamByName('Trazabilidad').value := pSalidas.FieldValues['sTrazabilidad'];
    connection.zCommand.ExecSQL;
    connection.zCommand.Connection.Commit;
  except
    on e:Exception do
      connection.zCommand.Connection.Rollback;
  end;

  SavePlace2 := Pedido.GetBookmark ;
  Pedido.Refresh;

  SavePlace := pSalidas.GetBookmark ;
  pSalidas.Refresh ;

  Try
    pSalidas.GotoBookmark(SavePlace);
  Except
  Else
    pSalidas.FreeBookmark(SavePlace);
  End;

  Try
    Pedido.GotoBookmark(SavePlace2);
  Except
  Else
    Pedido.FreeBookmark(SavePlace2);
  End;

  tsInsumo.SetFocus;
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmSalidaAlmacen.SpeedButton1Click(Sender: TObject);
begin
    Application.CreateForm(TfrmEntradaAnex, frmEntradaAnex);
    frmEntradaAnex.Show;
end;

procedure TfrmSalidaAlmacen.tsIsometricoReferenciaKeyPress(
  Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tmComentarios.SetFocus
end;


procedure TfrmSalidaAlmacen.GridPartidasTitleClick(Column: TColumn);
begin
   UtGrid3.DbGridTitleClick(Column);
end;

procedure TfrmSalidaAlmacen.frxReport50GetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'ANEXO') = 0 then
  Begin
      Connection.qryBusca.Active := False ;
      Connection.qryBusca.SQL.Clear ;
      Connection.qryBusca.SQL.Add('Select sAnexo From convenios Where sContrato = :Contrato And sIdConvenio = :convenio') ;
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := param_global_contrato ;
      Connection.qryBusca.Params.ParamByName('convenio').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('convenio').Value := global_convenio ;
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


procedure TfrmSalidaAlmacen.frmBarra2btnAddClick(Sender: TObject);
Var
  dFechaFinal : tDate ;
  iCheck      : Integer ;
  Maximo      : Integer;
begin
  //activapop(frmSalidaAlmacen,popupprincipal);
  Try
    OpcButton1 := 'New' ;
    OldFolio := -1;
    OldExt := '';
    frmBarra2.btnAddClick(Sender);
    frmBarra1.btnCancel.Click ;
    pgControl.ActivePageIndex := 0 ;
    PgControl.Pages[1].TabVisible := False;

    //BUSCAMOS SI EXISTE EL MATERIAL..
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select max(iFolioSalida) as Folio FROM almacen_salida ');
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
       Maximo := Connection.zCommand.FieldByName('Folio').AsInteger + 1
    else
       Maximo := 1;

    iFolio.Value := Maximo;

    ActivaBotones(True);
    tdIdFecha.Date     := global_fecha ;
    tmComentarios.Text := '' ;
    txtNombre.text     := '';
    tsNumeroOrden.ItemIndex := 0 ;
    tdIdFecha.Enabled := True;
    tdIdFecha.SetFocus;
    //anexo_suministro.Append;
    BotonPermiso.permisosBotones(frmBarra2);
    Grid_Entradas.Enabled := False;
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_EntradaAlmacen', 'Al agregar registro ', 0);
    end;
  end;
end;

procedure TfrmSalidaAlmacen.frmBarra2btnEditClick(Sender: TObject);
begin
    //activapop(frmSalidaAlmacen, popupprincipal);
    If anexo_suministro.RecordCount > 0 then
    Begin
         OldFolio := anexo_suministro.FieldByName('ifoliosalida').AsInteger;
         OldExt := anexo_suministro.FieldByName('sext').AsString;
         OpcButton1 := 'Edit' ;
         FolioAnt   := iFolio.Value;
         anexo_suministro.Edit;
         ActivaBotones(True);
         frmBarra2.btnEditClick(Sender);
         pgControl.ActivePageIndex := 0 ;
         PgControl.Pages[1].TabVisible := False;
    End;
    BotonPermiso.permisosBotones(frmBarra2);
    Grid_Entradas.Enabled := False;
end;

procedure TfrmSalidaAlmacen.frmBarra2btnPostClick(Sender: TObject);
begin
  //desactivapop(popupprincipal);
  if (iFolio.Value <> OldFolio) or (Trim(EdtExt.Text) <> OldExt) then
  begin
    connection.zCommand.Active := False ;
    connection.zCommand.SQL.Clear ;
    connection.zCommand.SQL.Add ('select * from almacen_salida where iFolioSalida =:Folio and sext = :ext and iFolioAnio =:Anio ');
    connection.zCommand.params.ParamByName('Folio').DataType    := ftInteger ;
    connection.zCommand.params.ParamByName('Folio').value       := iFolio.Value ;
    connection.zCommand.params.ParamByName('Anio').DataType     := ftInteger ;
    connection.zCommand.params.ParamByName('Anio').value        := yearOf(tdIdFecha.Date);
    connection.zCommand.params.ParamByName('ext').AsString      := Trim(EdtExt.Text);
    connection.zCommand.Open ;

    if connection.zCommand.RecordCount > 0  then
    begin
        MessageDlg('El Folio de Salida ya existe!', mtWarning, [mbOk],0) ;
        iFolio.SetFocus;
        exit;
    end;
  end;

  if Length(trim(iFolio.Text)) = 0 then
  begin
    MessageDlg('Es necesario establecer un numero de folio.', mtWarning, [mbOk],0) ;
    iFolio.SetFocus;
    exit;
  end;

  If OpcButton1 = 'New' then
  Begin
    try
      connection.zCommand.Active := False ;
      connection.zCommand.SQL.Clear ;
      connection.zCommand.SQL.Add ( 'INSERT INTO almacen_salida ( sContrato, iFolioSalida,sExt, dFechaSalida, sIdTipo, sNombre, sNumeroOrden, sNumeroActividad, sIdPlataforma, sIdUsuario, mComentarios, iFolioAnio ) ' +
                                    'VALUES (:Contrato, :Folio,:Ext, :FechaS, :Tipo, :Nombre, :Orden, :Actividad, :Plataforma, :Usuario, :Comentarios, :Anio )') ;
      connection.zCommand.params.ParamByName('Contrato').DataType    := ftString ;
      connection.zCommand.params.ParamByName('Contrato').value       := param_global_contrato ;
      connection.zCommand.params.ParamByName('Folio').DataType       := ftInteger ;
      connection.zCommand.params.ParamByName('Folio').value          := iFolio.Value ;
      connection.zCommand.params.ParamByName('FechaS').DataType      := ftDate ;
      connection.zCommand.params.ParamByName('FechaS').value         := tdIdFecha.Date ;
      connection.zCommand.params.ParamByName('Tipo').DataType        := ftString ;
      connection.zCommand.params.ParamByName('Tipo').value           := tsTipoMovimiento.KeyValue ;
      connection.zCommand.params.ParamByName('Nombre').DataType      := ftString ;
      connection.zCommand.params.ParamByName('Nombre').value         := txtNombre.Text ;
      connection.zCommand.params.ParamByName('Orden').DataType       := ftString ;
      connection.zCommand.params.ParamByName('Orden').value          := tsNumeroOrden.Text ;
      connection.zCommand.params.ParamByName('Actividad').DataType   := ftString ;
      connection.zCommand.params.ParamByName('Actividad').value      := tsPartida.KeyValue ;
      connection.zCommand.params.ParamByName('Plataforma').DataType  := ftString ;
      connection.zCommand.params.ParamByName('Plataforma').value     := tsPlataformaRef.KeyValue ;
      connection.zCommand.params.ParamByName('Usuario').DataType     := ftString ;
      connection.zCommand.params.ParamByName('Usuario').value        := Global_Usuario ;
      connection.zCommand.params.ParamByName('Comentarios').DataType := ftMemo ;
      connection.zCommand.params.ParamByName('Comentarios').value    := tmCOmentarios.Text ;
      connection.zCommand.params.ParamByName('Anio').DataType        := ftInteger ;
      connection.zCommand.params.ParamByName('Anio').value           := yearOf(tdIdFecha.Date);
      connection.zCommand.params.ParamByName('Ext').AsString := Trim(EdtExt.text);
      connection.zCommand.ExecSQL ;

      // Actualizo Kardex del Sistema ....
      connection.zCommand.Active := False ;
      connection.zCommand.SQL.Clear ;
      connection.zCommand.SQL.Add ( 'Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                    'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
      connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
      connection.zCommand.Params.ParamByName('Contrato').Value       := param_global_contrato ;
      connection.zCommand.Params.ParamByName('Usuario').DataType     := ftString ;
      connection.zCommand.Params.ParamByName('Usuario').Value        := Global_Usuario ;
      connection.zCommand.Params.ParamByName('Fecha').DataType       := ftDate ;
      connection.zCommand.Params.ParamByName('Fecha').Value          := Date ;
      connection.zCommand.Params.ParamByName('Hora').DataType        := ftString ;
      connection.zCommand.Params.ParamByName('Hora').value           := FormatDateTime('hh:mm:ss', Now) ;
      connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString ;
      connection.zCommand.Params.ParamByName('Descripcion').Value    := 'Registro de Salida al Almacen No. ' + FloatToStr(iFolio.Value) + ' El dia  ['+ DateToStr(Date())+ ']  Usuario [ ' + global_usuario + ']' ;
      connection.zCommand.Params.ParamByName('Origen').DataType      := ftString ;
      connection.zCommand.Params.ParamByName('Origen').Value         := 'Reporte Diario' ;
      connection.zCommand.ExecSQL ;
      ActivaBotones(False);
      frmBarra2.btnCancelClick(Sender);
     Except
       // MessageDlg('Ocurrio un error al actualizar el registro', mtWarning, [mbOk], 0);
     on e : exception do
     begin
       UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Salidas Almacen', 'Al salvar registro', 0);
     end;
    End
  End
  Else
  If OpcButton1 = 'Edit' then
  Begin
  try
    connection.zCommand.Connection.StartTransaction;
    if (iFolio.Value <> OldFolio) or (Trim(EdtExt.Text) <> OldExt) then
    begin
      //aqui debe editarse el id de material
      connection.zCommand.Active := False ;
      connection.zCommand.SQL.Clear ;
      connection.zCommand.SQL.Add ('UPDATE bitacoradesalida SET sExt= :Ext, iFolioSalida =:Folio  where sContrato =:Contrato and iFolioSalida =:FolioOld  '+
                                   'and sExt= :ExtOld ') ;
      connection.zCommand.Params.ParamByName('Folio').asinteger    :=   strtoint(iFolio.text) ;
      connection.zCommand.Params.ParamByName('Ext').AsString       := trim(EdtExt.Text);
      connection.zCommand.Params.ParamByName('Contrato').AsString     := param_global_contrato;
      connection.zCommand.Params.ParamByName('FolioOld').asinteger    := OldFolio;
      connection.zCommand.Params.ParamByName('ExtOld').AsString       := OldExt;
      connection.zCommand.ExecSQL ;
    end;

    connection.zCommand.Active := False ;
    connection.zCommand.SQL.Clear ;
    connection.zCommand.SQL.Add ( 'UPDATE almacen_salida SET  iFolioSalida =:Folio, sNumeroOrden = :Orden, sNumeroActividad =:Actividad, sIdPlataforma =:Plataforma, '+
                                  'sNombre =:Nombre,sExt = :Ext, sIdTipo =:Tipo, mComentarios = :Comentarios ' +
                                  'WHERE sContrato =:Contrato and iFolioSalida =:FolioAnt and ifolioAnio =:FolioAnio ') ;
    connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
    connection.zCommand.Params.ParamByName('Contrato').value       := param_global_contrato ;
    connection.zCommand.Params.ParamByName('Folio').DataType       := ftInteger ;
    connection.zCommand.Params.ParamByName('Folio').value          := iFolio.Value ;
    connection.zCommand.Params.ParamByName('FolioAnt').DataType    := ftFloat ;
    connection.zCommand.Params.ParamByName('FolioAnt').value       := FolioAnt;
    connection.zCommand.Params.ParamByName('Orden').DataType       := ftString ;
    connection.zCommand.Params.ParamByName('Orden').value          := tsNumeroOrden.Text ;
    connection.zCommand.params.ParamByName('Actividad').DataType   := ftString ;
    connection.zCommand.params.ParamByName('Actividad').value      := tsPartida.KeyValue ;
    connection.zCommand.params.ParamByName('Plataforma').DataType  := ftString ;
    connection.zCommand.params.ParamByName('Plataforma').value     := tsPlataformaRef.KeyValue ;
    connection.zCommand.Params.ParamByName('Tipo').DataType        := ftString ;
    connection.zCommand.Params.ParamByName('Tipo').value           := tsTipomovimiento.KeyValue;
    connection.zCommand.Params.ParamByName('Nombre').DataType      := ftString ;
    connection.zCommand.Params.ParamByName('Nombre').value         := txtNombre.Text ;
    connection.zCommand.Params.ParamByName('Comentarios').DataType := ftMemo ;
    connection.zCommand.Params.ParamByName('Comentarios').value    := tmCOmentarios.Text ;
    connection.zCommand.Params.ParamByName('FolioAnio').DataType   := ftInteger ;
    connection.zCommand.Params.ParamByName('FolioAnio').value      := anexo_suministro.FieldValues['iFolioAnio'];
    connection.zCommand.Params.ParamByName('Ext').asstring         := trim(edtExt.text);
    connection.zCommand.ExecSQL ;

    // Actualizo Kardex del Sistema ....
    connection.zCommand.Active := False ;
    connection.zCommand.SQL.Clear ;
    connection.zCommand.SQL.Add ( 'Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                  'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
    connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
    connection.zCommand.Params.ParamByName('Contrato').Value       := param_global_contrato ;
    connection.zCommand.Params.ParamByName('Usuario').DataType     := ftString ;
    connection.zCommand.Params.ParamByName('Usuario').Value        := Global_Usuario ;
    connection.zCommand.Params.ParamByName('Fecha').DataType       := ftDate ;
    connection.zCommand.Params.ParamByName('Fecha').Value          := Date ;
    connection.zCommand.Params.ParamByName('Hora').DataType        := ftString ;
    connection.zCommand.Params.ParamByName('Hora').value           := FormatDateTime('hh:mm:ss', Now) ;
    connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString ;
    connection.zCommand.Params.ParamByName('Descripcion').Value    := 'Modificación de Salida al Almacen No. ' + IntToStr(anexo_suministro.FieldValues['iFolioSalida']) + ' El día ['+ DateToStr(Date())+ '] Usuario [ ' + global_usuario + ']' ;
    connection.zCommand.Params.ParamByName('Origen').DataType      := ftString ;
    connection.zCommand.Params.ParamByName('Origen').Value         := 'Reporte Diario' ;
    connection.zCommand.ExecSQL ;
    ActivaBotones(False);
    frmBarra2.btnCancelClick(Sender);
    PgControl.Pages[1].TabVisible := True;
    connection.zCommand.Connection.Commit;
  except

    //  MessageDlg('Ocurrio un error al actualizar el registro', mtWarning, [mbOk], 0);
    on e : exception do
    begin
      if connection.zCommand.Connection.InTransaction then
        connection.zCommand.Connection.Rollback;
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Salidas Almacen', 'Al salvar registro', 0);
    end;
  End;
  End ;

  anexo_suministro.Cancel;

  SavePlace := anexo_suministro.GetBookmark ;
  Try
    anexo_suministro.Refresh ;
    anexo_suministro.GotoBookmark(SavePlace);
  Except
  Else
    anexo_suministro.FreeBookmark(SavePlace);
  End;

  OpcButton1 := '' ;
  BotonPermiso.permisosBotones(frmBarra2);
  Grid_Entradas.Enabled := True;
  PgControl.Pages[1].TabVisible := True;
end;

procedure TfrmSalidaAlmacen.frmBarra2btnPrinterClick(Sender: TObject);
begin
     If anexo_suministro.RecordCount > 0 Then
     begin
         Reporte.Active := False;
         Reporte.SQL.Clear;
         Reporte.SQL.Add('select s.*, concat(b.iFolioSalida,b.sext) as sFolioSalida ,b.*, i.dCantidadAnexo as dExistencia, i.dCostoMN, i.mDescripcion, i.sMedida, m.sDescripcion as Tipomovimiento from almacen_salida s '+
                         'inner join bitacoradesalida b '+
                         'on(b.sContrato = s.sContrato and b.iFolioSalida = s.iFolioSalida and b.iFolioAnio = s.iFolioAnio) '+
                         'inner join actividadesxanexo i on (i.sContrato = :contrato and i.sNumeroActividad = b.sIdInsumo  and i.sWbs = b.sTrazabilidad) '+
                         'inner join movimientosdealmacen m '+
                         'on (m.sIdTipo = s.sIdTipo) '+
                         'where s.sContrato =:Orden and s.iFolioSalida =:Folio and b.iFolioAnio =:FolioAnio ');
         Reporte.ParamByName('Contrato').DataType  := ftString ;
         Reporte.ParamByName('Contrato').Value     := global_contrato_barco ;
         Reporte.ParamByName('Orden').DataType     := ftString ;
         Reporte.ParamByName('Orden').Value        := param_global_contrato ;
         Reporte.ParamByName('Folio').DataType     := ftInteger;
         Reporte.ParamByName('Folio').Value        := anexo_suministro.FieldValues['iFolioSalida'];
         Reporte.ParamByName('FolioAnio').DataType := ftInteger;
         Reporte.ParamByName('FolioAnio').Value    := anexo_suministro.FieldValues['iFolioAnio'];
         Reporte.Open;

         frxEntrada.PreviewOptions.MDIChild := False ;
         frxEntrada.PreviewOptions.Modal := True ;
         frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
         frxEntrada.PreviewOptions.ShowCaptions := False ;
         frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
         frxEntrada.LoadFromFile (global_files + global_miReporte+'_Salida.fr3') ;
         frxEntrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
     end;
end;

procedure TfrmSalidaAlmacen.frmBarra2btnDeleteClick(Sender: TObject);
begin
     If anexo_suministro.RecordCount > 0 Then
        If MessageDlg('Desea eliminar el Folio seleccionado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
           Begin
              if pSalidas.RecordCount > 0 then
              begin
                   showmessage('No se puede Eliminar!, Existen Materiales para esta Salida.');
                   exit;
              end;
              // Actualizo Kardex del Sistema ....
              try
              connection.zCommand.Active := False ;
              connection.zCommand.SQL.Clear ;
              connection.zCommand.SQL.Add ('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                           'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
              connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
              connection.zCommand.Params.ParamByName('Contrato').Value       := param_global_contrato ;
              connection.zCommand.Params.ParamByName('Usuario').DataType     := ftString ;
              connection.zCommand.Params.ParamByName('Usuario').Value        := Global_Usuario ;
              connection.zCommand.Params.ParamByName('Fecha').DataType       := ftDate ;
              connection.zCommand.Params.ParamByName('Fecha').Value          := Date ;
              connection.zCommand.Params.ParamByName('Hora').DataType        := ftString ;
              connection.zCommand.Params.ParamByName('Hora').value           := FormatDateTime('hh:mm:ss', Now) ;
              connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString ;
              connection.zCommand.Params.ParamByName('Descripcion').Value    := 'Eliminación de Salida al Almacen ' + IntToStr(anexo_suministro.FieldValues['iFolioSalida']) + ' El día [' + DateToStr(Date())+ '] Usuario [ ' + global_usuario + ']' ;
              connection.zCommand.Params.ParamByName('Origen').DataType      := ftString ;
              connection.zCommand.Params.ParamByName('Origen').Value         := 'Reporte Diario' ;
              connection.zCommand.ExecSQL ;

              connection.zCommand.Active := False ;
              connection.zCommand.SQL.Clear ;
              connection.zCommand.SQL.Add ( 'Delete from almacen_salida where sContrato =:Contrato and iFolioSalida =:Folio ') ;
              connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
              connection.zcommand.Params.ParamByName('Contrato').value    := param_global_contrato ;
              connection.zcommand.Params.ParamByName('Folio').DataType    := ftInteger ;
              connection.zcommand.Params.ParamByName('Folio').value       := anexo_suministro.FieldValues['iFolioSalida'] ;
              connection.zCommand.ExecSQL ;

              SavePlace := anexo_suministro.GetBookmark ;
              anexo_suministro.Refresh ;

              Try
                 anexo_suministro.GotoBookmark(SavePlace);
              Except
              Else
                anexo_suministro.FreeBookmark(SavePlace);
              End;

              except
               on e : exception do begin
                  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Salidas Almacen', 'Al eliminar registro', 0);
                end;
              end;
          End
end;


procedure TfrmSalidaAlmacen.frmBarra2btnRefreshClick(Sender: TObject);
begin
    anexo_suministro.Active := False ;
    anexo_suministro.Open ;

    Pedido.Active := False;
    Pedido.Params.ParamByName('Contrato').DataType := ftString;
    Pedido.Params.ParamByName('Contrato').Value := global_Contrato_Barco;
    Pedido.Params.ParamByName('Orden').DataType := ftString;
    Pedido.Params.ParamByName('Orden').Value    := param_global_contrato;  
    Pedido.Open;

    tsNumeroOrden.Items.Clear ;
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select sNumeroOrden from ordenesdetrabajo where sContrato = :Contrato and ' +
                                'cIdStatus = :status order by sNumeroOrden') ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := param_global_contrato ;
    Connection.qryBusca.Params.ParamByName('status').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('status').Value :=  connection.configuracion.FieldValues [ 'cStatusProceso' ];
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 Then
        While NOT Connection.qryBusca.Eof Do
        Begin
            tsNumeroOrden.Items.Add(Connection.qryBusca.FieldValues['sNumeroOrden']) ;
            Connection.qryBusca.Next
        End ;
    tsNumeroOrden.ItemIndex := 0 ;

    anexo_suministro.Active := False ;
    anexo_suministro.Params.ParamByName('Contrato').DataType := ftString ;
    anexo_suministro.Params.ParamByName('Contrato').Value    := param_global_contrato;
    anexo_suministro.Open ;

    if anexo_suministro.RecordCount > 0 then
    begin
        pSalidas.Active := False;
        pSalidas.ParamByName('Contrato').DataType := ftString;
        pSalidas.ParamByName('Contrato').Value    := global_contrato_barco;
        pSalidas.ParamByName('Orden').DataType    := ftString;
        pSalidas.ParamByName('Orden').Value       := param_global_contrato;
        pSalidas.ParamByName('Folio').DataType    := ftInteger;
        pSalidas.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioSalida'];
        pSalidas.Open;
    end;

    zq_tipomovimiento.Active := False;
    zq_tipomovimiento.Open;

    qryPlataforma.Active := False;
    qryPlataforma.Open;
    
end;

procedure TfrmSalidaAlmacen.frmBarra2btnCancelClick(Sender: TObject);
begin
  //desactivapop(popupprincipal);
  ActivaBotones(False);
  frmBarra2.btnCancelClick(Sender);
  //Grid_Entradas.SetFocus ;
  BotonPermiso.permisosBotones(frmBarra2);
  Grid_Entradas.Enabled := True;
  PgControl.Pages[1].TabVisible := True;
  lblEstadoFolio.Visible := False;
end;

procedure TfrmSalidaAlmacen.frmBarra2btnExitClick(Sender: TObject);
begin
  frmBarra2.btnExitClick(Sender);
  close
end;

procedure TfrmSalidaAlmacen.tdIdFechaEnter(Sender: TObject);
begin
    tdIdFecha.Color := global_color_entrada
end;

procedure TfrmSalidaAlmacen.tdIdFechaExit(Sender: TObject);
begin
    tdIdFecha.Color := global_color_salida
end;

procedure TfrmSalidaAlmacen.tdIdFechaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsTipomovimiento.SetFocus
end;

procedure TfrmSalidaAlmacen.tsOrigenKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tmComentarios.SetFocus
end;

procedure TfrmSalidaAlmacen.tsPartidaEnter(Sender: TObject);
begin
    tsPartida.Color := global_color_entrada
end;

procedure TfrmSalidaAlmacen.tsPartidaExit(Sender: TObject);
begin
    tsPartida.Color := global_color_salida
end;

procedure TfrmSalidaAlmacen.tsPartidaKeyPress(Sender: TObject; var Key: Char);
begin
    if Key = #13 then
       tsPlataformaRef.SetFocus;
end;

procedure TfrmSalidaAlmacen.tsPlataformaRefKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key = #13 then
       txtNombre.SetFocus;
end;

procedure TfrmSalidaAlmacen.tsTipomovimientoEnter(Sender: TObject);
begin
 tsTipomovimiento.Color := global_color_entrada
end;

procedure TfrmSalidaAlmacen.tsTipomovimientoxit(Sender: TObject);
begin
    tsTipomovimiento.Color := global_color_salida
end;

procedure TfrmSalidaAlmacen.tsTrazabilidadEnter(Sender: TObject);
begin
   tsTrazabilidad.Color := global_color_entrada;
end;

procedure TfrmSalidaAlmacen.tsTrazabilidadExit(Sender: TObject);
begin
   tsTrazabilidad.Color := global_color_salida;
end;

procedure TfrmSalidaAlmacen.tsTrazabilidadKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin    if tsTrazabilidad.ReadOnly = False then
    begin
        filtro := 'Trazabilidad';
        filtra
    end;
end;

procedure TfrmSalidaAlmacen.tsTipomovimientoKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key = #13 then
       tsNumeroOrden.SetFocus;
end;

procedure TfrmSalidaAlmacen.txtNombreContextPopup(Sender: TObject;
  MousePos: TPoint; var Handled: Boolean);
begin
    txtNombre.Color := global_color_salida
end;

procedure TfrmSalidaAlmacen.txtNombreEnter(Sender: TObject);
begin
    txtNombre.Color := global_color_entrada
end;

procedure TfrmSalidaAlmacen.txtNombreExit(Sender: TObject);
begin
    txtNombre.Color := global_color_salida
end;

procedure TfrmSalidaAlmacen.txtNombreKeyPress(Sender: TObject; var Key: Char);
begin
      If Key = #13 Then
          tmComentarios.SetFocus
end;

procedure TfrmSalidaAlmacen.Grid_EntradasDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  orden, status :  string;
begin
//  QryBuscaFolio.Active := False;
//  QryBuscaFolio.Params.ParamByName('contrato').DataType := ftString ;
//  QryBuscaFolio.Params.ParamByName('contrato').Value := param_global_contrato ;
//  QryBuscaFolio.Params.ParamByName('orden').DataType := ftString;
//  QryBuscaFolio.Params.ParamByName('orden').Value := anexo_suministro.FieldByName('sNumeroOrden').AsString;
//  QryBuscaFolio.Open;
//
//  status := QryBuscaFolio.FieldByName('cIdStatus').AsString;
//
//  if status = 'T' then
//  begin
//    Grid_Entradas.Columns.Items[3].Color:= $00011421;
//    Grid_Entradas.Columns.Items[3].Font.Color := $00011421;
//  end;
//
//  if status = 'S' then
//  begin
//    Grid_Entradas.Columns.Items[3].Color := $0000D9D9;
//    Grid_Entradas.Columns.Items[3].Font.Color := $00002222;
//  end;
//
//  if status = 'P' then
//  begin
//    Grid_Entradas.Columns.Items[3].Color := $00FA9C29;
//    Grid_Entradas.Columns.items[3].Font.Color := $00481B02;
//  end;
//  Grid_Entradas.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmSalidaAlmacen.Grid_EntradasTitleClick(Column: TColumn);
begin
   UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmSalidaAlmacen.Grid_PedidoDblClick(Sender: TObject);
begin
  If Pedido.RecordCount > 0 Then
  begin
    if tdCantidad.Enabled  then
    begin
      tsInsumo.Text := Pedido.FieldByName('sidinsumo').AsString;
      tsTrazabilidad.Text := Pedido.FieldValues['sTrazabilidad'];
      tdCantidad.Value  := Pedido.FieldValues['dCantidadRestante'];
      tdCantidad.SetFocus;
    end;
  end;
end;

procedure TfrmSalidaAlmacen.Grid_PedidoKeyPress(Sender: TObject;
  var Key: Char);
begin
  If Key = #13 Then
  begin
    if tdCantidad.Enabled  then
    begin
        tdCantidad.Value  := Pedido.FieldValues['dCantidadRestante'];
        tdCantidad.SetFocus;
    end;
  end;
end;

procedure TfrmSalidaAlmacen.Grid_PedidoTitleClick(Column: TColumn);
begin
if grid_pedido.datasource.DataSet.IsEmpty=false  then
if grid_pedido.DataSource.DataSet.RecordCount>0 then
   UtGrid2.DbGridTitleClick(Column);
end;

procedure TfrmSalidaAlmacen.iFolioChange(Sender: TObject);
begin
  TCurrenCiEditChangef(iFolio, 'No. Salida');
end;

procedure TfrmSalidaAlmacen.iFolioEnter(Sender: TObject);
begin
  ifolio.Color:= global_color_entrada
end;

procedure TfrmSalidaAlmacen.iFolioExit(Sender: TObject);
begin
  ifolio.color:= global_color_salida
end;

procedure TfrmSalidaAlmacen.iFolioKeyPress(Sender: TObject; var Key: Char);
begin
if not keyFiltrotCurrencyEdit(iFolio,key) then
   key:=#0;
 if tdidfecha.Enabled=true then
  begin
  if key = #13 then
    tdidfecha.SetFocus
  end;
end;

procedure TfrmSalidaAlmacen.tmComentariosEnter(Sender: TObject);
begin
    tmComentarios.Color := global_color_entrada
end;

procedure TfrmSalidaAlmacen.tmComentariosExit(Sender: TObject);
begin
    tmComentarios.Color := global_color_salida
end;


procedure TfrmSalidaAlmacen.tmComentariosKeyPress(Sender: TObject;
  var Key: Char);
begin
       If Key = #13 Then
        ifolio.SetFocus
end;

procedure TfrmSalidaAlmacen.tdCantidadChange(Sender: TObject);
begin
  TRxCalcEditChangef(tdCantidad,'Cantidad');
end;

procedure TfrmSalidaAlmacen.tdCantidadEnter(Sender: TObject);
begin
    tdCantidad.Color := global_color_entrada
end;

procedure TfrmSalidaAlmacen.tdCantidadExit(Sender: TObject);
begin
    tdCantidad.Color := global_color_salida
end;

procedure TfrmSalidaAlmacen.tdCantidadKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(tdCantidad,key) then
   key:=#0;
  If Key = #13 Then
  Salvar.SetFocus;
end;

procedure TfrmSalidaAlmacen.tsAlmacenExit(Sender: TObject);
begin

    Pedido.Active := False ;
    Pedido.SQL.Clear;
    Pedido.SQL.Add('Select *, LEFT(mDescripcion, 200) as Descripcion from insumos where sContrato =:Contrato and sIdAlmacen =:Almacen ');
    Pedido.Params.ParamByName('Contrato').DataType := ftString ;
    Pedido.Params.ParamByName('Contrato').Value    := param_global_contrato;
    Pedido.Params.ParamByName('Almacen').DataType  := ftString;
    Pedido.Params.ParamByName('Almacen').Value     := 'ALM-001' ;
    Pedido.Open ;
end;

procedure TfrmSalidaAlmacen.tsAlmacenKeyPress(Sender: TObject; var Key: Char);
begin
 if tsinsumo.Enabled=true then
  begin
  if key = #13 then
    tsinsumo.SetFocus
  end;
end;

procedure TfrmSalidaAlmacen.tsArchivoChange(Sender: TObject);
begin
  if tsArchivo.Text <> '' then begin
    btnImportarVales.Enabled := True;
  end
  else begin
    btnImportarVales.Enabled := False;
  end;

end;

procedure TfrmSalidaAlmacen.tsFamiliaKeyPress(Sender: TObject; var Key: Char);
begin
     If Key = #13 Then
        tdCantidad.SetFocus
end;

procedure TfrmSalidaAlmacen.tsInsumoEnter(Sender: TObject);
begin
    tsInsumo.Color := global_color_entrada
end;

procedure TfrmSalidaAlmacen.tsInsumoExit(Sender: TObject);
begin
      tsInsumo.Color := global_color_salida;  
end;

procedure TfrmSalidaAlmacen.tsInsumoKeyPress(Sender: TObject; var Key: Char);
begin
    if key = #13 then
       grid_pedido.SetFocus;

end;

procedure TfrmSalidaAlmacen.tsInsumoKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
      filtra;
end;

procedure TfrmSalidaAlmacen.tsNumeroActividadKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tdCantidad.SetFocus
end;

procedure TfrmSalidaAlmacen.mComentariosKeyPress(Sender: TObject;
  var Key: Char);
begin
      If Key = #13 Then
        frmBarra1.btnPost.SetFocus
end;

procedure TfrmSalidaAlmacen.mDescripcionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tdcantidad.SetFocus
end;

procedure TfrmSalidaAlmacen.Paste1Click(Sender: TObject);
begin
  if grid_entradas.Focused=true then
    begin
     UtGrid.AddRowsFromClip;
    end;
  if grid_pedido.Focused=true then
    begin
      if grid_pedido.datasource.DataSet.IsEmpty=false  then
      if grid_pedido.DataSource.DataSet.RecordCount>0 then
      UtGrid2.AddRowsFromClip;
    end;
  if gridpartidas.Focused=true then
    begin
      UtGrid3.AddRowsFromClip;
    end;
end;

procedure TfrmSalidaAlmacen.PedidoAfterScroll(DataSet: TDataSet);
begin
    if Pedido.RecordCount > 0 then
    begin
        tsInsumo.Text       := Pedido.FieldValues['sIdInsumo'];
        tsTrazabilidad.Text := Pedido.FieldValues['sTrazabilidad'];
        tsMedida.Text       := Pedido.FieldValues['sMedida'];
        tdCantidad.Value    := Pedido.FieldValues['dCantidadRestante'];
        tmDescripcion.Text  := Pedido.FieldValues['mDescripcion'];
    end;
end;

procedure TfrmSalidaAlmacen.PgControlChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
     Grid_Entradas.Enabled := True;
     Grid_Entradas.SetFocus;
     //Grid_Pedido.Columns[1].Width := 64;
end;

procedure TfrmSalidaAlmacen.pSalidasAfterScroll(DataSet: TDataSet);
begin
     if pSalidas.RecordCount > 0 then
      begin
          try
           GridPartidas.Hint  := pSalidas.FieldValues['mDescripcion'];
           IdInsumo           := pSalidas.FieldValues['sIdInsumo'];
           tsTrazabilidad.Text :=pSalidas.FieldValues['sTrazabilidad'];
           Cantidad           := pSalidas.FieldValues['dCantidad'];
           tsInsumo.Text      := pSalidas.FieldValues['sIdInsumo'];
           tdCantidad.Value   := pSalidas.FieldValues['dCantidad'];
          Except

          end;
          Pedido.Locate('sIdInsumo', pSalidas.FieldByName('sIdInsumo').AsString, [loCaseInsensitive])
      end;
end;

procedure TfrmSalidaAlmacen.frxEntradaGetValue(const VarName: String; var Value: Variant);
var
  zConsulta : TZQuery;
  sSQL      : string;
begin
  If CompareText(VarName, 'TIPO_ENTRADA') = 0 then
      Value := '' ;

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

  sSQL:='SELECT * FROM firmas WHERE sContrato = :contrato AND dIdFecha <= :fecha ORDER BY dIdFecha DESC';
  zConsulta := TZQuery.Create(self);
  zConsulta.Connection := connection.zConnection;
  zConsulta.Active := False;
  zConsulta.SQL.Clear;
  zConsulta.SQL.Add(sSQL);
  zConsulta.Params.ParamByName('contrato').DataType := ftString;
  zConsulta.Params.ParamByName('contrato').Value := param_global_contrato;
  zConsulta.Params.ParamByName('fecha').DataType := ftDate;
  zConsulta.Params.ParamByName('fecha').Value := anexo_suministro.FieldValues['dFechaSalida'];
  zConsulta.Open;
  if zConsulta.RecordCount > 0 then begin
    If CompareText(VarName, 'ENTREGA_PUESTO') = 0 then
        Value := zConsulta.FieldValues['sPuesto17'] ;
    If CompareText(VarName, 'RECIBE_PUESTO') = 0 then
        Value := zConsulta.FieldValues['sPuesto18'] ;
    If CompareText(VarName, 'ENTREGA_FIRMA') = 0 then
        Value := zConsulta.FieldValues['sFirmante17'] ;
    If CompareText(VarName, 'RECIBE_FIRMA') = 0 then
        Value := zConsulta.FieldValues['sFirmante18'] ;
  end
  else
  begin
    If CompareText(VarName, 'ENTREGA_PUESTO') = 0 then
        Value := '*' ;
    If CompareText(VarName, 'RECIBE_PUESTO') = 0 then
        Value := '*' ;
    If CompareText(VarName, 'ENTREGA_FIRMA') = 0 then
        Value := '*' ;
    If CompareText(VarName, 'RECIBE_FIRMA') = 0 then
        Value := '*' ;
  end;

  zConsulta.free;
end;

procedure TfrmSalidaAlmacen.ComentariosAdicionalesClick(Sender: TObject);
begin
    Application.CreateForm(TfrmComentariosxAnexo, frmComentariosxAnexo);
    frmComentariosxAnexo.show ;
end;

procedure TfrmSalidaAlmacen.Copy1Click(Sender: TObject);
begin
  if grid_entradas.Focused=true then
    begin
      UtGrid.CopyRowsToClip;
    end;
  if grid_pedido.Focused=true then
    begin
      if grid_pedido.datasource.DataSet.IsEmpty=false  then
      if grid_pedido.DataSource.DataSet.RecordCount>0 then
      UtGrid2.CopyRowsToClip;
    end;
  if gridpartidas.Focused=true then
    begin
      UtGrid3.CopyRowsToClip;
    end;
end;

procedure TfrmSalidaAlmacen.dbPartidasKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    grid_pedido.SetFocus
end;

procedure TfrmSalidaAlmacen.tdFechaAvisoKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsNumeroOrden.SetFocus
end;

procedure TfrmSalidaAlmacen.tsNumeroOrdenChange(Sender: TObject);
var status : string;
begin
  buscaEstadoFolio;
end;

procedure TfrmSalidaAlmacen.tsNumeroOrdenEnter(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmSalidaAlmacen.tsNumeroOrdenExit(Sender: TObject);
begin
    qryPartidasRef.Active := False;
    qryPartidasRef.ParamByName('Contrato').DataType := ftString;
    qryPartidasRef.ParamByName('Contrato').Value    := param_global_contrato;
    qryPartidasRef.ParamByName('Orden').DataType    := ftString;
    qryPartidasRef.ParamByName('Orden').Value       := tsNumeroOrden.Text;
    qryPartidasRef.Open;

    tsNumeroOrden.Color := global_color_salida
end;

procedure TfrmSalidaAlmacen.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
       If Key = #13 Then
          tsPartida.SetFocus
end;

procedure TfrmSalidaAlmacen.ActivaBotones(Sender: Boolean);
begin
    if sender then
    begin
        iFolio.ReadOnly         := False;
        tsNumeroOrden.Enabled   := True ;
        tmComentarios.ReadOnly  := False ;
    end
    else
    begin
         iFolio.ReadOnly        := True ;
         tdIdFecha.Enabled      := False ;
         tsNumeroOrden.Enabled  := False ;
         tmComentarios.ReadOnly := True ;
         tdCantidad.ReadOnly    := True ;
    end;
end;

procedure TfrmSalidaAlmacen.ActivaBotones2(Sender: Boolean);
begin
    if sender then
    begin
        tdCantidad.Enabled := True;
        tsInsumo.Enabled   := True;
    end
    else
    begin
        tdCantidad.Enabled := False;
        tsInsumo.Enabled   := False;
    end;    
end;

procedure TfrmSalidaAlmacen.AgregarClick(Sender: TObject);
begin
     If (anexo_suministro.RecordCount > 0) Then
     Begin
          OpcButton := 'New';
          Agregar.Enabled  := False ;
          Editar.Enabled   := False ;
          Salvar.Enabled   := True ;
          Cancelar.Enabled := True ;
          Eliminar.Enabled := False ;
          //Imprimir.Enabled := False ;
          ActivaBotones2(true);
          tsInsumo.ReadOnly   := False;
          tdCantidad.ReadOnly := False;
          tsTrazabilidad.ReadOnly := False;
          tsInsumo.Text := '';
          tsInsumo.SetFocus;
          tdCantidad.Enabled  := True;
          tdCantidad.ReadOnly := False;
          ValorPrev := 0;
          PgControl.Pages[0].TabVisible := False;
          GBSuperior.Enabled := False;
    End;
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmSalidaAlmacen.anexo_suministroAfterScroll(DataSet: TDataSet);
var
  folio : string;
  Sender: TObject;
begin
     if anexo_suministro.RecordCount > 0 then
     begin
          iFolio.Value       := anexo_suministro.FieldByName('iFolioSalida').AsInteger;
          EdtExt.Text        := anexo_suministro.FieldByName('sext').AsString;
          tdIdFecha.Date     := anexo_suministro.FieldByName('dFechaSalida').AsDateTime;
          folio              := anexo_suministro.FieldByName('sNumeroOrden').AsString;
          txtNombre.Text     := anexo_suministro.FieldByName('sNombre').AsString;
          tmComentarios.Text := anexo_suministro.FieldByName('mComentarios').AsString;

          tsTipoMovimiento.KeyValue := anexo_suministro.FieldByName('sIdTipo').AsString;
          tsPlataformaRef.KeyValue  := anexo_suministro.FieldByName('sIdPlataforma').AsString;
          tsNumeroOrden.ItemIndex   := tsNumeroOrden.Items.IndexOf(anexo_suministro.FieldByName('sNumeroOrden').AsString);
          tsNumeroOrden.OnExit(sender);
          tsPartida.KeyValue        := anexo_suministro.FieldByName('sNumeroActividad').AsString;

          pSalidas.Active := False;
          pSalidas.ParamByName('Contrato').DataType  := ftString;
          pSalidas.ParamByName('Contrato').Value     := global_contrato_barco;
          pSalidas.ParamByName('Orden').DataType     := ftString;
          pSalidas.ParamByName('Orden').Value        := param_global_contrato;
          pSalidas.ParamByName('Folio').DataType     := ftInteger;
          pSalidas.ParamByName('Folio').Value        := anexo_suministro.FieldByName('iFolioSalida').AsInteger;
          pSalidas.ParamByName('FolioAnio').DataType := ftInteger;
          pSalidas.ParamByName('FolioAnio').Value    := anexo_suministro.FieldByName('iFolioAnio').AsInteger;
          pSalidas.ParamByName('Ext').AsString       := anexo_suministro.FieldByName('sExt').AsString;
          pSalidas.Open;
     end;
end;
procedure TfrmSalidaAlmacen.anexo_suministroCalcFields(DataSet: TDataSet);
var
  status : string;
begin
  QryBuscaFolio.Active := False;
  QryBuscaFolio.Params.ParamByName('contrato').DataType := ftString ;
  QryBuscaFolio.Params.ParamByName('contrato').Value := param_global_contrato ;
  QryBuscaFolio.Params.ParamByName('orden').DataType := ftString;
  QryBuscaFolio.Params.ParamByName('orden').Value := anexo_suministro.FieldByName('sNumeroOrden').AsString;
  QryBuscaFolio.Open;

  status := QryBuscaFolio.FieldByName('cIdStatus').AsString;

  if status = 'T' then
  begin
    anexo_suministro.FieldByName('EstadoFolio').Value := 'TERMINADO';
  end;

  if status = 'S' then
  begin
    anexo_suministro.FieldByName('EstadoFolio').Value := 'SUSPENDIDO';
  end;

  if status = 'P' then
  begin
    anexo_suministro.FieldByName('EstadoFolio').Value := 'PROCESO';
  end;
end;

End.
