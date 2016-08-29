unit frm_equipos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, global, ComCtrls, 
  StdCtrls, ExtCtrls, DBCtrls, Mask, frm_barra, db, Menus, frxClass,unitValidacion,
  frxDBSet, RXDBCtrl,  Buttons, RxMemDS, utilerias, RXSpin, DateUtils, ZDataset,
  ZAbstractRODataset,  rxCurrEdit, rxToolEdit,
  UnitExcepciones, UdbGrid, unittbotonespermisos, UnitValidaTexto, UnitExcel,
  ComObj, UnitTablasImpactadas, unitactivapop, UFunctionsGHH, DBDateTimePicker,
  ZAbstractDataset, RxLookup, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue,
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
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit,
  cxNavigator, cxDBData, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid;

type
  TfrmEquipos = class(TForm)
    frmBarra1: TfrmBarra;
    DBequipos: TfrxDBDataset;
    ds_equipos: TDataSource;
    pgPersonal: TPageControl;
    TabInformacion: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    tsIdEquipo: TDBEdit;
    tsDescripcion: TDBEdit;
    tsMedida: TDBEdit;
    tdCostoMn: TDBEdit;
    tdCostoDll: TDBEdit;
    tsIdTipodeEquipo: TDBLookupComboBox;
    tlProrrateo: TDBComboBox;
    tiJornada: TDBEdit;
    tlDistribuye: TDBComboBox;
    tdVentaMN: TDBEdit;
    tdVentaDLL: TDBEdit;
    Label8: TLabel;
    tiItemOrden: TDBEdit;
    Label14: TLabel;
    Label15: TLabel;
    ds_DistribuciondeEquipos: TDataSource;
    MemoryEquipos: TRxMemoryData;
    MemoryEquiposAnno: TIntegerField;
    MemoryEquiposMes: TStringField;
    MemoryEquiposCantidad: TFloatField;
    ds_Memory: TDataSource;
    frxEquipos: TfrxReport;
    MemoryEquiposiMes: TIntegerField;
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
    gbTarifaDiaria: TGroupBox;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    rxDistribucion: TRxDBGrid;
    tsMes: TComboBox;
    tiAnno: TRxSpinEdit;
    btnDistribuir: TBitBtn;
    tdCantidadMensual: TCurrencyEdit;
    DBAnexoDT4: TfrxDBDataset;
    AnexoDT41: TMenuItem;
    ActualizacindeCostos1: TMenuItem;
    frxAnexoDT4: TfrxReport;
    ds_tiposdeequipo: TDataSource;
    Equipos: TZQuery;
    MaximoItem: TZReadOnlyQuery;
    TiposdeEquipo: TZReadOnlyQuery;
    DistribuciondeEquipos: TZQuery;
    DistribuciondeEquipossIdEquipo: TStringField;
    DistribuciondeEquiposdIdFecha: TDateField;
    DistribuciondeEquiposdCantidad: TFloatField;
    DistribuciondeEquiposiAnno: TIntegerField;
    DistribuciondeEquipossMes: TStringField;
    AnexoDT4: TZReadOnlyQuery;
    EquipossContrato: TStringField;
    EquipossIdEquipo: TStringField;
    EquiposiItemOrden: TIntegerField;
    EquipossDescripcion: TStringField;
    EquipossIdTipoEquipo: TStringField;
    EquipossMedida: TStringField;
    EquiposdCantidad: TFloatField;
    EquiposdCostoMN: TFloatField;
    EquiposdCostoDLL: TFloatField;
    EquiposdVentaMN: TFloatField;
    EquiposdVentaDLL: TFloatField;
    EquiposdFechaInicio: TDateField;
    EquiposdFechaFinal: TDateField;
    EquiposlProrrateo: TStringField;
    EquiposiJornada: TIntegerField;
    EquiposlDistribuye: TStringField;
    Label16: TLabel;
    Label17: TLabel;
    tlImprime: TDBComboBox;
    tlCobro: TDBComboBox;
    EquiposiCobro: TStringField;
    EquiposlImprime: TStringField;
    ExportaaPlantillaExcel1: TMenuItem;
    SaveDialog1: TSaveDialog;
    tdFechaInicio: TDBDateTimePicker;
    tdFechaFinal: TDBDateTimePicker;
    tdCantidad: TDBEdit;
    Label18: TLabel;
    tCuadra: TDBComboBox;
    EquiposlCuadraEquipo: TStringField;
    DBCheckBox1: TDBCheckBox;
    EquiposlAplicaDiesel: TStringField;
    EquipossDescripcionDiesel: TStringField;
    dbedtsDescripcionDiesel: TDBEdit;
    Label22: TLabel;
    Label26: TLabel;
    dbcbblSumaSolicitado: TDBComboBox;
    EquiposlSumaSolicitado: TStringField;
    lbl_AgrupadorPersonal: TLabel;
    EquiposiId_AgrupadorEquipo: TIntegerField;
    zq_gruposdeequipo: TZReadOnlyQuery;
    ds_gruposdeequipo: TDataSource;
    dbedtNumeroSerie: TDBEdit;
    lbl1: TLabel;
    DBLookupComboBox1: TDBLookupComboBox;
    ZUnificadores: TZReadOnlyQuery;
    DsUnificadores: TDataSource;
    CbxUnificar: TCheckBox;
    EquiposiUnificador: TIntegerField;
    EquipossNumeroSerie: TStringField;
    raerequiposoptativosdemaestro1: TMenuItem;
    PnlProgreso: TPanel;
    BarraProgreso: TProgressBar;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label27: TLabel;
    TituloProgreso: TLabel;
    tsPaquete: TRxDBLookupCombo;
    cxgrdbtblvwGrid1DBTableView1: TcxGridDBTableView;
    cxgrdlvlGrid1Level1: TcxGridLevel;
    grid_Equipos: TcxGrid;
    tcxItem: TcxGridDBColumn;
    tcxEquipo: TcxGridDBColumn;
    tcxTipo: TcxGridDBColumn;
    tcxDescripcion: TcxGridDBColumn;
    tcxMedida: TcxGridDBColumn;
    tcxCantidad: TcxGridDBColumn;
    tcxFechaInicio: TcxGridDBColumn;
    tcxFechaFinal: TcxGridDBColumn;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure tsDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure tsMedidaKeyPress(Sender: TObject; var Key: Char);
    procedure tdCostoMnKeyPress(Sender: TObject; var Key: Char);
    procedure tdCostoDllKeyPress(Sender: TObject; var Key: Char);
    procedure tlProrrateoKeyPress(Sender: TObject; var Key: Char);
    procedure tiJornadaKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdTipodeEquipoKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdTipodeEquipoExit(Sender: TObject);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure tsIdEquipoKeyPress(Sender: TObject; var Key: Char);
    procedure Salir1Click(Sender: TObject);
    procedure tlDistribuyeKeyPress(Sender: TObject; var Key: Char);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure tdCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure grid_equiposEnter(Sender: TObject);
    procedure tsIdTipodeEquipoEnter(Sender: TObject);
    procedure tsIdEquipoEnter(Sender: TObject);
    procedure tsIdEquipoExit(Sender: TObject);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure tsMedidaEnter(Sender: TObject);
    procedure tsMedidaExit(Sender: TObject);
    procedure tdCantidadEnter(Sender: TObject);
    procedure tdCantidadExit(Sender: TObject);
    procedure tdCostoMnEnter(Sender: TObject);
    procedure tdCostoMnExit(Sender: TObject);
    procedure tdCostoDllEnter(Sender: TObject);
    procedure tdCostoDllExit(Sender: TObject);
    procedure tlProrrateoEnter(Sender: TObject);
    procedure tlProrrateoExit(Sender: TObject);
    procedure tiJornadaEnter(Sender: TObject);
    procedure tiJornadaExit(Sender: TObject);
    procedure tlDistribuyeEnter(Sender: TObject);
    procedure tlDistribuyeExit(Sender: TObject);
    procedure tdVentaDLLKeyPress(Sender: TObject; var Key: Char);
    procedure tdVentaMNEnter(Sender: TObject);
    procedure tdVentaMNExit(Sender: TObject);
    procedure tdVentaMNKeyPress(Sender: TObject; var Key: Char);
    procedure tdVentaDLLEnter(Sender: TObject);
    procedure tdVentaDLLExit(Sender: TObject);
    procedure frmBarra1btnPrinterClick(Sender: TObject);
    procedure tiItemOrdenEnter(Sender: TObject);
    procedure tiItemOrdenExit(Sender: TObject);
    procedure tiItemOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaFinalKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaInicioEnter(Sender: TObject);
    procedure tdFechaInicioExit(Sender: TObject);
    procedure tdFechaInicioKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaFinalEnter(Sender: TObject);
    procedure tdFechaFinalExit(Sender: TObject);
    procedure EquiposAfterScroll(DataSet: TDataSet);
    procedure Grid_EquiposTitleBtnClick(Sender: TObject; ACol: Integer;
      Field: TField);
    procedure DistribuciondeEquiposCalcFields(DataSet: TDataSet);
    procedure btnDistribuirClick(Sender: TObject);
    procedure tiAnnoEnter(Sender: TObject);
    procedure tiAnnoExit(Sender: TObject);
    procedure tiAnnoKeyPress(Sender: TObject; var Key: Char);
    procedure tsMesEnter(Sender: TObject);
    procedure tsMesExit(Sender: TObject);
    procedure tsMesKeyPress(Sender: TObject; var Key: Char);
    procedure tdCantidadMensualEnter(Sender: TObject);
    procedure tdCantidadMensualExit(Sender: TObject);
    procedure tdCantidadMensualKeyPress(Sender: TObject; var Key: Char);
    procedure AnexoDT41Click(Sender: TObject);
    procedure ActualizacindeCostos1Click(Sender: TObject);
    procedure rxDistribucionKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure rxDistribucionKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure tlImprimeEnter(Sender: TObject);
    procedure tlImprimeExit(Sender: TObject);
    procedure tlImprimeKeyPress(Sender: TObject; var Key: Char);
    procedure tlCobroEnter(Sender: TObject);
    procedure tlCobroExit(Sender: TObject);
    procedure tlCobroKeyPress(Sender: TObject; var Key: Char);
    //****************************BRITO 02/12/10****************************
    procedure actualizarEquipo(idOrig: string);
    function posibleBorrar(idOrig: string): boolean;
    procedure Grid_EquiposMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_EquiposMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_EquiposTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure ExportaaPlantillaExcel1Click(Sender: TObject);
    //****************************BRITO 02/12/10****************************
    procedure formatoEncabezado();
    function estaReportado(sIdEquipo: string): boolean;
    function estaDistribuido(sIdEquipo: string): boolean;
    function tienePersonal(sIdEquipo: string): boolean;
    procedure tdFechaFinalChange(Sender: TObject);
    procedure tdFechaInicioChange(Sender: TObject);
    procedure tdCantidadMensualChange(Sender: TObject);
    procedure EquiposdCantidadSetText(Sender: TField; const Text: string);
    procedure EquiposdCostoMNSetText(Sender: TField; const Text: string);
    procedure EquiposdCostoDLLSetText(Sender: TField; const Text: string);
    procedure EquiposdVentaMNSetText(Sender: TField; const Text: string);
    procedure EquiposBeforePost(DataSet: TDataSet);
    procedure EquiposPostError(DataSet: TDataSet; E: EDatabaseError;
      var Action: TDataAction);
    procedure tiJornadaChange(Sender: TObject);
    procedure tdCantidadChange(Sender: TObject);
    procedure tdVentaDLLChange(Sender: TObject);
    procedure tdCostoMnChange(Sender: TObject);
    procedure tdVentaMNChange(Sender: TObject);
    procedure tdCostoDllChange(Sender: TObject);
    procedure EquiposdVentaDLLSetText(Sender: TField; const Text: string);
    procedure EquiposiJornadaSetText(Sender: TField; const Text: string);
    procedure dbedtNumeroSerieKeyPress(Sender: TObject; var Key: Char);
    procedure dbedtNumeroSerieEnter(Sender: TObject);
    procedure dbedtNumeroSerieExit(Sender: TObject);
    procedure dbedtsDescripcionDieselKeyPress(Sender: TObject; var Key: Char);
    procedure dbedtsDescripcionDieselEnter(Sender: TObject);
    procedure dbedtsDescripcionDieselExit(Sender: TObject);
    procedure dbcbblSumaSolicitadoEnter(Sender: TObject);
    procedure dbcbblSumaSolicitadoExit(Sender: TObject);
    procedure dbcbblSumaSolicitadoKeyPress(Sender: TObject; var Key: Char);
    procedure tCuadraKeyPress(Sender: TObject; var Key: Char);
    procedure tCuadraEnter(Sender: TObject);
    procedure tCuadraExit(Sender: TObject);
    procedure DBCheckBox1Click(Sender: TObject);
    procedure raerequiposoptativosdemaestro1Click(Sender: TObject);
  private
  sMenuP: string;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  UtGrid:TicDbGrid;
  botonpermiso:tbotonespermisos;
  frmEquipos  : TfrmEquipos;
  sTipoEquipo : String ;
  Months: array[1..12] of string;
  //****************************BRITO 02/12/10****************************
  sOldEquipo  : String ;
  sOldDescripcion: string;
  //****************************BRITO 02/12/10****************************
  banderaAgregar:boolean;

  //Exporta elementos a Excel..
  Excel, Libro, Hoja: Variant;

implementation

{$R *.dfm}


procedure TfrmEquipos.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Equipos.Cancel ;
  action := cafree ;
  BotonPermiso.Free;
end;



procedure TfrmEquipos.FormShow(Sender: TObject);
begin
  sMenuP:=stMenu;

  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cEquipos', PopupPrincipal);
  OpcButton := '' ;
  sTipoEquipo := '' ;
  frmbarra1.btnCancel.Click ;

  zq_gruposdeequipo.Active := False;
  zq_gruposdeequipo.Open;
  
  Equipos.Active := False ;
  Equipos.Params.ParamByName('Contrato').DataType := ftString ;
  Equipos.Params.ParamByName('Contrato').Value := Global_Contrato ;
  Equipos.Params.ParamByName('Orden').DataType := ftString ;
  Equipos.Params.ParamByName('Orden').Value := 'iItemOrden' ;
  Equipos.Open ;

  Months[1] := 'ENERO';
  Months[2] := 'FEBRERO';
  Months[3] := 'MARZO';
  Months[4] := 'ABRIL';
  Months[5] := 'MAYO';
  Months[6] := 'JUNIO';
  Months[7] := 'JULIO';
  Months[8] := 'AGOSTO';
  Months[9] := 'SEPTIEMBRE';
  Months[10] := 'OCTUBRE';
  Months[11] := 'NOVIEMBRE';
  Months[12] := 'DICIEMBRE';

  TiposdeEquipo.Active := False ;
  TiposdeEquipo.Open ;
  //UtGrid:=TicdbGrid.create(grid_equipos);
  BotonPermiso.permisosBotones(frmBarra1);

  ZUnificadores.Active := False;
  ZUnificadores.ParamByName('sContrato').AsString := global_Contrato_Barco;
  ZUnificadores.Open;
  
end;


procedure TfrmEquipos.tsDescripcionKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsMedida.SetFocus
end;

procedure TfrmEquipos.tsMedidaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdCantidad.SetFocus 
end;

procedure TfrmEquipos.tdCostoMnKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTdbedit(tdCostoMn,key) then
      key:=#0;
  if key = #13 then
    tdCostoDll.SetFocus
end;

procedure TfrmEquipos.tdCostoDllKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTdbedit(tdCostoDll,key) then
      key:=#0;
  if key = #13 then
    tdVentaMN.SetFocus 
end;

procedure TfrmEquipos.tlProrrateoKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then
    tiJornada.SetFocus
end;


procedure TfrmEquipos.tiJornadaKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTdbedit(tiJornada,key) then
      key:=#0;
  if key = #13 then
    tlImprime.SetFocus
end;

procedure TfrmEquipos.tsIdTipodeEquipoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsIdEquipo.SetFocus ;
  If ( Equipos.State = dsEdit ) OR ( Equipos.State = dsInsert ) then
    tsIdEquipo.SelStart := length ( TiposdeEquipo.FieldValues[ 'sMascara' ] );
end;

procedure TfrmEquipos.tsIdTipodeEquipoExit(Sender: TObject);
begin
  If ( Equipos.State = dsEdit ) OR ( Equipos.State = dsInsert ) then
      If (OpcButton = 'New') then
          Equipos.fieldvalues[ 'sIdEquipo' ] := TiposdeEquipo.FieldValues[ 'sMascara' ] + trim(tiItemOrden.Text);
  tsIdTipodeEquipo.Color := global_color_salida
end;

function TfrmEquipos.estaDistribuido(sIdEquipo: string): boolean;
begin
  result := false;
  with connection.QryBusca do
  begin
    Active := false;
    Filtered := false;
    SQL.Clear;
    SQL.Add('SELECT sIdEquipo FROM distribuciondeequipos WHERE sContrato = :Contrato AND sIdEquipo = :sIdEquipo LIMIT 1');
    ParamByName('Contrato').Value := Global_Contrato;
    ParamByName('sIdEquipo').Value := sIdEquipo;
    Open;
    if RecordCount > 0 then
      result := true;
  end;
end;

function TfrmEquipos.estaReportado(sIdEquipo: string): boolean;
begin
  result := false;
  with connection.QryBusca do
  begin
    Active := false;
    Filtered := false;
    SQL.Clear;
    SQL.Add('SELECT sIdEquipo FROM bitacoradeequipos WHERE sContrato = :Contrato AND sIdEquipo = :sIdEquipo LIMIT 1');
    ParamByName('Contrato').Value := Global_Contrato;
    ParamByName('sIdEquipo').Value := sIdEquipo;
    Open;
    if RecordCount > 0 then
      result := true;
  end;
end;

function TfrmEquipos.tienePersonal(sIdEquipo: string): boolean;
begin
  result := false;
  with connection.QryBusca do
  begin
    Active := false;
    Filtered := false;
    SQL.Clear;
    SQL.Add('SELECT sIdPersonal FROM equiposxpersonal WHERE sContrato = :Contrato AND sIdEquipo = :sIdEquipo LIMIT 1');
    ParamByName('Contrato').Value := Global_Contrato;
    ParamByName('sIdEquipo').Value := sIdEquipo;
    Open;
    if RecordCount > 0 then
      result := true;
  end;
end;

procedure TfrmEquipos.frmBarra1btnAddClick(Sender: TObject);
begin
 banderaAgregar:=true;
 try
   activapop(frmEquipos, popupprincipal);
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   Equipos.Append ;

   MaximoItem.Active := False ;
   MaximoItem.SQL.Clear ;
   MaximoItem.SQL.Add('SELECT Max(iItemOrden) as Item FROM equipos where sContrato = :contrato Group By sContrato' ) ;
   MaximoItem.Params.ParamByName('Contrato').DataType := ftString ;
   MaximoItem.Params.ParamByName('Contrato').Value := Global_Contrato ;
   MaximoItem.Open ;
   If MaximoItem.FieldByName('Item').IsNull then
       Equipos.FieldValues [ 'iItemOrden' ] := 1
   else
       Equipos.FieldValues [ 'iItemOrden' ]  := MaximoItem.FieldValues['Item'] + 1 ;

   Equipos.FieldValues [ 'sContrato' ]  := Global_Contrato ;
   Equipos.FieldValues ['sMedida'] := 'Dia' ;
   Equipos.FieldValues ['iJornada'] := 24 ;
   Equipos.FieldValues['sIdTipoEquipo'] := sTipoEquipo ;
   Equipos.FieldValues ['lProrrateo'] := 'Si' ;
   Equipos.FieldValues ['lDistribuye'] := 'Si' ;
   Equipos.FieldValues ['dFechaInicio'] := Date ;
   Equipos.FieldValues ['dFechaFinal'] := Date ;
   Equipos.FieldValues [ 'dCantidad' ] := 0 ;
   Equipos.FieldValues [ 'dCostoMN' ] := 0 ;
   Equipos.FieldValues [ 'dCostoDLL' ] := 0 ;
   Equipos.FieldValues [ 'dVentaMN' ] := 0 ;
   Equipos.FieldValues [ 'dVentaDLL' ] := 0 ;
   Equipos.FieldValues [ 'lImprime' ] := 'Si' ;
   Equipos.FieldValues [ 'lCobro' ] := 'Si' ;
   Equipos.FieldValues [ 'lSumaSolicitado' ] := 'No' ;
   Equipos.FieldValues [ 'lCuadraEquipo' ] := 'No' ;
   //*******************************BRITO 02/12/10******************************
   sOldEquipo := '' ;
   sOldDescripcion := '';
   //*******************************BRITO 02/12/10******************************
   tsIdTipodeEquipo.SetFocus ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_equipos.Enabled:=false;
 Except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_equipos', 'Al agregar registro', 0);
  end;
 end;
end;

procedure TfrmEquipos.frmBarra1btnEditClick(Sender: TObject);
begin
   if estaDistribuido(Equipos.FieldByName('sIdEquipo').AsString) then
   begin
     MessageDlg('No se puede editar la categoría porque ya ha sido utilizada en la distribución', mtInformation, [mbOk], 0);
     exit;
   end;
   if tienePersonal(Equipos.FieldByName('sIdEquipo').AsString) then
   begin
     MessageDlg('No se puede editar la categoría porque esta ya ha sido asignada a personal', mtInformation, [mbOk], 0);
     exit;
   end;     
   banderaagregar:=false;
   frmBarra1.btnEditClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   try
       activapop(frmEquipos, popupprincipal);
       Equipos.Edit ;
       //*****************************BRITO 02/12/10****************************
       sOldEquipo := Equipos.FieldValues['sIdEquipo'] ;
       sOldDescripcion := Equipos.FieldValues['sDescripcion'] ;
       //*****************************BRITO 02/12/10****************************
   except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Equipos', 'Al editar registro', 0);
           frmbarra1.btnCancel.Click ;
       end;
   end;
   tsIdEquipo.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_equipos.Enabled:=false;
end;

procedure TfrmEquipos.actualizarEquipo(idOrig: string);
begin
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('Update bitacoradeequipos set sIdEquipo =:Equipo, sDescripcion =:Descripcion where sContrato =:contrato and sIdEquipo =:EquipoOld');
    connection.zCommand.ParamByName('Contrato').AsString    := global_contrato;
    connection.zCommand.ParamByName('Descripcion').AsString := equipos.FieldValues['sDescripcion'];
    connection.zCommand.ParamByName('Equipo').AsString      := equipos.FieldValues['sIdEquipo'];
    connection.zCommand.ParamByName('EquipoOld').AsString   := idOrig;
    connection.zCommand.ExecSQL;
end;

function TfrmEquipos.posibleBorrar(idOrig: string): boolean;
var
  ParamNamesWHERE,ParamValuesWHERE: TStringList;
begin
  result := true;
  ParamNamesWHERE:=TStringList.Create;ParamValuesWHERE:=TStringList.Create;
  ParamNamesWHERE.Add('sContrato');ParamValuesWHERE.Add(global_contrato);
  ParamNamesWHERE.Add('sIdEquipo');ParamValuesWHERE.Add(idOrig);
  result := not UnitTablasImpactadas.hayDependientes('equipos',ParamNamesWHERE,ParamValuesWHERE);
  if result then
  begin
    ParamNamesWHERE.Clear;ParamValuesWHERE.Clear;
    ParamNamesWHERE.Add('sContrato');ParamValuesWHERE.Add(global_contrato);
    ParamNamesWHERE.Add('sEquipo');ParamValuesWHERE.Add(idOrig);
    result := not UnitTablasImpactadas.hayDependientes('equipos1',ParamNamesWHERE,ParamValuesWHERE);
  end;
end;
procedure TfrmEquipos.raerequiposoptativosdemaestro1Click(Sender: TObject);
var
  ZRqEq:TZReadOnlyQuery;
  ZqTeOpM:TZQuery;
  Campo:string;
  Cont:Integer;
  CurOld:TCursor;
begin
  CurOld := screen.Cursor;
  Screen.Cursor := crAppStart;
  try
    ZRqEq:=TZReadOnlyQuery.Create(nil);
    try
      ZRqEq.Connection := connection.zConnection;
      ZRqEq.Active := False;
      ZRqEq.SQL.Clear;
      ZRqEq.SQL.Text := 'SELECT * FROM equipos WHERE sContrato = :CBarco and lAplicaDiesel = "Si"';
      ZRqEq.ParamByName('CBarco').AsString := global_Contrato_Barco;
      ZRqEq.Open;

      if ZRqEq.RecordCount > 0 then
      begin
        TituloProgreso.Caption := 'Procesando '+inttostr(ZRqEq.RecordCount)+' Equipos.';
        PnlProgreso.Visible := True;
        BarraProgreso.Max := ZRqEq.RecordCount;
        ZqTeOpM := TZQuery.Create(nil);
        try
          ZqTeOpM.Connection := connection.zConnection;
          ZqTeOpM.Active := False;
          ZqTeOpM.SQL.Clear;
          ZqTeOpM.SQL.Text := 'SELECT * FROM equipos WHERE scontrato = :contrato and lAplicaDiesel = "Si"';
          ZqTeOpM.ParamByName('contrato').AsString := global_contrato;
          ZqTeOpM.Open;
          ZRqEq.First;
          while not ZRqEq.Eof do
          begin
            BarraProgreso.Position := ZRqEq.RecNo+1;
            //si no esta ya dado de alta entonces  agregarlo
            if not ZqTeOpM.Locate('sidequipo',ZRqEq.FieldByName('sidequipo').AsString,[]) then
            begin
              ZqTeOpM.Append;
              for cont := 0 to ZRqEq.FieldList.count-1 do
              begin
                Campo := ZRqEq.FieldList[cont].FieldName;
                if lowercase(Campo) <> 'scontrato' then
                  ZqTeOpM.FieldByName(Campo).AsString := ZRqEq.FieldByName(Campo).AsString
                else
                  ZqTeOpM.FieldByName(Campo).AsString := global_contrato;
              end;
              ZqTeOpM.Post;
              ZqTeOpM.Refresh;
            end;
            ZRqEq.Next;
          end;
        finally
          ZqTeOpM.Free;
        end;
      end;
    finally
      ZRqEq.Free;
    end;
  finally
    Screen.Cursor := CurOld;
    if PnlProgreso.Visible then
      PnlProgreso.Visible := False;
  end;

end;

//****************************BRITO 02/12/10************************************

procedure TfrmEquipos.frmBarra1btnPostClick(Sender: TObject);
//*******************************BRITO 02/12/10*********************************
var
    lEdicion: boolean;
    sIdKardex, opcKardex: string;
//*******************************BRITO 02/12/10*********************************
  nombres, cadenas: TStringList;
  i : integer;
begin
  //empieza validacion
  nombres:=TStringList.Create;cadenas:=TStringList.Create;
  nombres.Add('Item Orden');nombres.Add('T equipo'); nombres.Add('Descripcion');
  nombres.Add('U medidas');;nombres.Add('Costo MN');
  nombres.Add('Precio MN');nombres.Add('Prorrateo?'); nombres.Add('Jornada');
  nombres.Add('Se imprime?');nombres.Add('Se cobra?');nombres.Add('Distribuye');
  nombres.Add('Cantidad');nombres.Add('Costo DLL');
  nombres.Add('Precio DLL');
  cadenas.Add(tiItemorden.Text);cadenas.Add(tsIdTipodeEquipo.Text);cadenas.Add(tsDescripcion.Text);
  cadenas.Add(tsMedida.Text);cadenas.Add(tdCostomn.text);
  cadenas.Add(tdVentaMn.Text);cadenas.Add(tlProrrateo.Text);cadenas.Add(tiJornada.Text);
  cadenas.Add(tlImprime.Text);cadenas.Add(tlcobro.Text);cadenas.Add(tlDistribuye.Text);
  cadenas.Add(tdCantidad.Text);cadenas.Add(tdCostoDLL.Text);
  cadenas.Add(tdVentaDLL.Text);
  if not validaTexto(nombres, cadenas, 'Equipo id', tsIdEquipo.Text) then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;

   //Verifica que la fecha final no sea menor que la fecha inicio
   if tdFechaFinal.Date<tdFechaInicio.Date then
   begin
     showmessage('la fecha final es menor a la fecha inicial' );
     tdFechaFinal.SetFocus;
     exit;
   end;

  //Validacion de unidad de medida..
  try
      i := StrToInt(tsMedida.Text);
      messageDLg('No se aceptan numeros como Unidad de Medida.', mtInformation, [mbOk], 0);
      exit;
  Except
      //Continua...
  end;
    Equipos.fieldvalues['dFechaInicio'] :=tdfechaInicio.date;
    Equipos.fieldvalues['dFechaFinal'] :=tdfechaFinal.date;
   try
      lEdicion := false;//**********************BRITO 02/12/10******************
      If OpcButton = 'Edit' Then
      Begin
          lEdicion := true;//********************BRITO 02/12/10*****************
      End ;
      sTipoEquipo := Equipos.FieldValues['sIdTipoEquipo'] ;
      if not CbxUnificar.Checked then
        Equipos.FieldByName('iunificador').Clear;

      Equipos.Post ;

      //Aqui solo hacemos una excepcion de edicion si el equipo fue reportado nadamas si permite editarlo..
      actualizarEquipo(sOldEquipo);
      //registrar la operacion en el kardex
      sIdKardex := Equipos.FieldByName('sIdEquipo').AsString;
      if lEdicion then
          opcKardex := 'Edita Equipo'
      else
          opcKardex := 'Crea Equipo';
      Kardex('Otros Movimientos', opcKardex, sIdKardex, 'ID Equipo', '', '', '','Tarifa Diaria','Catalogo de Equipos' );


      Insertar1.Enabled := True ;
      Editar1.Enabled := True ;
      Registrar1.Enabled := False ;
      Can1.Enabled := False ;
      Eliminar1.Enabled := True ;
      Refresh1.Enabled := True ;
      Salir1.Enabled := True ;
      frmBarra1.btnPostClick(Sender);
   except
      on e : exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Equipos', 'Al salvar registro', 0);
          frmbarra1.btnCancel.Click ;
          lEdicion := false;//cancelar la actualizacion de tablas dependientes
      end;
   end;
  //en caso de ser edicion, aplicar cambio de ID en tablas dependientes
  if (lEdicion) and (Equipos.FieldByName('sIdEquipo').AsString <> sOldEquipo) then
  begin
    actualizarEquipo(sOldEquipo);
  end;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  grid_equipos.Enabled:=true;
  frmbarra1.btnCancel.Click;
  if banderaAgregar then
  frmbarra1.btnAdd.Click;
end;

procedure TfrmEquipos.frmBarra1btnCancelClick(Sender: TObject);
begin
   desactivapop(popupprincipal);
   frmBarra1.btnCancelClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   Equipos.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_equipos.Enabled:=true;
end;

procedure TfrmEquipos.frmBarra1btnDeleteClick(Sender: TObject);
var
  //*********************BRITO 02/12/10*************************
  lSigue: boolean;
  sIdKardex, opcKardex: string;
  //*********************BRITO 02/12/10*************************
begin
  If Equipos.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if not posibleBorrar(Equipos.FieldByName('sIdEquipo').AsString) then
      begin
        MessageDlg('No es posible eliminar el registro, existen registros dependientes.', mtInformation, [mbOk], 0);
        exit;
      end;

      //**************************BRITO 02/12/10********************************
      //verificar si el equipo se encuentra en alguna de las siguientes tablas antes de borrar
      lSigue := true;
      //tabla bitacoradeequipos
      Connection.QryBusca.Active := False ;
      Connection.QryBusca.SQL.Clear ;
      Connection.QryBusca.SQL.Add('SELECT sIdEquipo FROM bitacoradeequipos ' +
                                  'WHERE sContrato = :Contrato AND sIdEquipo = :Equipo') ;
      Connection.QryBusca.ParamByName('Contrato').Value := global_contrato ;
      Connection.QryBusca.ParamByName('Equipo').Value := Equipos.FieldByName('sIdEquipo').AsString;
      Connection.QryBusca.Open ;
      If Connection.QryBusca.RecordCount > 0 Then Begin
          MessageDlg('El equipo que desea eliminar a sido utilizado en reportes diarios, por lo tanto no puede eliminarse.', mtInformation, [mbOk], 0);
          lSigue := false;
      End;
      If lSigue Then Begin
          //tabla distribuciondeequipos y equiposxpersonal
          Connection.QryBusca.Active := False ;
          Connection.QryBusca.SQL.Clear ;
          Connection.QryBusca.SQL.Add('SELECT sIdEquipo FROM distribuciondeequipos '+
                                      'WHERE sContrato = :Contrato AND sIdEquipo = :Equipo '+
                                      'UNION SELECT sIdEquipo FROM equiposxpersonal '+
                                      'WHERE sContrato = :Contrato AND sIdEquipo = :Equipo');
          Connection.QryBusca.ParamByName('Contrato').Value := Global_Contrato ;
          Connection.QryBusca.ParamByName('Equipo').Value := Equipos.FieldByName('sIdEquipo').AsString;
          Connection.QryBusca.Open ;
          If Connection.QryBusca.RecordCount > 0 Then Begin
              MessageDlg('El equipo que desea eliminar ha sido utilizado en distribución de personal, no puede eliminarse.', mtInformation, [mbOk], 0);
              lSigue := false;
          End;
      End;
      If lSigue Then Begin
          //tablas paquetesdeequipo, recursosequipo y recursosequiposnuevos
          Connection.QryBusca.Active := False ;
          Connection.QryBusca.SQL.Clear ;
          Connection.QryBusca.SQL.Add('SELECT sIdEquipo FROM paquetesdeequipo '+
                                      'WHERE sContrato = :Contrato AND sIdEquipo = :Equipo '+
                                      'UNION SELECT sIdEquipo FROM recursosequipo '+
                                      'WHERE sContrato = :Contrato AND sIdEquipo = :Equipo '+
                                      'UNION SELECT sIdEquipo FROM recursosequiposnuevos '+
                                      'WHERE sContrato = :Contrato AND sIdEquipo = :Equipo');
          Connection.QryBusca.ParamByName('Contrato').Value := Global_Contrato ;
          Connection.QryBusca.ParamByName('Equipo').Value := Equipos.FieldByName('sIdEquipo').AsString;
          Connection.QryBusca.Open ;
          If Connection.QryBusca.RecordCount > 0 Then Begin
              MessageDlg('El equipo que desea eliminar ha sido utilizado en recursos de equipo, no puede eliminarse.', mtInformation, [mbOk], 0);
              lSigue := false;
          End;
      End;
      If lSigue Then Begin
          //tablas estimacionxequipo
          Connection.QryBusca.Active := False ;
          Connection.QryBusca.SQL.Clear ;
          Connection.QryBusca.SQL.Add('SELECT sIdEquipo FROM estimacionxequipo '+
                                      'WHERE sContrato = :Contrato AND sIdEquipo = :Equipo');
          Connection.QryBusca.ParamByName('Contrato').Value := Global_Contrato ;
          Connection.QryBusca.ParamByName('Equipo').Value := Equipos.FieldByName('sIdEquipo').AsString;
          Connection.QryBusca.Open ;
          If Connection.QryBusca.RecordCount > 0 Then Begin
              MessageDlg('El equipo que desea eliminar ha sido utilizado en estimaciones de equipo, no puede eliminarse.', mtInformation, [mbOk], 0);
              lSigue := false;
          End;
      End;
      If lSigue Then
      //Else
      Begin  
          //**************************BRITO 02/12/10****************************
          
          try
               sIdKardex := Equipos.FieldByName('sIdEquipo').AsString;
               Equipos.Delete ;

               //registrar la operacion en el kardex
               opcKardex := 'Borra Equipo';
               Kardex('Otros Movimientos', opcKardex, sIdKardex, 'ID Equipo', '', '', '','Tarifa Diaria','Catalogo de Equipos' );
          except
              on e : exception do begin
                  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Equipos', 'Al eliminar registro', 0);
              end;
          end;
      End;
    end;
end;

procedure TfrmEquipos.frmBarra1btnRefreshClick(Sender: TObject);
begin
 try
  Equipos.refresh ;
  TiposdeEquipo.refresh
 except
 on e : exception do begin
 UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_equipos', 'Al actualizar registro', 0);
 end;
 end;
end;

procedure TfrmEquipos.frmBarra1btnExitClick(Sender: TObject);
begin
   frmBarra1.btnExitClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   close
end;

procedure TfrmEquipos.tsIdEquipoKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsDescripcion.SetFocus
end;

procedure TfrmEquipos.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;


procedure TfrmEquipos.tlDistribuyeKeyPress(Sender: TObject; var Key: Char);
begin
 if key = #13 then
    dbedtNumeroSerie.SetFocus
end;


procedure TfrmEquipos.tlImprimeEnter(Sender: TObject);
begin
    tlImprime.Color := global_color_entrada
end;

procedure TfrmEquipos.tlImprimeExit(Sender: TObject);
begin
    tlImprime.Color := global_color_salida
end;

procedure TfrmEquipos.tlImprimeKeyPress(Sender: TObject; var Key: Char);
begin
    if Key = #13 then
        tlCobro.SetFocus
end;

procedure TfrmEquipos.Imprimir1Click(Sender: TObject);
begin
    frmbarra1.btnPrinter.Click;
end;

procedure TfrmEquipos.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure TfrmEquipos.Paste1Click(Sender: TObject);
begin
try
utGrid.AddRowsFromClip;
except
on e : exception do begin
UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_Equipos', 'Al pegar registro', 0);
end;
end;
end;

procedure TfrmEquipos.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click
end;

procedure TfrmEquipos.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click
end;

procedure TfrmEquipos.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click
end;

procedure TfrmEquipos.Copy1Click(Sender: TObject);
begin
  UtGrid.CopyRowsToClip;
end;

procedure TfrmEquipos.dbcbblSumaSolicitadoEnter(Sender: TObject);
begin
    dbcbblSumaSolicitado.Color := global_color_entrada
end;

procedure TfrmEquipos.dbcbblSumaSolicitadoExit(Sender: TObject);
begin
    dbcbblSumaSolicitado.Color := global_color_salida
end;

procedure TfrmEquipos.dbcbblSumaSolicitadoKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key = #13 then
       tCuadra.SetFocus
end;

procedure TfrmEquipos.DBCheckBox1Click(Sender: TObject);
begin
  if (Equipos.State = dsInsert) and (DBCheckBox1.Checked) then
  begin
    if Length(Trim(Equipos.FieldByName('sDescripcionDiesel').AsString)) = 0 then
      Equipos.FieldByName('sDescripcionDiesel').AsString :=  Equipos.FieldByName('sDescripcion').AsString;
  end;
end;

procedure TfrmEquipos.dbedtNumeroSerieEnter(Sender: TObject);
begin
    dbedtNumeroSerie.Color := global_color_entrada;
end;

procedure TfrmEquipos.dbedtNumeroSerieExit(Sender: TObject);
begin
    dbedtNumeroSerie.Color := global_color_salida;
end;

procedure TfrmEquipos.dbedtNumeroSerieKeyPress(Sender: TObject; var Key: Char);
begin
   if key = #13 then
      dbedtsDescripcionDiesel.SetFocus
end;

procedure TfrmEquipos.dbedtsDescripcionDieselEnter(Sender: TObject);
begin
   dbedtsDescripcionDiesel.Color := global_color_entrada;
end;

procedure TfrmEquipos.dbedtsDescripcionDieselExit(Sender: TObject);
begin
    dbedtsDescripcionDiesel.Color := global_color_salida;
    if Length(Trim(Equipos.FieldByName('sDescripcion').AsString)) = 0 then
      Equipos.FieldByName('sDescripcion').AsString :=  Equipos.FieldByName('sDescripcionDiesel').AsString;
end;

procedure TfrmEquipos.dbedtsDescripcionDieselKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key = #13 then
      dbcbblSumaSolicitado.SetFocus
end;

procedure TfrmEquipos.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click
end;

procedure TfrmEquipos.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmEquipos.tdCantidadKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTdbedit(tdCantidad,key) then
      key:=#0;
  if key = #13 then
    tdFechaInicio.SetFocus 
end;

procedure TfrmEquipos.grid_equiposEnter(Sender: TObject);
begin
    If frmBarra1.btnCancel.Enabled = True Then
      frmbarra1.btnCancel.Click ;
end;

procedure TfrmEquipos.tsIdTipodeEquipoEnter(Sender: TObject);
begin
    tsIdTipodeEquipo.Color := global_color_entrada
end;

procedure TfrmEquipos.tsIdEquipoEnter(Sender: TObject);
begin
    tsIdEquipo.Color := global_color_entrada
end;

procedure TfrmEquipos.tsIdEquipoExit(Sender: TObject);
begin
    tsIdEquipo.Color := global_color_salida
end;

procedure TfrmEquipos.tsDescripcionEnter(Sender: TObject);
begin
    tsDescripcion.Color := global_color_entrada
end;

procedure TfrmEquipos.tsDescripcionExit(Sender: TObject);
begin
    tsDescripcion.Color := global_color_salida
end;

procedure TfrmEquipos.tsMedidaEnter(Sender: TObject);
begin
    tsMedida.Color := global_color_entrada
end;

procedure TfrmEquipos.tsMedidaExit(Sender: TObject);
begin
    tsMedida.Color := global_color_salida
end;

procedure TfrmEquipos.tCuadraEnter(Sender: TObject);
begin
    tCuadra.Color := global_color_entrada
end;

procedure TfrmEquipos.tCuadraExit(Sender: TObject);
begin
    tCuadra.Color := global_color_salida
end;

procedure TfrmEquipos.tCuadraKeyPress(Sender: TObject; var Key: Char);
begin
   if key = #13 then
      tiAnno.SetFocus;
end;

procedure TfrmEquipos.tdCantidadChange(Sender: TObject);
begin
tdbeditchangef(tdCantidad,'Cantidad');
end;

procedure TfrmEquipos.tdCantidadEnter(Sender: TObject);
begin
    tdCantidad.Color := global_color_entrada
end;

procedure TfrmEquipos.tdCantidadExit(Sender: TObject);
begin
    tdCantidad.Color := global_color_salida
end;

procedure TfrmEquipos.tdCostoMnChange(Sender: TObject);
begin
  tdbeditchangef(tdCostoMn, 'Costo MN');
end;

procedure TfrmEquipos.tdCostoMnEnter(Sender: TObject);
begin
    tdCostoMN.Color := global_color_entrada
end;

procedure TfrmEquipos.tdCostoMnExit(Sender: TObject);
begin
    tdCostoMN.Color := global_color_salida
end;

procedure TfrmEquipos.tdCostoDllChange(Sender: TObject);
begin
  tdbeditchangef(tdCostoDll,'Costo DLL');
end;

procedure TfrmEquipos.tdCostoDllEnter(Sender: TObject);
begin
    tdCostoDll.Color := global_color_entrada
end;

procedure TfrmEquipos.tdCostoDllExit(Sender: TObject);
begin
    tdCostoDll.Color := global_color_salida
end;

procedure TfrmEquipos.tlProrrateoEnter(Sender: TObject);
begin
    tlProrrateo.Color := global_color_entrada
end;

procedure TfrmEquipos.tlProrrateoExit(Sender: TObject);
begin
    tlProrrateo.Color := global_color_salida
end;

procedure TfrmEquipos.tiJornadaChange(Sender: TObject);
begin
tdbeditchangei(tijornada,'Jornada');
end;

procedure TfrmEquipos.tiJornadaEnter(Sender: TObject);
begin
    tiJornada.Color := global_color_entrada
end;

procedure TfrmEquipos.tiJornadaExit(Sender: TObject);
begin
    tiJornada.Color := global_color_salida
end;

procedure TfrmEquipos.tlCobroEnter(Sender: TObject);
begin
    tlCobro.Color := global_color_entrada
end;

procedure TfrmEquipos.tlCobroExit(Sender: TObject);
begin
    tlCobro.Color := global_color_salida
end;

procedure TfrmEquipos.tlCobroKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tlDistribuye.SetFocus
end;

procedure TfrmEquipos.tlDistribuyeEnter(Sender: TObject);
begin
    tlDistribuye.Color := global_color_entrada
end;

procedure TfrmEquipos.tlDistribuyeExit(Sender: TObject);
begin
    tlDistribuye.Color := global_color_salida
end;

procedure TfrmEquipos.tdVentaDLLKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTdbedit(tdVentaDLL,key) then
      key:=#0;
  if key = #13 then
    tlProrrateo.SetFocus
end;

procedure TfrmEquipos.tdVentaMNChange(Sender: TObject);
begin
  tdbeditchangef(tdVentaMN, 'Precio MN');
end;

procedure TfrmEquipos.tdVentaMNEnter(Sender: TObject);
begin
    tdVentaMN.Color := global_color_entrada
end;

procedure TfrmEquipos.tdVentaMNExit(Sender: TObject);
begin
    tdVentaMN.Color := global_color_salida
end;

procedure TfrmEquipos.tdVentaMNKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTdbedit(tdVentaMN,key) then
      key:=#0;
    If Key = #13 Then
        tdVentaDLL.SetFocus 
end;

procedure TfrmEquipos.tdVentaDLLChange(Sender: TObject);
begin
  tdbeditchangef(tdVentaDLL,'Precio DLL');
end;

procedure TfrmEquipos.tdVentaDLLEnter(Sender: TObject);
begin
    tdVentaDLL.Color := global_color_entrada
end;

procedure TfrmEquipos.tdVentaDLLExit(Sender: TObject);
begin
    tdVentaDLL.Color := global_color_salida
end;

procedure TfrmEquipos.frmBarra1btnPrinterClick(Sender: TObject);
begin
    if Equipos.RecordCount > 0 then
    begin
        frxEquipos.PreviewOptions.MDIChild := False ;
        frxEquipos.PreviewOptions.Modal := True ;
        frxEquipos.PreviewOptions.Maximized := lCheckMaximized () ;
        frxEquipos.PreviewOptions.ShowCaptions := False ;
        frxEquipos.Previewoptions.ZoomMode := zmPageWidth ;
        frxEquipos.LoadFromFile(Global_Files+'CatalogoEquipos.fr3') ;
        //<ROJAS>
        frxEquipos.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP)) ;
        //
     end
     else
        messageDLG('No se encontro informacion a Imprimir', mtInformation, [mbOk], 0);
end;

procedure TfrmEquipos.tiItemOrdenEnter(Sender: TObject);
begin
    tiItemOrden.Color := global_color_Entrada    
end;

procedure TfrmEquipos.tiItemOrdenExit(Sender: TObject);
begin
    tiItemOrden.Color := global_color_salida
end;



procedure TfrmEquipos.tiItemOrdenKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsIdTipodeEquipo.SetFocus 
end;

procedure TfrmEquipos.tdFechaFinalKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdCostoMN.SetFocus
end;

procedure TfrmEquipos.tdFechaInicioChange(Sender: TObject);
begin
  if tdFechaFinal.Date<tdFechaInicio.Date then
   tdFechaFinal.Date:=tdFechainicio.Date;
end;

procedure TfrmEquipos.tdFechaInicioEnter(Sender: TObject);
begin
    tdFechaInicio.Color := global_color_entrada
end;

procedure TfrmEquipos.tdFechaInicioExit(Sender: TObject);
begin
    tdFechaInicio.Color := global_color_salida
end;


procedure TfrmEquipos.tdFechaInicioKeyPress(Sender: TObject;
  var Key: Char);
begin

   If Key = #13 Then
    tdFechaFinal.SetFocus
end;

procedure TfrmEquipos.tdFechaFinalChange(Sender: TObject);
begin
 // tdFechaFinal.MinDate:=tdFechainicio.Date;
end;

procedure TfrmEquipos.tdFechaFinalEnter(Sender: TObject);
begin
    tdFechaFinal.Color := global_color_entrada
end;

procedure TfrmEquipos.tdFechaFinalExit(Sender: TObject);
begin
  tdFechaFinal.Color := global_color_salida
end;

procedure TfrmEquipos.EquiposAfterScroll(DataSet: TDataSet);
begin
  if Equipos.State=dsbrowse then
  begin
    If frmBarra1.btnCancel.Enabled = False Then
    Begin
        DistribuciondeEquipos.Active := False ;
        DistribuciondeEquipos.Params.ParamByName('Contrato').DataType := ftString ;
        DistribuciondeEquipos.Params.ParamByName('Contrato').Value := global_contrato ;
        DistribuciondeEquipos.Params.ParamByName('Equipo').DataType := ftString ;
        DistribuciondeEquipos.Params.ParamByName('Equipo').Value := Equipos.FieldByName('sIdEquipo').AsString  ;
        DistribuciondeEquipos.Open ;

        tiAnno.Value := YearOf(Equipos.FieldByName('dFechaInicio').AsDateTime) ;
        tsMes.ItemIndex := MonthOf(Equipos.FieldByName('dFechaInicio').AsDateTime)  ;
        tdCantidadMensual.Value := 0 ;

        DistribuciondeEquipos.Active := False ;
        DistribuciondeEquipos.Params.ParamByName('Contrato').DataType := ftString ;
        DistribuciondeEquipos.Params.ParamByName('Contrato').Value := global_contrato ;
        DistribuciondeEquipos.Params.ParamByName('Equipo').DataType := ftString ;
        DistribuciondeEquipos.Params.ParamByName('Equipo').Value := Equipos.FieldByName('sIdEquipo').AsString ;
        DistribuciondeEquipos.Open ;
    End
  end;
  if Equipos.RecordCount > 0 then
    CbxUnificar.Checked := length(trim(Equipos.FieldByName('iUnificador').AsString)) <> 0;
  
end;

procedure TfrmEquipos.EquiposBeforePost(DataSet: TDataSet);
begin
 // zeoverififloat(tdcantidad,'cantidad',equipos,'dcantidad');
 PCAbsoluto(Equipos,'iItemOrden');
 PCAbsoluto(Equipos,'iJornada');
end;

procedure TfrmEquipos.EquiposdCantidadSetText(Sender: TField;
  const Text: string);
begin
  sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmEquipos.EquiposdCostoDLLSetText(Sender: TField;
  const Text: string);
begin
  sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmEquipos.EquiposdCostoMNSetText(Sender: TField;
  const Text: string);
begin
  sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmEquipos.EquiposdVentaDLLSetText(Sender: TField;
  const Text: string);
begin
 Sender.Value:=(strtofloatdef(text,0));
end;

procedure TfrmEquipos.EquiposdVentaMNSetText(Sender: TField;
  const Text: string);
begin
  sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmEquipos.EquiposiJornadaSetText(Sender: TField;
  const Text: string);
begin
  Sender.Value:=(strtointdef(text,0));
end;

procedure TfrmEquipos.EquiposPostError(DataSet: TDataSet; E: EDatabaseError;
  var Action: TDataAction);
begin
showmessage('error al enviar');
end;

procedure TfrmEquipos.ExportaaPlantillaExcel1Click(Sender: TObject);
Var
  CadError, OrdenVigencia: String;
//////////////////////////////////// PLANTILAS DE IMPORTACION //////////////////
Function GenerarPlantilla: Boolean;
Var
  Resultado: Boolean;

Procedure DatosPlantilla;
Var
  CadFecha, tmpNombre, cadena : String;
  fs: tStream;
  Alto : Extended;
  Ren, nivel : integer;
Begin
    Ren := 2;
  // Realizar los ajustes visuales y de formato de hoja
    Excel.ActiveWindow.Zoom := 100;
//  if rAnexoC.Checked then
//  begin
      Excel.Columns['A:A'].ColumnWidth := 20;
      Excel.Columns['B:B'].ColumnWidth := 10;
      Excel.Columns['C:C'].ColumnWidth := 15;
      Excel.Columns['D:D'].ColumnWidth := 40;
      Excel.Columns['E:M'].ColumnWidth := 12;


      // Colocar los encabezados de la plantilla...
      Hoja.Range['A1:A1'].Select;
      Excel.Selection.Value := 'Contrato';
      FormatoEncabezado;
      Hoja.Range['B1:B1'].Select;
      Excel.Selection.Value := 'Id_Equipo';
      FormatoEncabezado;
      Hoja.Range['C1:C1'].Select;
      Excel.Selection.Value := 'Ordenamiento';
      FormatoEncabezado;
      Hoja.Range['D1:D1'].Select;
      Excel.Selection.Value := 'Descripcion';
      FormatoEncabezado;
      Hoja.Range['E1:E1'].Select;
      Excel.Selection.Value := 'Medida';
      FormatoEncabezado;
      Hoja.Range['F1:F1'].Select;
      Excel.Selection.Value := 'Cantidad';
      FormatoEncabezado;
      Hoja.Range['G1:G1'].Select;
      Excel.Selection.Value := 'Costo MN';
      FormatoEncabezado;
      Hoja.Range['H1:H1'].Select;
      Excel.Selection.Value := 'Costo DLL';
      FormatoEncabezado;
      Hoja.Range['I1:I1'].Select;
      Excel.Selection.Value := 'Venta MN';
      FormatoEncabezado;
      Hoja.Range['J1:J1'].Select;
      Excel.Selection.Value := 'Venta DLL';
      FormatoEncabezado;
      Hoja.Range['K1:K1'].Select;
      Excel.Selection.Value := 'Fecha_Inicio';
      FormatoEncabezado;
      Hoja.Range['L1:L1'].Select;
      Excel.Selection.Value := 'Fecha_Final';
      FormatoEncabezado;
      Hoja.Range['M1:M1'].Select;
      Excel.Selection.Value := 'sIdTipoEquipo';
      FormatoEncabezado;
      Hoja.Range['N1:N1'].Select;
      Excel.Selection.Value := 'Jornada';
      FormatoEncabezado;

      connection.QryBusca.Active := False ;
      Connection.QryBusca.Filtered := False;
      connection.QryBusca.SQL.Clear ;
      connection.QryBusca.SQL.Add('select * from equipos where sContrato =:Contrato order by iItemOrden');
      connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      connection.QryBusca.Params.ParamByName('Contrato').Value    := global_contrato ;
      connection.QryBusca.Open ;

      if connection.QryBusca.RecordCount > 0 then
      begin
           while not connection.QryBusca.Eof do
           begin
                Hoja.Cells[Ren,1].Select;
                Excel.Selection.Value := global_contrato;
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Font.Size := 11;
                Excel.Selection.Font.Bold := False;
                Excel.Selection.Font.Name := 'Calibri';

                Hoja.Cells[Ren,2].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['sIdEquipo'];
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;

                Hoja.Cells[Ren,3].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['iItemOrden'];

                Hoja.Cells[Ren,4].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['sDescripcion'];
                Alto := Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight;
                Hoja.Cells[Ren,4].Value := '';

                if Alto > 15 then
                   Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := Alto
                Else
                   Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := 15;

                Excel.Selection.Value := connection.QryBusca.FieldValues['sDescripcion'];

                Hoja.Cells[Ren,5].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['sMedida'];
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;

                Hoja.Cells[Ren,6].Select;
                Excel.Selection.NumberFormat := '@';
                Excel.Selection.Value := connection.QryBusca.FieldValues['dCantidad'];
                Excel.Selection.HorizontalAlignment := xlRight;
                Excel.Selection.VerticalAlignment   := xlCenter;

                Hoja.Cells[Ren,7].Select;
                Excel.Selection.NumberFormat := '@';
                Excel.Selection.Value := connection.QryBusca.FieldValues['dCostoMN'];
                Excel.Selection.HorizontalAlignment := xlRight;
                Excel.Selection.VerticalAlignment   := xlCenter;

                Hoja.Cells[Ren,8].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['dCostoDLL'];

                Hoja.Cells[Ren,9].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['dVentaMN'];

                Hoja.Cells[Ren,10].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['dVentaDLL'];

                Hoja.Cells[Ren,11].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['dFechaInicio'];

                Hoja.Cells[Ren,12].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['dFechaFinal'];

                Hoja.Cells[Ren,13].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['sIdTipoEquipo'];

                Hoja.Cells[Ren,14].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['iJornada'];

                connection.QryBusca.Next;
                Inc(Ren);
           end;
      end;
      Hoja.Cells[2,2].Select;


  Hoja.Range['A1:N1'].Select;
  // Formato general de encabezado de datos..
  Excel.Selection.HorizontalAlignment                   := xlCenter;
  Excel.Selection.VerticalAlignment                     := xlCenter;
  Excel.Selection.Interior.ColorIndex := 5;
  Excel.Selection.Font.color          := clWhite;
  Excel.Selection.Interior.Pattern    := xlSolid;

  Hoja.Range['A1:A1'].Select;
End;

Begin
  Resultado := True;
  Try
    Hoja := Libro.Sheets[1];
    Hoja.Select;
    try
       Hoja.Name := 'EQUIPO '+ global_contrato;
    Except
       Hoja.Name := 'EQUIPO '+ global_contrato;
    end;
    DatosPlantilla;
  Except
    on e:exception do
    Begin
      Resultado := False;
      CadError := 'Se ha producido el siguiente error al generar la Plantilla de Equipos' + #10 + #10 + e.Message;
    End;
  End;

  Result := Resultado;
End;

begin
  // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
  If Not SaveDialog1.Execute Then
    Exit;

  // Generar el ambiente de excel
  try
    Excel := CreateOleObject('Excel.Application');
  except
    FreeAndNil(Excel);
    showmessage('No es posible generar el ambiente de EXCEL, informe de esto al administrador del sistema.');
    Exit;
  end;

  Excel.Visible := True;
  Excel.DisplayAlerts := False;
  Excel.ScreenUpdating := True;

  Libro := Excel.Workbooks.Add;    // Crear el libro sobre el que se ha de trabajar

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
    // Grabar el archivo de excel con el nombre dado
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;

      Excel := '';

  if CadError <> '' then
    showmessage(CadError);

end;

procedure TfrmEquipos.Grid_EquiposMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmEquipos.Grid_EquiposMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmEquipos.Grid_EquiposTitleBtnClick(Sender: TObject;
  ACol: Integer; Field: TField);
Var
  sCampo : String ;
begin
  sCampo := Field.FieldName ;
  Equipos.Active := False ;
  Equipos.Params.ParamByName('Contrato').DataType := ftString ;
  Equipos.Params.ParamByName('Contrato').Value := Global_Contrato ;
  Equipos.Params.ParamByName('Orden').DataType := ftString ;
  Equipos.Params.ParamByName('Orden').Value := sCampo ;
  Equipos.Open ;
end;

procedure TfrmEquipos.Grid_EquiposTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmEquipos.DistribuciondeEquiposCalcFields(DataSet: TDataSet);
begin
    Case MonthOf(DistribuciondeEquipos.FieldValues['dIdFecha']) Of
        1 : DistribuciondeEquipossMes.Value := 'ENERO' ;
        2 : DistribuciondeEquipossMes.Value := 'FEBRERO' ;
        3 : DistribuciondeEquipossMes.Value := 'MARZO' ;
        4 : DistribuciondeEquipossMes.Value := 'ABRIL' ;
        5 : DistribuciondeEquipossMes.Value := 'MAYO' ;
        6 : DistribuciondeEquipossMes.Value := 'JUNIO' ;
        7 : DistribuciondeEquipossMes.Value := 'JULIO' ;
        8 : DistribuciondeEquipossMes.Value := 'AGOSTO' ;
        9 : DistribuciondeEquipossMes.Value := 'SEPTIEMBRE' ;
        10 : DistribuciondeEquipossMes.Value := 'OCTUBRE' ;
        11 : DistribuciondeEquipossMes.Value := 'NOVIEMBRE' ;
        12 : DistribuciondeEquipossMes.Value := 'DICIEMBRE' ;
    End ;
    DistribuciondeEquiposiAnno.Value := YearOf (DistribuciondeEquipos.FieldValues['dIdFecha'])
end;

procedure TfrmEquipos.btnDistribuirClick(Sender: TObject);
Var
    sFecha : String ;
    iMes   : Integer ;
begin
    If tsMes.Text = 'ENERO' Then iMes := 1
    Else If tsMes.Text = 'FEBRERO' Then iMes := 2
    Else If tsMes.Text = 'MARZO' Then iMes := 3
    Else If tsMes.Text = 'ABRIL' Then iMes := 4
    Else If tsMes.Text = 'MAYO' Then iMes := 5
    Else If tsMes.Text = 'JUNIO' Then iMes := 6
    Else If tsMes.Text = 'JULIO' Then iMes := 7
    Else If tsMes.Text = 'AGOSTO' Then iMes := 8
    Else If tsMes.Text = 'SEPTIEMBRE' Then iMes := 9
    Else If tsMes.Text = 'OCTUBRE' Then iMes := 10
    Else If tsMes.Text = 'NOVIEMBRE' Then iMes := 11
    Else If tsMes.Text = 'DICIEMBRE' Then iMes := 12 ;
    If iMes < 9 Then
        sFecha := '01/0' + Trim(IntToStr(iMes + 1)) + '/' + tiAnno.Text
    Else
        If iMes < 12 Then
            sFecha := '01/' + Trim(IntTostr(iMes + 1)) + '/' + tiAnno.Text
        Else
        Begin
            tiAnno.Value := tiAnno.Value + 1 ;
            sFecha := '01/01/' + tiAnno.Text
        End ;
    connection.qryBusca.Active := False ;
    connection.qryBusca.SQL.Clear ;
    connection.qryBusca.SQL.Add('Select sContrato From distribuciondeequipos Where ' +
                                'sContrato = :Contrato and sIdEquipo = :Equipo and dIdFecha = :Fecha') ;
    connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
    connection.qryBusca.Params.ParamByName('Equipo').DataType := ftString ;
    connection.qryBusca.Params.ParamByName('Equipo').Value := Equipos.FieldValues['sIdEquipo'] ;
    connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
    connection.qryBusca.Params.ParamByName('Fecha').Value := StrToDate(sFecha) - 1 ;
    connection.qryBusca.Open ;
    If connection.qryBusca.RecordCount > 0 then
    Begin
        try
            //************************BRITO 16/05/2011**************************
            //averiguar la suma de todo menos el registro que se estara editando
            Connection.QryBusca.Active := False ;
            Connection.QryBusca.SQL.Clear ;
            Connection.QryBusca.SQL.Add('Select sum(dCantidad) as suma From distribuciondeequipos ' +
                                        'Where sContrato = :Contrato and sIdEquipo = :Equipo and dIdFecha <> :Fecha') ;
            Connection.QryBusca.ParamByName('Contrato').Value := global_contrato ;
            Connection.QryBusca.ParamByName('Equipo').Value := Equipos.FieldByName('sIdEquipo').AsString ;
            Connection.QryBusca.ParamByName('Fecha').Value := StrToDate(sFecha) - 1 ;
            Connection.QryBusca.Open ;
            if Connection.QryBusca.RecordCount > 0 then
            begin
              if (Connection.QryBusca.FieldByName('suma').AsFloat + tdCantidadMensual.Value)
                > (Equipos.FieldByName('dCantidad').AsFloat) then
              begin
                //no es posible distribuir mas de la cantidad asignada a la categoria de personal
                showmessage('No se puede distribuir más de lo asignado a la categoría de equipo');
                exit;
              end;
            end;
            //************************BRITO 16/05/2011**************************

            connection.zCommand.Active := False ;
            connection.zCommand.SQL.Clear ;
            connection.zCommand.SQL.Add ( 'update distribuciondeequipos SET dCantidad = :Cantidad ' +
                                          'Where sContrato = :Contrato And sIdEquipo = :Equipo And dIdFecha = :Fecha') ;
            connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
            connection.zCommand.Params.ParamByName('Equipo').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Equipo').Value := Equipos.FieldValues ['sIdEquipo'] ;
            connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate ;
            connection.zCommand.Params.ParamByName('Fecha').Value := StrToDate(sFecha) - 1 ;
            connection.zCommand.Params.ParamByName('Cantidad').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Cantidad').Value := tdCantidadMensual.Value ;
            connection.zCommand.ExecSQL () ;
          except
              on e : exception do begin
                  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Equipos', 'Al editar distribucion de equipos', 0);
              end;
          end;
    End
    Else
    Begin
        try
            //************************BRITO 16/05/2011**************************
            //averiguar la suma de todo
            Connection.QryBusca.Active := False ;
            Connection.QryBusca.SQL.Clear ;
            Connection.QryBusca.SQL.Add('Select sum(dCantidad) as suma From distribuciondeequipos ' +
                                        'Where sContrato = :Contrato and sIdEquipo = :Equipo') ;
            Connection.QryBusca.ParamByName('Contrato').Value := global_contrato ;
            Connection.QryBusca.ParamByName('Equipo').Value := Equipos.FieldByName('sIdEquipo').AsString ;
            Connection.QryBusca.Open ;
            if Connection.QryBusca.RecordCount > 0 then
            begin
              if (Connection.QryBusca.FieldByName('suma').AsFloat + tdCantidadMensual.Value)
                > (Equipos.FieldByName('dCantidad').AsFloat) then
              begin
                //no es posible distribuir mas de la cantidad asignada a la categoria de personal
                showmessage('No se puede distribuir más de lo asignado a la categoría de equipo');
                exit;
              end;
            end;
            //************************BRITO 16/05/2011**************************

            connection.zCommand.Active := False ;
            connection.zCommand.SQL.Clear ;
            connection.zCommand.SQL.Add ( 'INSERT INTO distribuciondeequipos (sContrato, sIdEquipo, dIdFecha, dCantidad) ' +
                                          'VALUES (:Contrato, :Equipo, :Fecha, :Cantidad)') ;
            connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
            connection.zCommand.Params.ParamByName('Equipo').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Equipo').Value := Equipos.FieldValues ['sIdEquipo'] ;
            connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate ;
            connection.zCommand.Params.ParamByName('Fecha').Value := StrToDate(sFecha) - 1 ;
            connection.zCommand.Params.ParamByName('Cantidad').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Cantidad').Value := tdCantidadMensual.Value ;
            connection.zCommand.ExecSQL () ;
          except
              on e : exception do begin
                  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Equipos', 'Al agregar en distribucion de equipos', 0);
              end;
          end;
    End ;
    DistribuciondeEquipos.refresh ;
    DistribuciondeEquipos.Last ;
    If tsMes.Text <> 'DICIEMBRE' Then
        tsMes.ItemIndex := tsMes.ItemIndex + 1
    Else
    Begin
        //tiAnno.Value := tiAnno.Value + 1 ;
        tsMes.ItemIndex := 0 ;
    End ;
    tdCantidadMensual.SetFocus
end;

procedure TfrmEquipos.tiAnnoEnter(Sender: TObject);
begin
    tiAnno.Color := global_color_entrada
end;

procedure TfrmEquipos.tiAnnoExit(Sender: TObject);
begin
    tiAnno.Color := global_color_salida
end;

procedure TfrmEquipos.tiAnnoKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsMes.SetFocus
end;

procedure TfrmEquipos.tsMesEnter(Sender: TObject);
begin
    tsMes.Color := global_color_entrada
end;

procedure TfrmEquipos.tsMesExit(Sender: TObject);
begin
    tsMes.Color := global_color_salida
end;

procedure TfrmEquipos.tsMesKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tdCantidadMensual.SetFocus
end;

procedure TfrmEquipos.tdCantidadMensualChange(Sender: TObject);
begin
tCurrenciEditChangef(tdCantidadMensual,'Cantidad Mensual');
end;

procedure TfrmEquipos.tdCantidadMensualEnter(Sender: TObject);
begin
    tdCantidadMensual.Color := global_color_entrada
end;

procedure TfrmEquipos.tdCantidadMensualExit(Sender: TObject);
begin
    tdCantidadMensual.Color := global_color_salida
end;

procedure TfrmEquipos.tdCantidadMensualKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltrotCurrencyEdit(tdCantidadMensual, key) then
    key:=#0;
    If Key = #13 Then
        btnDistribuir.SetFocus 
end;

procedure TfrmEquipos.AnexoDT41Click(Sender: TObject);
begin
 If Equipos.RecordCount > 0 Then
 begin
      AnexoDT4.Active := False ;
      AnexoDT4.Params.ParamByName('Contrato').DataType := ftString ;
      AnexoDT4.Params.ParamByName('Contrato').Value := global_Contrato ;
      AnexoDT4.Open ;

      if AnexoDT4.RecordCount > 0 then
      begin
        frxAnexoDT4.PreviewOptions.MDIChild := False ;
        frxAnexoDT4.PreviewOptions.Modal := True ;
        frxAnexoDT4.PreviewOptions.Maximized := lCheckMaximized () ;
        frxAnexoDT4.PreviewOptions.ShowCaptions := False ;
        frxAnexoDT4.Previewoptions.ZoomMode := zmPageWidth ;
        frxAnexoDT4.LoadFromFile(Global_Files+'DmoEquipo.fr3') ;
        //<ROJAS>
        frxAnexoDT4.ShowReport;   //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP)) ;
        //
     end
     else
        messageDLG('No se encontro informacion a Imprimir', mtInformation, [mbOk], 0);
 end
 else
     ShowMessage('No existen registros para Generar el anexo DME');
end;

procedure TfrmEquipos.ActualizacindeCostos1Click(Sender: TObject);
begin
    If MessageDlg('Desea actualizar el costo de todo el equipo existente en reportes diarios?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
         connection.zCommand.Active := False ;
         connection.zCommand.SQL.Clear ;
         connection.zCommand.SQL.Add ( 'UPDATE bitacoradeequipos b, Equipos p SET b.dCostoMN = p.dCostoMN, b.dCostoDLL = p.dCostoDLL Where b.sContrato = p.sContrato And b.sIdEquipo = p.sIdEquipo And p.sContrato = :Contrato') ;
         connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
         connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
         connection.zCommand.ExecSQL () ;
         messageDLG('Proceso Terminado con Exito!', mtInformation, [mbOk], 0);
    End;
end;

procedure TfrmEquipos.rxDistribucionKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    If DistribuciondeEquipos.RecordCount > 0 Then
    Begin
        tsMes.ItemIndex := MonthOf(DistribuciondeEquipos.FieldValues['dIdFecha']) - 1 ;
        tiAnno.Value := DistribuciondeEquiposiAnno.Value ;
        tdCantidadMensual.Value := DistribuciondeEquipos.FieldValues['dCantidad'] ;
    End
    Else
    Begin
        tsMes.ItemIndex := MonthOf(Equipos.FieldValues['dFechaInicio'])  ;
        tiAnno.Value := YearOf(Equipos.FieldValues['dFechaInicio']) ;
        tdCantidadMensual.Value := 0  ;
    End
end;

procedure TfrmEquipos.rxDistribucionKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    If DistribuciondeEquipos.RecordCount > 0 Then
    Begin
        tsMes.ItemIndex := MonthOf(DistribuciondeEquipos.FieldValues['dIdFecha']) - 1 ;
        tiAnno.Value := DistribuciondeEquiposiAnno.Value ;
        tdCantidadMensual.Value := DistribuciondeEquipos.FieldValues['dCantidad'] ;
    End
    Else
    Begin
        tsMes.ItemIndex := MonthOf(Equipos.FieldValues['dFechaInicio'])  ;
        tiAnno.Value := YearOf(Equipos.FieldValues['dFechaInicio']) ;
        tdCantidadMensual.Value := 0  ;
    End
end;

procedure TfrmEquipos.formatoEncabezado;
begin
      Excel.Selection.MergeCells := False;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment := xlCenter;
      Excel.Selection.Font.Size := 12;
      Excel.Selection.Font.Bold := False;
      Excel.Selection.Font.Name := 'Calibri';
end;

end.
