unit frm_EntradaAlmacen;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, DB, DBCtrls, global, 
  Mask, OleCtrls, Grids, DBGrids, frm_barra, ExtCtrls, Utilerias,
  Menus, frxClass, frxDBSet, RXDBCtrl,  RxLookup, 
  RXCtrls, CheckLst, RxMemDS, ZAbstractRODataset, ZDataset,
  rxCurrEdit, rxToolEdit, AdvGlowButton,
  ZAbstractDataset, udbgrid, unitexcepciones, unittbotonespermisos, unitactivapop,
  UFunctionsGHH, jpeg;
  function IsDate(ADate: string): Boolean;
type
  TfrmEntradaAlmacen = class(TForm)
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
    ds_proveedores: TDataSource;
    ds_FolioCompra: TDataSource;
    frxEntrada: TfrxReport;
    Pedido: TZReadOnlyQuery;
    Proveedores: TZReadOnlyQuery;
    FolioCompra: TZReadOnlyQuery;
    Label1: TLabel;
    Almacen: TZReadOnlyQuery;
    ds_almacen: TDataSource;
    tsAlmacen: TDBLookupComboBox;
    ds_pEntradas: TDataSource;
    pEntradas: TZReadOnlyQuery;
    ds_familia: TDataSource;
    Familia: TZReadOnlyQuery;
    Reporte: TZReadOnlyQuery;
    PedidosContrato: TStringField;
    PedidoiFolioPedido: TIntegerField;
    PedidosIdInsumo: TStringField;
    PedidosMedida: TStringField;
    PedidodCantidad: TFloatField;
    PedidodCosto: TFloatField;
    PedidosNumeroActividad: TStringField;
    PedidosNumeroOrden: TStringField;
    PedidosStatus: TStringField;
    GroupBox3: TGroupBox;
    frmBarra2: TfrmBarra;
    PgControl: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    tsPlataforma: TLabel;
    imgNotas: TImage;
    frmBarra1: TfrmBarra;
    Grid_Entradas: TRxDBGrid;
    Label7: TLabel;
    tdFechaAviso: TDateTimePicker;
    Label16: TLabel;
    Label17: TLabel;
    tsNumeroOrden: TComboBox;
    tmComentarios: TMemo;
    Label6: TLabel;
    Label2: TLabel;
    tsInsumo: TEdit;
    mDescripcion: TMemo;
    Label8: TLabel;
    Label5: TLabel;
    tsFamilia: TRxDBLookupCombo;
    Label14: TLabel;
    tdCantidad: TRxCalcEdit;
    mComentarios: TMemo;
    Label9: TLabel;
    GridPartidas: TRxDBGrid;
    Grid_Pedido: TRxDBGrid;
    lblEncabezado: TStaticText;
    Agregar: TAdvGlowButton;
    Editar: TAdvGlowButton;
    Salvar: TAdvGlowButton;
    Cancelar: TAdvGlowButton;
    Eliminar: TAdvGlowButton;
    Imprimir: TAdvGlowButton;
    PedidoDescripcion: TStringField;
    SeguimientoMaterialGeneral1: TMenuItem;
    SeguimientoMaterialxPartida1: TMenuItem;
    SeguimientoMaterialxPartidaDetalle1: TMenuItem;
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
    pEntradassContrato: TStringField;
    pEntradasiItem: TIntegerField;
    pEntradassIdInsumo: TStringField;
    pEntradasdFechaEntrega: TDateField;
    pEntradasdCantidad: TFloatField;
    pEntradasdPrecio: TFloatField;
    pEntradasdNuevoPrecio: TFloatField;
    pEntradasdCantidadAnterior: TFloatField;
    pEntradassIdAlmacen: TStringField;
    pEntradassNumeroActividad: TStringField;
    pEntradassIdUsuario: TStringField;
    pEntradasmComentarios: TMemoField;
    pEntradassUbicacion: TStringField;
    pEntradassIdFamilia: TStringField;
    pEntradasAnterior: TFloatField;
    pEntradasdExistencia: TFloatField;
    pEntradasdVentaMN: TFloatField;
    pEntradasmDescripcion: TStringField;
    pEntradasdPendiente: TFloatField;
    pEntradasiFolioEntrada: TIntegerField;
    Label3: TLabel;
    tsTipomovimiento: TDBLookupComboBox;
    Label4: TLabel;
    iFolio: TCurrencyEdit;
    ds_tipomovimiento: TDataSource;
    zq_tipomovimiento: TZReadOnlyQuery;
    anexo_suministro: TZQuery;
    tsFolioMovimiento: TDBLookupComboBox;
    ReportesContrato: TStringField;
    ReporteiFolioEntrada: TIntegerField;
    ReporteiFolioMovimiento: TIntegerField;
    ReportesNumeroOrden: TStringField;
    ReportedFecha: TDateField;
    ReportesIdUsuario: TStringField;
    ReportemComentarios: TStringField;
    ReportesIdInsumo: TStringField;
    ReportemDescripcion: TMemoField;
    ReportesMedida: TStringField;
    ReportedCantidad: TFloatField;
    ReportedCosto: TFloatField;
    ReportesStatus: TStringField;
    ReportedDescuento: TFloatField;
    ReportedExistencia: TFloatField;
    ReportedCostoMN: TFloatField;
    Reportealmacen: TStringField;
    Label10: TLabel;
    Image1: TImage;
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
    procedure tdIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure tsOrigenKeyPress(Sender: TObject; var Key: Char);
    procedure tmComentariosEnter(Sender: TObject);
    procedure tmComentariosExit(Sender: TObject);
    procedure tdCantidadEnter(Sender: TObject);
    procedure tdCantidadExit(Sender: TObject);
    procedure tdCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroActividadKeyPress(Sender: TObject; var Key: Char);
    procedure frxEntradaGetValue(const VarName: String;
      var Value: Variant);
    procedure ComentariosAdicionalesClick(Sender: TObject);
    procedure tdFechaAvisoEnter(Sender: TObject);
    procedure tdFechaAvisoExit(Sender: TObject);
    procedure tdFechaAvisoKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure ActivaBotones(Sender :boolean);
    procedure ActivaBotones2(Sender :boolean);
    procedure tsFamiliaEnter(Sender: TObject);
    procedure tsFamiliaExit(Sender: TObject);
    procedure mComentariosEnter(Sender: TObject);
    procedure mComentariosExit(Sender: TObject);
    procedure mComentariosKeyPress(Sender: TObject; var Key: Char);
    procedure tsFamiliaKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure Grid_PedidoKeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra2btnPrinterClick(Sender: TObject);
    procedure ImprimirClick(Sender: TObject);
    procedure tsAlmacenExit(Sender: TObject);
    procedure AgregarClick(Sender: TObject);
    procedure EditarClick(Sender: TObject);
    procedure SalvarClick(Sender: TObject);
    procedure CancelarClick(Sender: TObject);
    procedure EliminarClick(Sender: TObject);
    procedure SeguimientoMaterialGeneral1Click(Sender: TObject);
    procedure Seguimiento_Material(dParamActividad : string);
    procedure SeguimientoMaterialxPartida1Click(Sender: TObject);
    procedure SeguimientoMaterialxPartidaDetalle1Click(Sender: TObject);
    procedure tsAlmacenEnter(Sender: TObject);
    procedure tsInsumoEnter(Sender: TObject);
    procedure tsInsumoExit(Sender: TObject);
    procedure mDescripcionEnter(Sender: TObject);
    procedure mDescripcionExit(Sender: TObject);
    procedure mDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure Grid_EntradasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_EntradasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_EntradasTitleClick(Column: TColumn);
    procedure Grid_PedidoMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_PedidoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_PedidoTitleClick(Column: TColumn);
    procedure GridPartidasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure GridPartidasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure GridPartidasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure tdFechaAvisoChange(Sender: TObject);
    procedure tdIdFechaChange(Sender: TObject);
    procedure tdCantidadChange(Sender: TObject);
    procedure PgControlChange(Sender: TObject);
    procedure anexo_suministroAfterScroll(DataSet: TDataSet);
    procedure dbFolioMovimientoEnter(Sender: TObject);
    procedure dbFolioMovimientoExit(Sender: TObject);
    procedure pEntradasAfterScroll(DataSet: TDataSet);
    procedure pEntradasCalcFields(DataSet: TDataSet);
    procedure PedidoAfterScroll(DataSet: TDataSet);
    procedure tsTipomovimientoExit(Sender: TObject);
    procedure tsFolioMovimientoEnter(Sender: TObject);
    procedure tsTipomovimientoEnter(Sender: TObject);
    procedure tsFolioMovimientoExit(Sender: TObject);
    procedure iFolioKeyPress(Sender: TObject; var Key: Char);
    procedure tsTipomovimientoKeyPress(Sender: TObject; var Key: Char);
    procedure tsFolioMovimientoKeyPress(Sender: TObject; var Key: Char);
  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmEntradaAlmacen: TfrmEntradaAlmacen;
  SavePlace     : TBookmark;
  sDescripcion  : String ;
  txtAux        : String ;
  lNuevo        : Boolean ;
  OpcButton1    : String ;
  FechaEAnt     : String ;
  sBackup,
  IdInsumo      : String ;
  Cantidad      : Double;
  TipoExplosion : string;
   utgrid:ticdbgrid;
   utgrid2:ticdbgrid;
   utgrid3:ticdbgrid;
   botonpermiso:tbotonespermisos;
  BanderaAgregar : Boolean;
implementation

uses frm_connection , frm_comentariosxanexo, UnitValidaTexto;

{$R *.dfm}

procedure TfrmEntradaAlmacen.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  botonpermiso.Free;
  action := cafree ;
  utgrid.Destroy;
  utgrid2.Destroy;
  utgrid3.Destroy;
end;

function IsDate(ADate: string): Boolean;
 var
  Dummy: TDateTime;
begin
  IsDate := TryStrToDate(ADate, Dummy);
end;

procedure TfrmEntradaAlmacen.FormShow(Sender: TObject);
begin
  try
    sMenuP:=stMenu;
    BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'MnuEntAlmace', PopupPrincipal);
    UtGrid:=TicdbGrid.create(grid_entradas);
    UtGrid2:=TicdbGrid.create(grid_pedido);
    UtGrid3:=TicdbGrid.create(gridpartidas);
    ActivaBotones(False);

    Almacen.Active := False;
    Almacen.Open;
    if Almacen.RecordCount > 0 then
       tsAlmacen.KeyValue := Almacen.FieldValues['sIdAlmacen'];
       
    Familia.Active := False;
    Familia.Open;
    if Familia.RecordCount > 0 then
       tsFamilia.KeyValue := Familia.FieldValues['sIdFamilia'];

    tsNumeroOrden.Items.Clear ;
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select sNumeroOrden from ordenesdetrabajo where sContrato = :Contrato and ' +
                                'cIdStatus = :status order by sNumeroOrden') ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := Global_Contrato ;
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
    anexo_suministro.Params.ParamByName('Contrato').Value    := global_contrato ;
    anexo_suministro.Params.ParamByName('Almacen').DataType  := ftString ;
    anexo_suministro.Params.ParamByName('Almacen').Value     := tsAlmacen.KeyValue ;
    anexo_suministro.Open ;

    pEntradas.Active := False ;
    pEntradas.Params.ParamByName('Contrato').DataType := ftString ;
    pEntradas.Params.ParamByName('Contrato').Value    := global_contrato ;
    pEntradas.Params.ParamByName('Folio').DataType    := ftInteger ;
    pEntradas.Params.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioEntrada'] ;
    pEntradas.Params.ParamByName('Almacen').DataType  := ftString ;
    pEntradas.Params.ParamByName('Almacen').Value     := tsAlmacen.KeyValue ;
    pEntradas.Open ;

    zq_tipomovimiento.Active := False;
    zq_tipomovimiento.Open;

    anexo_suministro.Refresh;
    grid_entradas.SetFocus;

    if connection.configuracion.FieldValues['sExplosion'] = 'Recursos por Concepto/Partida' then
       TipoExplosion := 'recursosanexo'
    else
       TipoExplosion := 'recursosanexosnuevos';
    BotonPermiso.permisosBotones(frmBarra1);
    BotonPermiso.permisosBotones(frmBarra2);

  except
  on e : exception do
      begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Entrada de materiales almacen', 'Al iniciar el formulario', 0);
      end;
  end;
end;

procedure TfrmEntradaAlmacen.BtnExitClick(Sender: TObject);
begin
    Close ;
end;

procedure TfrmEntradaAlmacen.frmBarra1btnExitClick(Sender: TObject);
begin
  Insertar1.Enabled := True ;
  Editar1.Enabled := True ;
  Registrar1.Enabled := False ;
  Can1.Enabled := False ;
  Eliminar1.Enabled := True ;
  Refresh1.Enabled := True ;
  frmBarra1.btnExitClick(Sender);
end;

procedure TfrmEntradaAlmacen.Insertar1Click(Sender: TObject);
begin
    frmBarra2.btnAdd.Click
end;

procedure TfrmEntradaAlmacen.Editar1Click(Sender: TObject);
begin
    frmBarra2.btnEdit.Click
end;

procedure TfrmEntradaAlmacen.EditarClick(Sender: TObject);
begin
     If anexo_suministro.RecordCount > 0 Then
     Begin
         Showmessage('No se Pueden editar las entradas.. Se recomienda eliminarlas e Insertar nuevamente. ');
     End;
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmEntradaAlmacen.Registrar1Click(Sender: TObject);
begin
    frmBarra2.btnPost.Click 
end;

procedure TfrmEntradaAlmacen.Can1Click(Sender: TObject);
begin
    frmBarra2.btnCancel.Click
end;

procedure TfrmEntradaAlmacen.CancelarClick(Sender: TObject);
begin
    Agregar.Enabled  := True ;
    Editar.Enabled   := True;
    Salvar.Enabled   := False ;
    Cancelar.Enabled := False ;
    Eliminar.Enabled := True ;
    Imprimir.Enabled := True ;
    ActivaBotones2(False);
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmEntradaAlmacen.Eliminar1Click(Sender: TObject);
begin
    frmBarra2.btnDelete.Click
end;

procedure TfrmEntradaAlmacen.EliminarClick(Sender: TObject);
begin
     If pEntradas.RecordCount > 0 Then
     Begin
//          try
             //soad -> Actualizacion de los insumos de la Orden de Compra
             //***************************************************************

             connection.zCommand.Active := False;
             connection.zCommand.SQL.Clear;

             if (pos('COMPRA',tsTipoMovimiento.Text) > 0) OR (pos('O.C.',tsTipoMovimiento.Text) > 0) then
                 connection.zCommand.SQL.Add('UPDATE anexo_ppedido SET sStatus = "Pendiente" '+
                                         'WHERE sContrato =:Contrato and iFolioPedido =:Folio And sIdInsumo =:Insumo ');

             if (pos('TRASPASO',tsTipoMovimiento.Text) > 0) OR (pos('EMBARQUE',tsTipoMovimiento.Text) > 0) OR (pos('DESEMBARQUE',tsTipoMovimiento.Text) > 0) then
                 connection.zCommand.SQL.Add('UPDATE bitacoradesalida SET sStatus = "Pendiente" '+
                                         'WHERE sContrato =:Contrato and iFolioSalida =:Folio And sIdInsumo =:Insumo ');

             connection.zCommand.ParamByName('Contrato').DataType := ftString;
             connection.zCommand.ParamByName('Contrato').Value    := global_contrato;
             connection.zCommand.ParamByName('Folio').DataType    := ftInteger;
             connection.zCommand.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioMovimiento'];
             connection.zCommand.ParamByName('Insumo').DataType   := ftString ;
             connection.zCommand.ParamByName('Insumo').value      := pEntradas.FieldValues['sIdInsumo'] ;
             connection.zCommand.ExecSQL;
             Pedido.Refresh;

             // Consulta de Insumo antes de Actualizar...
             //*************************************************
             connection.QryBusca.Active := False;
             connection.QryBusca.SQL.Clear;
             connection.QryBusca.SQL.Add('select dExistencia from insumos where sContrato =:Contrato and sIdInsumo =:Insumo and sIdAlmacen =:Almacen ');
             connection.QryBusca.ParamByName('Contrato').DataType := ftString;
             connection.QryBusca.ParamByName('Contrato').Value    := global_contrato;
             connection.QryBusca.ParamByName('Insumo').DataType   := ftString;
             connection.QryBusca.ParamByName('Insumo').Value      := pEntradas.FieldValues['sIdInsumo'] ; ;
             connection.QryBusca.ParamByName('Almacen').DataType  := ftString;
             connection.QryBusca.ParamByName('Almacen').Value     := tsAlmacen.KeyValue ;
             connection.QryBusca.Open;

             //soad -> Actualizacion de los insumos...
             //**************************************************
             connection.QryBusca2.Active := False;
             connection.QryBusca2.SQL.Clear;
             connection.QryBusca2.SQL.Add('UPDATE insumos SET dExistencia =:Cantidad '+
                                          'WHERE sContrato =:Contrato And sIdInsumo =:Insumo and sIdAlmacen =:Almacen ');
             connection.QryBusca2.ParamByName('Contrato').DataType  := ftString;
             connection.QryBusca2.ParamByName('Contrato').Value     := global_contrato;
             connection.QryBusca2.ParamByName('Insumo').DataType    := ftString;
             connection.QryBusca2.ParamByName('Insumo').Value       := pEntradas.FieldValues['sIdInsumo'] ; ;
             connection.QryBusca2.ParamByName('Cantidad').DataType  := ftFloat ;
             connection.QryBusca2.ParamByName('Cantidad').value     := connection.QryBusca.FieldValues['dExistencia'] - pEntradas.FieldValues['dCantidad'];
             connection.QryBusca2.ParamByName('Almacen').DataType   := ftString;
             connection.QryBusca2.ParamByName('Almacen').Value      := tsAlmacen.KeyValue ;
             connection.QryBusca2.ExecSQL;

             //Eliminamos registro....
             connection.zCommand.Active := False ;
             connection.zCommand.SQL.Clear ;
             connection.zCommand.SQL.Add ('Delete from bitacoradeentrada where sContrato = :Contrato ' +
                                          'and iFolioEntrada =:Folio And sIdInsumo =:Insumo ') ;
             connection.zcommand.Params.ParamByName('Contrato').DataType   := ftString ;
             connection.zcommand.Params.ParamByName('Contrato').value      := Global_Contrato ;
             connection.zcommand.Params.ParamByName('Folio').DataType      := ftInteger ;
             connection.zcommand.Params.ParamByName('Folio').value         := pEntradas.FieldValues['iFolioEntrada'] ;
             connection.zcommand.Params.ParamByName('Insumo').DataType     := ftString ;
             connection.zcommand.Params.ParamByName('Insumo').value        := pEntradas.FieldValues['sIdInsumo'] ;
             connection.zCommand.ExecSQL  ;

             SavePlace := pEntradas.GetBookmark ;
             pEntradas.Refresh ;

             Try
                 pEntradas.GotoBookmark(SavePlace);
             Except
             Else
                pEntradas.FreeBookmark(SavePlace);
             End;
//          Except
//               MessageDlg('Ocurrio un error al eliminar el registro.', mtInformation, [mbOk], 0);
//          End
     End
end;

procedure TfrmEntradaAlmacen.Refresh1Click(Sender: TObject);
begin
    frmBarra2.btnRefresh.Click
end;

procedure TfrmEntradaAlmacen.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure TfrmEntradaAlmacen.ImprimirClick(Sender: TObject);
begin
     If pEntradas.RecordCount > 0  Then
     begin
         Reporte.Active := False;
         Reporte.SQL.Clear;
        if (pos('COMPRA',tsTipoMovimiento.Text) > 0) OR (pos('O.C.',tsTipoMovimiento.Text) > 0) then
             Reporte.SQL.Add('select e.sContrato, e.iFolioEntrada, e.iFolioMovimiento, e.sNumeroOrden, e.dFecha, e.sIdUsuario, e.mComentarios, '+
                        'ped.sIdInsumo, ped.mDescripcion, ped.sMedida, ped.dCantidad, ped.dCosto, ped.sStatus, ped.dDescuento, i.dExistencia, i.dCostoMN, alm.sDescripcion as almacen from almacen_entrada e '+
                        'inner join anexo_ppedido ped '+
                        'on (ped.sContrato = e.sContrato and ped.iFolioPedido = e.iFolioMovimiento) '+
                        'inner join almacenes alm '+
                        'on (alm.sIdAlmacen = e.sIdAlmacen) '+
                        'inner join insumos i '+
                        'on (i.sContrato = e.sContrato and i.sIdInsumo = ped.sIdInsumo and i.sIdAlmacen = e.sIdAlmacen) '+
                        'where e.sContrato =:Contrato and e.sIdAlmacen =:Almacen and e.iFolioEntrada =:Folio ');

         if (pos('TRASPASO',tsTipoMovimiento.Text) > 0) OR (pos('EMBARQUE',tsTipoMovimiento.Text) > 0) OR (pos('DESEMBARQUE',tsTipoMovimiento.Text) > 0) then
             Reporte.SQL.Add('select e.sContrato, e.iFolioEntrada, e.iFolioMovimiento, e.sNumeroOrden, e.dFecha, e.sIdUsuario, e.mComentarios, '+
                        'ped.sIdInsumo, i.mDescripcion, i.sMedida, ped.dCantidad, i.dNuevoPrecio as dCosto, ped.sStatus, 0.0 as dDescuento, i.dExistencia, i.dCostoMN, alm.sDescripcion as almacen from almacen_entrada e '+
                        'inner join bitacoradesalida ped '+
                        'on (ped.sContrato = e.sContrato and ped.iFolioSalida = e.iFolioMovimiento) '+
                        'inner join almacenes alm '+
                        'on (alm.sIdAlmacen = e.sIdAlmacen) '+
                        'inner join insumos i '+
                        'on (i.sContrato = e.sContrato and i.sIdInsumo = ped.sIdInsumo and i.sIdAlmacen = e.sIdAlmacen) '+
                        'where e.sContrato =:Contrato and e.sIdAlmacen =:Almacen and e.iFolioEntrada =:Folio ');
         Reporte.ParamByName('Contrato').DataType := ftString ;
         Reporte.ParamByName('Contrato').Value    := global_contrato ;
         Reporte.ParamByName('Almacen').DataType  := ftString;
         Reporte.ParamByName('Almacen').Value     := tsAlmacen.KeyValue;
         Reporte.ParamByName('Folio').DataType    := ftInteger;
         Reporte.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioEntrada'];

         Reporte.ExecSQL;

         frxEntrada.PreviewOptions.MDIChild := False ;
         frxEntrada.PreviewOptions.Modal := True ;
         frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
         frxEntrada.PreviewOptions.ShowCaptions := False ;
         frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
         frxEntrada.LoadFromFile (global_files + 'Entrada2.fr3') ;
         frxEntrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
     end
     else
        showmessage('No han entrado Materiales al Almacen ');

end;

procedure TfrmEntradaAlmacen.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

procedure TfrmEntradaAlmacen.SalvarClick(Sender: TObject);
var
   CantidadPedido, CantidadEntrada, total : double;
   Nombres,Cadenas:TStringList;
  I: Integer;
begin
      nombres:=TStringList.Create;cadenas:=TStringList.Create;
      nombres.Add('Descripción');nombres.Add('Familia');
      //nombres.Add('Comentarios');

      cadenas.Add(mDescripcion.Text);cadenas.Add(tsFamilia.Text);
      //cadenas.Add(mComentarios.Text);

      if not validaTexto(nombres, cadenas, 'Id Insumo', tsInsumo.Text) then
      begin
          MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
          exit;
      end;

        //Antes de actualizar el Status.. veirifcar si esta todo el material comprado dentro del alamacen..
        connection.QryBusca2.Active := False;
        connection.QryBusca2.SQL.Clear;

        if (pos('COMPRA',tsTipoMovimiento.Text) > 0) OR (pos('O.C.',tsTipoMovimiento.Text) > 0) then
            connection.QryBusca2.SQL.Add('select dCantidad from anexo_ppedido '+
                                         'WHERE sContrato =:Contrato and iFolioPedido =:Folio And sIdInsumo =:Insumo and sStatus = "Pendiente" ');

        if (pos('TRASPASO',tsTipoMovimiento.Text) > 0) OR (pos('EMBARQUE',tsTipoMovimiento.Text) > 0) OR (pos('DESEMBARQUE',tsTipoMovimiento.Text) > 0) then
            connection.QryBusca2.SQL.Add('select dCantidad from bitacoradesalida '+
                                         'WHERE sContrato =:Contrato and iFolioSalida =:Folio And sIdInsumo =:Insumo and sStatus = "Pendiente" ');

        connection.QryBusca2.ParamByName('Contrato').DataType := ftString;
        connection.QryBusca2.ParamByName('Contrato').Value    := global_contrato;
        connection.QryBusca2.ParamByName('Folio').DataType    := ftInteger;
        connection.QryBusca2.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioMovimiento'];
        connection.QryBusca2.ParamByName('Insumo').DataType   := ftString ;
        connection.QryBusca2.ParamByName('Insumo').value      := pedido.FieldValues['sIdInsumo'];
        connection.QryBusca2.Open;

        CantidadPedido := 0;
        if connection.QryBusca2.RecordCount > 0 then
           CantidadPedido := connection.QryBusca2.FieldValues['dCantidad'];

        //Verificamos cuanto entro al almacen.. y lo que resta por entrar..
        connection.QryBusca2.Active := False;
        connection.QryBusca2.SQL.Clear;
        connection.QryBusca2.SQL.Add('select dCantidad from bitacoradeentrada '+
                                     'WHERE sContrato =:Contrato and iFolioEntrada =:Folio And sIdInsumo =:Insumo ');
        connection.QryBusca2.ParamByName('Contrato').DataType := ftString;
        connection.QryBusca2.ParamByName('Contrato').Value    := global_contrato;
        connection.QryBusca2.ParamByName('Folio').DataType    := ftInteger;
        connection.QryBusca2.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioEntrada'];
        connection.QryBusca2.ParamByName('Insumo').DataType   := ftString ;
        connection.QryBusca2.ParamByName('Insumo').value      := pedido.FieldValues['sIdInsumo'];
        connection.QryBusca2.Open;

        CantidadEntrada := 0;
        if connection.QryBusca2.RecordCount > 0 then
           CantidadEntrada := connection.QryBusca2.FieldValues['dCantidad'];

        if (tdCantidad.Value + CantidadEntrada) > CantidadPedido then
        begin
             messageDLG('No se puede Recibir mas Material que lo Comprado, Favor de Verificar.', mtInformation, [mbOk], 0);
             exit;
        end;

        If OpcButton = 'New' then
        Begin
                // Consulta de Insumo antes de insertar.
                //*************************************************
                connection.QryBusca.Active := False;
                connection.QryBusca.SQL.Clear;
                connection.QryBusca.SQL.Add('select dCostoMN, dExistencia from insumos where sContrato =:Contrato and sIdInsumo =:Insumo and sIdAlmacen =:Almacen ');
                connection.QryBusca.ParamByName('Contrato').DataType := ftString;
                connection.QryBusca.ParamByName('Contrato').Value    := global_contrato;
                connection.QryBusca.ParamByName('Insumo').DataType   := ftString;
                connection.QryBusca.ParamByName('Insumo').Value      := pedido.FieldValues['sIdInsumo'];
                connection.QryBusca.ParamByName('Almacen').DataType  := ftString;
                connection.QryBusca.ParamByName('Almacen').Value     := tsAlmacen.KeyValue ;
                connection.QryBusca.Open;

                if connection.QryBusca.RecordCount = 0 then
                begin
                     messageDLG('No se encontro el Insumo '+tsInsumo.Text+' en el Almacen '+ tsAlmacen.KeyValue, mtInformation, [mbOk], 0);
                     exit;
                end;

                try
                  // soad -> Inbsercion de los datos en la bitacora de Entrada....
                  //****************************************************************
                  if CantidadEntrada = 0 then
                  begin
                      connection.zCommand.Active := False ;
                      connection.zCommand.SQL.Clear ;
                      connection.zCommand.SQL.Add ( 'INSERT INTO bitacoradeentrada ( sContrato, iItem, iFolioEntrada, sIdInsumo, dFechaEntrega, dCantidad, dPrecio, dNuevoPrecio, dCantidadAnterior, sIdAlmacen, sNumeroActividad, sIdUsuario, mComentarios, sUbicacion, sIdFamilia ) ' +
                                                    'VALUES (:Contrato, :Item, :Folio, :Insumo, :FechaE, :Cantidad, :Precio, :NvoPrecio, :CantidadAnt, :IdAlmacen, :Actividad, :Usuario, :Comentario, :Ubicacion, :Familia  )') ;
                      connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
                      connection.zCommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                      connection.zCommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                      connection.zCommand.Params.ParamByName('Folio').value          := anexo_suministro.FieldValues['iFolioEntrada'];
                      connection.zCommand.Params.ParamByName('Insumo').DataType      := ftString ;
                      connection.zCommand.Params.ParamByName('Insumo').value         := pedido.FieldValues['sIdInsumo'];
                      connection.zCommand.Params.ParamByName('Item').DataType        := ftInteger ;
                      connection.zCommand.Params.ParamByName('Item').value           := 0;
                      connection.zCommand.Params.ParamByName('FechaE').DataType      := ftDate ;
                      connection.zCommand.Params.ParamByName('FechaE').value         := anexo_suministro.FieldValues['dFecha'];
                      connection.zCommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
                      connection.zCommand.Params.ParamByName('Cantidad').value       := tdCantidad.Value ;
                      connection.zCommand.Params.ParamByName('Precio').DataType      := ftFloat ;
                      connection.zCommand.Params.ParamByName('Precio').value         := pedido.FieldValues['dCosto'];
                      connection.zCommand.Params.ParamByName('NvoPrecio').DataType   := ftFloat ;
                      connection.zCommand.Params.ParamByName('CantidadAnt').DataType := ftFloat ;
                      connection.zCommand.Params.ParamByName('CantidadAnt').value    := tdCantidad.Value ;
                      if connection.QryBusca.RecordCount > 0 then
                      begin
                           if connection.QryBusca.FieldValues['dCostoMN'] <> pedido.FieldValues['dCosto'] then
                              connection.zCommand.Params.ParamByName('NvoPrecio').value  := pedido.FieldValues['dCosto']
                          else
                              connection.zCommand.Params.ParamByName('NvoPrecio').value  := connection.QryBusca.FieldValues['dCostoMN'] ;
                          connection.zCommand.Params.ParamByName('CantidadAnt').value     := connection.QryBusca.FieldValues['dExistencia'] ;
                      end;
                      connection.zCommand.Params.ParamByName('IdAlmacen').DataType   := ftString ;
                      connection.zCommand.Params.ParamByName('IdAlmacen').value      := tsAlmacen.KeyValue;
                      connection.zCommand.Params.ParamByName('Usuario').DataType     := ftString ;
                      connection.zCommand.Params.ParamByName('Usuario').value        := anexo_suministro.FieldValues['sIdUsuario'];
                      connection.zCommand.Params.ParamByName('Comentario').DataType  := ftString ;
                      connection.zCommand.Params.ParamByName('Comentario').value     := mComentarios.Text;
                      connection.zCommand.Params.ParamByName('Actividad').DataType   := ftString ;
                      connection.zCommand.Params.ParamByName('Actividad').value      := pedido.FieldValues['sNumeroActividad'];
                      connection.zCommand.Params.ParamByName('Familia').DataType     := ftString ;
                      connection.zCommand.Params.ParamByName('Familia').value        := tsFamilia.KeyValue;
                      connection.zCommand.ExecSQL ;
                  end
                  else
                  begin
                      connection.zCommand.Active := False ;
                      connection.zCommand.SQL.Clear ;
                      connection.zCommand.SQL.Add ( 'Update bitacoradeentrada set dCantidad =:Cantidad, dCantidadAnterior =:Anterior ' +
                                                    'where sContrato =:Contrato and iFolioEntrada =:Folio and sIdInsumo =:Insumo ') ;
                      connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
                      connection.zCommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
                      connection.zCommand.Params.ParamByName('Folio').DataType       := ftInteger ;
                      connection.zCommand.Params.ParamByName('Folio').value          := anexo_suministro.FieldValues['iFolioEntrada'];
                      connection.zCommand.Params.ParamByName('Insumo').DataType      := ftString ;
                      connection.zCommand.Params.ParamByName('Insumo').value         := pedido.FieldValues['sIdInsumo'];
                      connection.zCommand.Params.ParamByName('Cantidad').DataType    := ftFloat ;
                      connection.zCommand.Params.ParamByName('Cantidad').value       := CantidadEntrada + tdCantidad.Value ;
                      connection.zCommand.Params.ParamByName('Anterior').DataType    := ftFloat ;
                      connection.zCommand.Params.ParamByName('Anterior').value       := CantidadEntrada - tdCantidad.Value ;
                      connection.zCommand.ExecSQL ;
                  end;
                except
                     begin
                         MessageDlg('Ocurrio un error al actualizar el registro', mtWarning, [mbOk], 0);
                         exit;
                     end;
                End;
                //soad -> Actualizacion de los insumos...
                //**************************************************
                connection.QryBusca2.Active := False;
                connection.QryBusca2.SQL.Clear;
                connection.QryBusca2.SQL.Add('UPDATE insumos SET dExistencia =:Cantidad '+
                                            'WHERE sContrato =:Contrato And sIdInsumo =:Insumo and sIdAlmacen =:Almacen ');
                connection.QryBusca2.ParamByName('Contrato').DataType  := ftString;
                connection.QryBusca2.ParamByName('Contrato').Value     := global_contrato;
                connection.QryBusca2.ParamByName('Insumo').DataType    := ftString;
                connection.QryBusca2.ParamByName('Insumo').Value       := pedido.FieldValues['sIdInsumo'];
                connection.QryBusca2.ParamByName('Almacen').DataType   := ftString;
                connection.QryBusca2.ParamByName('Almacen').Value      := tsAlmacen.KeyValue ;
                connection.QryBusca2.ParamByName('Cantidad').DataType  := ftFloat ;
                connection.QryBusca2.ParamByName('Cantidad').value     := (connection.QryBusca.FieldValues['dExistencia'] + tdCantidad.Value) ;
                connection.QryBusca2.ExecSQL;


                //soad -> Actualizacion de los insumos de la Orden de Compra
                //***************************************************************

                if (tdCantidad.Value + CantidadEntrada) = CantidadPedido then
                begin
                     connection.QryBusca2.Active := False;
                     connection.QryBusca2.SQL.Clear;
                     if (pos('COMPRA',tsTipoMovimiento.Text) > 0) OR (pos('O.C.',tsTipoMovimiento.Text) > 0) then
                         connection.QryBusca2.SQL.Add('UPDATE anexo_ppedido SET sStatus = "Entregado" '+
                                                      'WHERE sContrato =:Contrato and iFolioPedido =:Folio And sIdInsumo =:Insumo ');

                     if (pos('TRASPASO',tsTipoMovimiento.Text) > 0) OR (pos('EMBARQUE',tsTipoMovimiento.Text) > 0) OR (pos('DESEMBARQUE',tsTipoMovimiento.Text) > 0) then
                        connection.QryBusca2.SQL.Add('UPDATE bitacoradesalida SET sStatus = "Entregado" '+
                                                     'WHERE sContrato =:Contrato and iFolioSalida =:Folio And sIdInsumo =:Insumo  ');
                     connection.QryBusca2.ParamByName('Contrato').DataType := ftString;
                     connection.QryBusca2.ParamByName('Contrato').Value    := global_contrato;
                     connection.QryBusca2.ParamByName('Folio').DataType    := ftInteger;
                     connection.QryBusca2.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioMovimiento'];
                     connection.QryBusca2.ParamByName('Insumo').DataType   := ftString ;
                     connection.QryBusca2.ParamByName('Insumo').value      := pedido.FieldValues['sIdInsumo'];
                     connection.QryBusca2.ExecSQL;
                     Pedido.Refresh;
                end ;
                if (tdCantidad.Value + CantidadEntrada) < CantidadPedido then
                    messageDLG('Queda '+FloatToStr(CantidadPedido - (tdCantidad.Value + CantidadEntrada) )+ ' de '+pedido.FieldValues['mDescripcion']+' por recibir. El material esta en estado de "Pendiente" ', mtInformation, [mbOk], 0 );

                if Pedido.RecordCount > 0 then
                begin
                      Pedido.First;
                      tsInsumo.Text     := Pedido.FieldValues['sIdInsumo'];
                      tdCantidad.Value  := Pedido.FieldValues['dCantidad'];
                      mDescripcion.Text := Pedido.FieldValues['Descripcion'];
                end;
            End;

        Agregar.Enabled  := True ;
        Editar.Enabled   := True ;
        Salvar.Enabled   := False ;
        Cancelar.Enabled := False ;
        Eliminar.Enabled := True ;
        Imprimir.Enabled := True ;
        pEntradas.Refresh;
        Grid_pedido.SetFocus;
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmEntradaAlmacen.SeguimientoMaterialGeneral1Click(Sender: TObject);
begin
    Seguimiento_Material('');
    frxEntrada.PreviewOptions.MDIChild := False ;
    frxEntrada.PreviewOptions.Modal := True ;
    frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
    frxEntrada.PreviewOptions.ShowCaptions := False ;
    frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
    frxEntrada.LoadFromFile (global_files + 'seguimiento_materialxpartida.fr3') ;
    frxEntrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
end;

procedure TfrmEntradaAlmacen.SeguimientoMaterialxPartida1Click(Sender: TObject);
begin
    if pEntradas.RecordCount > 0 then
       Seguimiento_Material(pEntradas.FieldValues['sNumeroActividad'])
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
end;

procedure TfrmEntradaAlmacen.SeguimientoMaterialxPartidaDetalle1Click(
  Sender: TObject);
var
   x, y, num, i, contador : integer;
   SumCantidad, SumTotal, SumExcedente, SumRestante : double;
   Cadena   : string;
begin
    if pEntradas.RecordCount = 0 then
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
    connection.zCommand.ParamByName('Actividad').AsString := pEntradas.FieldValues['sNumeroActividad'];
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
                                          'where ra.sContrato =:Contrato and ra.sNumeroActividad =:Actividad group by ra.sIdInsumo, ap.iFolioEntrada, ap.iItem order by ra.sIdInsumo, ap.iFolioPedido, ap.iItem');
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

end;

procedure TfrmEntradaAlmacen.tsIsometricoReferenciaKeyPress(
  Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tmComentarios.SetFocus
end;


procedure TfrmEntradaAlmacen.GridPartidasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid3.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmEntradaAlmacen.GridPartidasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid3.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmEntradaAlmacen.GridPartidasTitleClick(Column: TColumn);
begin
  UtGrid3.DbGridTitleClick(Column);
end;

procedure TfrmEntradaAlmacen.frxReport50GetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'ANEXO') = 0 then
  Begin
      Connection.qryBusca.Active := False ;
      Connection.qryBusca.SQL.Clear ;
      Connection.qryBusca.SQL.Add('Select sAnexo From convenios Where sContrato = :Contrato And sIdConvenio = :convenio') ;
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
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


procedure TfrmEntradaAlmacen.frmBarra2btnAddClick(Sender: TObject);
Var
  dFechaFinal : tDate ;
  iCheck      : Integer ;
begin
activapop(frmEntradaAlmacen, popupprincipal);
 Try
    banderaAgregar:=true;
    OpcButton1 := 'New' ;
    frmBarra2.btnAddClick(Sender);
    frmBarra1.btnCancel.Click ;
    pgControl.ActivePageIndex := 0 ;

    ActivaBotones(True);
    tdFechaAviso.Date  := Date ;
    tmComentarios.Text := '' ;
    tsTipoMovimiento.SetFocus;
    tsFolioMovimiento.KeyValue := Null;
    Anexo_suministro.Append;
    Grid_Entradas.Enabled := False;
    BotonPermiso.permisosBotones(frmBarra1);
    BotonPermiso.permisosBotones(frmBarra2);
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_EntradaAlmacen', 'Al agregar registro ', 0);
    end;
  end;
end;

procedure TfrmEntradaAlmacen.frmBarra2btnEditClick(Sender: TObject);
begin
activapop(frmEntradaAlmacen, popupprincipal);
    if pEntradas.RecordCount > 0 then
    begin
         showmessage('No se puede Editar la Informacion, Existen Materiales para esta Entrada al Almecen.');
         exit;
    end;
    If anexo_suministro.RecordCount > 0 then
    Begin
         OpcButton1 := 'Edit' ;
         ActivaBotones(True);
         frmBarra2.btnEditClick(Sender);
         pgControl.ActivePageIndex := 0 ;
         tdFechaAviso.Enabled  := False ;
         FechaEAnt := anexo_suministro.FieldValues['dFecha'];
         anexo_suministro.Edit;
    End
    Else
         MessageDlg('Folio de Entrada Aplicada no se pueden realizar cambios', mtWarning, [mbOk], 0);
BotonPermiso.permisosBotones(frmBarra1);
BotonPermiso.permisosBotones(frmBarra2);
Grid_Entradas.Enabled := False;

end;

procedure TfrmEntradaAlmacen.frmBarra2btnPostClick(Sender: TObject);
Var
  Nombres,Cadenas:TStringList;
begin
  nombres:=TStringList.Create;cadenas:=TStringList.Create;
  nombres.Add('Tipo de Movimiento');
  nombres.Add('Folio de Movimiento');
  nombres.Add('No. de Orden');
  nombres.Add('Comentarios');
  
  cadenas.Add(tsTipoMovimiento.Text);
  cadenas.Add(tsFolioMovimiento.Text);
  cadenas.Add(tsNumeroOrden.Text);
  cadenas.Add(tmComentarios.Text);

  if not validaTexto(nombres, cadenas, '', '') then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;

  //Continua insercion de datos...
  desactivapop(popupprincipal);
  If OpcButton1 = 'New' then
  Begin
      try
          connection.zCommand.Active := False ;
          connection.zCommand.SQL.Clear ;
          connection.zCommand.SQL.Add ( 'INSERT INTO almacen_entrada ( sContrato, sIdAlmacen, iFolioMovimiento, sIdTipo, sNumeroOrden, dFecha, sIdUsuario, mComentarios ) ' +
                                        'VALUES (:Contrato, :IdAlmacen, :Folio, :Tipo, :Orden, :Fecha, :Usuario, :Comentarios )') ;
          connection.zCommand.params.ParamByName('Contrato').DataType    := ftString ;
          connection.zCommand.params.ParamByName('Contrato').value       := Global_Contrato ;
          connection.zCommand.params.ParamByName('IdAlmacen').DataType   := ftString ;
          connection.zCommand.params.ParamByName('IdAlmacen').value      := tsAlmacen.KeyValue ;
          connection.zCommand.params.ParamByName('Folio').DataType       := ftInteger ;
          connection.zCommand.params.ParamByName('Folio').value          := tsFolioMovimiento.KeyValue ;
          connection.zCommand.params.ParamByName('Tipo').DataType        := ftString ;
          connection.zCommand.params.ParamByName('Tipo').value           := tsTipoMovimiento.KeyValue ;
          connection.zCommand.params.ParamByName('Orden').DataType       := ftString ;
          connection.zCommand.params.ParamByName('Orden').value          := tsNumeroOrden.Text ;
          connection.zCommand.params.ParamByName('Fecha').DataType       := ftDate ;
          connection.zCommand.params.ParamByName('Fecha').value          := tdFechaAviso.Date ;
          connection.zCommand.params.ParamByName('Usuario').DataType     := ftString ;
          connection.zCommand.params.ParamByName('Usuario').value        := Global_Usuario ;
          connection.zCommand.params.ParamByName('Comentarios').DataType := ftMemo ;
          connection.zCommand.params.ParamByName('Comentarios').value    := tmCOmentarios.Text ;
          connection.zCommand.ExecSQL ;

          // Actualizo Kardex del Sistema ....
          connection.zCommand.Active := False ;
          connection.zCommand.SQL.Clear ;
          connection.zCommand.SQL.Add ( 'Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                        'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
          connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
          connection.zCommand.Params.ParamByName('Contrato').Value       := Global_Contrato ;
          connection.zCommand.Params.ParamByName('Usuario').DataType     := ftString ;
          connection.zCommand.Params.ParamByName('Usuario').Value        := Global_Usuario ;
          connection.zCommand.Params.ParamByName('Fecha').DataType       := ftDate ;
          connection.zCommand.Params.ParamByName('Fecha').Value          := Date ;
          connection.zCommand.Params.ParamByName('Hora').DataType        := ftString ;
          connection.zCommand.Params.ParamByName('Hora').value           := FormatDateTime('hh:mm:ss', Now) ;
          connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString ;
          connection.zCommand.Params.ParamByName('Descripcion').Value    := 'Registro de Aviso de Embarque No. ' + ' ' + ' Recibido el día [' + DateToStr(tdFechaAviso.Date) + '] Registrado el [' + DateToStr(tdFechaAviso.Date) +  '] Usuario [ ' + global_usuario + ']' ;
          connection.zCommand.Params.ParamByName('Origen').DataType      := ftString ;
          connection.zCommand.Params.ParamByName('Origen').Value         := 'Reporte Diario' ;
          connection.zCommand.ExecSQL ;
          ActivaBotones(False);
          frmBarra2.btnCancelClick(Sender);

          anexo_suministro.Cancel;

          SavePlace := anexo_suministro.GetBookmark ;
          anexo_suministro.Refresh ;

          Try
             anexo_suministro.GotoBookmark(SavePlace);
          Except
          Else
            anexo_suministro.FreeBookmark(SavePlace);
          End;

      Except
         on e : exception do begin
           anexo_suministro.Cancel;
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Entradas Almacen', 'Al salvar registro', 0)
         end;
      End
  End
  Else
  If OpcButton1 = 'Edit' then
  Begin
      try
            connection.zCommand.Active := False ;
            connection.zCommand.SQL.Clear ;
            connection.zCommand.SQL.Add ( 'UPDATE almacen_entrada SET dFecha = :Fecha, sNumeroOrden = :Orden, mComentarios = :Comentarios ' +
                                          'WHERE sContrato =:Contrato And sIdAlmacen =:IdAlmacen and iFolioEntrada =:Folio and dFecha =:Fecha and sNumeroOrden =:Orden ') ;
            connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
            connection.zCommand.Params.ParamByName('Contrato').value       := Global_Contrato ;
            connection.zCommand.Params.ParamByName('IdAlmacen').DataType  := ftString ;
            connection.zCommand.Params.ParamByName('IdAlmacen').value      := tsAlmacen.KeyValue ;
            connection.zCommand.Params.ParamByName('Folio').DataType       := ftInteger ;
            connection.zCommand.Params.ParamByName('Folio').value          := anexo_suministro.FieldValues['iFolioEntrada'] ;
            connection.zCommand.Params.ParamByName('Orden').DataType       := ftString ;
            connection.zCommand.Params.ParamByName('Orden').value          := tsNumeroOrden.Text ;
            connection.zCommand.Params.ParamByName('Fecha').DataType       := ftDate ;
            connection.zCommand.Params.ParamByName('Fecha').value          := StrToDate(FechaEAnt);
            connection.zCommand.Params.ParamByName('Proveedor').DataType   := ftString ;
            connection.zCommand.Params.ParamByName('Proveedor').value      := '' ;
            connection.zCommand.Params.ParamByName('Comentarios').DataType := ftMemo ;
            connection.zCommand.Params.ParamByName('Comentarios').value    := tmCOmentarios.Text ;
            connection.zCommand.ExecSQL ;

            // Actualizo Kardex del Sistema ....
            connection.zCommand.Active := False ;
            connection.zCommand.SQL.Clear ;
            connection.zCommand.SQL.Add ( 'Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                          'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
            connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
            connection.zCommand.Params.ParamByName('Contrato').Value       := Global_Contrato ;
            connection.zCommand.Params.ParamByName('Usuario').DataType     := ftString ;
            connection.zCommand.Params.ParamByName('Usuario').Value        := Global_Usuario ;
            connection.zCommand.Params.ParamByName('Fecha').DataType       := ftDate ;
            connection.zCommand.Params.ParamByName('Fecha').Value          := Date ;
            connection.zCommand.Params.ParamByName('Hora').DataType        := ftString ;
            connection.zCommand.Params.ParamByName('Hora').value           := FormatDateTime('hh:mm:ss', Now) ;
            connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Descripcion').Value    := 'Modificación de Aviso de Embarque No. ' + ''  + ' Recibido el día [' + DateToStr(tdFechaAviso.Date) + '] Registrado el [' + DateToStr(tdFechaAviso.Date) +  '] Usuario [ ' + global_usuario + ']' ;
            connection.zCommand.Params.ParamByName('Origen').DataType      := ftString ;
            connection.zCommand.Params.ParamByName('Origen').Value         := 'Reporte Diario' ;
            connection.zCommand.ExecSQL ;
            ActivaBotones(False);
            frmBarra2.btnCancelClick(Sender);
            
            SavePlace := anexo_suministro.GetBookmark ;
            anexo_suministro.Refresh ;

            Try
               anexo_suministro.GotoBookmark(SavePlace);
            Except
            Else
              anexo_suministro.FreeBookmark(SavePlace);
            End;
      except
          //  MessageDlg('Ocurrio un error al actualizar el registro', mtWarning, [mbOk], 0);
         on e : exception do begin
         UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Entradas Almacen', 'Al editar registro', 0);
         end;
      End;
  End ;
  Grid_Entradas.Enabled:=True;
  OpcButton1 := '' ;
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmEntradaAlmacen.frmBarra2btnPrinterClick(Sender: TObject);
begin
   try
     If anexo_suministro.RecordCount > 0 Then
     begin
         Reporte.Active := False;
         Reporte.SQL.Clear;

         if (pos('COMPRA',tsTipoMovimiento.Text) > 0) OR (pos('O.C.',tsTipoMovimiento.Text) > 0) then
             Reporte.SQL.Add('select e.sContrato, e.iFolioEntrada, e.iFolioMovimiento, e.sNumeroOrden, e.dFecha, e.sIdUsuario, e.mComentarios, '+
                        'ped.sIdInsumo, ped.mDescripcion, ped.sMedida, ped.dCantidad, ped.dCosto, ped.sStatus, ped.dDescuento, i.dExistencia, i.dCostoMN, alm.sDescripcion as almacen from almacen_entrada e '+
                        'inner join anexo_ppedido ped '+
                        'on (ped.sContrato = e.sContrato and ped.iFolioPedido = e.iFolioMovimiento) '+
                        'inner join almacenes alm '+
                        'on (alm.sIdAlmacen = e.sIdAlmacen) '+
                        'inner join insumos i '+
                        'on (i.sContrato = e.sContrato and i.sIdInsumo = ped.sIdInsumo and i.sIdAlmacen = e.sIdAlmacen) '+
                        'where e.sContrato =:Contrato and e.sIdAlmacen =:Almacen and e.iFolioEntrada =:Folio ');

         if (pos('TRASPASO',tsTipoMovimiento.Text) > 0) OR (pos('EMBARQUE',tsTipoMovimiento.Text) > 0) OR (pos('DESEMBARQUE',tsTipoMovimiento.Text) > 0) then
             Reporte.SQL.Add('select e.sContrato, e.iFolioEntrada, e.iFolioMovimiento, e.sNumeroOrden, e.dFecha, e.sIdUsuario, e.mComentarios, '+
                        'ped.sIdInsumo, i.mDescripcion, i.sMedida, ped.dCantidad, i.dNuevoPrecio as dCosto, ped.sStatus, 0.0 as dDescuento, i.dExistencia, i.dCostoMN, alm.sDescripcion as almacen from almacen_entrada e '+
                        'inner join bitacoradesalida ped '+
                        'on (ped.sContrato = e.sContrato and ped.iFolioSalida = e.iFolioMovimiento) '+
                        'inner join almacenes alm '+
                        'on (alm.sIdAlmacen = e.sIdAlmacen) '+
                        'inner join insumos i '+
                        'on (i.sContrato = e.sContrato and i.sIdInsumo = ped.sIdInsumo and i.sIdAlmacen = e.sIdAlmacen) '+
                        'where e.sContrato =:Contrato and e.sIdAlmacen =:Almacen and e.iFolioEntrada =:Folio ');

         Reporte.ParamByName('Contrato').DataType := ftString ;
         Reporte.ParamByName('Contrato').Value    := global_contrato ;
         Reporte.ParamByName('Almacen').DataType  := ftString;
         Reporte.ParamByName('Almacen').Value     := tsAlmacen.KeyValue;
         Reporte.ParamByName('Folio').DataType    := ftInteger;
         Reporte.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioEntrada'];
         Reporte.ExecSQL;

         frxEntrada.PreviewOptions.MDIChild := False ;
         frxEntrada.PreviewOptions.Modal := True ;
         frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
         frxEntrada.PreviewOptions.ShowCaptions := False ;
         frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
         frxEntrada.LoadFromFile (global_files + global_miReporte+ '_Entrada.fr3') ;
         frxEntrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
   if not FileExists(global_files + 'Entrada.fr3') then
       showmessage('El archivo de reporte Entrada.fr3 no existe, notifique al administrador del sistema');
     end;
   except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Entrada de materiales almacen', 'Al imprimir', 0);
    end;
   end;
end;

procedure TfrmEntradaAlmacen.frmBarra2btnDeleteClick(Sender: TObject);
begin
     If anexo_suministro.RecordCount > 0 Then
        If MessageDlg('Desea eliminar el folio seleccionado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
           Begin
              if pEntradas.RecordCount > 0 then
              begin
                   showmessage('Existen Materiales para esta Entrada al Almecen, Favor de Devolverlos a la Orden de Compra.');
                   exit;
              end;
              // Actualizo Kardex del Sistema ....
              try
                  connection.zCommand.Active := False ;
                  connection.zCommand.SQL.Clear ;
                  connection.zCommand.SQL.Add ('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                               'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
                  connection.zCommand.Params.ParamByName('Contrato').DataType    := ftString ;
                  connection.zCommand.Params.ParamByName('Contrato').Value       := Global_Contrato ;
                  connection.zCommand.Params.ParamByName('Usuario').DataType     := ftString ;
                  connection.zCommand.Params.ParamByName('Usuario').Value        := Global_Usuario ;
                  connection.zCommand.Params.ParamByName('Fecha').DataType       := ftDate ;
                  connection.zCommand.Params.ParamByName('Fecha').Value          := Date ;
                  connection.zCommand.Params.ParamByName('Hora').DataType        := ftString ;
                  connection.zCommand.Params.ParamByName('Hora').value           := FormatDateTime('hh:mm:ss', Now) ;
                  connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString ;
                  connection.zCommand.Params.ParamByName('Descripcion').Value    := 'Eliminación de Salida Alamacen No. + ' + ' Recibido el día [' + DateToStr(tdFechaAviso.Date) + '] Registrado el [' + DateToStr(tdFechaAviso.Date) +  '] Usuario [ ' + global_usuario + ']' ;
                  connection.zCommand.Params.ParamByName('Origen').DataType      := ftString ;
                  connection.zCommand.Params.ParamByName('Origen').Value         := 'Entrada Almacen' ;
                  connection.zCommand.ExecSQL ;

                  connection.zCommand.Active := False ;
                  connection.zCommand.SQL.Clear ;
                  connection.zCommand.SQL.Add ( 'Delete from almacen_entrada where sContrato =:Contrato And sIdAlmacen =:Almacen and iFolioEntrada =:Folio ') ;
                  connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                  connection.zcommand.Params.ParamByName('Contrato').value    := Global_Contrato ;
                  connection.zcommand.Params.ParamByName('Almacen').DataType  := ftString ;
                  connection.zcommand.Params.ParamByName('Almacen').value     := tsAlmacen.KeyValue ;
                  connection.zcommand.Params.ParamByName('Folio').DataType    := ftInteger ;
                  connection.zcommand.Params.ParamByName('Folio').value       := anexo_suministro.FieldValues['iFolioEntrada'] ;
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
                 UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Entradas Almacen', 'Al eliminar registro', 0);
                end;
              end;
          End
end;


procedure TfrmEntradaAlmacen.frmBarra2btnRefreshClick(Sender: TObject);
begin
    anexo_suministro.Active := False ;
    anexo_suministro.Open ;
end;

procedure TfrmEntradaAlmacen.frmBarra2btnCancelClick(Sender: TObject);
begin
  desactivapop(popupprincipal);
  ActivaBotones(False);
  frmBarra2.btnCancelClick(Sender);
  //Grid_Entradas.SetFocus ;
  anexo_suministro.Cancel;
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);
  Grid_Entradas.Enabled:=TRUE;
end;

procedure TfrmEntradaAlmacen.frmBarra2btnExitClick(Sender: TObject);
begin
  frmBarra2.btnExitClick(Sender);
  close
end;

procedure TfrmEntradaAlmacen.tdIdFechaChange(Sender: TObject);
begin
//  if tdFechaAviso.Date<tdidFecha.Date then
//    tdFechaAviso.MinDate:=tdIdFecha.Date;
end;

procedure TfrmEntradaAlmacen.tdIdFechaKeyPress(Sender: TObject;
  var Key: Char);
begin

  If Key = #13 Then
    tdFechaAviso.SetFocus
end;

procedure TfrmEntradaAlmacen.tsOrigenKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tmComentarios.SetFocus
end;

procedure TfrmEntradaAlmacen.tsTipomovimientoEnter(Sender: TObject);
begin
    tsTipoMovimiento.Color := global_color_entrada;
end;

procedure TfrmEntradaAlmacen.tsTipomovimientoExit(Sender: TObject);
begin
    tsTipoMovimiento.Color := global_color_salida;
    //Con esto desplegamos todas las ordenes de compra existentes,
    if (pos('COMPRA',tsTipoMovimiento.Text) > 0) OR (pos('O.C.',tsTipoMovimiento.Text) > 0) then
    begin
        FolioCompra.Active := False;
        FolioCompra.SQL.Clear;
        FolioCompra.SQL.Add('Select iFolioPedido as Folio, sOrdenCompra as Detalle from anexo_pedidos  '+
                            'where sContrato =:Contrato and sStatus ="AUTORIZADO" ');
        FolioCompra.ParamByName('Contrato').AsString := global_contrato;
        FolioCompra.Open;

        if FolioCompra.RecordCount = 0 then
           messageDLG('No se encontraron Ordenes de Compra Autorizadas!', mtInformation, [mbOk], 0);
    end;

    //Con esto desplegamos todos los avsisos de embarque, desembarque, traspasos de materiales existentes..
    if (pos('TRASPASO',tsTipoMovimiento.Text) > 0) OR (pos('EMBARQUE',tsTipoMovimiento.Text) > 0) OR (pos('DESEMBARQUE',tsTipoMovimiento.Text) > 0) then
    begin
        FolioCompra.Active := False;
        FolioCompra.SQL.Clear;
        FolioCompra.SQL.Add('Select iFolioSalida as Folio, CONCAT("Traspaso No. ",iFolioSalida) as Detalle from almacen_salida '+
                            'where sContrato =:Contrato ');
        FolioCompra.ParamByName('Contrato').AsString := global_contrato;
        FolioCompra.Open;

        if FolioCompra.RecordCount = 0 then
           messageDLG('No se encontraron Traspasos Autorizados!', mtInformation, [mbOk], 0);
    end;
end;

procedure TfrmEntradaAlmacen.tsTipomovimientoKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key = #13 then
    tsFolioMovimiento.SetFocus;
end;

procedure TfrmEntradaAlmacen.Grid_EntradasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmEntradaAlmacen.Grid_EntradasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmEntradaAlmacen.Grid_EntradasTitleClick(Column: TColumn);
begin
 UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmEntradaAlmacen.Grid_PedidoKeyPress(Sender: TObject;
  var Key: Char);
begin
       If Key = #13 Then
       begin
            tsInsumo.Text     := Pedido.FieldValues['sIdInsumo'];

            tdCantidad.Value  := Pedido.FieldValues['dCantidad'];
            mDescripcion.Text := Pedido.FieldValues['mDescripcion'];
            mComentarios.Text := '';
            tsFamilia.SetFocus
       end;
end;

procedure TfrmEntradaAlmacen.Grid_PedidoMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
   if grid_pedido.datasource.DataSet.IsEmpty=false  then
     if grid_pedido.DataSource.DataSet.RecordCount>0 then
        UtGrid2.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmEntradaAlmacen.Grid_PedidoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
    if grid_pedido.datasource.DataSet.IsEmpty=false  then
       if grid_pedido.DataSource.DataSet.RecordCount>0 then
          UtGrid2.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmEntradaAlmacen.Grid_PedidoTitleClick(Column: TColumn);
begin
if grid_pedido.datasource.DataSet.IsEmpty=false  then
if grid_pedido.DataSource.DataSet.RecordCount>0 then
 UtGrid2.DbGridTitleClick(Column);
end;

procedure TfrmEntradaAlmacen.iFolioKeyPress(Sender: TObject; var Key: Char);
begin
     if key = #13 then
        tsTipoMovimiento.SetFocus
end;

procedure TfrmEntradaAlmacen.tmComentariosEnter(Sender: TObject);
begin
    tmComentarios.Color := global_color_entrada
end;

procedure TfrmEntradaAlmacen.tmComentariosExit(Sender: TObject);
begin
    tmComentarios.Color := global_color_salida
end;


procedure TfrmEntradaAlmacen.tdCantidadChange(Sender: TObject);
begin
   //TRxCalcEditChangef(tdCantidad, 'Cantidad');
end;

procedure TfrmEntradaAlmacen.tdCantidadEnter(Sender: TObject);
begin
    tdCantidad.Color := global_color_entrada
end;

procedure TfrmEntradaAlmacen.tdCantidadExit(Sender: TObject);
begin
    tdCantidad.Color := global_color_salida
end;

procedure TfrmEntradaAlmacen.tdCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
    //if keyFiltroTRxCalcEdit(tdCantidad, key) then
    //key:=#0;
    If Key = #13 Then
    mComentarios.SetFocus
end;

procedure TfrmEntradaAlmacen.tsAlmacenEnter(Sender: TObject);
begin
tsalmacen.Color:=global_color_entrada
end;

procedure TfrmEntradaAlmacen.tsAlmacenExit(Sender: TObject);
begin
    tsalmacen.Color:=global_color_salida;
    anexo_suministro.Active := False ;
    anexo_suministro.Params.ParamByName('Contrato').DataType := ftString ;
    anexo_suministro.Params.ParamByName('Contrato').Value    := global_contrato ;
    anexo_suministro.Params.ParamByName('Almacen').DataType  := ftString ;
    anexo_suministro.Params.ParamByName('Almacen').Value     := tsAlmacen.KeyValue ;
    anexo_suministro.Open ;
end;

procedure TfrmEntradaAlmacen.tsFamiliaEnter(Sender: TObject);
begin
      tsFamilia.Color := global_color_entrada;
      tsFamilia.Hint  := tsFamilia.Text;
end;

procedure TfrmEntradaAlmacen.tsFamiliaExit(Sender: TObject);
begin
      tsFamilia.Color := global_color_salida;
end;

procedure TfrmEntradaAlmacen.tsFamiliaKeyPress(Sender: TObject; var Key: Char);
begin
     If Key = #13 Then
        tdCantidad.SetFocus
end;

procedure TfrmEntradaAlmacen.tsFolioMovimientoEnter(Sender: TObject);
begin
     tsFolioMovimiento.Color := global_color_entrada;
end;

procedure TfrmEntradaAlmacen.tsFolioMovimientoExit(Sender: TObject);
begin
     tsFolioMovimiento.Color := global_color_salida;
end;

procedure TfrmEntradaAlmacen.tsFolioMovimientoKeyPress(Sender: TObject;
  var Key: Char);
begin
    if key = #13 then
       tdFechaAviso.SetFocus;
end;

procedure TfrmEntradaAlmacen.tsInsumoEnter(Sender: TObject);
begin
tsinsumo.Color:=global_color_entrada;
end;

procedure TfrmEntradaAlmacen.tsInsumoExit(Sender: TObject);
begin
tsinsumo.Color:=global_color_salida;
end;

procedure TfrmEntradaAlmacen.tsNumeroActividadKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tdCantidad.SetFocus
end;

procedure TfrmEntradaAlmacen.mComentariosEnter(Sender: TObject);
begin
     mComentarios.Color := global_color_entrada;
end;

procedure TfrmEntradaAlmacen.mComentariosExit(Sender: TObject);
begin
     mComentarios.Color := global_color_salida;
end;

procedure TfrmEntradaAlmacen.mComentariosKeyPress(Sender: TObject; var Key: Char);
begin
      If Key = #13 Then
        tsinsumo.SetFocus;
end;

procedure TfrmEntradaAlmacen.mDescripcionEnter(Sender: TObject);
begin
mdescripcion.Color:=global_color_entrada;
end;

procedure TfrmEntradaAlmacen.mDescripcionExit(Sender: TObject);
begin
mdescripcion.Color:=global_color_salida;
end;

procedure TfrmEntradaAlmacen.mDescripcionKeyPress(Sender: TObject;
  var Key: Char);
begin
       If Key = #13 Then
        tsfamilia.SetFocus 
end;

procedure TfrmEntradaAlmacen.Paste1Click(Sender: TObject);
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

procedure TfrmEntradaAlmacen.PedidoAfterScroll(DataSet: TDataSet);
begin
     if pedido.RecordCount > 0 then
     begin
         if Not pedido.FieldByName('Descripcion').IsNull then
             mDescripcion.Text := pedido.FieldValues['Descripcion'];

         if Not pedido.FieldByName('sIdInsumo').IsNull then
             tsInsumo.Text := pedido.FieldValues['sIdInsumo'] ;

         if Not pedido.FieldByName('dCantidad').IsNull then
             tdCantidad.Text := pedido.FieldValues['dCantidad'] ;
     end;
end;

procedure TfrmEntradaAlmacen.pEntradasAfterScroll(DataSet: TDataSet);
begin
     if pEntradas.RecordCount > 0 then
     begin
         if Not pEntradas.FieldByName('mDescripcion').IsNull then
         begin
             GridPartidas.Hint := pEntradas.FieldValues['mDescripcion'];
             mDescripcion.Text := pEntradas.FieldValues['mDescripcion'];
         end;
         if Not pEntradas.FieldByName('sIdInsumo').IsNull then
             tsInsumo.Text := pEntradas.FieldValues['sIdInsumo'] ;
     end;
end;

procedure TfrmEntradaAlmacen.pEntradasCalcFields(DataSet: TDataSet);
begin
    if pEntradas.RecordCount > 0 then
    begin
        connection.QryBusca2.Active := False;
        connection.QryBusca2.SQL.Clear;
        connection.QryBusca2.SQL.Add('select dCantidad from anexo_ppedido '+
                                     'WHERE sContrato =:Contrato and iFolioPedido =:Folio And sIdInsumo =:Insumo and iItem =:Item ');
        connection.QryBusca2.ParamByName('Contrato').DataType := ftString;
        connection.QryBusca2.ParamByName('Contrato').Value    := global_contrato;
        connection.QryBusca2.ParamByName('Folio').DataType    := ftInteger;
        connection.QryBusca2.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioMovimiento'];
        connection.QryBusca2.ParamByName('Insumo').DataType   := ftString ;
        connection.QryBusca2.ParamByName('Insumo').value      := pEntradas.FieldValues['sIdInsumo'];
        connection.QryBusca2.ParamByName('Item').DataType     := ftInteger ;
        connection.QryBusca2.ParamByName('Item').value        := pEntradas.FieldValues['iItem'] ;
        connection.QryBusca2.Open;

        if connection.QryBusca2.RecordCount > 0 then
           pEntradas.FieldValues['dPendiente'] := connection.QryBusca2.FieldValues['dCantidad'] - pEntradas.FieldValues['dCantidad']
        else
           pEntradas.FieldValues['dPendiente'] := 0;
    end;
end;

procedure TfrmEntradaAlmacen.PgControlChange(Sender: TObject);
begin
     if anexo_suministro.RecordCount > 0 then
     begin
          tsFolioMovimiento.KeyValue        := anexo_suministro.FieldValues['iFolioMovimiento'];
          tdFechaAviso.Date      := anexo_suministro.FieldValues['dFecha'];
          tsNumeroOrden.Text     := anexo_suministro.FieldValues['sNumeroOrden'];
          tmComentarios.Text     := anexo_suministro.FieldValues['mComentarios'];
          if pgControl.ActivePageIndex = 1 then
          begin
             lblEncabezado.Caption  := 'MATERIALES PARA LA ORDEN DE COMPRA NO. '+ IntToStr(anexo_suministro.FieldValues['iFolioEntrada']);
            // lblEncabezado.Color    := cl3DDkShadow ;
          end
          else
          begin
             lblEncabezado.Caption  := '';
             //lblEncabezado.Color    := $00D7D7D7 ;
          end;

          if (pos('COMPRA',tsTipoMovimiento.Text) > 0) OR (pos('O.C.',tsTipoMovimiento.Text) > 0) then
          begin
              Pedido.Active := False;
              Pedido.SQL.Clear;
              Pedido.SQL.Add('Select sContrato, iFolioPedido, sIdInsumo, sMedida, dCantidad, dCosto, sNumeroActividad, sNumeroOrden, sStatus, LEFT(mDescripcion,200) as Descripcion from anexo_ppedido '+
                             'where sContrato =:Contrato and iFolioPedido =:Folio and sStatus = "Pendiente"');
              Pedido.ParamByName('Contrato').DataType := ftString;
              Pedido.ParamByName('Contrato').Value    := global_contrato;
              Pedido.ParamByName('Folio').DataType    := ftInteger;
              Pedido.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioMovimiento'];
              Pedido.Open;
          end;

          if (pos('TRASPASO',tsTipoMovimiento.Text) > 0) OR (pos('EMBARQUE',tsTipoMovimiento.Text) > 0) OR (pos('DESEMBARQUE',tsTipoMovimiento.Text) > 0) then
          begin
              Pedido.Active := False;
              Pedido.SQL.Clear;
              Pedido.SQL.Add('Select s.sContrato, s.iFolioSalida as iFolioPedido, s.sIdInsumo, i.sMedida, s.dCantidad, i.dNuevoPrecio as dCosto, s.sNumeroActividad, s.sNumeroOrden, s.sStatus, LEFT(i.mDescripcion,200) as Descripcion from bitacoradesalida s '+
                             'inner join insumos i on (i.sContrato = s.sContrato and i.sIdAlmacen = s.sIdAlmacen and i.sIdInsumo = s.sIdInsumo) '+
                             'where s.sContrato =:Contrato and s.iFoliosalida =:Folio ');
              Pedido.ParamByName('Contrato').DataType := ftString;
              Pedido.ParamByName('Contrato').Value    := global_contrato;
              Pedido.ParamByName('Folio').DataType    := ftInteger;
              Pedido.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioMovimiento'];
              Pedido.Open;
          end;

          pEntradas.Active := False;
          pEntradas.ParamByName('Contrato').DataType := ftString;
          pEntradas.ParamByName('Contrato').Value    := global_contrato;
          pEntradas.ParamByName('Folio').DataType    := ftInteger;
          pEntradas.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioEntrada'];
          pEntradas.ParamByName('Almacen').DataType  := ftString ;
          pEntradas.ParamByName('Almacen').Value     := tsAlmacen.KeyValue ;
          pEntradas.Open;
     end;
end;

procedure TfrmEntradaAlmacen.frxEntradaGetValue(const VarName: String; var Value: Variant);
var
  zConsulta:TZQuery;
  sSQL:string;
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
  zConsulta.Params.ParamByName('contrato').Value := global_contrato;
  zConsulta.Params.ParamByName('fecha').DataType := ftDate;
  zConsulta.Params.ParamByName('fecha').Value := anexo_suministro.FieldValues['dFecha'];
  zConsulta.Open;
  if zConsulta.RecordCount>0 then
  begin
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
        Value := 'Sin puesto';
    If CompareText(VarName, 'RECIBE_PUESTO') = 0 then
        Value := 'Sin Puesto';
    If CompareText(VarName, 'ENTREGA_FIRMA') = 0 then
        Value := 'Sin Firmante';
    If CompareText(VarName, 'RECIBE_FIRMA') = 0 then
        Value := 'Sin Firmante';
  end;
  zConsulta.free;
end;

procedure TfrmEntradaAlmacen.dbFolioMovimientoEnter(Sender: TObject);
begin
    iFolio.Color := global_Color_entrada;
end;

procedure TfrmEntradaAlmacen.dbFolioMovimientoExit(Sender: TObject);
begin
    iFolio.Color := global_color_salida;
end;

procedure TfrmEntradaAlmacen.ComentariosAdicionalesClick(Sender: TObject);
begin
    Application.CreateForm(TfrmComentariosxAnexo, frmComentariosxAnexo);
    frmComentariosxAnexo.show ;
end;
//codigo cpl para corregir el error del segundo grid********************
procedure TfrmEntradaAlmacen.Copy1Click(Sender: TObject);
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

procedure TfrmEntradaAlmacen.tdFechaAvisoChange(Sender: TObject);
begin
//  tdFechaAviso.MinDate:=tdidFecha.Date;
end;

procedure TfrmEntradaAlmacen.tdFechaAvisoEnter(Sender: TObject);
begin
    tdFechaAviso.Color := global_color_entrada
end;

procedure TfrmEntradaAlmacen.tdFechaAvisoExit(Sender: TObject);
begin
//    If frmBarra2.btnCancel.Enabled = True  Then
//        If tsReferencia.Text = '' Then
//            tsReferencia.Text := 'CAL' + FormatDateTime('yymmdd' , tdFechaAviso.Date) ;
    tdFechaAviso.Color := global_color_salida
end;

procedure TfrmEntradaAlmacen.tdFechaAvisoKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsNumeroOrden.SetFocus
end;

procedure TfrmEntradaAlmacen.tsNumeroOrdenEnter(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmEntradaAlmacen.tsNumeroOrdenExit(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_salida
end;

procedure TfrmEntradaAlmacen.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
       If Key = #13 Then
          tmComentarios.SetFocus
end;

procedure TfrmEntradaAlmacen.ActivaBotones(Sender: Boolean);
begin
    if sender then
    begin
        iFolio.ReadOnly         := False;
        tdFechaAviso.Enabled    := True ;
        tsNumeroOrden.Enabled   := True ;
        tmComentarios.ReadOnly  := False ;
    end
    else
    begin
         iFolio.ReadOnly        := True ;
         //tdIdFecha.Enabled      := False ;
         tdFechaAviso.Enabled   := False ;
         tsNumeroOrden.Enabled  := False ;
         tmComentarios.ReadOnly := True ;
         //tdCantidad.ReadOnly    := True ;
    end;
end;

procedure TfrmEntradaAlmacen.ActivaBotones2(Sender: Boolean);
begin
    if sender then
    begin
        tdCantidad.Enabled   := True ;
        tsFamilia.ReadOnly   := False ;
        mComentarios.Enabled := True ;
    end
    else
    begin
        tdCantidad.Enabled   := False ;
        tsFamilia.ReadOnly   := True ;
        mComentarios.Enabled := False ;
    end;
end;

procedure TfrmEntradaAlmacen.AgregarClick(Sender: TObject);
begin
     If (anexo_suministro.RecordCount > 0) Then
     Begin
          if tsInsumo.Text = '' then
          begin
               ShowMessage(' Seleccione un Material.. ' );
               exit;
          end;
          OpcButton := 'New';
          Agregar.Enabled := False ;
          Editar.Enabled := False ;
          Salvar.Enabled := True ;
          Cancelar.Enabled := True ;
          Eliminar.Enabled := False ;
          Imprimir.Enabled := False ;
          ActivaBotones2(true);
          tsFamilia.SetFocus ;
    End ;

    BotonPermiso.permisosBotones(frmBarra1);
    BotonPermiso.permisosBotones(frmBarra2);
end;

procedure TfrmEntradaAlmacen.anexo_suministroAfterScroll(DataSet: TDataSet);
begin
     if anexo_suministro.RecordCount > 0 then
     begin
          if Not anexo_suministro.FieldByName('iFolioEntrada').IsNull then
             iFolio.Value := anexo_suministro.FieldValues['iFolioEntrada'] ;
          if Not anexo_suministro.FieldByName('iFolioMovimiento').IsNull then
             tsFolioMovimiento.KeyValue := anexo_suministro.FieldValues['iFolioMovimiento'] ;
          if Not anexo_suministro.FieldByName('dFecha').IsNull then
             tdFechaAviso.Date         := anexo_suministro.FieldValues['dFecha'] ;
          if Not anexo_suministro.FieldByName('sNumeroOrden').IsNull then
             tsNumeroOrden.Text := anexo_suministro.FieldValues['sNumeroOrden'] ;
          if Not anexo_suministro.FieldByName('mComentarios').IsNull then
             tmComentarios.Text := anexo_suministro.FieldValues['mComentarios'] ;


          if (pos('COMPRA',tsTipoMovimiento.Text) > 0) OR (pos('O.C.',tsTipoMovimiento.Text) > 0) then
          begin
              Pedido.Active := False;
              Pedido.SQL.Clear;
              Pedido.SQL.Add('Select sContrato, iFolioPedido, sIdInsumo, sMedida, dCantidad, dCosto, sNumeroActividad, sNumeroOrden, sStatus, LEFT(mDescripcion,200) as Descripcion from anexo_ppedido '+
                             'where sContrato =:Contrato and iFolioPedido =:Folio and sStatus = "Pendiente"');
              Pedido.ParamByName('Contrato').DataType := ftString;
              Pedido.ParamByName('Contrato').Value    := global_contrato;
              Pedido.ParamByName('Folio').DataType    := ftInteger;
              Pedido.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioMovimiento'];
              Pedido.Open;
          end;

          if (pos('TRASPASO',tsTipoMovimiento.Text) > 0) OR (pos('EMBARQUE',tsTipoMovimiento.Text) > 0) OR (pos('DESEMBARQUE',tsTipoMovimiento.Text) > 0) then
          begin
              Pedido.Active := False;
              Pedido.SQL.Clear;
              Pedido.SQL.Add('Select s.sContrato, s.iFolioSalida as iFolioPedido, s.sIdInsumo, i.sMedida, s.dCantidad, i.dNuevoPrecio as dCosto, s.sNumeroActividad, s.sNumeroOrden, s.sStatus, LEFT(i.mDescripcion,200) as Descripcion from bitacoradesalida s '+
                             'inner join insumos i on (i.sContrato = s.sContrato and i.sIdAlmacen = s.sIdAlmacen and i.sIdInsumo = s.sIdInsumo) '+
                             'where s.sContrato =:Contrato and s.iFoliosalida =:Folio and sStatus = "Pendiente"');
              Pedido.ParamByName('Contrato').DataType := ftString;
              Pedido.ParamByName('Contrato').Value    := global_contrato;
              Pedido.ParamByName('Folio').DataType    := ftInteger;
              Pedido.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioMovimiento'];
              Pedido.Open;
          end;

          pEntradas.Active := False;
          pEntradas.ParamByName('Contrato').DataType := ftString;
          pEntradas.ParamByName('Contrato').Value    := global_contrato;
          pEntradas.ParamByName('Folio').DataType    := ftInteger;
          pEntradas.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolioEntrada'];
          pEntradas.ParamByName('Almacen').DataType  := ftString ;
          pEntradas.ParamByName('Almacen').Value     := tsAlmacen.KeyValue ;
          pEntradas.Open;
     end;
end;

procedure TfrmEntradaAlmacen.Seguimiento_Material(dParamActividad: string);
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
                                          'on (i.sContrato = ra.sContrato and i.sIdInsumo = ra.sIdInsumo ) '+
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
                        if not connection.QryBusca.Fieldbyname('dCantidadOC').IsNull then
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
                                          'left join bitacoradeentrada  ap '+
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

End.
