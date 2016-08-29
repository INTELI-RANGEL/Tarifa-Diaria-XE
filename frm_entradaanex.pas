unit frm_entradaanex;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, DB, DBCtrls, global, 
  Mask, OleCtrls, Grids, DBGrids, frm_barra, ExtCtrls, Utilerias,
  Menus, frxClass, frxDBSet, RXDBCtrl, RxLookup, unitactivapop,
  RXCtrls, CheckLst, RxMemDS, ZAbstractRODataset, ZDataset, jpeg,
  Newpanel, rxCurrEdit, rxToolEdit, unitexcepciones, udbgrid,
  unittbotonespermisos, UnitValidaTexto, UFunctionsGHH, UnitValidacion, frm_BusquedaGeneralizada,
  ExtDlgs, masUtilerias, NxCollection, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, dxSkinsCore, dxSkinDevExpressStyle,
  dxSkinFoggy, dxSkinscxPCPainter, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxEdit, cxNavigator, cxDBData, cxGridLevel,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxClasses,
  cxGridCustomView, cxGrid, cxMemo, dxSkinBlack, dxSkinBlue, dxSkinBlueprint,
  dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide,
  dxSkinDevExpressDarkStyle, dxSkinGlassOceans, dxSkinHighContrast,
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
function IsDate(ADate: string): Boolean;
type

  Tfrmentradaanex = class(TForm)
    frmBarra2: TfrmBarra;
    anexo_psuministro: TZReadOnlyQuery;
    ds_anexo_psuministro: TDataSource;
    ds_anexo_suministro: TDataSource;
    anexo_suministro: TZReadOnlyQuery;
    ds_MovimientosdeAlmacen: TDataSource;
    MovimientosdeAlmacen: TZReadOnlyQuery;
    Proveedores: TZReadOnlyQuery;
    ds_proveedores: TDataSource;
    QrySuministro: TZReadOnlyQuery;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    N4: TMenuItem;
    ComentariosAdicionales: TMenuItem;
    KardexdelInventario1: TMenuItem;
    N3: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    MaxFolio: TZReadOnlyQuery;
    ReporteDiario: TZReadOnlyQuery;
    ds_Historico: TDataSource;
    Historico: TZReadOnlyQuery;
    frxDBReporte: TfrxDBDataset;
    Reporte: TZReadOnlyQuery;
    PgControl: TPageControl;
    TabSheet1: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label7: TLabel;
    Label3: TLabel;
    Label6: TLabel;
    Label4: TLabel;
    Label8: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label10: TLabel;
    PaintBox1: TPaintBox;
    tmComentarios: TMemo;
    tsIdTipo: TRxDBLookupCombo;
    tiFolio: TCurrencyEdit;
    tdIdFecha: TDateTimePicker;
    tsReferencia: TEdit;
    tsOrigen: TEdit;
    tsIdProveedor: TRxDBLookupCombo;
    tdFechaAviso: TDateTimePicker;
    tsNumeroOrden: TComboBox;
    TabSheet2: TTabSheet;
    Label11: TLabel;
    Label14: TLabel;
    Label5: TLabel;
    Label19: TLabel;
    tsPlataforma: TLabel;
    imgNotas: TImage;
    frmBarra1: TfrmBarra;
    tsMedida: TEdit;
    GridPartidas: TRxDBGrid;
    tdCantidad: TRxCalcEdit;
    PanelHistorico: tNewGroupBox;
    Grid_Historico: TRxDBGrid;
    frxEntrada: TfrxReport;
    QryPartidasEfectivas: TZReadOnlyQuery;
    ds_PartidasEfectivas: TDataSource;
    zqMateriales: TZReadOnlyQuery;
    ds_materiales: TDataSource;
    RxAviso: TRxMemoryData;
    RxAvisodCantidad: TFloatField;
    RxAvisosContrato: TStringField;
    RxAvisoiFolio: TIntegerField;
    RxAvisosNumeroOrden: TStringField;
    RxAvisodIdFecha: TDateField;
    RxAvisosIdTipo: TStringField;
    RxAvisosReferencia: TStringField;
    RxAvisosIdProveedor: TStringField;
    RxAvisosOrigen: TStringField;
    RxAvisomComentarios: TMemoField;
    RxAvisosNumeroActividad: TStringField;
    RxAvisomDescripcion: TMemoField;
    RxAvisodCantidadAnexo: TFloatField;
    RxAvisodVentaMN: TFloatField;
    RxAvisodVentaDLL: TFloatField;
    RxAvisodAcumulado: TFloatField;
    RxAvisodFechaAviso: TDateField;
    RxAvisosMedida: TStringField;
    RxAvisosTipoActividad: TStringField;
    Panel: tNewGroupBox;
    Label13: TLabel;
    btnGrabar: TBitBtn;
    btnEliminar: TBitBtn;
    btnExaminar: TBitBtn;
    btnEditar: TBitBtn;
    btnCancelar: TBitBtn;
    GroupBox3: TGroupBox;
    bImagen: TImage;
    btnSaveImage: TBitBtn;
    btnNext: TBitBtn;
    btnPrevious: TBitBtn;
    tiRegistro: TCurrencyEdit;
    QryImgAvisos: TZReadOnlyQuery;
    OpenPicture: TOpenPictureDialog;
    SaveImage: TSaveDialog;
    SoporteAvisoEmbarque1: TMenuItem;
    lbl1: TLabel;
    ds_plataformas: TDataSource;
    qryPlataforma: TZReadOnlyQuery;
    qryPlataformasIdPlataforma: TStringField;
    qryPlataformasDescripcion: TStringField;
    cbbPlataforma: TComboBox;
    tsIdMaterial: TEdit;
    tsTrazabilidad: TEdit;
    RxAvisosTrazabilidad: TStringField;
    PanelMateriales: tNewGroupBox;
    PanelDatos: TNxFlipPanel;
    lbl2: TLabel;
    lbl3: TLabel;
    lbl4: TLabel;
    ts1: TEdit;
    tm1: TMemo;
    btnAceptar: TButton;
    btnCancelar1: TButton;
    lbl5: TLabel;
    tsTrazabilidadInsumo: TEdit;
    pmMaterial: TPopupMenu;
    NuevoMaterial: TMenuItem;
    EliminarMaterial1: TMenuItem;
    RxAvisosIdInsumo: TStringField;
    Label9: TLabel;
    tsIdLabel: TEdit;
    chkUnica: TCheckBox;
    Label15: TLabel;
    ds_ProveedorCatalogo: TDataSource;
    zqProveedores: TZReadOnlyQuery;
    cmbProveedores: TDBLookupComboBox;
    Label18: TLabel;
    Grid_entradas: TcxGrid;
    BView_entradas: TcxGridDBTableView;
    sReferencia: TcxGridDBColumn;
    dFecha: TcxGridDBColumn;
    dRecepcion: TcxGridDBColumn;
    stipo: TcxGridDBColumn;
    sOrigen: TcxGridDBColumn;
    Grid_entradasLevel1: TcxGridLevel;
    panelanexo: tNewGroupBox;
    SpeedButton1: TSpeedButton;
    grid_iguales: TcxGrid;
    view_iguales: TcxGridDBTableView;
    Id: TcxGridDBColumn;
    sMedida: TcxGridDBColumn;
    sDescripcion: TcxGridDBColumn;
    cxGridLevel_iguales: TcxGridLevel;
    cmdSalir: TBitBtn;
    function lExisteActividad(sActividad: string): Boolean;
    procedure frmBarra2btnExitClick(Sender: TObject);
    procedure frmBarra2btnAddClick(Sender: TObject);
    procedure frmBarra2btnEditClick(Sender: TObject);
    procedure frmBarra2btnPostClick(Sender: TObject);
    procedure frmBarra2btnCancelClick(Sender: TObject);
    procedure frmBarra2btnDeleteClick(Sender: TObject);
    procedure frmBarra2btnRefreshClick(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure anexo_suministroAfterScroll(DataSet: TDataSet);
    procedure tdCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure tdCantidadEnter(Sender: TObject);
    procedure tdCantidadExit(Sender: TObject);
    procedure tdIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure tdIdFechaEnter(Sender: TObject);
    procedure tdIdFechaExit(Sender: TObject);
    procedure tdFechaAvisoKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaAvisoEnter(Sender: TObject);
    procedure tdFechaAvisoExit(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tsIdTipoKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdTipoEnter(Sender: TObject);
    procedure tsIdTipoExit(Sender: TObject);
    procedure tsReferenciaKeyPress(Sender: TObject; var Key: Char);
    procedure tsReferenciaEnter(Sender: TObject);
    procedure tsReferenciaExit(Sender: TObject);
    procedure tsIdProveedorKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdProveedorEnter(Sender: TObject);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure GridPartidasEnter(Sender: TObject);
    procedure GridPartidasTitleBtnClick(Sender: TObject; ACol: Integer;
      Field: TField);
    procedure KardexdelInventario1Click(Sender: TObject);
    procedure Grid_HistoricoDblClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra2btnPrinterClick(Sender: TObject);
    procedure ComentariosAdicionalesClick(Sender: TObject);
    procedure Grid_EntradasTitleBtnClick(Sender: TObject; ACol: Integer;
      Field: TField);
    procedure anexo_psuministroAfterScroll(DataSet: TDataSet);
    procedure tsIdProveedorExit(Sender: TObject);
    procedure frxEntradaGetValue(const VarName: string; var Value: Variant);
    procedure Grid_EntradasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_EntradasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_EntradasTitleClick(Column: TColumn);
    procedure GridPartidasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure tdCantidadChange(Sender: TObject);
    procedure tiFolioChange(Sender: TObject);
    procedure btnExaminarClick(Sender: TObject);
    procedure SoporteAvisoEmbarque1Click(Sender: TObject);
    procedure btnEditarClick(Sender: TObject);
    procedure btnGrabarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnEliminarClick(Sender: TObject);
    procedure btnPreviousClick(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure QryImgAvisosAfterScroll(DataSet: TDataSet);
    procedure tsIdMaterialKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdMaterialEnter(Sender: TObject);
    procedure tsTrazabilidadEnter(Sender: TObject);
    procedure tsTrazabilidadExit(Sender: TObject);
    procedure tsTrazabilidadKeyPress(Sender: TObject; var Key: Char);
    procedure btnAceptarClick(Sender: TObject);
    procedure NuevoMaterialClick(Sender: TObject);
    procedure tm1Enter(Sender: TObject);
    procedure tm1Exit(Sender: TObject);
    procedure ts1Enter(Sender: TObject);
    procedure ts1Exit(Sender: TObject);
    procedure tsTrazabilidadInsumoEnter(Sender: TObject);
    procedure tsTrazabilidadInsumoExit(Sender: TObject);
    procedure ts1KeyPress(Sender: TObject; var Key: Char);
    procedure tsTrazabilidadInsumoKeyPress(Sender: TObject; var Key: Char);
    procedure btnCancelar1Click(Sender: TObject);
    procedure EliminarMaterial1Click(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure tsIdLabelKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdLabelEnter(Sender: TObject);
    procedure tsIdLabelExit(Sender: TObject);
    procedure cmbProveedoresEnter(Sender: TObject);
    procedure cmbProveedoresExit(Sender: TObject);
    procedure tsOrigenKeyPress(Sender: TObject; var Key: Char);
    procedure view_igualesDblClick(Sender: TObject);
    procedure view_igualesKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdMaterialExit(Sender: TObject);
    procedure cmdSalirClick(Sender: TObject);

  private
    sMenuP: string;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmentradaanex: Tfrmentradaanex;
  sDescripcion: string;
  txtAux: string;
  lNuevo, DebeMostrar: Boolean;
  OpcButton1: string;
  botonpermiso: tbotonespermisos;
  utgrid: ticdbgrid;
  utgrid2: ticdbgrid;
  sBackup: string;
  sSwbs: string;
  sOpcion         : String ;
  sArchivo,
  trazabilidad    : String ;
implementation

uses frm_connection, frm_comentariosxanexo, frm_SalidaAlmacen;


{$R *.dfm}

procedure Tfrmentradaanex.anexo_psuministroAfterScroll(DataSet: TDataSet);
begin
  ImgNotas.Visible := False;
  if anexo_psuministro.RecordCount > 0 then
  begin
    tsIdMaterial.Text := anexo_psuministro.FieldValues['sIdInsumo'];
    tsMedida.Text := anexo_psuministro.FieldValues['sMedida'];
    tsTrazabilidad.Text := anexo_psuministro.FieldValues['sTrazabilidad'];
    tdCantidad.Value := anexo_psuministro.FieldValues['dCantidad'];

    Connection.qryBusca.Active := False;
    Connection.qryBusca.SQL.Clear;
    Connection.qryBusca.SQL.Add('Select sContrato From comentariosxanexo Where sContrato = :Contrato And sNumeroActividad = :Actividad');
    Connection.qryBusca.Params.ParamByName('Contrato').DataType  := ftString;
    Connection.qryBusca.Params.ParamByName('Contrato').Value     := param_global_contrato;
    Connection.qryBusca.Params.ParamByName('actividad').DataType := ftString;
    Connection.qryBusca.Params.ParamByName('actividad').Value    := tsIdMaterial.Text;
    Connection.qryBusca.Open;
    if Connection.qryBusca.RecordCount > 0 then
      imgNotas.Visible := True;
  end
  else
  begin

    Historico.Active := False;
    Historico.Params.ParamByName('Contrato').DataType := ftString;
    Historico.Params.ParamByName('Contrato').Value := param_global_contrato;
    Historico.Params.ParamByName('Actividad').DataType := ftString;
    Historico.Params.ParamByName('Actividad').Value := '';
    Historico.Open;

    tsMedida.Text := '';
    tsTrazabilidad.Text := '';
    tdCantidad.Value := 0;
  end
end;

function IsDate(ADate: string): Boolean;
var
  Dummy: TDateTime;
begin
  IsDate := TryStrToDate(ADate, Dummy);
end;

procedure Tfrmentradaanex.anexo_suministroAfterScroll(DataSet: TDataSet);
begin
  if Self.Visible and (frmbarra2.btnCancel.Enabled = False) then
    frmBarra2.btnCancel.Click;

  if anexo_suministro.RecordCount > 0 then
  begin
    tiFolio.Value  := anexo_suministro.FieldValues['iFolio'];
    tdIdFecha.Date := anexo_suministro.FieldValues['dIdFecha'];
    tdFechaAviso.Date := anexo_suministro.FieldValues['dFechaAviso'];
    tsIdTipo.KeyValue := anexo_suministro.FieldValues['sIdTipo'];
    tsIdProveedor.KeyValue := anexo_suministro.FieldValues['sIdProveedor'];
    tsReferencia.Text := anexo_suministro.FieldValues['sReferencia'];
    tsOrigen.Text := anexo_suministro.FieldValues['sOrigen'];
    tsNumeroOrden.ItemIndex := tsNumeroOrden.Items.IndexOf(anexo_suministro.FieldByName('sNumeroOrden').AsString);
    tmComentarios.Text := anexo_suministro.FieldValues['mComentarios'];

    anexo_psuministro.Active := False;
    anexo_psuministro.Params.ParamByName('Contrato').DataType := ftString;
    anexo_psuministro.Params.ParamByName('Contrato').Value := global_Contrato_Barco;
    anexo_psuministro.Params.ParamByName('Orden').DataType := ftString;
    anexo_psuministro.Params.ParamByName('Orden').Value    := param_global_contrato;
    anexo_psuministro.Params.ParamByName('Folio').DataType := ftInteger;
    anexo_psuministro.Params.ParamByName('Folio').Value := anexo_suministro.FieldValues['iFolio'];
    anexo_psuministro.Open;

    if anexo_psuministro.RecordCount > 0 then
    begin
      tsMedida.Text := anexo_psuministro.FieldValues['sMedida'];
      tsTrazabilidad.Text:= anexo_psuministro.FieldValues['sTrazabilidad'];
      tdCantidad.Value := anexo_psuministro.FieldValues['dCantidad'];

      Connection.qryBusca.Active := False;
      Connection.qryBusca.SQL.Clear;
      Connection.qryBusca.SQL.Add('Select sContrato From comentariosxanexo Where sContrato = :Contrato And sNumeroActividad = :Actividad');
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := param_global_contrato;
      Connection.qryBusca.Params.ParamByName('actividad').DataType := ftString;
      Connection.qryBusca.Params.ParamByName('actividad').Value := tsIdMaterial.Text;
      Connection.qryBusca.Open;

      if Connection.qryBusca.RecordCount > 0 then
        imgNotas.Visible := True;
    end
    else
    begin
      tsMedida.Text := '';
      tsTrazabilidad.Text := '';
      tdCantidad.Value := 0;
    end
  end
  else
  begin
    tiFolio.Value := 0;
    tdIdFecha.Date := Date;
    tdFechaAviso.Date := Date;
    tsReferencia.Text := '';
    tsOrigen.Text := '';
    tmComentarios.Text := '';
    tdCantidad.Value := 0;
    tsNumeroOrden.Text := '';

    Historico.Active := False;
    Historico.Params.ParamByName('Contrato').DataType := ftString;
    Historico.Params.ParamByName('Contrato').Value := param_global_contrato;
    Historico.Params.ParamByName('Actividad').DataType := ftString;
    Historico.Params.ParamByName('Actividad').Value := '';
    Historico.Open;

    anexo_psuministro.Active := False;
    anexo_psuministro.Params.ParamByName('Contrato').DataType := ftString;
    anexo_psuministro.Params.ParamByName('Contrato').Value := global_Contrato_Barco;
    anexo_psuministro.Params.ParamByName('Orden').DataType := ftString;
    anexo_psuministro.Params.ParamByName('Orden').Value    := param_global_contrato;
    anexo_psuministro.Params.ParamByName('Folio').DataType := ftInteger;
    anexo_psuministro.Params.ParamByName('Folio').Value := anexo_suministro.FieldValues['iFolio'];
    anexo_psuministro.Open;

    anexo_psuministro.Open;
  end ;

     QryImgAvisos.Active := False ;
     QryImgAvisos.SQL.Clear ;
     QryImgAvisos.SQL.Add('Select iImagen, bImagen From avisosembarques_adjuntos ' +
                                'Where sContrato = :Contrato And iFolio = :Folio Order By iImagen') ;
     QryImgAvisos.Params.ParamByName('Contrato').DataType  := ftString ;
     QryImgAvisos.Params.ParamByName('Contrato').Value     := param_global_contrato ;
     QryImgAvisos.Params.ParamByName('Folio').DataType     := ftString ;
     QryImgAvisos.Params.ParamByName('Folio').Value        := anexo_suministro.FieldValues['iFolio'] ;
     QryImgAvisos.Open ;
     If QryImgAvisos.RecordCount > 1 Then
            btnNext.Enabled := True
        Else
            bImagen.Picture := Nil ;
     tiRegistro.Value := 1 ;
end;

procedure Tfrmentradaanex.btnAceptarClick(Sender: TObject);
var
   Id : string;
   numero : integer;
begin
      if tsTrazabilidadInsumo.Text = '' then
      begin
          MessageDlg('Indique una Trazabilidad!', mtInformation, [mbOk],0);
          tsTrazabilidadInsumo.SetFocus;
          Exit;
      end;

      if chkUnica.Checked then
      begin
          //Validamos si ya existe la trazabilidad con el mismo insumo..
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('select sTrazabilidad FROM  insumos Where sTrazabilidad = :trazabilidad and sTrazabilidad <> "ST" ');
          connection.zCommand.ParamByName('Trazabilidad').AsString := tsTrazabilidadInsumo.Text;
          connection.zCommand.Open;

           if connection.zCommand.RecordCount > 0 then
          begin
              MessageDlg('La Trazabilidad '+connection.zCommand.FieldValues['sTrazabilidad']+ ' ya Existe. Indique Otro codigo.', mtInformation, [mbOk],0);
              tsTrazabilidadInsumo.SetFocus;
              Exit;
          end;
      end;

      //BUSCAMOS SI EXISTE EL MATERIAL..
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('select max(sIdInsumo) as sIdMaterial FROM insumos Where sContrato = :contrato and sIdInsumo like "MA.%" group by sContrato ');
      Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
      Connection.zCommand.Params.ParamByName('Contrato').Value    := Global_Contrato_barco;
      connection.zCommand.Open;

      //Sino existe lo damos de alta..
      if connection.zCommand.RecordCount > 0 then
         Id := connection.zCommand.FieldValues['sIdMaterial'];

      try
         numero := StrToInt(copy(Id, pos('.', Id) + 1, length(Id)));
      Except
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select ifnull(count(sIdInsumo),1) as sIdMaterial FROM insumos Where sContrato = :contrato group by sContrato ');
        Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
        Connection.zCommand.Params.ParamByName('Contrato').Value    := Global_Contrato_barco;
        connection.zCommand.Open;
        numero := connection.zCommand.FieldByName('sidmaterial').AsInteger;
      end;

      if numero >= 10000 then
         inc(numero)
      else
         numero := 10000;

      //Ahora insertamos el material..
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add('INSERT INTO insumos ( sContrato, sIdInsumo, sIdProveedor, sIdAlmacen, sTipoActividad, mDescripcion, dFechaInicio, dCostoMN, dCostoDLL, dVentaMN, dVentaDLL, sMedida, '+
      ' dCantidad, dInstalado, sIdGrupo, dNuevoPrecio, sIdFase, sTrazabilidad, sLabelIdMaterial, sColumnaAux ) '+
      ' VALUES (:contrato, :insumo, :proveedor, :almacen, :tipoactividad, :Descripcion, :fechai, 0, 0, 0, 0, :medida, 0, 0, null, 0, null, :trazabilidad, :labelmaterial, :insumo)');
      connection.zCommand.Params.ParamByName('contrato').DataType := ftString;
      connection.zCommand.Params.ParamByName('contrato').value    := global_contrato_barco;
      connection.zCommand.Params.ParamByName('insumo').DataType   := ftString;
      connection.zCommand.Params.ParamByName('insumo').value      := 'MA.'+IntToStr(numero);
      connection.zCommand.Params.ParamByName('almacen').DataType  := ftString;
      connection.zCommand.Params.ParamByName('almacen').value     := 'ALM-01';
      connection.zCommand.Params.ParamByName('tipoactividad').DataType := ftString;
      connection.zCommand.Params.ParamByName('tipoactividad').value  := 'Material';
      connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString;
      connection.zCommand.Params.ParamByName('Descripcion').value    := tm1.Text ;
      connection.zCommand.Params.ParamByName('fechai').DataType      := ftDate;
      connection.zCommand.Params.ParamByName('fechai').value         := date;
      connection.zCommand.Params.ParamByName('medida').DataType      := ftString;
      connection.zCommand.Params.ParamByName('medida').value         := ts1.Text;
      connection.zCommand.Params.ParamByName('trazabilidad').DataType:= ftString;
      connection.zCommand.Params.ParamByName('trazabilidad').value   := tsTrazabilidadInsumo.Text;
      connection.zCommand.Params.ParamByName('labelmaterial').DataType := ftString;
      connection.zCommand.Params.ParamByName('labelmaterial').value    := tsIdLabel.Text;
      connection.zCommand.Params.ParamByName('proveedor').DataType := ftString;
      if connection.QryBusca2.RecordCount > 0 then
         connection.zCommand.Params.ParamByName('proveedor').value  := zqProveedores.FieldByName('sIdProveedor').AsString
      else
         connection.zCommand.Params.ParamByName('proveedor').value  := Null;
      connection.zCommand.ExecSQL;

      PanelMateriales.Visible := False;
      zqMateriales.Refresh;
      zqMateriales.Locate('sIdInsumo', 'MA.'+IntToStr(numero), [loCaseInsensitive])
end;

procedure Tfrmentradaanex.btnCancelar1Click(Sender: TObject);
begin
    PanelMateriales.Visible := False;
end;

procedure Tfrmentradaanex.btnCancelarClick(Sender: TObject);
begin
    btnCancelar.Enabled := False ;
    btnExaminar.Enabled := True ;
    btnEditar.Enabled := True ;
    btnGrabar.Enabled := False ;
    If QryImgAvisos.State <> dsInactive Then
        If QryImgAvisos.RecordCount > 0 Then
         Begin
             btnSaveImage.Enabled := True ;
             btnEliminar.Enabled := True ;
         End
         Else
         Begin
             btnSaveImage.Enabled := False ;
             btnEliminar.Enabled := False ;
         End
end;

procedure Tfrmentradaanex.btnEditarClick(Sender: TObject);
begin
  If frmBarra2.btnCancel.Enabled = False Then
    Begin
          sOpcion := 'Edit' ;
          sArchivo := '' ;
          btnGrabar.Enabled := True ;
          btnCancelar.Enabled := True ;
          btnExaminar.Enabled := False ;
          btnEditar.Enabled := False ;
          btnSaveImage.Enabled := False ;
          btnEliminar.Enabled := False ;
    End
end;

procedure Tfrmentradaanex.btnEliminarClick(Sender: TObject);
begin
 If anexo_suministro.RecordCount > 0 Then
     If frmBarra2.btnCancel.Enabled = False Then
          If QryImgAvisos.RecordCount > 0 Then
          Begin
              connection.zCommand.Active := False ;
              connection.zCommand.SQL.Clear ;
              connection.zCommand.SQL.Add ( 'Delete From avisosembarques_adjuntos Where sContrato = :Contrato And iFolio = :Folio And iImagen = :Item') ;
              connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
              connection.zCommand.Params.ParamByName('Contrato').Value    := param_Global_Contrato ;
              connection.zCommand.Params.ParamByName('Folio').DataType    := ftString ;
              connection.zCommand.Params.ParamByName('Folio').Value       := anexo_suministro.FieldValues['sIdConvenio'] ;
              connection.zCommand.Params.ParamByName('Item').DataType     := ftInteger ;
              connection.zCommand.Params.ParamByName('Item').Value        := QryImgAvisos.FieldValues['iImagen'] ;
              connection.zcommand.ExecSQL ;
              bImagen.Picture.Bitmap := Nil ;
//              bImagen.Picture.LoadFromFile('') ;
              QryImgAvisos.Active := False ;
              QryImgAvisos.Open ;
              If QryImgAvisos.RecordCount > 0 Then
                   btnEliminar.Enabled := True
              Else
                   btnEliminar.Enabled := False ;
          End
end;

procedure Tfrmentradaanex.btnExaminarClick(Sender: TObject);

Var
   size: Real ;
begin
  If anexo_suministro.RecordCount > 0 Then
     If frmBarra2.btnCancel.Enabled = False Then
     Begin
          sOpcion := 'New' ;
          bImagen.Picture.Bitmap := Nil ;
          btnGrabar.Enabled     := True ;
          btnCancelar.Enabled   := True ;
          btnSaveImage.Enabled  := False ;
          btnExaminar.Enabled   := False ;
          btnEditar.Enabled     := False ;
          OpenPicture.Title := 'Inserta Imagen';
          sArchivo := '' ;
          If OpenPicture.Execute then
          begin
              try
                  sArchivo := OpenPicture.FileName ;
                  size := Tamanyo (sArchivo) ;
                  If size <= 100 Then
                      bImagen.Picture.LoadFromFile(OpenPicture.FileName)
                  Else
                      MessageDlg('La imagen a adjuntar no debe ser mayor a 100 kb.', mtInformation, [mbOk], 0);
              except
//                  bImagen.Picture.LoadFromFile('') ;
              end
          end
     end
end;

procedure Tfrmentradaanex.btnGrabarClick(Sender: TObject);
Var
  iItem    : Integer ;
Begin
  If anexo_suministro.RecordCount > 0 Then
      If sOpcion = 'New' Then
      Begin
          If sArchivo <> '' Then
          Begin
              iItem := 1 ;
              If QryImgAvisos.RecordCount > 0 Then
              Begin
                  QryImgAvisos.Last ;
                  iItem := QryImgAvisos.FieldValues['iImagen'] + 1;
              End ;
              connection.zCommand.Active := False ;
              connection.zCommand.SQL.Clear ;
              connection.zCommand.SQL.Add ( 'Insert Into avisosembarques_adjuntos (sContrato, iFolio, iImagen, bImagen) ' +
                                            'Values (:Contrato, :Folio, :Item, :Imagen)') ;
              connection.zCommand.Params.ParamByName('Contrato').DataType  := ftString ;
              connection.zCommand.Params.ParamByName('Contrato').Value     := param_Global_Contrato ;
              connection.zCommand.Params.ParamByName('Folio').DataType     := ftString ;
              connection.zCommand.Params.ParamByName('Folio').Value        := anexo_suministro.FieldValues['iFolio'] ;
              connection.zCommand.Params.ParamByName('Item').DataType      := ftInteger ;
              connection.zCommand.Params.ParamByName('Item').Value         := iItem ;
              connection.zCommand.Params.ParamByName('Imagen').LoadFromFile(sArchivo, ftGraphic) ;
              connection.zcommand.ExecSQL
          End
      End
      Else
      Begin
          If sArchivo <> '' Then
          Begin
                  connection.zCommand.Active := False ;
                  connection.zCommand.SQL.Clear ;
                  connection.zCommand.SQL.Add ( 'Update avisosembarques_Adjuntos SET bImagen = :Imagen ' +
                                                    'Where sContrato = :contrato And iFolio = :Folio And iImagen = :Item') ;
                  connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
                  connection.zCommand.Params.ParamByName('Contrato').Value    := param_Global_Contrato ;
                  connection.zCommand.Params.ParamByName('Folio').DataType    := ftString ;
                  connection.zCommand.Params.ParamByName('Folio').Value       := anexo_suministro.FieldValues['iFolio'] ;
                  connection.zCommand.Params.ParamByName('Item').DataType     := ftInteger ;
                  connection.zCommand.Params.ParamByName('Item').Value        := QryImgAvisos.FieldValues['iImagen'] ;
                  connection.zCommand.Params.ParamByName('Imagen').LoadFromFile(sArchivo, ftGraphic) ;
                  connection.zcommand.ExecSQL
          End
      End ;
      QryImgAvisos.refresh ;
      If QryImgAvisos.RecordCount > 1 Then
      btnNext.Enabled       := True ;
      btnGrabar.Enabled     := False ;
      btnCancelar.Enabled   := False ;
      btnSaveImage.Enabled  := True ;
      btnExaminar.Enabled   := True ;
      btnEditar.Enabled     := True ;
      btnEliminar.Enabled   := True ;
end;

procedure Tfrmentradaanex.btnNextClick(Sender: TObject);
begin
    btnPrevious.Enabled := True ;
    If NOT QryImgAvisos.eof Then
        QryImgAvisos.Next ;
    If QryImgAvisos.Eof Then
        btnNext.Enabled := False ;
end;

procedure Tfrmentradaanex.btnPreviousClick(Sender: TObject);
begin
    btnNext.Enabled := True ;
    If NOT QryImgAvisos.Bof Then
        QryImgAvisos.Prior ;
    If QryImgAvisos.Bof Then
        btnPrevious.Enabled := False ;
end;

procedure Tfrmentradaanex.Can1Click(Sender: TObject);
begin
  frmBarra2.btnCancel.Click
end;

procedure Tfrmentradaanex.cmbProveedoresEnter(Sender: TObject);
begin
    cmbProveedores.Color := global_color_entrada;
end;

procedure Tfrmentradaanex.cmbProveedoresExit(Sender: TObject);
begin
    cmbProveedores.Color := global_color_salida;
end;

procedure Tfrmentradaanex.cmdSalirClick(Sender: TObject);
begin
    panelAnexo.Visible := False;
    frmbarra1.btnCancel.Click;
end;

procedure Tfrmentradaanex.ComentariosAdicionalesClick(Sender: TObject);
begin
  global_partida := tsIdMaterial.Text;
  global_orden   := tsNumeroOrden.Text;
  Application.CreateForm(TfrmComentariosxAnexo, frmComentariosxAnexo);
  frmComentariosxAnexo.Show;
end;

procedure Tfrmentradaanex.Copy1Click(Sender: TObject);
begin
  UtGrid.CopyRowsToClip;
end;

procedure Tfrmentradaanex.view_igualesDblClick(Sender: TObject);
begin
    if view_iguales.OptionsView.CellAutoHeight then
       view_iguales.OptionsView.CellAutoHeight := False
    else
       view_iguales.OptionsView.CellAutoHeight := True;
end;

procedure Tfrmentradaanex.view_igualesKeyPress(Sender: TObject; var Key: Char);
begin
    if key = #13 then
    begin
        tdCantidad.SetFocus;
        if zqMateriales.RecordCount > 0 then
        begin
            tsIdMaterial.Text   := zqMateriales.FieldValues['sIdInsumo'];
            //Buscamos la trazabilidad
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select strazabilidad from anexo_psuministro where sContrato =:Contrato and sIdInsumo =:Insumo group by sTrazabilidad ');
            connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
            connection.zCommand.ParamByName('Insumo').AsString   := zqMateriales.FieldValues['sIdInsumo'];
            connection.zCommand.Open;

            if connection.zCommand.RecordCount > 0 then
               tstrazabilidad.Text := connection.zCommand.FieldByName('strazabilidad').AsString;

            tsMedida.Text       := zqMateriales.FieldValues['sMedida'];
        end;
        panelanexo.Visible := False;
    end;
end;

procedure Tfrmentradaanex.Editar1Click(Sender: TObject);
begin
  frmBarra2.btnEdit.Click
end;

procedure Tfrmentradaanex.Eliminar1Click(Sender: TObject);
begin
  frmBarra2.btnDelete.Click
end;

procedure Tfrmentradaanex.EliminarMaterial1Click(Sender: TObject);
var
    SavePlace: TBookmark;
begin
    //Eliminacion de materiales,
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select * from anexo_psuministro where sIdInsumo =:Insumo');
    connection.zCommand.ParamByName('Insumo').AsString := zqMateriales.FieldValues['sIdInsumo'];
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
    begin
        MSG_W('No se puede eliminar el Material, existen regitros en Avisos de Embarque');
        Exit;
    end
    else
    begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('delete from insumos where sContrato =:Contrato and sIdInsumo =:Insumo and sTrazabilidad =:Trazabilidad');
        connection.zCommand.ParamByName('Contrato').AsString     := global_Contrato_Barco;
        connection.zCommand.ParamByName('Insumo').AsString       := zqMateriales.FieldValues['sIdInsumo'];
        connection.zCommand.ParamByName('Trazabilidad').AsString := zqMateriales.FieldValues['sTrazabilidad'];
        connection.zCommand.ExecSQL;

        SavePlace := zqMateriales.GetBookmark;
        zqMateriales.Active := False;

        zqMateriales.Open;
        try
           zqMateriales.GotoBookmark(SavePlace);
        except
        else
           zqMateriales.FreeBookmark(SavePlace);
        end;
    end;
end;

procedure Tfrmentradaanex.FormActivate(Sender: TObject);
begin
  MovimientosdeAlmacen.Active := False;
  MovimientosdeAlmacen.Open;
  
  Proveedores.Active := False;
  Proveedores.Open;
end;

procedure Tfrmentradaanex.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := cafree;
  botonpermiso.free;
end;

procedure Tfrmentradaanex.FormShow(Sender: TObject);
begin
  sMenuP := stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'opAvisodeEmb', PopupPrincipal);
  Label10.Caption := '';

  tdIdFecha.Enabled := False;
  tdFechaAviso.Enabled := False;
  tsNumeroOrden.Enabled := False;
  cbbPlataforma.Enabled := False;
  tsIdTipo.ReadOnly := True;
  tsIdProveedor.ReadOnly := True;
  tsReferencia.ReadOnly := True;
  tsOrigen.ReadOnly := True;
  tmComentarios.ReadOnly := True;
  tsIdMaterial.ReadOnly := True;
  tdCantidad.ReadOnly := True;

  MovimientosdeAlmacen.Active := False;
  MovimientosdeAlmacen.Open;

  Proveedores.Active := False;
  Proveedores.Open;

  qryPlataforma.Active := False;
  qryPlataforma.Open;

  cbbPlataforma.Clear;
  cbbPlataforma.Items.Add('S/N');
  while not qryPlataforma.Eof do
  begin
      cbbPlataforma.Items.Add(qryPlataforma.FieldValues['sIdPlataforma']);
      qryPlataforma.Next;
  end;

  QryPartidasEfectivas.Active := False;
  QryPartidasEfectivas.ParamByName('Contrato').AsString := param_global_Contrato;
  QryPartidasEfectivas.ParamByName('Convenio').AsString := convenio_reporte;
  QryPartidasEfectivas.Open;

  zqMateriales.Active := False;
  zqMateriales.ParamByName('Contrato').AsString  := global_Contrato_Barco;
  zqMateriales.Open;

  tsNumeroOrden.Items.Clear;
  Connection.qryBusca.Active := False;
  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('Select sNumeroOrden from ordenesdetrabajo where sContrato = :Contrato and ' +
    'cIdStatus = :status order by sNumeroOrden');
  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value := param_Global_Contrato;
  Connection.qryBusca.Params.ParamByName('status').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('status').Value := connection.configuracion.FieldValues['cStatusProceso'];
  Connection.qryBusca.Open;

  tsNumeroOrden.Items.Add('S/N');
  if Connection.qryBusca.RecordCount >= 1 then
    while not Connection.qryBusca.Eof do
    begin
      tsNumeroOrden.Items.Add(Connection.qryBusca.FieldValues['sNumeroOrden']);
      Connection.qryBusca.Next
    end;
    tsNumeroOrden.ItemIndex := 0;

    anexo_suministro.Active := False;
    anexo_suministro.Params.ParamByName('Contrato').DataType := ftString;
    anexo_suministro.Params.ParamByName('Contrato').Value := param_global_contrato;
    anexo_suministro.Open;

    anexo_psuministro.Active := False;
    anexo_psuministro.Params.ParamByName('Contrato').DataType := ftString;
    anexo_psuministro.Params.ParamByName('Contrato').Value := global_Contrato_Barco;
    anexo_psuministro.Params.ParamByName('Orden').DataType := ftString;
    anexo_psuministro.Params.ParamByName('Orden').Value    := param_global_contrato;
    anexo_psuministro.Params.ParamByName('Folio').DataType := ftInteger;
    anexo_psuministro.Params.ParamByName('Folio').Value := anexo_suministro.FieldValues['iFolio'];
    anexo_psuministro.Open;

  if anexo_suministro.RecordCount > 0 then
    anexo_psuministro.Params.ParamByName('Folio').Value := anexo_suministro.FieldValues['iFolio']
  else
    anexo_psuministro.Params.ParamByName('Folio').Value := 0;
  anexo_psuministro.Open;

  zqProveedores.Active := False;
  zqProveedores.Open;

  frmBarra2.btnAdd.Enabled := true;
  frmBarra2.btnAdd.Enabled := true;
  frmBarra2.btnEdit.Enabled := true;
  frmBarra2.btnDelete.Enabled := true;
  frmBarra2.btnPrinter.Enabled := true;
  BotonPermiso.permisosBotones(frmBarra1);
  BotonPermiso.permisosBotones(frmBarra2);

  PgControl.ActivePageIndex := 0;
  PgControl.ActivePageIndex := 1;
  Grid_Entradas.Enabled := True;
end;

procedure Tfrmentradaanex.frmBarra1btnAddClick(Sender: TObject);
var
  lValido: Boolean;
begin
  lValido := True;
  if global_grupo <> 'INTEL-CODE' then
  begin
    ReporteDiario.Active := False;
    ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
    ReporteDiario.Params.ParamByName('contrato').Value := param_global_contrato;
    ReporteDiario.Params.ParamByName('turno').DataType := ftString;
    ReporteDiario.Params.ParamByName('turno').Value := global_turno_reporte;
    ReporteDiario.Params.ParamByName('Fecha').DataType := ftDate;
    ReporteDiario.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
    ReporteDiario.Open;
    if ReporteDiario.RecordCount > 0 then
      while not ReporteDiario.Eof and lValido do
      begin
        if ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' then
        begin
          MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
          lValido := False;
        end;
        ReporteDiario.Next
      end
  end;

  if lValido then
    if (anexo_suministro.RecordCount > 0) then
    begin
      Insertar1.Enabled := False;
      Editar1.Enabled := False;
      Registrar1.Enabled := True;
      Can1.Enabled := True;
      Eliminar1.Enabled := False;
      Refresh1.Enabled := False;
      frmBarra1.btnAddClick(Sender);
      tsIdMaterial.ReadOnly := False;
      tdCantidad.ReadOnly := False;
      tsTrazabilidad.Enabled := True;
      tsMedida.Text := '';
      tsTrazabilidad.Text := '';
      tdCantidad.Value := 0;
      tsIdMaterial.SetFocus;
      sSwbs := '';
    end;

  panelanexo.Visible := True;
  panelanexo.Width  := 727;
  panelanexo.Height := 366;
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure Tfrmentradaanex.frmBarra1btnCancelClick(Sender: TObject);
begin
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  tsIdMaterial.ReadOnly := True;
  tdCantidad.ReadOnly := True;
  tdCantidad.Enabled := True;
  if GridPartidas.CanFocus then
    GridPartidas.SetFocus;
  frmBarra1.btnCancelClick(Sender);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure Tfrmentradaanex.frmBarra1btnDeleteClick(Sender: TObject);
var
  SavePlace: TBookmark;
  lValido: Boolean;
begin
  lValido := True;
  if global_grupo <> 'INTEL-CODE' then
  begin
    ReporteDiario.Active := False;
    ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
    ReporteDiario.Params.ParamByName('contrato').Value := param_global_contrato;
    ReporteDiario.Params.ParamByName('turno').DataType := ftString;
    ReporteDiario.Params.ParamByName('turno').Value := global_turno_reporte;
    ReporteDiario.Params.ParamByName('Fecha').DataType := ftDate;
    ReporteDiario.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
    ReporteDiario.Open;
    if ReporteDiario.RecordCount > 0 then
      while not ReporteDiario.Eof and lValido do
      begin
        if ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' then
        begin
          MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
          lValido := False;
        end;
        ReporteDiario.Next
      end
  end;

  if lValido then
    if anexo_psuministro.RecordCount > 0 then
    begin
      try
        {Validamos sino está reportado el material en una salida..}
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select iFolioAviso FROM bitacoradesalida ' +
          'WHERE sContrato = :Contrato And iFolioAviso = :Folio And sIdInsumo = :Insumo and sTrazabilidad =:Trazabilidad ');
        connection.zCommand.Params.ParamByName('Contrato').DataType        := ftString;
        connection.zCommand.Params.ParamByName('Contrato').value           := param_Global_Contrato;
        connection.zCommand.Params.ParamByName('Folio').DataType           := ftInteger;
        connection.zCommand.Params.ParamByName('Folio').value              := anexo_suministro.FieldValues['iFolio'];
        connection.zCommand.Params.ParamByName('Insumo').DataType          := ftString;
        connection.zCommand.Params.ParamByName('Insumo').value             := anexo_psuministro.FieldValues['sIdInsumo'];
        connection.zCommand.Params.ParamByName('Trazabilidad').DataType    := ftString;
        connection.zCommand.Params.ParamByName('Trazabilidad').value       := anexo_psuministro.FieldValues['sTrazabilidad'];
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
           messageDLG('No se puede Eliminar, Existen Salidas de Mateiales con Folio ['+IntToStr(anexo_psuministro.FieldValues['iFolio'])+ '] Id Mat. <<'+ anexo_psuministro.FieldValues['sIdInsumo'] +'>>', mtInformation, [mbOk], 0)
        else
        begin
            {Sino existe ningun registro en la salida de materiales..}
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('Delete from anexo_psuministro where sContrato = :Contrato And ' +
                      'iFolio = :Folio and sIdInsumo =:Insumo And sTrazabilidad = :trazabilidad ');
            connection.zcommand.Params.ParamByName('Contrato').DataType     := ftString;
            connection.zcommand.Params.ParamByName('Contrato').value        := param_Global_Contrato;
            connection.zcommand.Params.ParamByName('Folio').DataType        := ftInteger;
            connection.zcommand.Params.ParamByName('Folio').value           := anexo_suministro.FieldValues['iFolio'];
            connection.zcommand.Params.ParamByName('Insumo').DataType       := ftString;
            connection.zcommand.Params.ParamByName('Insumo').value          := anexo_psuministro.FieldValues['sIdInsumo'];
            connection.zcommand.Params.ParamByName('trazabilidad').DataType := ftString;
            connection.zcommand.Params.ParamByName('trazabilidad').value    := anexo_psuministro.FieldValues['sTrazabilidad'];
            connection.zCommand.ExecSQL;
        end;

        SavePlace := anexo_psuministro.GetBookmark;
        anexo_psuministro.Active := False;

        anexo_psuministro.Open;
        try
          anexo_psuministro.GotoBookmark(SavePlace);
        except
        else
          anexo_psuministro.FreeBookmark(SavePlace);
        end;
        GridPartidas.SetFocus
      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Avisos de Embarque', 'Al eliminar registro', 0);
        end;
      end
    end
end;

procedure Tfrmentradaanex.frmBarra1btnEditClick(Sender: TObject);
var
  lValido: Boolean;
begin
  lValido := True;
  if global_grupo <> 'INTEL-CODE' then
  begin
    ReporteDiario.Active := False;
    ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
    ReporteDiario.Params.ParamByName('contrato').Value := param_global_contrato;
    ReporteDiario.Params.ParamByName('turno').DataType := ftString;
    ReporteDiario.Params.ParamByName('turno').Value := global_turno_reporte;
    ReporteDiario.Params.ParamByName('Fecha').DataType := ftDate;
    ReporteDiario.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
    ReporteDiario.Open;
    if ReporteDiario.RecordCount > 0 then
      while not ReporteDiario.Eof and lValido do
      begin
        if ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' then
        begin
          MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
          lValido := False;
        end;
        ReporteDiario.Next
      end
  end;

  if lValido then
    if anexo_suministro.RecordCount > 0 then
    begin    
      Insertar1.Enabled := False;
      Editar1.Enabled := False;
      Registrar1.Enabled := True;
      Can1.Enabled := True;
      Eliminar1.Enabled := False;
      Refresh1.Enabled := False;
      frmBarra1.btnEditClick(Sender);
      tsIdMaterial.ReadOnly := False;
      tdCantidad.ReadOnly := False;
      trazabilidad := tsTrazabilidad.Text;
      tsIdMaterial.SetFocus;
    end;
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure Tfrmentradaanex.frmBarra1btnPostClick(Sender: TObject);
var
  SavePlace: TBookmark;
  nombres, cadenas: TStringList;
begin
  {Validacion de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('Id Material'); nombres.Add('Medida'); nombres.Add('Trazabilidad');
  cadenas.Add(tsIdMaterial.Text); cadenas.Add(tsMedida.Text); cadenas.Add(tsTrazabilidad.Text);
  if not validaTexto(nombres, cadenas, '', '') then
  begin
      MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
      if  tsTrazabilidad.Text = '' then
          tsTrazabilidad.SetFocus;
      exit;
  end;
  {Continua insercion de datos}

  if tsIdMaterial.Text <> '' then
    begin
      try
        if OpcButton = 'New' then
        begin
          connection.zCommand.Active := False;
          connection.zCommand.SQL.Clear;
          connection.zCommand.SQL.Add('INSERT INTO anexo_psuministro ( sContrato, iFolio ,swbs, sNumeroActividad, dCantidad, dCantidadRestante, sIdInsumo, sTrazabilidad ) ' +
            'VALUES (:Contrato, :Folio,:swbs, :Actividad, :Cantidad, :Cantidad, :Insumo, :trazabilidad )');
          connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
          connection.zCommand.Params.ParamByName('Contrato').value := param_Global_Contrato;
          connection.zCommand.Params.ParamByName('Folio').DataType := ftInteger;
          connection.zCommand.Params.ParamByName('Folio').value := anexo_suministro.FieldValues['iFolio'];
          connection.zCommand.Params.ParamByName('Actividad').DataType := ftString;
          connection.zCommand.Params.ParamByName('Actividad').value := '';
          connection.zCommand.Params.ParamByName('Cantidad').DataType := ftFloat;
          connection.zCommand.Params.ParamByName('Cantidad').value := tdCantidad.Value;
          connection.zCommand.Params.ParamByName('swbs').AsString := '';
          connection.zCommand.Params.ParamByName('Insumo').DataType := ftString;
          connection.zCommand.Params.ParamByName('Insumo').value := zqMateriales.FieldValues['sIdInsumo'];
          connection.zCommand.Params.ParamByName('Trazabilidad').DataType := ftString;
          connection.zCommand.Params.ParamByName('Trazabilidad').value := tsTrazabilidad.Text;
          connection.zCommand.ExecSQL;
        end
        else
        begin
            {Validamos que la edicion no se menor a la cantidad requerida..}
            if tdCantidad.Value < anexo_psuministro.FieldByName('dCantidadRestante').AsFloat then
               messageDLG('No se puede Editar la Cantidad. La Cantidad es menor a la Cantidad Restante', mtInformation , [mbOk], 0)
            else
            begin
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('UPDATE anexo_psuministro SET dCantidad = :Cantidad, dCantidadRestante = dCantidadRestante + :CantidadAdic, sTrazabilidad =:trazabilidad ' +
                  'WHERE sContrato = :Contrato And iFolio = :Folio And sIdInsumo = :Insumo and sTrazabilidad =:TrazabilidadAnt ');
                connection.zCommand.Params.ParamByName('Contrato').DataType     := ftString;
                connection.zCommand.Params.ParamByName('Contrato').value        := param_Global_Contrato;
                connection.zCommand.Params.ParamByName('Folio').DataType        := ftInteger;
                connection.zCommand.Params.ParamByName('Folio').value           := anexo_suministro.FieldValues['iFolio'];
                connection.zCommand.Params.ParamByName('Insumo').DataType       := ftString;
                connection.zCommand.Params.ParamByName('Insumo').value          := anexo_psuministro.FieldValues['sIdInsumo'];
                connection.zCommand.Params.ParamByName('Cantidad').DataType     := ftFloat;
                connection.zCommand.Params.ParamByName('Cantidad').value        := tdCantidad.Value;
                connection.zCommand.Params.ParamByName('CantidadAdic').DataType := ftFloat;
                if (tdCantidad.Value > anexo_psuministro.FieldByName('dCantidad').AsFloat) and (tdCantidad.Value > anexo_psuministro.FieldByName('dCantidadRestante').AsFloat) then
                   connection.zCommand.Params.ParamByName('CantidadAdic').value    := tdCantidad.Value - anexo_psuministro.FieldByName('dCantidad').AsFloat
                else
                   connection.zCommand.Params.ParamByName('CantidadAdic').value    := 0;
                connection.zCommand.Params.ParamByName('TrazabilidadAnt').DataType := ftString;
                connection.zCommand.Params.ParamByName('TrazabilidadAnt').value    := trazabilidad;
                connection.zCommand.Params.ParamByName('Trazabilidad').DataType    := ftString;
                connection.zCommand.Params.ParamByName('Trazabilidad').value       := tsTrazabilidad.Text;
                connection.zCommand.ExecSQL;
            end;

        end
      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Avisos de Embarque', 'Al actualizar registro', 0);
        end;
      end;
      SavePlace := anexo_psuministro.GetBookmark;
      anexo_psuministro.Active := False;

      anexo_psuministro.Open;
      try
        anexo_psuministro.GotoBookmark(SavePlace);
      except
      else
        anexo_psuministro.FreeBookmark(SavePlace);
      end;

      tsIdMaterial.ReadOnly := True;
      tdCantidad.ReadOnly := True;
      tdCantidad.Enabled := True;
      Insertar1.Enabled := True;
      Editar1.Enabled := True;
      Registrar1.Enabled := False;
      Can1.Enabled := False;
      Eliminar1.Enabled := True;
      Refresh1.Enabled := True;
      frmBarra1.btnPostClick(Sender);
    end;

  BotonPermiso.permisosBotones(frmBarra1);

end;

procedure Tfrmentradaanex.frmBarra1btnRefreshClick(Sender: TObject);
begin
  anexo_psuministro.Active := False;
  anexo_psuministro.Open;
end;

procedure Tfrmentradaanex.frmBarra2btnAddClick(Sender: TObject);
var
  dFechaFinal: tDate;
  iCheck: Integer;
begin
  try
    OpcButton1 := 'New';
    frmBarra2.btnAddClick(Sender);
    frmBarra1.btnCancel.Click;
    pgControl.ActivePageIndex := 0;
    tdIdFecha.Enabled := True;
    tdFechaAviso.Enabled := True;
    tsIdTipo.ReadOnly := False;
    tsIdProveedor.ReadOnly := False;
    tsReferencia.ReadOnly := False;
    tsOrigen.ReadOnly := False;
    tmComentarios.ReadOnly := False;
    tsNumeroOrden.Enabled := True;
    cbbPlataforma.Enabled := True;

    tiFolio.Value := 0;
    tdIdFecha.Date := Date;
    tdFechaAviso.Date := Date;
    tsReferencia.Text := '';
    tsOrigen.Text := '';
    tmComentarios.Text := '';
    tdCantidad.Value := 0;

    MaxFolio.Active := False;
    MaxFolio.Params.ParamByName('Contrato').DataType := ftString;
    MaxFolio.Params.ParamByName('Contrato').Value := param_global_contrato;
    MaxFolio.Open;

    if MaxFolio.RecordCount > 0 then
      tiFolio.Value := MaxFolio.FieldValues['iFolio'] + 1
    else
      tiFolio.Value := 1;

    tsNumeroOrden.ItemIndex := 0;
    tdFechaAviso.SetFocus;
    //activapop(frmentradaanex, popupprincipal);
    BotonPermiso.permisosBotones(frmBarra2);
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_entradaanex', 'Al iniciar el formulario', 0);
    end;
  end;
end;

procedure Tfrmentradaanex.frmBarra2btnCancelClick(Sender: TObject);
begin
  tdIdFecha.Enabled := False;
  tdFechaAviso.Enabled := False;
  tsNumeroOrden.Enabled := False;
  cbbPlataforma.Enabled := False;
  tsIdTipo.ReadOnly := True;
  tsIdProveedor.ReadOnly := True;
  tsReferencia.ReadOnly := True;
  tsOrigen.ReadOnly := True;
  tmComentarios.ReadOnly := True;
  frmBarra2.btnCancelClick(Sender);
  Grid_Entradas.SetFocus;
  DebeMostrar := False;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure Tfrmentradaanex.frmBarra2btnDeleteClick(Sender: TObject);
var
  lValido: Boolean;
  SavePlace: TBookmark;

begin
  lValido := True;
  if global_grupo <> 'INTEL-CODE' then
  begin
    ReporteDiario.Active := False;
    ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
    ReporteDiario.Params.ParamByName('contrato').Value := param_global_contrato;
    ReporteDiario.Params.ParamByName('turno').DataType := ftString;
    ReporteDiario.Params.ParamByName('turno').Value := global_turno_reporte;
    ReporteDiario.Params.ParamByName('Fecha').DataType := ftDate;
    ReporteDiario.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
    ReporteDiario.Open;
    if ReporteDiario.RecordCount > 0 then
      while not ReporteDiario.Eof and lValido do
      begin
        if ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' then
        begin
          MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
          lValido := False;
        end;
        ReporteDiario.Next
      end
  end;

  if lValido then
    if anexo_suministro.RecordCount > 0 then
      if MessageDlg('Desea eliminar el folio seleccionado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin

        Connection.QryBusca.Active := False;
        Connection.QryBusca.SQL.Clear;
        Connection.QryBusca.SQL.Add('Select * from anexo_psuministro Where sContrato =:Contrato And iFolio = :Folio ');
        Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
        Connection.QryBusca.Params.ParamByName('Contrato').Value := param_Global_Contrato;
        Connection.QryBusca.Params.ParamByName('Folio').DataType := ftInteger;
        Connection.QryBusca.Params.ParamByName('Folio').Value := anexo_suministro.FieldValues['iFolio'];
        Connection.QryBusca.Open;
        if Connection.QryBusca.RecordCount > 0 then
          Connection.QryBusca.First;

        while not Connection.QryBusca.Eof do
        begin
          Connection.qryBusca2.Active := False;
          Connection.QryBusca2.SQL.Clear;
          Connection.QryBusca2.SQL.Add('Select sReferenciaDetalle, sFolioDetalle From bitacoradealcances ' +
            'Where sContrato = :Contrato And sNumeroActividad = :Actividad');
          Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString;
          Connection.QryBusca2.Params.ParamByName('Contrato').Value := param_Global_Contrato;
          Connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString;
          Connection.QryBusca2.Params.ParamByName('Actividad').Value := Connection.QryBusca.FieldValues['sNumeroActividad'];
          Connection.QryBusca2.Open;
          if Connection.QryBusca2.RecordCount > 0 then
          begin
            MessageDlg('No se Puede Borrar el Aviso Embarque. Existen Suministros de Partidas Reportadas', mtError, [mbOk], 0);
            exit;
          end;
          Connection.QryBusca.Next;
        end;

              // Actualizo Kardex del Sistema ....
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
          'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)');
        connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('Contrato').Value := param_Global_Contrato;
        connection.zCommand.Params.ParamByName('Usuario').DataType := ftString;
        connection.zCommand.Params.ParamByName('Usuario').Value := Global_Usuario;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := Date;
        connection.zCommand.Params.ParamByName('Hora').DataType := ftString;
        connection.zCommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now);
        connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString;
        connection.zCommand.Params.ParamByName('Descripcion').Value := 'Eliminación del Aviso de Embarque No. ' + tsReferencia.Text + ' Recibido el día [' + DateToStr(tdFechaAviso.Date) + '] Registrado el [' + DateToStr(tdIdFecha.Date) + '] Usuario [ ' + global_usuario + ']';
        connection.zCommand.Params.ParamByName('Origen').DataType := ftString;
        connection.zCommand.Params.ParamByName('Origen').Value := 'Reporte Diario';
        connection.zCommand.ExecSQL;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Delete from anexo_psuministro where sContrato = :Contrato And iFolio = :Folio');
        connection.zcommand.Params.ParamByName('Contrato').DataType := ftString;
        connection.zcommand.Params.ParamByName('Contrato').value := param_Global_Contrato;
        connection.zcommand.Params.ParamByName('Folio').DataType := ftInteger;
        connection.zcommand.Params.ParamByName('Folio').value := anexo_suministro.FieldValues['iFolio'];
        connection.zCommand.ExecSQL;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Delete from anexo_suministro where sContrato = :Contrato And iFolio = :Folio');
        connection.zcommand.Params.ParamByName('Contrato').DataType := ftString;
        connection.zcommand.Params.ParamByName('Contrato').value := param_Global_Contrato;
        connection.zcommand.Params.ParamByName('Folio').DataType := ftInteger;
        connection.zcommand.Params.ParamByName('Folio').value := anexo_suministro.FieldValues['iFolio'];
        connection.zCommand.ExecSQL;

        SavePlace := anexo_suministro.GetBookmark;
        anexo_suministro.Active := False;
        anexo_suministro.Open;
        try
          anexo_suministro.GotoBookmark(SavePlace);
        except
        else
          anexo_suministro.FreeBookmark(SavePlace);
        end;
      end
end;

procedure Tfrmentradaanex.frmBarra2btnEditClick(Sender: TObject);
var
  lValido: Boolean;
begin
  lValido := True;
  if global_grupo <> 'INTEL-CODE' then
  begin
    ReporteDiario.Active := False;
    ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
    ReporteDiario.Params.ParamByName('contrato').Value := param_global_contrato;
    ReporteDiario.Params.ParamByName('turno').DataType := ftString;
    ReporteDiario.Params.ParamByName('turno').Value := global_turno_reporte;
    ReporteDiario.Params.ParamByName('Fecha').DataType := ftDate;
    ReporteDiario.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
    ReporteDiario.Open;
    if ReporteDiario.RecordCount > 0 then
      while not ReporteDiario.Eof and lValido do
      begin
        if ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' then
        begin
          MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
          lValido := False;
        end;
        ReporteDiario.Next
      end
  end;

  if lValido then
    if anexo_suministro.RecordCount > 0 then
    begin
            // Verificar si ya se ha reportado un consumo de esta partida en específico
      Connection.qryBusca.Active := False;
      Connection.qryBusca.Sql.Text := 'select a.snumeroorden, a.sCantidadDetalle, a.sFolioDetalle, a.sReferenciaDetalle ' +
        'from bitacoradealcances a where a.scontrato = :contrato and ' +
        'concat(CHAR(254), a.sReferenciaDetalle, char(254)) like concat("%", Char(254), :referencia, Char(254), "%") limit 1';
      Connection.qryBusca.ParamByName('contrato').AsString := param_global_contrato;
      Connection.qryBusca.ParamByName('referencia').AsString := anexo_suministro.FieldByName('sReferencia').AsString;
      Connection.qryBusca.Open;
      activapop(frmentradaanex, popupprincipal);
      if Connection.QryBusca.RecordCount = 1 then
      begin
        DebeMostrar := True;
             // MuestraGlobo;
        Label10.Caption := 'Campo cancelado para edición';
        tsReferencia.ReadOnly := True;
      end
      else
      begin
        DebeMostrar := False;
        Label10.Caption := '';
        tsReferencia.ReadOnly := False;
      end;
      Connection.qryBusca.Close;

      OpcButton1 := 'Edit';
      frmBarra2.btnEditClick(Sender);
      pgControl.ActivePageIndex := 0;
      tdIdFecha.Enabled := True;
      tdFechaAviso.Enabled := True;
      tsNumeroOrden.Enabled := True;
      cbbPlataforma.Enabled := True;
      tsIdTipo.ReadOnly := False;
      tsIdProveedor.ReadOnly := False;
      tsOrigen.ReadOnly := False;
      tmComentarios.ReadOnly := False;
      trazabilidad := tsTrazabilidad.Text;
      tiFolio.SetFocus
    end
    else
      MessageDlg('Folio de Entrada Aplicada no se pueden realizar cambios', mtWarning, [mbOk], 0);

  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure Tfrmentradaanex.frmBarra2btnExitClick(Sender: TObject);
begin
  //frmBarra2.btnExitClick(Sender);
  close
end;

procedure Tfrmentradaanex.frmBarra2btnPostClick(Sender: TObject);
var
  SavePlace: TBookmark;
  lValido: Boolean;
  nombres, cadenas: TStringList;
begin
  {Validacion de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('No. de Orden'); nombres.Add('Tipo Movimiento'); nombres.Add('Referencia');
  cadenas.Add(tsNumeroOrden.Text); cadenas.Add(tsIdTipo.Text); cadenas.Add(tsReferencia.Text);
  if not validaTexto(nombres, cadenas, 'Folio', tiFolio.Text) then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
   //Verifica que la fecha recepcion no sea menor que la fecha de aviso
  if tdFechaAviso.Date < tdIdFecha.Date then
  begin
    showmessage('la fecha de recepción es menor a la fecha de captura');
    tdFechaAviso.SetFocus;
    exit;
  end;
  {Continua insercion de datos}
  DebeMostrar := False;

  if OpcButton1 = 'New' then
  begin
    lValido := True;
    if global_grupo <> 'INTEL-CODE' then
    begin
      ReporteDiario.Active := False;
      ReporteDiario.Params.ParamByName('contrato').DataType := ftString;
      ReporteDiario.Params.ParamByName('contrato').Value := param_global_contrato;
      ReporteDiario.Params.ParamByName('turno').DataType := ftString;
      ReporteDiario.Params.ParamByName('turno').Value := global_turno_reporte;
      ReporteDiario.Params.ParamByName('Fecha').DataType := ftDate;
      ReporteDiario.Params.ParamByName('Fecha').Value := tdIdFecha.Date;
      ReporteDiario.Open;
      if ReporteDiario.RecordCount > 0 then
        while not ReporteDiario.Eof and lValido do
        begin
          if ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' then
          begin
            MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
            lValido := False;
          end;
          ReporteDiario.Next
        end
    end;

    if lValido then
    begin
      try
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('INSERT INTO anexo_suministro ( sContrato, iFolio, sNumeroOrden, dIdFecha, dFechaAviso, sIdTipo, sReferencia, sIdProveedor, sOrigen, mComentarios, sIdPlataforma ) ' +
          'VALUES (:Contrato, :Folio, :Orden, :Fecha, :FechaAviso, :Tipo, :Referencia, :Proveedor, :Origen, :Comentarios, :Plataforma )');
        connection.zCommand.params.ParamByName('Contrato').DataType := ftString;
        connection.zCommand.params.ParamByName('Contrato').value := param_Global_Contrato;
        connection.zCommand.params.ParamByName('Folio').DataType := ftInteger;
        connection.zCommand.params.ParamByName('Folio').value := tiFolio.Value;
        connection.zCommand.params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.params.ParamByName('Orden').value := tsNumeroOrden.Text;
        connection.zCommand.params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.params.ParamByName('Fecha').value := tdIdFecha.Date;
        connection.zCommand.params.ParamByName('FechaAviso').DataType := ftDate;
        connection.zCommand.params.ParamByName('FechaAviso').value := tdFechaAviso.Date;
        connection.zCommand.params.ParamByName('Tipo').DataType := ftString;
        connection.zCommand.params.ParamByName('Tipo').value := tsIdTipo.KeyValue;
        connection.zCommand.params.ParamByName('Referencia').DataType := ftString;
        connection.zCommand.params.ParamByName('Referencia').value := tsReferencia.Text;
        connection.zCommand.params.ParamByName('Proveedor').DataType := ftString;
        connection.zCommand.params.ParamByName('Proveedor').value := tsIdProveedor.KeyValue;
        connection.zCommand.params.ParamByName('Plataforma').DataType := ftString;
        connection.zCommand.params.ParamByName('Plataforma').value := cbbPlataforma.Text;
        connection.zCommand.params.ParamByName('Origen').DataType := ftString;
        connection.zCommand.params.ParamByName('Origen').value := tsOrigen.Text;
        connection.zCommand.params.ParamByName('Comentarios').DataType := ftMemo;
        connection.zCommand.params.ParamByName('Comentarios').value := tmCOmentarios.Text;
        connection.zCommand.ExecSQL;

                // Actualizo Kardex del Sistema ....
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
          'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)');
        connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('Contrato').Value := param_Global_Contrato;
        connection.zCommand.Params.ParamByName('Usuario').DataType := ftString;
        connection.zCommand.Params.ParamByName('Usuario').Value := Global_Usuario;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := Date;
        connection.zCommand.Params.ParamByName('Hora').DataType := ftString;
        connection.zCommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now);
        connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString;
        connection.zCommand.Params.ParamByName('Descripcion').Value := 'Registro de Aviso de Embarque No. ' + tsReferencia.Text + ' Recibido el día [' + DateToStr(tdFechaAviso.Date) + '] Registrado el [' + DateToStr(tdIdFecha.Date) + '] Usuario [ ' + global_usuario + ']';
        connection.zCommand.Params.ParamByName('Origen').DataType := ftString;
        connection.zCommand.Params.ParamByName('Origen').Value := 'Reporte Diario';
        connection.zCommand.ExecSQL;

        SavePlace := anexo_suministro.GetBookmark;
        anexo_suministro.Active := False;
        anexo_suministro.Open;
        try
          anexo_suministro.GotoBookmark(SavePlace);
        except
        else
          anexo_suministro.FreeBookmark(SavePlace);
        end;

        tdIdFecha.Enabled := False;
        tdFechaAviso.Enabled := False;
        tsNumeroOrden.Enabled := False;
        cbbPlataforma.Enabled := False;
        tsIdTipo.ReadOnly := True;
        tsIdProveedor.ReadOnly := True;
        tsReferencia.ReadOnly := True;
        tsOrigen.ReadOnly := True;
        tmComentarios.ReadOnly := True;
        frmBarra2.btnCancelClick(Sender);
      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Avisos de Embarque', 'Al actualizar registro', 0);
        end;
      end
    end
  end
  else
    if OpcButton1 = 'Edit' then
    begin
      try
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('UPDATE anexo_suministro SET dIdFecha = :Fecha, dFechaAviso = :FechaAviso, sNumeroOrden = :Orden, sIdTipo = :Tipo, sReferencia = :Referencia, sIdProveedor = :Proveedor, '+
          'sOrigen = :Origen, mComentarios = :Comentarios, sIdPlataforma =:Plataforma ' +
          'WHERE sContrato = :Contrato And iFolio = :Folio ');
        connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('Contrato').value := param_Global_Contrato;
        connection.zCommand.Params.ParamByName('Folio').DataType := ftInteger;
        connection.zCommand.Params.ParamByName('Folio').value := anexo_suministro.FieldValues['iFolio'];
        connection.zCommand.Params.ParamByName('Orden').DataType := ftString;
        connection.zCommand.Params.ParamByName('Orden').value := tsNumeroOrden.Text;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').value := tdIdFecha.Date;
        connection.zCommand.Params.ParamByName('FechaAviso').DataType := ftDate;
        connection.zCommand.Params.ParamByName('FechaAviso').value := tdFechaAviso.Date;
        connection.zCommand.Params.ParamByName('Tipo').DataType := ftString;
        connection.zCommand.Params.ParamByName('Tipo').value := tsIdTipo.KeyValue;
        connection.zCommand.Params.ParamByName('Referencia').DataType := ftString;
        connection.zCommand.Params.ParamByName('Referencia').value := tsReferencia.Text;
        connection.zCommand.Params.ParamByName('Proveedor').DataType := ftString;
        connection.zCommand.Params.ParamByName('Proveedor').value := tsIdProveedor.KeyValue;
        connection.zCommand.params.ParamByName('Plataforma').DataType := ftString;
        connection.zCommand.params.ParamByName('Plataforma').value := cbbPlataforma.Text;
        connection.zCommand.Params.ParamByName('Origen').DataType := ftString;
        connection.zCommand.Params.ParamByName('Origen').value := tsOrigen.Text;
        connection.zCommand.Params.ParamByName('Comentarios').DataType := ftMemo;
        connection.zCommand.Params.ParamByName('Comentarios').value := tmCOmentarios.Text;
        connection.zCommand.ExecSQL;

            // Actualizo Kardex del Sistema ....
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
          'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)');
        connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
        connection.zCommand.Params.ParamByName('Contrato').Value := param_Global_Contrato;
        connection.zCommand.Params.ParamByName('Usuario').DataType := ftString;
        connection.zCommand.Params.ParamByName('Usuario').Value := Global_Usuario;
        connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate;
        connection.zCommand.Params.ParamByName('Fecha').Value := Date;
        connection.zCommand.Params.ParamByName('Hora').DataType := ftString;
        connection.zCommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now);
        connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString;
        connection.zCommand.Params.ParamByName('Descripcion').Value := 'Modificación de Aviso de Embarque No. ' + tsReferencia.Text + ' Recibido el día [' + DateToStr(tdFechaAviso.Date) + '] Registrado el [' + DateToStr(tdIdFecha.Date) + '] Usuario [ ' + global_usuario + ']';
        connection.zCommand.Params.ParamByName('Origen').DataType := ftString;
        connection.zCommand.Params.ParamByName('Origen').Value := 'Reporte Diario';
        connection.zCommand.ExecSQL;

        SavePlace := anexo_suministro.GetBookmark;
        anexo_suministro.Active := False;
        anexo_suministro.Open;
        try
          anexo_suministro.GotoBookmark(SavePlace);
        except
        else
          anexo_suministro.FreeBookmark(SavePlace);
        end;
        tdIdFecha.Enabled := False;
        tdFechaAviso.Enabled := False;
        tsNumeroOrden.Enabled := False;
        cbbPlataforma.Enabled := False;
        tsIdTipo.ReadOnly := True;
        tsIdProveedor.ReadOnly := True;
        tsReferencia.ReadOnly := True;
        tsOrigen.ReadOnly := True;
        tmComentarios.ReadOnly := True;
        frmBarra2.btnCancelClick(Sender);
      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Avisos de Embarque', 'Al actualizar registro', 0);
        end;
      end
    end;
  desactivapop(popupprincipal);
  OpcButton1 := '';
  BotonPermiso.permisosBotones(frmBarra2);
end;

procedure Tfrmentradaanex.frmBarra2btnPrinterClick(Sender: TObject);
var
    WbsAnterior, WbsAnteriorPaquete, WbsActual: string;
    num, iNivel: integer;
    ArrayPaquetes: array[1..10, 1..3] of string;
begin
  try
    if anexo_suministro.RecordCount > 0 then
    begin
      pgControl.ActivePageIndex := 0;
     
      if (tsNumeroOrden.Text = ('CONTRATO No. ' + param_global_contrato)) then
        rDiarioFirmas(param_global_contrato, '', global_turno_reporte, anexo_suministro.FieldValues['dFechaAviso'], frmEntradaAnex)
      else
        rDiarioFirmas(param_global_contrato, anexo_suministro.FieldValues['sNumeroOrden'], global_turno_reporte, anexo_suministro.FieldValues['dFechaAviso'], frmEntradaAnex);


      Reporte.Active := False;
      Reporte.SQL.Clear;
      Reporte.SQL.Add('Select a1.*, a2.sNumeroOrden, a2.dIdFecha, a2.dFechaAviso, a2.sIdTipo, a2.sReferencia, a2.sIdProveedor, a2.sOrigen, a2.mComentarios, a2.sIdPlataforma, i.sMedida, i.mDescripcion '+
                      'from anexo_psuministro a1 '+
                      'inner join anexo_suministro a2  on (a1.sContrato = a2.sContrato And a1.iFolio = a2.iFolio) '+
                      'inner join actividadesxanexo i on (i.sContrato = :contratoBarco and i.sNumeroActividad = a1.sIdInsumo ) '+
                      'Where a1.sContrato =:contrato And a1.iFolio =:folio ');
      Reporte.Params.ParamByName('Contrato').DataType := ftString;
      Reporte.Params.ParamByName('Contrato').Value := param_global_contrato;
      Reporte.Params.ParamByName('ContratoBarco').DataType := ftString;
      Reporte.Params.ParamByName('ContratoBarco').Value := global_contrato_barco;
      Reporte.Params.ParamByName('Folio').DataType := ftInteger;
      Reporte.Params.ParamByName('Folio').Value := anexo_suministro.FieldValues['iFolio'];
      Reporte.Open;

      rxAviso.EmptyTable;
      while not Reporte.Eof do
      begin
          rxAviso.Append;
          rxAviso.FieldValues['sTipoActividad'] := 'Actividad';
          rxAviso.FieldValues['dCantidad'] := Reporte.FieldValues['dCantidad'];
          rxAviso.FieldValues['sContrato'] := Reporte.FieldValues['sContrato'];
          rxAviso.FieldValues['iFolio'] := Reporte.FieldValues['iFolio'];
          rxAviso.FieldValues['sNumeroOrden'] := Reporte.FieldValues['sNumeroOrden'];
          rxAviso.FieldValues['dIdFecha'] := Reporte.FieldValues['dIdFecha'];
          rxAviso.FieldValues['dFechaAviso'] := Reporte.FieldValues['dFechaAviso'];
          rxAviso.FieldValues['sIdTipo'] := Reporte.FieldValues['sIdTipo'];
          rxAviso.FieldValues['sReferencia'] := Reporte.FieldValues['sReferencia'];
          rxAviso.FieldValues['sIdProveedor'] := Reporte.FieldValues['sIdProveedor'];
          rxAviso.FieldValues['sOrigen'] := Reporte.FieldValues['sOrigen'];
          rxAviso.FieldValues['mDescripcion'] := Reporte.FieldValues['mDescripcion'];
          rxAviso.FieldValues['mComentarios'] := Reporte.FieldValues['mComentarios'];
          rxAviso.FieldValues['sNumeroActividad'] := Reporte.FieldValues['sNumeroActividad'];
          rxAviso.FieldValues['sMedida'] := Reporte.FieldValues['sMedida'];
          rxAviso.FieldValues['sTrazabilidad'] := Reporte.FieldValues['sTrazabilidad'];
          rxAviso.FieldValues['sIdInsumo'] := Reporte.FieldValues['sIdInsumo'];
          rxAviso.Post;

          Reporte.Next;
      end;

      if Reporte.RecordCount > 0 then
      begin
          //Traemos las descripciones y descripciones de contrato de acuerdo al parametro..
          ActualizaConfiguraciones(param_global_contrato);

          frxEntrada.PreviewOptions.MDIChild := False ;
          frxEntrada.PreviewOptions.Modal := True ;
          frxEntrada.PreviewOptions.Maximized := lCheckMaximized () ;
          frxEntrada.PreviewOptions.ShowCaptions := False ;
          frxEntrada.Previewoptions.ZoomMode := zmPageWidth ;
          frxEntrada.LoadFromFile (global_files +global_miReporte+ '_AvisoEmbarque.fr3') ;
          frxentrada.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP)); //listaVerificacion.fr3

          //Regresamos las descripciones y descripciones de contrato de acuerdo al parametro..
          ActualizaConfiguraciones(global_contrato);
      end
      else begin
        showmessage('El aviso de embarque seleccionado no tiene Partidas Anexo');
      end;

    end;
  except
    on e: exception do
    begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Captura de Avisos de Embarque', 'Al imprimir', 0);
    end;
  end;
end;

procedure Tfrmentradaanex.frmBarra2btnRefreshClick(Sender: TObject);
begin
  anexo_suministro.Active := False;
  anexo_suministro.Open;

  MovimientosdeAlmacen.Active := False;
  MovimientosdeAlmacen.Open;

  Proveedores.Active := False;
  Proveedores.Open;

  zqMateriales.Active := False;
  zqMateriales.ParamByName('Contrato').AsString  := global_Contrato_Barco;
  zqMateriales.Open;

  cbbPlataforma.Clear;
  cbbPlataforma.Items.Add('S/N');
  while not qryPlataforma.Eof do
  begin
      cbbPlataforma.Items.Add(qryPlataforma.FieldValues['sIdPlataforma']);
      qryPlataforma.Next;
  end;

  tsNumeroOrden.Items.Clear;
  Connection.qryBusca.Active := False;
  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('Select sNumeroOrden from ordenesdetrabajo where sContrato = :Contrato and ' +
    'cIdStatus = :status order by sNumeroOrden');
  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value := param_Global_Contrato;
  Connection.qryBusca.Params.ParamByName('status').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('status').Value := connection.configuracion.FieldValues['cStatusProceso'];
  Connection.qryBusca.Open;
  if Connection.qryBusca.RecordCount >= 1 then
    while not Connection.qryBusca.Eof do
    begin
      tsNumeroOrden.Items.Add(Connection.qryBusca.FieldValues['sNumeroOrden']);
      Connection.qryBusca.Next
    end;
  tsNumeroOrden.ItemIndex := 0;

end;



procedure Tfrmentradaanex.frxEntradaGetValue(const VarName: string;
  var Value: Variant);
begin
  if CompareText(VarName, 'TIPO_ENTRADA') = 0 then
    Value := tsIdTipo.Text;

  if CompareText(VarName, 'SUPERINTENDENTE') = 0 then
    Value := sSuperIntendente;
  if CompareText(VarName, 'SUPERVISOR') = 0 then
    Value := sSupervisor;
  if CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
    Value := sSupervisorTierra;

  if CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
    Value := sPuestoSuperIntendente;
  if CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
    Value := sPuestoSupervisor;
  if CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
    Value := sPuestoSupervisorTierra;
end;

procedure Tfrmentradaanex.GridPartidasEnter(Sender: TObject);
begin
  if frmBarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;


  if anexo_psuministro.state <> dsInactive then
  begin
    if anexo_psuministro.RecordCount > 0 then
    begin
      tsMedida.Text := anexo_psuministro.FieldValues['sMedida'];
      tsTrazabilidad.Text := anexo_psuministro.FieldValues['sTrazabilidad'];
      tdCantidad.Value := anexo_psuministro.FieldValues['dCantidad'];
    end
    else
    begin

      Historico.Active := False;
      Historico.Params.ParamByName('Contrato').DataType := ftString;
      Historico.Params.ParamByName('Contrato').Value := param_global_contrato;
      Historico.Params.ParamByName('Actividad').DataType := ftString;
      Historico.Params.ParamByName('Actividad').Value := '';
      Historico.Open;

      tsMedida.Text := '';
      tsTrazabilidad.Text := '';
      tdCantidad.Value := 0;
    end
  end;
end;

procedure Tfrmentradaanex.GridPartidasTitleBtnClick(Sender: TObject;
  ACol: Integer; Field: TField);
var
  sCampo: string;
begin
    sCampo := Field.FieldName;
    anexo_psuministro.Active := False;
    anexo_psuministro.Params.ParamByName('Contrato').DataType := ftString;
    anexo_psuministro.Params.ParamByName('Contrato').Value := global_Contrato_Barco;
    anexo_psuministro.Params.ParamByName('Orden').DataType := ftString;
    anexo_psuministro.Params.ParamByName('Orden').Value    := param_global_contrato;
    anexo_psuministro.Params.ParamByName('Folio').DataType := ftInteger;
    anexo_psuministro.Params.ParamByName('Folio').Value := anexo_suministro.FieldValues['iFolio'];
    anexo_psuministro.Open;

  anexo_psuministro.Open;
end;

procedure Tfrmentradaanex.GridPartidasTitleClick(Column: TColumn);
begin
  UtGrid2.DbGridTitleClick(Column);
end;

procedure Tfrmentradaanex.Grid_EntradasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x, y);
end;

procedure Tfrmentradaanex.Grid_EntradasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender, Button, Shift, X, Y);
end;

procedure Tfrmentradaanex.Grid_EntradasTitleBtnClick(Sender: TObject;
  ACol: Integer; Field: TField);
var
  sOrdenado: string;
begin
  sOrdenado := Field.FieldName;
  anexo_suministro.Active := False;
  anexo_suministro.Params.ParamByName('Contrato').DataType := ftString;
  anexo_suministro.Params.ParamByName('Contrato').Value := param_global_contrato;
  anexo_suministro.Open;
end;

procedure Tfrmentradaanex.Grid_EntradasTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure Tfrmentradaanex.Grid_HistoricoDblClick(Sender: TObject);
begin
  PanelHistorico.Visible := not PanelHistorico.Visible;
end;

procedure Tfrmentradaanex.Insertar1Click(Sender: TObject);
begin
  frmBarra2.btnAdd.Click
end;

procedure Tfrmentradaanex.KardexdelInventario1Click(Sender: TObject);
begin
  PanelHistorico.Visible := not PanelHistorico.Visible;
end;

procedure Tfrmentradaanex.tdCantidadChange(Sender: TObject);
begin
  TRxCalcEditChangef(tdCantidad, 'Cantidad');
end;

procedure Tfrmentradaanex.tdCantidadEnter(Sender: TObject);
begin
  tdCantidad.Color := global_color_entrada
end;

procedure Tfrmentradaanex.tdCantidadExit(Sender: TObject);
begin
  tdCantidad.Color := global_color_salida
end;

procedure Tfrmentradaanex.tdCantidadKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTRxCalcEdit(tdCantidad, key) then
    key := #0;
  if Key = #13 then
    frmBarra1.btnPost.SetFocus
end;

procedure Tfrmentradaanex.tdFechaAvisoEnter(Sender: TObject);
begin
  tdFechaAviso.Color := global_color_entrada
end;

procedure Tfrmentradaanex.tdFechaAvisoExit(Sender: TObject);
begin
  if frmBarra2.btnCancel.Enabled = True then
    if tsReferencia.Text = '' then
      tsReferencia.Text := 'CAL' + FormatDateTime('yymmdd', tdFechaAviso.Date);
  tdFechaAviso.Color := global_color_salida
end;

procedure Tfrmentradaanex.tdFechaAvisoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsReferencia.SetFocus
end;

procedure Tfrmentradaanex.tdIdFechaEnter(Sender: TObject);
begin
  tdIdFecha.Color := global_color_entrada
end;

procedure Tfrmentradaanex.tdIdFechaExit(Sender: TObject);
begin
  tdIdFecha.Color := global_color_salida
end;


procedure Tfrmentradaanex.tdIdFechaKeyPress(Sender: TObject; var Key: Char);
begin

  if Key = #13 then
    tdFechaAviso.SetFocus
end;

procedure Tfrmentradaanex.tiFolioChange(Sender: TObject);
begin
  TCurrenCyEdit(sender).Value := abs(TCurrenCyEdit(sender).Value);
end;


procedure Tfrmentradaanex.tm1Enter(Sender: TObject);
begin
   tm1.Color := global_color_entrada;
end;

procedure Tfrmentradaanex.tm1Exit(Sender: TObject);
begin
   tm1.Color := global_color_salida;
end;

procedure Tfrmentradaanex.tsIdProveedorEnter(Sender: TObject);
begin
  tsIdProveedor.Color := global_color_entrada
end;

procedure Tfrmentradaanex.tsIdProveedorExit(Sender: TObject);
begin
  tsIdProveedor.Color := global_color_salida
end;

procedure Tfrmentradaanex.tsIdProveedorKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsOrigen.SetFocus
end;

procedure Tfrmentradaanex.tsIdTipoEnter(Sender: TObject);
begin
  tsIdTipo.Color := global_color_entrada
end;

procedure Tfrmentradaanex.tsIdTipoExit(Sender: TObject);
begin
  tsIdTipo.Color := global_color_salida
end;

procedure Tfrmentradaanex.tsIdTipoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsOrigen.SetFocus
end;


procedure Tfrmentradaanex.ts1Enter(Sender: TObject);
begin
    ts1.Color := global_color_entrada;
end;

procedure Tfrmentradaanex.ts1Exit(Sender: TObject);
begin
   ts1.Color := global_color_salida;
end;

procedure Tfrmentradaanex.ts1KeyPress(Sender: TObject; var Key: Char);
begin
   if Key = #13  then
      tsTrazabilidadInsumo.SetFocus

end;

procedure Tfrmentradaanex.tsIdLabelEnter(Sender: TObject);
begin
    tsIdLabel.Color := global_color_entrada;
end;

procedure Tfrmentradaanex.tsIdLabelExit(Sender: TObject);
begin
        tsIdLabel.Color := global_color_salida;
end;

procedure Tfrmentradaanex.tsIdLabelKeyPress(Sender: TObject; var Key: Char);
begin
   if Key = #13 then
      tsTrazabilidadInsumo.SetFocus;
end;

procedure Tfrmentradaanex.tsIdMaterialEnter(Sender: TObject);
begin
    imgNotas.Visible := False;
    tsIdMaterial.Color := global_Color_Entrada
end;

procedure Tfrmentradaanex.tsIdMaterialExit(Sender: TObject);
begin
   tsIdMaterial.Color := global_Color_salida;
end;

procedure Tfrmentradaanex.tsIdMaterialKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
  begin

    Grid_iguales.SetFocus;
  end;

end;

procedure Tfrmentradaanex.tsNumeroOrdenEnter(Sender: TObject);
begin
  tsNumeroOrden.Color := global_color_entrada
end;

procedure Tfrmentradaanex.tsNumeroOrdenExit(Sender: TObject);
begin
    if tsNumeroOrden.Text <> 'S/N' then
    begin
        //Consultaos la plataforma donde se encuntra el folio..
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select * from ordenesdetrabajo where sContrato =:Contrato and sNumeroOrden =:Folio ');
        connection.zCommand.ParamByName('contrato').AsString := param_global_contrato;
        connection.zCommand.ParamByName('Folio').AsString    := tsNumeroOrden.Text;
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
           cbbPlataforma.ItemIndex := cbbPlataforma.Items.IndexOf( connection.zCommand.FieldValues['sIdPlataforma']);
    end
    else
       cbbPlataforma.ItemIndex := cbbPlataforma.Items.IndexOf('S/N');

    tsNumeroOrden.Color := global_color_salida;
end;

procedure Tfrmentradaanex.tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsIdTipo.SetFocus
end;

procedure Tfrmentradaanex.tsOrigenKeyPress(Sender: TObject; var Key: Char);
begin
   if key =#13 then
      tmComentarios.SetFocus
end;

procedure Tfrmentradaanex.tsReferenciaEnter(Sender: TObject);
begin
  tsReferencia.Color := global_color_entrada;
end;

procedure Tfrmentradaanex.tsReferenciaExit(Sender: TObject);
begin
  tsReferencia.Color := global_color_salida
end;

procedure Tfrmentradaanex.tsReferenciaKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsIdTipo.SetFocus
end;

procedure Tfrmentradaanex.tsTrazabilidadEnter(Sender: TObject);
begin
    tsTrazabilidad.Color := global_color_entrada;
end;

procedure Tfrmentradaanex.tsTrazabilidadExit(Sender: TObject);
begin
       tsTrazabilidad.Color := global_color_salida;
end;

procedure Tfrmentradaanex.tsTrazabilidadInsumoEnter(Sender: TObject);
begin
    tsTrazabilidadInsumo.Color := global_color_entrada;
end;

procedure Tfrmentradaanex.tsTrazabilidadInsumoExit(Sender: TObject);
begin
    tsTrazabilidadInsumo.Color := global_color_salida;
end;

procedure Tfrmentradaanex.tsTrazabilidadInsumoKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key = #13 then
       btnAceptar.SetFocus
end;

procedure Tfrmentradaanex.tsTrazabilidadKeyPress(Sender: TObject;
  var Key: Char);
begin
     if key = #13 then
        tdCantidad.SetFocus;
end;

function TfrmEntradaAnex.lExisteActividad(sActividad: string): Boolean;
var
  sBotonPresionado: string;
begin
 // sSwbs := '';
  connection.qryBusca.Active := False;
  connection.qryBusca.SQL.Clear;
  connection.qryBusca.SQL.Add('select mDescripcion, dCantidadAnexo,swbs, sMedida from actividadesxanexo where sContrato = :Contrato ' +
    'And sIdConvenio = :Convenio and sNumeroActividad = :Actividad');
  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value := param_global_contrato;
  Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Convenio').Value := convenio_reporte;
  Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Actividad').Value := sActividad;
  connection.qryBusca.Open;
  if connection.qryBusca.RecordCount > 0 then
  begin
    if sSwbs <> '' then
      lExisteActividad := True
    else
    begin
      if connection.qryBusca.RecordCount > 1 then
      begin
        while sSwbs = '' do
        begin
          {si hay mas de un concepto con el mismo numero de actividad, permitir seleccionar cual es la deseada.}
          Application.CreateForm(TfrmBusquedaGeneralizada, frmBusquedaGeneralizada);
          frmBusquedaGeneralizada.Titulo.Caption := 'PARTIDAS EFECTIVAS';
          frmBusquedaGeneralizada.Qry.Active := false;
          frmBusquedaGeneralizada.Qry.SQL.Clear;
          frmBusquedaGeneralizada.Qry.SQL.Add('select a.sNumeroActividad as Id,a.dCantidadAnexo as CantidadAnexo, a.sMedida as Medida,' +
            '  (select substr(ta.mDescripcion,1,400) as Paquete from actividadesxanexo ta where ta.sContrato=a.sContrato and ta.sWbs=a.sWbsAnterior and ' +
            ' ta.sIdConvenio=a.sIdConvenio  ) as Paquete ,  ' +
            ' a.sWbs as Wbs,substr(a.mDescripcion,1,400) as Concepto ' +
            '  from actividadesxanexo a where a.sContrato = :Contrato ' +
            ' And a.sIdConvenio = :Convenio and a.sNumeroActividad = :Actividad ');
          frmBusquedaGeneralizada.Qry.Params.ParamByName('contrato').DataType := ftString;
          frmBusquedaGeneralizada.Qry.Params.ParamByName('contrato').Value := param_global_contrato;
          frmBusquedaGeneralizada.Qry.Params.ParamByName('convenio').DataType := ftString;
          frmBusquedaGeneralizada.Qry.Params.ParamByName('Convenio').Value := convenio_reporte;
          frmBusquedaGeneralizada.Qry.Params.ParamByName('Actividad').DataType := ftString;
          frmBusquedaGeneralizada.Qry.Params.ParamByName('Actividad').Value := sActividad;
          frmBusquedaGeneralizada.Qry.Open;
          frmBusquedaGeneralizada.ShowModal;

          sBotonPresionado := frmBusquedaGeneralizada.BotonPulsado;

          if sBotonPresionado = 'ACEPTAR' then
          begin
            tsTrazabilidad.Text := frmBusquedaGeneralizada.Qry.FieldValues['sTrazabilidad'];
            sSwbs := frmBusquedaGeneralizada.Qry.FieldValues['Wbs'];
          end;
          frmBusquedaGeneralizada.Destroy;
          if sSwbs = '' then
            MessageDlg('Debe seleccionar una actividad, de lo contrario no podra salir de esta ventana', mtInformation, [mbOk], 0)
          else
            lExisteActividad := True;
        end;
      end
      else
      begin {si solo es una partida en el anexo, no preguntar y tomar la unica que existe}
        tsTrazabilidad.Text := connection.qryBusca.FieldValues['sTrazabilidad'];
        sSwbs := connection.qryBusca.FieldByName('swbs').AsString;
        lExisteActividad := True;
      end;
    end;
  end
  else
  begin
    tsTrazabilidad.Text := '';
    lExisteActividad := False
  end
end;

procedure Tfrmentradaanex.NuevoMaterialClick(Sender: TObject);
var
   numero : Integer;
   id : string;
begin
    //BUSCAMOS SI EXISTE EL MATERIAL..
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select max(sIdInsumo) as sIdMaterial FROM insumos Where sContrato = :contrato and sIdInsumo like "MA.%" group by sContrato ');
    Connection.zCommand.Params.ParamByName('Contrato').DataType := ftString;
    Connection.zCommand.Params.ParamByName('Contrato').Value    := Global_Contrato_barco;
    connection.zCommand.Open;

    //Sino existe lo damos de alta..
    if connection.zCommand.RecordCount > 0 then
       Id := connection.zCommand.FieldValues['sIdMaterial'];

    try
       numero := StrToInt(copy(Id, pos('.', Id) + 1, length(Id)));
       if numero >= 10000 then
           inc(numero)
       else
           numero := 10000;

       tsIdLabel.Text := 'MA.'+IntToStr(numero);
    Except
        tsIdLabel.Text := 'S/N';
    end;

    PanelMateriales.Visible := True;
    PanelMateriales.Height  := 218;
    PanelMateriales.Width   := 509;
    tm1.Text := '';
    ts1.Text := '';
    tsTrazabilidadInsumo.Text := '';
    tm1.SetFocus;

    if zqProveedores.RecordCount > 0 then
    begin
        zqProveedores.First;
        cmbProveedores.KeyValue := zqProveedores.FieldByName('sIdProveedor').AsString;
    end;
    
end;

procedure Tfrmentradaanex.Paste1Click(Sender: TObject);
begin
  UtGrid.AddRowsFromClip;
end;

procedure Tfrmentradaanex.QryImgAvisosAfterScroll(DataSet: TDataSet);
var
   bS  : TStream;
   Pic : TJpegImage;
   BlobField : tField ;
begin
    tiRegistro.Value := QryImgAvisos.RecNo ;
//    bImagen.Picture.LoadFromFile('') ;
    If QryImgAvisos.RecordCount > 0 then
    Begin
        BlobField := QryImgAvisos.FieldByName('bImagen') ;
        BS := QryImgAvisos.CreateBlobStream ( BlobField , bmRead ) ;
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
        End
    End
end;

procedure Tfrmentradaanex.Refresh1Click(Sender: TObject);
begin
  frmBarra2.btnRefresh.Click
end;

procedure Tfrmentradaanex.Registrar1Click(Sender: TObject);
begin
  frmBarra2.btnPost.Click
end;

procedure Tfrmentradaanex.SoporteAvisoEmbarque1Click(Sender: TObject);
begin
  Panel.Visible := Not Panel.Visible
end;


procedure Tfrmentradaanex.SpeedButton1Click(Sender: TObject);
begin
    try
      param_global_contrato := param_Global_Contrato ;
      Application.CreateForm(TfrmSalidaAlmacen, frmSalidaAlmacen);
      frmSalidaAlmacen.ShowModal;
       Close;
    except
        MSG_OK('No se puede acceder a las Salidas de Materiales!');
    end;
end;

end.

