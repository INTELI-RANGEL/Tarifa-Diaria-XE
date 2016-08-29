unit Frm_BuscaPersonal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint,
  dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide,
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
  cxNavigator, DB, cxDBData, StdCtrls, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid, ExtCtrls,
  NxPageControl, ZAbstractRODataset, ZAbstractDataset, ZDataset, cxTextEdit,
  cxCheckBox,frm_Listado_Personal, DBCtrls, Mask, Grids, DBGrids, frm_barra,
  UnitExcepciones, AdvCombo, rxToolEdit, RXDBCtrl;

type
  TFrmBuscaPersonal = class(TForm)
    NxPCPersonal: TNxPageControl;
    NxTabSheet1: TNxTabSheet;
    NxTabSheet2: TNxTabSheet;
    Panel1: TPanel;
    cxDbGridListadoDBTable: TcxGridDBTableView;
    cxgrdListadoLevel1: TcxGridLevel;
    cxgrdListado: TcxGrid;
    QListado: TZQuery;
    dsListado: TDataSource;
    cxDbGridListadoDBTableColumn1: TcxGridDBColumn;
    cxDbGridListadoDBTableColumn2: TcxGridDBColumn;
    cxDbGridListadoDBTableColumn3: TcxGridDBColumn;
    cxDbGridListadoDBTableColumn4: TcxGridDBColumn;
    cxDbGridListadoDBTableColumn5: TcxGridDBColumn;
    cxDbGridListadoDBTableColumn6: TcxGridDBColumn;
    cxDbGridListadoDBTableColumn7: TcxGridDBColumn;
    Panel2: TPanel;
    btnCerrar: TButton;
    btnAgregar: TButton;
    Panel3: TPanel;
    btnNuevo: TButton;
    QExterno: TZQuery;
    pnlContenedor: TPanel;
    zq_Esp: TZQuery;
    ds_Esp: TDataSource;
    ds_categoria: TDataSource;
    zq_categoria: TZQuery;
    zq_compania: TZQuery;
    ds_compania: TDataSource;
    ds_listadoper: TDataSource;
    zq_listadoper: TZQuery;
    Panel4: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    edtFicha: TDBEdit;
    edtNombre: TDBEdit;
    edtApellidoP: TDBEdit;
    edtApellidoM: TDBEdit;
    edtRfc: TDBEdit;
    lkcbCompania: TDBLookupComboBox;
    lkcbEsp: TDBLookupComboBox;
    lkcbCategoria: TDBLookupComboBox;
    Panel5: TPanel;
    frmBarra1: TfrmBarra;
    GridListPer: TDBGrid;
    Panel6: TPanel;
    edtBuscar: TEdit;
    cbbBuscar: TAdvComboBox;
    btnBuscar: TButton;
    Label9: TLabel;
    edtLibreta: TDBEdit;
    dedtVigencia: TDBDateEdit;
    Label10: TLabel;
    procedure FormShow(Sender: TObject);
    procedure btnAgregarClick(Sender: TObject);
    procedure btnCerrarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure cxDbGridListadoDBTableDblClick(Sender: TObject);
    procedure cxDbGridListadoDBTableCellDblClick(Sender: TcxCustomGridTableView;
      ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
      AShift: TShiftState; var AHandled: Boolean);
    procedure cxDbGridListadoDBTableEditDblClick(Sender: TcxCustomGridTableView;
      AItem: TcxCustomGridTableItem; AEdit: TcxCustomEdit);
    procedure btnNuevoClick(Sender: TObject);
    procedure NxPCPersonalChange(Sender: TObject);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure edtFichaKeyPress(Sender: TObject; var Key: Char);
    procedure edtFichaEnter(Sender: TObject);
    procedure edtFichaExit(Sender: TObject);
    procedure edtNombreKeyPress(Sender: TObject; var Key: Char);
    procedure edtNombreEnter(Sender: TObject);
    procedure edtNombreExit(Sender: TObject);
    procedure edtApellidoPKeyPress(Sender: TObject; var Key: Char);
    procedure edtApellidoPEnter(Sender: TObject);
    procedure edtApellidoPExit(Sender: TObject);
    procedure edtApellidoMKeyPress(Sender: TObject; var Key: Char);
    procedure edtApellidoMEnter(Sender: TObject);
    procedure edtApellidoMExit(Sender: TObject);
    procedure edtRfcKeyPress(Sender: TObject; var Key: Char);
    procedure edtRfcEnter(Sender: TObject);
    procedure edtRfcExit(Sender: TObject);
    procedure lkcbCategoriaKeyPress(Sender: TObject; var Key: Char);
    procedure lkcbCategoriaEnter(Sender: TObject);
    procedure lkcbCategoriaExit(Sender: TObject);
    procedure lkcbEspKeyPress(Sender: TObject; var Key: Char);
    procedure lkcbEspEnter(Sender: TObject);
    procedure lkcbEspExit(Sender: TObject);
    procedure lkcbCompaniaEnter(Sender: TObject);
    procedure lkcbCompaniaExit(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure zq_listadoperAfterScroll(DataSet: TDataSet);
    procedure ds_categoriaDataChange(Sender: TObject; Field: TField);
    procedure cbbBuscarKeyPress(Sender: TObject; var Key: Char);
    procedure edtBuscarKeyPress(Sender: TObject; var Key: Char);
    procedure btnBuscarClick(Sender: TObject);
    procedure lkcbCompaniaKeyPress(Sender: TObject; var Key: Char);
    procedure edtLibretaKeyPress(Sender: TObject; var Key: Char);
    procedure edtLibretaEnter(Sender: TObject);
    procedure edtLibretaExit(Sender: TObject);
    procedure dedtVigenciaEnter(Sender: TObject);
    procedure dedtVigenciaExit(Sender: TObject);
  private
    { Private declarations }
    procedure cambio_estado;
  public
    { Public declarations }
    ParamFecha:TDate;
    sCategoria:String;
  end;

var
  FrmBuscaPersonal: TFrmBuscaPersonal;

implementation

uses frm_connection, global;

{$R *.dfm}

procedure TFrmBuscaPersonal.cambio_estado;
begin
  if zq_listadoper.State in [dsInsert,dsEdit] then
  begin
    frmBarra1.btnAdd.Enabled      :=False;
    frmBarra1.btnEdit.Enabled        :=False;
    frmBarra1.btnPost.Enabled       :=True;
    frmBarra1.btnCancel.Enabled      :=True;
    frmBarra1.btnDelete.Enabled      :=False;
    frmBarra1.btnPrinter.Enabled      :=False;
    frmBarra1.btnRefresh.Enabled     :=False;
    frmBarra1.btnExit.Enabled         :=False;
    GridListPer.Enabled    :=False;
  end
  else
  if zq_listadoper.State in [dsBrowse] then
  begin
    frmBarra1.btnAdd.Enabled      :=True;
    frmBarra1.btnEdit.Enabled        :=True;
    frmBarra1.btnPost.Enabled       :=False;
    frmBarra1.btnCancel.Enabled      :=False;
    frmBarra1.btnDelete.Enabled      :=True;
    frmBarra1.btnPrinter.Enabled      :=True;
    frmBarra1.btnRefresh.Enabled     :=True;
    frmBarra1.btnExit.Enabled         :=True;
    GridListPer.Enabled    :=True;
  end;
end;

procedure TFrmBuscaPersonal.cbbBuscarKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    edtbuscar.SetFocus
end;

procedure TFrmBuscaPersonal.btnAgregarClick(Sender: TObject);
begin
  QExterno.Append;
  if connection.contrato.FieldByName('sTipoObra').AsString='BARCO' then
    QExterno.fieldByname('sContrato').AsString:=global_contrato
  else
    QExterno.fieldByname('scontrato').AsString:= global_contrato_Barco;

  QExterno.fieldByname('sIdTurno').AsString:=Global_turno;
  QExterno.fieldByname('dIdFecha').AsdateTime:=ParamFecha;
  QExterno.fieldByname('sIdCategoria').AsString:=sCategoria;
  QExterno.fieldByname('sIdTripulacion').AsString:=QListado.FieldByName('sIdTripulacion').AsString;
  QExterno.fieldByname('sOrden').AsString:='';
  QExterno.fieldByname('sNombre').AsString:=QListado.FieldByName('NombreCompleto').AsString;
  QExterno.fieldByname('sIdpersonal').AsString:=QListado.FieldByName('sIdPersonal').AsString;
  QExterno.fieldByname('sDescripcion').AsString:=QListado.FieldByName('Categoria').AsString;
  QExterno.fieldByname('sNumeroCabina').AsString:='';
  QExterno.fieldByname('sNacionalidad').AsString:='';
  QExterno.fieldByname('sIdCompania').AsString:=QListado.FieldByName('sIdCompania').AsString;
  QExterno.fieldByname('sRfc').AsString:=QListado.FieldByName('sRfc').AsString;
  QExterno.fieldByname('sIdCuenta').AsString:='1';
  QExterno.fieldByname('iNacionales').AsInteger:=1;
  QExterno.fieldByname('iExtranjeros').AsInteger:=0;
  QExterno.fieldByname('Compania').AsString:=QListado.FieldByName('Compania').AsString;
  QExterno.fieldByname('sidPernocta').AsString:='';
  QExterno.fieldByname('pernocta').AsString:='';
  QExterno.fieldByname('lImprimeListado').AsString:='Si';
  //QExterno.fieldByname('lPernocta').AsString:='Si';
  //Compania,c.sidPernocta,c.sdescripcion as pernocta


  QExterno.post;
  QListado.Refresh;
end;

procedure TFrmBuscaPersonal.btnBuscarClick(Sender: TObject);
begin
  if zq_listadoper.RecordCount>0 then
    if cbbBuscar.ItemIndex=0 then
    begin
      if not zq_listadoper.Locate('sIdTripulacion',trim(edtBuscar.Text),[LoCaseInsensitive]) then
        MessageDlg('No se Encontro la Ficha Buscada',mtInformation,[MbOk],0);
    end
    else
    begin
      if not zq_listadoper.Locate('sNombre',trim(edtBuscar.Text),[LoCaseInsensitive]) then
        MessageDlg('No se Encontro el Nombre Buscado',mtInformation,[MbOk],0);
    end;
end;

procedure TFrmBuscaPersonal.btnCerrarClick(Sender: TObject);
begin
  close;
end;

procedure TFrmBuscaPersonal.btnNuevoClick(Sender: TObject);
begin
//  frmListado_Personal: TfrmListado_Personal
  NxPcPersonal.ActivePageIndex:=1;
  frmBarra1.btnAdd.Click;


end;

procedure TFrmBuscaPersonal.cxDbGridListadoDBTableCellDblClick(
  Sender: TcxCustomGridTableView; ACellViewInfo: TcxGridTableDataCellViewInfo;
  AButton: TMouseButton; AShift: TShiftState; var AHandled: Boolean);
begin
  if ssCtrl in AShift then
    btnAgregarClick(Sender);
end;

procedure TFrmBuscaPersonal.cxDbGridListadoDBTableDblClick(Sender: TObject);
begin
  //btnAgregarClick(Sender);
end;

procedure TFrmBuscaPersonal.cxDbGridListadoDBTableEditDblClick(
  Sender: TcxCustomGridTableView; AItem: TcxCustomGridTableItem;
  AEdit: TcxCustomEdit);
begin
 // btnAgregarClick(Sender);
end;

procedure TFrmBuscaPersonal.dedtVigenciaEnter(Sender: TObject);
begin
  dedtVigencia.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.dedtVigenciaExit(Sender: TObject);
begin
  dedtVigencia.color := global_color_salida
end;

procedure TFrmBuscaPersonal.ds_categoriaDataChange(Sender: TObject;
  Field: TField);
begin
  zq_Esp.Active:=False;
  zq_Esp.ParamByName('Contrato').AsString:=global_contrato_barco;
  zq_Esp.ParamByName('TipoPer').AsString:=zq_categoria.FieldByName('sIdTipoPersonal').AsString;
  zq_Esp.ParamByName('Per').AsString:='-1';
  zq_Esp.Open;
end;

procedure TFrmBuscaPersonal.edtApellidoMEnter(Sender: TObject);
begin
  edtApellidoM.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.edtApellidoMExit(Sender: TObject);
begin
  edtApellidoM.color := global_color_salida
end;

procedure TFrmBuscaPersonal.edtApellidoMKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    edtRfc.SetFocus
end;

procedure TFrmBuscaPersonal.edtApellidoPEnter(Sender: TObject);
begin
  edtApellidoP.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.edtApellidoPExit(Sender: TObject);
begin
  edtApellidoP.color := global_color_salida
end;

procedure TFrmBuscaPersonal.edtApellidoPKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    edtApellidoM.SetFocus
end;

procedure TFrmBuscaPersonal.edtBuscarKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
  begin
    btnBuscar.SetFocus;
    btnBuscar.Click;
  end;
end;

procedure TFrmBuscaPersonal.edtFichaEnter(Sender: TObject);
begin
  edtFicha.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.edtFichaExit(Sender: TObject);
begin
  edtFicha.color := global_color_salida
end;

procedure TFrmBuscaPersonal.edtFichaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    edtNombre.SetFocus
end;

procedure TFrmBuscaPersonal.edtLibretaEnter(Sender: TObject);
begin
  edtLibreta.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.edtLibretaExit(Sender: TObject);
begin
  edtLibreta.color := global_color_salida
end;

procedure TFrmBuscaPersonal.edtLibretaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    dedtVigencia.SetFocus
end;

procedure TFrmBuscaPersonal.edtNombreEnter(Sender: TObject);
begin
  edtNombre.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.edtNombreExit(Sender: TObject);
begin
  edtNombre.color := global_color_salida
end;

procedure TFrmBuscaPersonal.edtNombreKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    edtApellidoP.SetFocus
end;

procedure TFrmBuscaPersonal.edtRfcEnter(Sender: TObject);
begin
  edtRfc.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.edtRfcExit(Sender: TObject);
begin
  edtRfc.color := global_color_salida
end;

procedure TFrmBuscaPersonal.edtRfcKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    lkcbCategoria.SetFocus
end;

procedure TFrmBuscaPersonal.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  QExterno:=nil;
end;

procedure TFrmBuscaPersonal.FormCreate(Sender: TObject);
begin
  NxPCPersonal.ShowTabs:=false;
  NxPCPersonal.ActivePageIndex:=0;
end;

procedure TFrmBuscaPersonal.FormShow(Sender: TObject);
begin
  QListado.Active:=false;
  if connection.contrato.FieldByName('sTipoObra').AsString='BARCO' then
    QListado.ParamByName('Contrato').AsString:=global_contrato
  else
    QListado.ParamByName('contrato').AsString:= global_contrato_Barco;

  QListado.ParamByName('fecha').AsDate:=ParamFecha;
  QListado.Open;
end;

procedure TFrmBuscaPersonal.frmBarra1btnAddClick(Sender: TObject);
begin
  //frmBarra1.btnAddClick(Sender);
  zq_listadoper.Append;
  zq_listadoper.FieldByName('sContrato').AsString:=global_contrato_barco;
  zq_listadoper.FieldByName('sLibretadeMar').AsString:='';
  cambio_estado;
  frmBarra1.btnAddClick(Sender);
  edtFicha.SetFocus;
end;

procedure TFrmBuscaPersonal.frmBarra1btnCancelClick(Sender: TObject);
begin
//  frmBarra1.btnCancelClick(Sender);
  zq_listadoper.Cancel;
  cambio_estado;
  frmBarra1.btnCancelClick(Sender);
end;

procedure TFrmBuscaPersonal.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If zq_listadoper.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Seleccionado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        zq_listadoper.Delete ;
      except
        on e : exception do
        begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Categorias de Personal', 'Al eliminar registro', 0);
        end;
      end
    end
end;

procedure TFrmBuscaPersonal.frmBarra1btnEditClick(Sender: TObject);
begin
//  frmBarra1.btnEditClick(Sender);
  if zq_listadoper.RecordCount>0 then
  begin
    zq_listadoper.Edit;
    cambio_estado;
    frmBarra1.btnEditClick(Sender);
  end;
end;

procedure TFrmBuscaPersonal.frmBarra1btnExitClick(Sender: TObject);
begin
  //frmBarra1.btnExitClick(Sender);
  Qlistado.Refresh;
  NxPcPersonal.ActivePageIndex:=0;
  try
    Qlistado.locate('sIdTripulacion',zq_listadoper.FieldByName('sIdTripulacion').AsString,[loCaseInsensitive]);
  except

  end;
end;

procedure TFrmBuscaPersonal.frmBarra1btnPostClick(Sender: TObject);
begin
//  frmBarra1.btnPostClick(Sender);
  try
    if Length(Trim(edtFicha.Text)) = 0  then
    begin
      MessageDlg('El campo Ficha debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtFicha.CanFocus then
        edtFicha.SetFocus;
      Exit;
    end;
    if Length(Trim(edtNombre.Text)) = 0  then
    begin
      MessageDlg('El campo Nombre debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtNombre.CanFocus then
        edtNombre.SetFocus;
      Exit;
    end;
    if Length(Trim(edtApellidoP.Text)) = 0  then
    begin
      MessageDlg('El campo Apellido Paterno debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtApellidoP.CanFocus then
        edtApellidoP.SetFocus;
      Exit;
    end;
    if Length(Trim(edtApellidoM.Text)) = 0  then
    begin
      MessageDlg('El campo Apellido Materno debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtApellidoM.CanFocus then
        edtApellidoM.SetFocus;
      Exit;
    end;
    if Length(Trim(edtRfc.Text)) = 0  then
    begin
      MessageDlg('El campo Rfc debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtRfc.CanFocus then
        edtRfc.SetFocus;
      Exit;
    end;
    if Length(Trim(lkcbEsp.Text)) = 0  then
    begin
      MessageDlg('El campo Categoria debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if lkcbEsp.CanFocus then
        lkcbEsp.SetFocus;
      Exit;
    end;
    if Length(Trim(lkcbCompania.Text)) = 0  then
    begin
      MessageDlg('El campo Compañia debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if lkcbCompania.CanFocus then
        lkcbCompania.SetFocus;
      Exit;
    end;

    zq_listadoper.Post;
    cambio_estado;
    //frmBarra1.btnPostClick(Sender);
  except
    on e:Exception do
    begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Categorias de Personal', 'Al guardar registro', 0);
      frmbarra1.btnCancel.Click ;
    end;

  end;
end;

procedure TFrmBuscaPersonal.frmBarra1btnRefreshClick(Sender: TObject);
begin
  try
    zq_listadoper.Refresh;
  except
    on e : exception do
    begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Categorias de Personal', 'Al actualizar los Datos',0);    end;
  end;
end;

procedure TFrmBuscaPersonal.lkcbCategoriaEnter(Sender: TObject);
begin
  lkcbCategoria.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.lkcbCategoriaExit(Sender: TObject);
begin
  lkcbCategoria.color := global_color_salida
end;

procedure TFrmBuscaPersonal.lkcbCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    lkcbEsp.SetFocus
end;

procedure TFrmBuscaPersonal.lkcbCompaniaEnter(Sender: TObject);
begin
  lkcbCompania.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.lkcbCompaniaExit(Sender: TObject);
begin
  lkcbCompania.color := global_color_salida
end;

procedure TFrmBuscaPersonal.lkcbCompaniaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    edtLibreta.SetFocus
end;

procedure TFrmBuscaPersonal.lkcbEspEnter(Sender: TObject);
begin
  lkcbEsp.color := global_color_entrada
end;

procedure TFrmBuscaPersonal.lkcbEspExit(Sender: TObject);
begin
  lkcbEsp.color := global_color_salida
end;

procedure TFrmBuscaPersonal.lkcbEspKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    lkcbCompania.SetFocus
end;

procedure TFrmBuscaPersonal.NxPCPersonalChange(Sender: TObject);
begin
  case  NxPcpersonal.ActivePageIndex of
    0:begin
        self.width:=1141;//1053
      end;

    1:begin
        self.width:=700;
        try
          IsOpen:=False;
          zq_listadoper.Active:=False;
          zq_listadoper.Open;

          zq_compania.Active:=False;
          zq_compania.Open;

          zq_categoria.Active:=False;
          zq_categoria.Open;
        finally
          IsOpen:=True;
          zq_listadoper.First;
        end;
      end;

  end;
  self.position:=poOwnerFormCenter;
end;

procedure TFrmBuscaPersonal.zq_listadoperAfterScroll(DataSet: TDataSet);
begin
  if IsOpen then
  begin
    zq_Esp.Active:=False;
    zq_Esp.ParamByName('Contrato').AsString:=global_contrato_barco;
    zq_Esp.ParamByName('TipoPer').AsString:='-1';
    zq_Esp.ParamByName('Per').AsString:=zq_listadoper.FieldByName('sIdPersonal').AsString;
    zq_Esp.Open;

    lkcbCategoria.keyvalue:=zq_Esp.FieldByName('sIdTipoPersonal').AsString;
  end;
end;

end.
