unit frm_firmantes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,frm_CatNomFirmantes,
  Dialogs, frm_connection, StdCtrls, Mask, DBCtrls, frm_barra, Grids,
  DBGrids, ComCtrls, global, DB, Menus,  ExtCtrls, 
  ZDataset, ZAbstractRODataset, ZAbstractDataset, rxToolEdit,
  UnitTBotonesPermisos, UnitExcepciones, udbgrid,UnitValidaTexto,
  unitactivapop, 
  
  DBDateTimePicker, JvExComCtrls, JvDateTimePicker, JvDBDateTimePicker, AdvEdit,
  AdvEdBtn, DBAdvEdBtn, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, dxSkinsCore, dxSkinDevExpressStyle,
  dxSkinOffice2010Silver, dxSkinscxPCPainter, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxEdit, cxNavigator, cxDBData, cxGridCustomTableView,
  cxGridTableView, cxGridDBTableView, cxGridLevel, cxClasses, cxGridCustomView,
  cxGrid, dxSkinFoggy;
  function IsDate(ADate: string): Boolean;
  type
  TfrmFirmas = class(TForm)
    Label1: TLabel;
    frmBarra1: TfrmBarra;
    Label20: TLabel;
    tsNumeroOrden: TDBLookupComboBox;
    ds_ordenesdetrabajo: TDataSource;
    dsFirmas: TDataSource;
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
    Salir1: TMenuItem;
    Firmas: TZQuery;
    ordenesdetrabajo: TZReadOnlyQuery;
    ReporteDiario: TZReadOnlyQuery;
    pgControl: TPageControl;
    TabDependenca: TTabSheet;
    TabContratista: TTabSheet;
    Label10: TLabel;
    Label11: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    tsPuesto5: TDBEdit;
    tsPuesto6: TDBEdit;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    tsPuesto2: TDBEdit;
    tsPuesto3: TDBEdit;
    tsPuesto4: TDBEdit;
    Label14: TLabel;
    Label15: TLabel;
    tsPuesto8: TDBEdit;
    Label2: TLabel;
    Label3: TLabel;
    tsPuesto1: TDBEdit;
    Label12: TLabel;
    Label13: TLabel;
    tsPuesto7: TDBEdit;
    Label16: TLabel;
    Label17: TLabel;
    tsPuesto9: TDBEdit;
    Label18: TLabel;
    Label19: TLabel;
    tsPuesto10: TDBEdit;
    cbActualizarOrdenes: TCheckBox;
    TabRequisiciones: TTabSheet;
    TabOrdenesCompra: TTabSheet;
    TabEntradasSalidas: TTabSheet;
    Label21: TLabel;
    tsPuesto11: TDBEdit;
    Label22: TLabel;
    Label23: TLabel;
    tsPuesto12: TDBEdit;
    Label24: TLabel;
    Label25: TLabel;
    tsPuesto13: TDBEdit;
    Label28: TLabel;
    Label29: TLabel;
    tsPuesto14: TDBEdit;
    Label30: TLabel;
    Label31: TLabel;
    tsPuesto15: TDBEdit;
    Label32: TLabel;
    Label33: TLabel;
    tsPuesto16: TDBEdit;
    Label34: TLabel;
    Label35: TLabel;
    tsPuesto17: TDBEdit;
    Label36: TLabel;
    Label37: TLabel;
    tsPuesto18: TDBEdit;
    Label38: TLabel;
    Label39: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Label43: TLabel;
    Label44: TLabel;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    TabEstimaciones: TTabSheet;
    Label49: TLabel;
    Label50: TLabel;
    Label51: TLabel;
    tsPuesto19: TDBEdit;
    Label52: TLabel;
    Label53: TLabel;
    tsPuesto20: TDBEdit;
    Label54: TLabel;
    Label55: TLabel;
    Label56: TLabel;
    tsPuesto21: TDBEdit;
    Label57: TLabel;
    Label58: TLabel;
    Label59: TLabel;
    Label60: TLabel;
    tsPuesto22: TDBEdit;
    tdIdFecha: TJvDBDateTimePicker;
    ds_turnos: TDataSource;
    QryTurnos: TZQuery;
    Label61: TLabel;
    cmbTurnos: TDBLookupComboBox;
    tsFirma7: TDBAdvEditBtn;
    tsFirma10: TDBAdvEditBtn;
    tsFirma1: TDBAdvEditBtn;
    tsFirma9: TDBAdvEditBtn;
    tsFirma2: TDBAdvEditBtn;
    tsFirma11: TDBAdvEditBtn;
    tsFirma14: TDBAdvEditBtn;
    tsFirma17: TDBAdvEditBtn;
    tsFirma19: TDBAdvEditBtn;
    tsFirma5: TDBAdvEditBtn;
    tsFirma3: TDBAdvEditBtn;
    tsFirma6: TDBAdvEditBtn;
    tsFirma4: TDBAdvEditBtn;
    tsFirma8: TDBAdvEditBtn;
    tsFirma12: TDBAdvEditBtn;
    tsFirma13: TDBAdvEditBtn;
    tsFirma15: TDBAdvEditBtn;
    tsFirma16: TDBAdvEditBtn;
    tsFirma18: TDBAdvEditBtn;
    tsFirma20: TDBAdvEditBtn;
    tsFirma21: TDBAdvEditBtn;
    tsFirma22: TDBAdvEditBtn;
    ts1: TTabSheet;
    cxgrdbtblvwGrid1DBTableView1: TcxGridDBTableView;
    cxgrdlvlGrid1Level1: TcxGridLevel;
    grid_Firmantes: TcxGrid;
    tcxdIdFecha: TcxGridDBColumn;
    tcxsIdTurno: TcxGridDBColumn;
    tcxFirmante1: TcxGridDBColumn;
    tcxFirmante2: TcxGridDBColumn;
    tcxFirmante3: TcxGridDBColumn;
    tcxFirmante4: TcxGridDBColumn;
    tcxFirmante5: TcxGridDBColumn;
    tcxFirmante10: TcxGridDBColumn;
    dbedtsPuesto30: TDBEdit;
    lbl1: TLabel;
    lbl2: TLabel;
    tsFirma30: TDBAdvEditBtn;
    lbl3: TLabel;
    tcxFirmante30: TcxGridDBColumn;
    Label62: TLabel;
    tsCapitan: TDBEdit;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure grid_firmantesEnter(Sender: TObject);
    procedure grid_firmantesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grid_firmantesKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure grid_firmantesCellClick(Column: TColumn);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure tdIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma1KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma2KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma3KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma4KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma5KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma6KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma7KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma8KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma9KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma10KeyPress(Sender: TObject; var Key: Char);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tsPuesto1KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma1Enter(Sender: TObject);
    procedure tsFirma1Exit(Sender: TObject);
    procedure tsPuesto2KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma2Enter(Sender: TObject);
    procedure tsFirma2Exit(Sender: TObject);
    procedure tsPuesto3KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma3Enter(Sender: TObject);
    procedure tsFirma3Exit(Sender: TObject);
    procedure tsPuesto4KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma4Enter(Sender: TObject);
    procedure tsFirma4Exit(Sender: TObject);
    procedure tsPuesto5KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma5Enter(Sender: TObject);
    procedure tsFirma5Exit(Sender: TObject);
    procedure tsPuesto6KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma6Enter(Sender: TObject);
    procedure tsFirma6Exit(Sender: TObject);
    procedure tsPuesto7KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma7Enter(Sender: TObject);
    procedure tsFirma7Exit(Sender: TObject);
    procedure tsPuesto8KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma8Enter(Sender: TObject);
    procedure tsFirma8Exit(Sender: TObject);
    procedure tsPuesto9KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma9Enter(Sender: TObject);
    procedure tsFirma9Exit(Sender: TObject);
    procedure tsPuesto10KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma10Enter(Sender: TObject);
    procedure tsFirma10Exit(Sender: TObject);
    procedure tdIdFechaEnter(Sender: TObject);
    procedure tdIdFechaExit(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure FirmasAfterPost(DataSet: TDataSet);
    procedure tsPuesto11KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma11KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto12KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto13KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma12KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma13KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto14KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto15KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto16KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma14KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma15KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma16KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto17KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto18KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma17KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma18KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma11Enter(Sender: TObject);
    procedure tsFirma12Enter(Sender: TObject);
    procedure tsFirma13Enter(Sender: TObject);
    procedure tsFirma14Enter(Sender: TObject);
    procedure tsFirma15Enter(Sender: TObject);
    procedure tsFirma16Enter(Sender: TObject);
    procedure tsFirma17Enter(Sender: TObject);
    procedure tsFirma18Enter(Sender: TObject);
    procedure tsFirma11Exit(Sender: TObject);
    procedure tsFirma12Exit(Sender: TObject);
    procedure tsFirma13Exit(Sender: TObject);
    procedure tsFirma14Exit(Sender: TObject);
    procedure tsFirma15Exit(Sender: TObject);
    procedure tsFirma16Exit(Sender: TObject);
    procedure tsFirma18Exit(Sender: TObject);
    procedure tsFirma17Exit(Sender: TObject);
    procedure tsPuesto19KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma19KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto20KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma20KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto21KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma21KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto22KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma22KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma20Enter(Sender: TObject);
    procedure tsFirma20Exit(Sender: TObject);
    procedure tsFirma21Exit(Sender: TObject);
    procedure tsFirma22Exit(Sender: TObject);
    procedure tsFirma19Enter(Sender: TObject);
    procedure tsFirma21Enter(Sender: TObject);
    procedure tsFirma22Enter(Sender: TObject);
    procedure tsFirma19Exit(Sender: TObject);
    procedure grid_firmantesMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_firmantesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_firmantesTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure tsFirma7ClickBtn(Sender: TObject);
  private
    procedure TraerNombre(Comp: tdbadveditbtn);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmFirmas: TfrmFirmas;
  Opcion : String ;
  Registro_Actual : String ;
  BotonPermiso: TBotonesPermisos;  
  utgrid:ticdbgrid;
  banderaagregar:boolean;
  
implementation

{$R *.dfm}

procedure TfrmFirmas.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  firmas.Cancel;
  action := cafree;
end;

procedure TfrmFirmas.grid_firmantesEnter(Sender: TObject);
begin
  If frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;

end;

procedure TfrmFirmas.grid_firmantesKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  If frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;

end;

procedure TfrmFirmas.grid_firmantesKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;

end;


procedure TfrmFirmas.grid_firmantesMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmFirmas.grid_firmantesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmFirmas.grid_firmantesTitleClick(Column: TColumn);
begin
 UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmFirmas.FormShow(Sender: TObject);
begin
   try
    if Connection.Contrato.FieldByName('sIdResidencia').AsString = '03' then
    begin
      cbActualizarOrdenes.visible := False;
      cbActualizarOrdenes.Enabled := False;
    end;
    //UtGrid:=TicdbGrid.create(grid_firmantes);
    BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'opFirmas', PopupPrincipal);
    OpcButton := '' ;
    frmBarra1.btnCancel.Click ;
    pgcontrol.ActivePageIndex := 0 ;
    TabContratista.Caption := 'Firmantes por parte de : ' + connection.configuracion.FieldValues['sNombre'] ;
    OrdenesdeTrabajo.Active := False ;

    If global_orden = '' then
    begin
        param_global_contrato := global_contrato;
        global_turno_reporte  := global_turno;
    end;

    OrdenesdeTrabajo.Active := False;
    OrdenesdeTrabajo.Params.ParamByName('Codigo').DataType := ftString ;
    OrdenesdeTrabajo.Params.ParamByName('Codigo').Value    := global_Contrato_Barco ;
    OrdenesdeTrabajo.Open ;

    If global_orden <> '' Then
    Begin
        tsNumeroOrden.KeyValue := global_orden ;
        Firmas.Active := False ;
        Firmas.Params.ParamByName('Contrato').DataType := ftString ;
        Firmas.Params.ParamByName('Contrato').Value   := param_Global_Contrato ;
        Firmas.Open ;
    End
    Else
    Begin
        Firmas.Active := False ;
        Firmas.Params.ParamByName('Contrato').DataType := ftString ;
        Firmas.Params.ParamByName('Contrato').Value := param_Global_Contrato ;
        Firmas.Open ;
    End ;

    If OrdenesdeTrabajo.RecordCount > 0 Then
       tsNumeroOrden.KeyValue := param_Global_Contrato ;

    QryTurnos.Active := False;
    Qryturnos.ParamByName('Contrato').AsString := param_global_contrato;
    QryTurnos.Open;

    cmbTurnos.KeyValue := global_turno_reporte;
    Grid_Firmantes.SetFocus;
    BotonPermiso.permisosBotones(frmBarra1);
    frmBarra1.btnPrinter.Enabled := False;
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_firmantes', 'Al iniciar registro', 0);
  end;
 end;
 frmbarra1.btnPrinter.Enabled:=false;
end;

procedure TfrmFirmas.grid_firmantesCellClick(Column: TColumn);
begin
  If frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;

end;

procedure TfrmFirmas.frmBarra1btnAddClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
  banderaAgregar:=true;
  // grid_firmantes.Enabled:=false;
  //empieza validacion
  nombres:=TStringList.Create;cadenas:=TStringList.Create;
  nombres.Add('Frente de trabajo');cadenas.Add(tsNumeroOrden.Text);

  if not validaTexto(nombres, cadenas, '', '') then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
  //continuainserccion de datos
  try
    If tsNumeroOrden.Text <> '' Then
    Begin
      frmBarra1.btnAddClick(Sender);
      Insertar1.Enabled  := False;
      Editar1.Enabled    := False;
      Registrar1.Enabled :=  True;
      Can1.Enabled       :=  True;
      Eliminar1.Enabled  := False;
      Refresh1.Enabled   := False;
      Salir1.Enabled     := False;

      connection.qryBusca.Active := False ;
      connection.qryBusca.SQL.Clear ;
      connection.qryBusca.SQL.Add('Select * from firmas where sContrato = :contrato Order By dIdFecha Desc') ;
      connection.qryBusca.Params.ParamByName('Contrato').DataType :=           ftString ;
      connection.qryBusca.Params.ParamByName('Contrato').Value := param_global_contrato ;
      connection.qryBusca.Open ;

      Firmas.Append ;
      If connection.qryBusca.RecordCount > 0 Then
      Begin
        firmas.FieldValues [ 'sContrato' ]  := param_Global_Contrato ;
        Firmas.FieldValues [ 'dIdFecha' ]   := connection.qryBusca.FieldValues [ 'dIdFecha' ] ;
        Firmas.FieldValues [ 'sIdTurno' ]   := global_turno_reporte ;
        tdIdFecha.Date := connection.qryBusca.Fieldbyname('didfecha').AsDateTime ;
        Firmas.FieldValues [ 'sFirmante1' ]  :=  connection.qryBusca.FieldValues [ 'sFirmante1' ] ;
        Firmas.FieldValues [ 'sFirmante2' ]  :=  connection.qryBusca.FieldValues [ 'sFirmante2' ] ;
        Firmas.FieldValues [ 'sFirmante3' ]  :=  connection.qryBusca.FieldValues [ 'sFirmante3' ] ;
        Firmas.FieldValues [ 'sFirmante4' ]  :=  connection.qryBusca.FieldValues [ 'sFirmante4' ] ;
        Firmas.FieldValues [ 'sFirmante5' ]  :=  connection.qryBusca.FieldValues [ 'sFirmante5' ] ;
        Firmas.FieldValues [ 'sFirmante6' ]  :=  connection.qryBusca.FieldValues [ 'sFirmante6' ] ;
        Firmas.FieldValues [ 'sFirmante7' ]  :=  connection.qryBusca.FieldValues [ 'sFirmante7' ] ;
        Firmas.FieldValues [ 'sFirmante8' ]  :=  connection.qryBusca.FieldValues [ 'sFirmante8' ] ;
        Firmas.FieldValues [ 'sFirmante9' ]  :=  connection.qryBusca.FieldValues [ 'sFirmante9' ] ;
        Firmas.FieldValues [ 'sFirmante10' ] := connection.qryBusca.FieldValues [ 'sFirmante10' ] ;
        Firmas.FieldValues [ 'sFirmante11' ] := connection.qryBusca.FieldValues [ 'sFirmante11' ] ;
        Firmas.FieldValues [ 'sFirmante12' ] := connection.qryBusca.FieldValues [ 'sFirmante12' ] ;
        Firmas.FieldValues [ 'sFirmante13' ] := connection.qryBusca.FieldValues [ 'sFirmante13' ] ;
        Firmas.FieldValues [ 'sFirmante14' ] := connection.qryBusca.FieldValues [ 'sFirmante14' ] ;
        Firmas.FieldValues [ 'sFirmante15' ] := connection.qryBusca.FieldValues [ 'sFirmante15' ] ;
        Firmas.FieldValues [ 'sFirmante16' ] := connection.qryBusca.FieldValues [ 'sFirmante16' ] ;
        Firmas.FieldValues [ 'sFirmante17' ] := connection.qryBusca.FieldValues [ 'sFirmante17' ] ;
        Firmas.FieldValues [ 'sFirmante18' ] := connection.qryBusca.FieldValues [ 'sFirmante18' ] ;
        Firmas.FieldValues [ 'sFirmante19' ] := connection.qryBusca.FieldValues [ 'sFirmante19' ] ;
        Firmas.FieldValues [ 'sFirmante20' ] := connection.qryBusca.FieldValues [ 'sFirmante20' ] ;
        Firmas.FieldValues [ 'sFirmante21' ] := connection.qryBusca.FieldValues [ 'sFirmante21' ] ;
        Firmas.FieldValues [ 'sFirmante22' ] := connection.qryBusca.FieldValues [ 'sFirmante22' ] ;
        Firmas.FieldValues [ 'sFirmante31' ] := connection.qryBusca.FieldValues [ 'sFirmante31' ] ;
        Firmas.FieldValues [ 'sPuesto1' ]    :=    connection.qryBusca.FieldValues [ 'sPuesto1' ] ;
        Firmas.FieldValues [ 'sPuesto2' ]    :=    connection.qryBusca.FieldValues [ 'sPuesto2' ] ;
        Firmas.FieldValues [ 'sPuesto3' ]    :=    connection.qryBusca.FieldValues [ 'sPuesto3' ] ;
        Firmas.FieldValues [ 'sPuesto4' ]    :=    connection.qryBusca.FieldValues [ 'sPuesto4' ] ;
        Firmas.FieldValues [ 'sPuesto5' ]    :=    connection.qryBusca.FieldValues [ 'sPuesto5' ] ;
        Firmas.FieldValues [ 'sPuesto6' ]    :=    connection.qryBusca.FieldValues [ 'sPuesto6' ] ;
        Firmas.FieldValues [ 'sPuesto7' ]    :=    connection.qryBusca.FieldValues [ 'sPuesto7' ] ;
        Firmas.FieldValues [ 'sPuesto8' ]    :=    connection.qryBusca.FieldValues [ 'sPuesto8' ] ;
        Firmas.FieldValues [ 'sPuesto9' ]    :=    connection.qryBusca.FieldValues [ 'sPuesto9' ] ;
        Firmas.FieldValues [ 'sPuesto10' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto10' ] ;
        Firmas.FieldValues [ 'sPuesto11' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto11' ] ;
        Firmas.FieldValues [ 'sPuesto12' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto12' ] ;
        Firmas.FieldValues [ 'sPuesto13' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto13' ] ;
        Firmas.FieldValues [ 'sPuesto14' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto14' ] ;
        Firmas.FieldValues [ 'sPuesto15' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto15' ] ;
        Firmas.FieldValues [ 'sPuesto16' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto16' ] ;
        Firmas.FieldValues [ 'sPuesto17' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto17' ] ;
        Firmas.FieldValues [ 'sPuesto18' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto18' ] ;
        Firmas.FieldValues [ 'sPuesto19' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto19' ] ;
        Firmas.FieldValues [ 'sPuesto20' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto20' ] ;
        Firmas.FieldValues [ 'sPuesto21' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto21' ] ;
        Firmas.FieldValues [ 'sPuesto22' ]   :=   connection.qryBusca.FieldValues [ 'sPuesto22' ] ;
      End
      Else
      Begin
        tdIdFecha.Date := Date ;
        Firmas.FieldValues ['dIdFecha'] := Date;
        Firmas.FieldValues [ 'sFirmante1' ]  := '*';
        Firmas.FieldValues [ 'sFirmante2' ]  := '*';
        Firmas.FieldValues [ 'sFirmante3' ]  := '*';
        Firmas.FieldValues [ 'sFirmante4' ]  := '*';
        Firmas.FieldValues [ 'sFirmante5' ]  := '*';
        Firmas.FieldValues [ 'sFirmante6' ]  := '*';
        Firmas.FieldValues [ 'sFirmante7' ]  := '*';
        Firmas.FieldValues [ 'sFirmante8' ]  := '*';
        Firmas.FieldValues [ 'sFirmante9' ]  := '*';
        Firmas.FieldValues [ 'sFirmante10' ] := '*';
        Firmas.FieldValues [ 'sFirmante11' ] := '*';
        Firmas.FieldValues [ 'sFirmante12' ] := '*';
        Firmas.FieldValues [ 'sFirmante13' ] := '*';
        Firmas.FieldValues [ 'sFirmante14' ] := '*';
        Firmas.FieldValues [ 'sFirmante15' ] := '*';
        Firmas.FieldValues [ 'sFirmante16' ] := '*';
        Firmas.FieldValues [ 'sFirmante17' ] := '*';
        Firmas.FieldValues [ 'sFirmante18' ] := '*';
        Firmas.FieldValues [ 'sFirmante19' ] := '*';
        Firmas.FieldValues [ 'sFirmante20' ] := '*';
        Firmas.FieldValues [ 'sFirmante21' ] := '*';
        Firmas.FieldValues [ 'sFirmante22' ] := '*';
        Firmas.FieldValues [ 'sFirmante31' ] := '*';
        Firmas.FieldValues [ 'sPuesto1' ]    := '*';
        Firmas.FieldValues [ 'sPuesto2' ]    := '*';
        Firmas.FieldValues [ 'sPuesto3' ]    := '*';
        Firmas.FieldValues [ 'sPuesto4' ]    := '*';
        Firmas.FieldValues [ 'sPuesto5' ]    := '*';
        Firmas.FieldValues [ 'sPuesto6' ]    := '*';
        Firmas.FieldValues [ 'sPuesto7' ]    := '*';
        Firmas.FieldValues [ 'sPuesto8' ]    := '*';
        Firmas.FieldValues [ 'sPuesto9' ]    := '*';
        Firmas.FieldValues [ 'sPuesto10' ]   := '*';
        Firmas.FieldValues [ 'sPuesto11' ]   := '*';
        Firmas.FieldValues [ 'sPuesto12' ]   := '*';
        Firmas.FieldValues [ 'sPuesto13' ]   := '*';
        Firmas.FieldValues [ 'sPuesto14' ]   := '*';
        Firmas.FieldValues [ 'sPuesto15' ]   := '*';
        Firmas.FieldValues [ 'sPuesto16' ]   := '*';
        Firmas.FieldValues [ 'sPuesto17' ]   := '*';
        Firmas.FieldValues [ 'sPuesto18' ]   := '*';
        Firmas.FieldValues [ 'sPuesto19' ]   := '*';
        Firmas.FieldValues [ 'sPuesto20' ]   := '*';
        Firmas.FieldValues [ 'sPuesto21' ]   := '*';
        Firmas.FieldValues [ 'sPuesto22' ]   := '*';
      End;
      pgControl.ActivePageIndex := 0;
      tdIdFecha.SetFocus;
    End;
    BotonPermiso.permisosBotones(frmBarra1);
    frmBarra1.btnPrinter.Enabled := False;
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_firmantes', 'Al agregar registro', 0);
    end;
  end;
  frmbarra1.btnPrinter.Enabled:=false;
end;

procedure TfrmFirmas.frmBarra1btnEditClick(Sender: TObject);
Var
    lValido : Boolean ;
begin
  try
    banderaAgregar:=false;
    grid_firmantes.Enabled:=false;
    lValido := False ;
    If Trim(tsNumeroOrden.Text) <> '' Then
    If global_grupo <> 'INTEL-CODE' Then
    Begin
      lValido := True ;
      ReporteDiario.Active := False ;
      ReporteDiario.Params.ParamByName('contrato').DataType :=           ftString;
      ReporteDiario.Params.ParamByName('contrato').Value := global_contrato_barco;
      ReporteDiario.Params.ParamByName('turno').DataType :=              ftString;
      ReporteDiario.Params.ParamByName('turno').Value    :=  global_turno_reporte;
      ReporteDiario.Params.ParamByName('Fecha').DataType :=                ftDate;
      ReporteDiario.Params.ParamByName('Fecha').Value    :=        tdIdFecha.Date;
      ReporteDiario.Params.ParamByName('Orden').DataType :=              ftString;
      ReporteDiario.Params.ParamByName('Orden').Value    := param_global_contrato;
      ReporteDiario.Open ;

      If ReporteDiario.RecordCount > 0 then
      Begin
        If ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' Then
        Begin
          MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
          lValido := False;
          frmBarra1.btnCancel.Click
        End
      end
    End
    Else
      lValido := True ;
    If lValido Then
    Begin
      frmBarra1.btnEditClick(Sender);
      Insertar1.Enabled  := False;
      Editar1.Enabled    := False;
      Registrar1.Enabled :=  True;
      Can1.Enabled       :=  True;
      Eliminar1.Enabled  := False;
      Refresh1.Enabled   := False;
      Salir1.Enabled     := False;
      try
        Firmas.Edit;
      except
        on e : exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Firmantes', 'Al editar registro', 0);
          frmbarra1.btnCancel.Click ;
        end;
      end ;
      tdIdFecha.SetFocus
    End;
    BotonPermiso.permisosBotones(frmBarra1);
    frmBarra1.btnPrinter.Enabled := False;
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_firmantes', 'Al editar registro', 0);
    end;
  end;
  frmbarra1.btnPrinter.Enabled:=false;
end;

procedure TfrmFirmas.frmBarra1btnPostClick(Sender: TObject);
Var
    lValido : Boolean ;
    nombres, cadenas: TStringList;
    QryFirmas : TZReadOnlyQuery;
begin
  QryFirmas := TZReadOnlyQuery.Create(self);
  QryFirmas.Connection := connection.zConnection;

  grid_firmantes.Enabled:=true;
  try
    lValido := False ;
    If Trim(tsNumeroOrden.Text) <> '' Then
    If global_grupo <> 'INTEL-CODE' Then
    Begin
      lValido := True ;
      ReporteDiario.Active := False ;
      ReporteDiario.Params.ParamByName('contrato').DataType :=              ftString;
      ReporteDiario.Params.ParamByName('contrato').Value    := global_contrato_barco;
      ReporteDiario.Params.ParamByName('turno').DataType    :=              ftString;
      ReporteDiario.Params.ParamByName('turno').Value       :=  global_turno_reporte;
      ReporteDiario.Params.ParamByName('Fecha').DataType    :=                ftDate;
      ReporteDiario.Params.ParamByName('Fecha').Value       :=        tdIdFecha.Date;
      ReporteDiario.Params.ParamByName('Orden').DataType    :=              ftString;
      ReporteDiario.Params.ParamByName('Orden').Value       := param_global_contrato;
      ReporteDiario.Open;
      If ReporteDiario.RecordCount > 0 then
      Begin
        If ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' Then
        Begin
          MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
          lValido := False ;
        End
      end
    End
    Else
    lValido := True;
    If lValido Then
    Begin
      firmas.FieldValues ['sContrato']    := tsNumeroOrden.Text;
      firmas.FieldValues ['sNumeroOrden'] := tsNumeroOrden.Text;
      firmas.FieldValues ['sIdTurno']     := QryTurnos.FieldValues['sIdTurno'];
      firmas.Post;

      QryFirmas.Active := False;
      QryFirmas.SQL.Clear;
      QryFirmas.SQL.Add('select * from firmas where sContrato =:Contrato and sIdTurno =:Turno and sNumeroOrden =:Orden and dIdFecha =:Fecha ');
      QryFirmas.ParamByName('Contrato').AsString := tsNumeroOrden.Text;
      QryFirmas.ParamByName('Turno').AsString    := QryTurnos.FieldValues['sIdTurno'];
      QryFirmas.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
      QryFirmas.ParamByName('Fecha').AsDate      :=        tdIdFecha.Date;
      QryFirmas.Open;

      if QryFirmas.RecordCount > 0 then
      begin
        QryTurnos.First;
        while not QryTurnos.Eof do
        begin
          try
            if QryFirmas.FieldValues['sIdTurno'] <> QryTurnos.FieldValues['sIdTurno'] then
            begin
              firmas.Append;
              firmas.FieldValues [ 'sContrato' ]   :=      QryFirmas.FieldValues['sContrato'];
              firmas.FieldValues [ 'sNumeroOrden'] :=   QryFirmas.FieldValues['sNumeroOrden'];
              firmas.FieldValues [ 'sIdTurno']     :=       QryTurnos.FieldValues['sIdTurno'];
              Firmas.FieldValues [ 'dIdFecha' ]    :=    QryFirmas.FieldValues [ 'dIdFecha' ];
              Firmas.FieldValues [ 'sFirmante1' ]  :=  QryFirmas.FieldValues [ 'sFirmante1' ];
              Firmas.FieldValues [ 'sFirmante2' ]  :=  QryFirmas.FieldValues [ 'sFirmante2' ];
              Firmas.FieldValues [ 'sFirmante3' ]  :=  QryFirmas.FieldValues [ 'sFirmante3' ];
              Firmas.FieldValues [ 'sFirmante4' ]  :=  QryFirmas.FieldValues [ 'sFirmante4' ];
              Firmas.FieldValues [ 'sFirmante5' ]  :=  QryFirmas.FieldValues [ 'sFirmante5' ];
              Firmas.FieldValues [ 'sFirmante6' ]  :=  QryFirmas.FieldValues [ 'sFirmante6' ];
              Firmas.FieldValues [ 'sFirmante7' ]  :=  QryFirmas.FieldValues [ 'sFirmante7' ];
              Firmas.FieldValues [ 'sFirmante8' ]  :=  QryFirmas.FieldValues [ 'sFirmante8' ];
              Firmas.FieldValues [ 'sFirmante9' ]  :=  QryFirmas.FieldValues [ 'sFirmante9' ];
              Firmas.FieldValues [ 'sFirmante10' ] := QryFirmas.FieldValues [ 'sFirmante10' ];
              Firmas.FieldValues [ 'sFirmante11' ] := QryFirmas.FieldValues [ 'sFirmante11' ];
              Firmas.FieldValues [ 'sFirmante12' ] := QryFirmas.FieldValues [ 'sFirmante12' ];
              Firmas.FieldValues [ 'sFirmante13' ] := QryFirmas.FieldValues [ 'sFirmante13' ];
              Firmas.FieldValues [ 'sFirmante14' ] := QryFirmas.FieldValues [ 'sFirmante14' ];
              Firmas.FieldValues [ 'sFirmante15' ] := QryFirmas.FieldValues [ 'sFirmante15' ];
              Firmas.FieldValues [ 'sFirmante16' ] := QryFirmas.FieldValues [ 'sFirmante16' ];
              Firmas.FieldValues [ 'sFirmante17' ] := QryFirmas.FieldValues [ 'sFirmante17' ];
              Firmas.FieldValues [ 'sFirmante18' ] := QryFirmas.FieldValues [ 'sFirmante18' ];
              Firmas.FieldValues [ 'sFirmante19' ] := QryFirmas.FieldValues [ 'sFirmante19' ];
              Firmas.FieldValues [ 'sFirmante20' ] := QryFirmas.FieldValues [ 'sFirmante20' ];
              Firmas.FieldValues [ 'sFirmante21' ] := QryFirmas.FieldValues [ 'sFirmante21' ];
              Firmas.FieldValues [ 'sFirmante22' ] := QryFirmas.FieldValues [ 'sFirmante22' ];
              Firmas.FieldValues [ 'sPuesto1' ]    :=    QryFirmas.FieldValues [ 'sPuesto1' ];
              Firmas.FieldValues [ 'sPuesto2' ]    :=    QryFirmas.FieldValues [ 'sPuesto2' ];
              Firmas.FieldValues [ 'sPuesto3' ]    :=    QryFirmas.FieldValues [ 'sPuesto3' ];
              Firmas.FieldValues [ 'sPuesto4' ]    :=    QryFirmas.FieldValues [ 'sPuesto4' ];
              Firmas.FieldValues [ 'sPuesto5' ]    :=    QryFirmas.FieldValues [ 'sPuesto5' ];
              Firmas.FieldValues [ 'sPuesto6' ]    :=    QryFirmas.FieldValues [ 'sPuesto6' ];
              Firmas.FieldValues [ 'sPuesto7' ]    :=    QryFirmas.FieldValues [ 'sPuesto7' ];
              Firmas.FieldValues [ 'sPuesto8' ]    :=    QryFirmas.FieldValues [ 'sPuesto8' ];
              Firmas.FieldValues [ 'sPuesto9' ]    :=    QryFirmas.FieldValues [ 'sPuesto9' ];
              Firmas.FieldValues [ 'sPuesto10' ]   :=   QryFirmas.FieldValues [ 'sPuesto10' ];
              Firmas.FieldValues [ 'sPuesto11' ]   :=   QryFirmas.FieldValues [ 'sPuesto11' ];
              Firmas.FieldValues [ 'sPuesto12' ]   :=   QryFirmas.FieldValues [ 'sPuesto12' ];
              Firmas.FieldValues [ 'sPuesto13' ]   :=   QryFirmas.FieldValues [ 'sPuesto13' ];
              Firmas.FieldValues [ 'sPuesto14' ]   :=   QryFirmas.FieldValues [ 'sPuesto14' ];
              Firmas.FieldValues [ 'sPuesto15' ]   :=   QryFirmas.FieldValues [ 'sPuesto15' ];
              Firmas.FieldValues [ 'sPuesto16' ]   :=   QryFirmas.FieldValues [ 'sPuesto16' ];
              Firmas.FieldValues [ 'sPuesto17' ]   :=   QryFirmas.FieldValues [ 'sPuesto17' ];
              Firmas.FieldValues [ 'sPuesto18' ]   :=   QryFirmas.FieldValues [ 'sPuesto18' ];
              Firmas.FieldValues [ 'sPuesto19' ]   :=   QryFirmas.FieldValues [ 'sPuesto19' ];
              Firmas.FieldValues [ 'sPuesto20' ]   :=   QryFirmas.FieldValues [ 'sPuesto20' ];
              Firmas.FieldValues [ 'sPuesto21' ]   :=   QryFirmas.FieldValues [ 'sPuesto21' ];
              Firmas.FieldValues [ 'sPuesto22' ]   :=   QryFirmas.FieldValues [ 'sPuesto22' ];
              Firmas.Post;
            end;
          Except;
          end;
          QryTurnos.Next;
        end;
      end;

      Insertar1.Enabled  :=  True;
      Editar1.Enabled    :=  True;
      Registrar1.Enabled := False;
      Can1.Enabled       := False;
      Eliminar1.Enabled  :=  True;
      Refresh1.Enabled   :=  True;
      Salir1.Enabled     :=  True;
      frmBarra1.btnCancelClick(Sender);
    End;
    desactivapop(popupprincipal);
    BotonPermiso.permisosBotones(frmBarra1);
    frmBarra1.btnPrinter.Enabled := False;
  except
    on e : exception do
    begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_firmantes', 'Al salvar registro', 0);
      frmbarra1.btnCancel.Click;
      banderaagregar:=false;
    end;
  end;
  if banderaAgregar then
  frmbarra1.btnAdd.Click;
end;

procedure TfrmFirmas.frmBarra1btnCancelClick(Sender: TObject);
begin
  try
    grid_firmantes.Enabled:=true;
    frmBarra1.btnCancelClick(Sender);
    Insertar1.Enabled  :=  True;
    Editar1.Enabled    :=  True;
    Registrar1.Enabled := False;
    Can1.Enabled       := False;
    Eliminar1.Enabled  :=  True;
    Refresh1.Enabled   :=  True;
    Salir1.Enabled     :=  True;
    firmas.Cancel;
    desactivapop(popupprincipal);
    BotonPermiso.permisosBotones(frmBarra1);
    frmBarra1.btnPrinter.Enabled :=   False;
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_firmantes', 'Al cancelar', 0);
    end;
  end;
  frmbarra1.btnPrinter.Enabled:=false;
end;

procedure TfrmFirmas.frmBarra1btnDeleteClick(Sender: TObject);
Var
    lValido : Boolean ;
begin
  try
    lValido := False ;
    If Trim(tsNumeroOrden.Text) <> '' Then
    If global_grupo <> 'INTEL-CODE' Then
    Begin
      lValido := True;
      ReporteDiario.Active := False;
      ReporteDiario.Params.ParamByName('contrato').DataType :=              ftString;
      ReporteDiario.Params.ParamByName('contrato').Value    := global_Contrato_Barco;
      ReporteDiario.Params.ParamByName('turno').DataType    :=              ftString;
      ReporteDiario.Params.ParamByName('turno').Value       :=  global_turno_reporte;
      ReporteDiario.Params.ParamByName('Fecha').DataType    :=                ftDate;
      ReporteDiario.Params.ParamByName('Fecha').Value       :=        tdIdFecha.Date;
      ReporteDiario.Params.ParamByName('Orden').DataType    :=              ftString;
      ReporteDiario.Params.ParamByName('Orden').Value       := param_global_contrato;
      ReporteDiario.Open;

      If ReporteDiario.RecordCount > 0 then
      Begin
        If ReporteDiario.FieldValues['lStatus'] <> 'Pendiente' Then
        Begin
          MessageDlg('El reporte diario del dia seleccionado se encuentra en proceso de VALIDACIÓN/AUTORIZACIÓN por lo tanto no podra hacer modificaciones al dia seleccionado.', mtWarning, [mbOk], 0);
          lValido := False;
          frmBarra1.btnCancel.Click
        End
      end;
    End
    Else
    lValido := True;
    If lValido Then
    Begin
      If Firmas.RecordCount  > 0 then
        If MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        begin
          Firmas.Delete;
        end
    End
  except
  on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_firmantes', 'Al eliminar registro', 0);
  end;
  end;
end;

procedure TfrmFirmas.frmBarra1btnRefreshClick(Sender: TObject);
begin
  try
    Firmas.active := False;
    Firmas.Open;
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_firmantes', 'Al actualizar grid', 0);
    end;
  end;
end;

procedure TfrmFirmas.frmBarra1btnExitClick(Sender: TObject);
begin
  frmBarra1.btnExitClick(Sender);
  Insertar1.Enabled  :=  True;
  Editar1.Enabled    :=  True;
  Registrar1.Enabled := False;
  Can1.Enabled       := False;
  Eliminar1.Enabled  :=  True;
  Refresh1.Enabled   :=  True;
  Salir1.Enabled     :=  True;
  close
end;
function IsDate(ADate: string): Boolean;
var
  Dummy: TDateTime;
begin
  IsDate := TryStrToDate(ADate, Dummy);
end;
procedure TfrmFirmas.tdIdFechaKeyPress(Sender: TObject; var Key: Char);
begin
   if Key = #13 then
  Begin
     tsPuesto1.SetFocus
  End
end;

procedure TfrmFirmas.tsFirma1KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsFirma7.SetFocus 
end;

procedure TfrmFirmas.tsFirma2KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsPuesto3.SetFocus
end;

procedure TfrmFirmas.tsFirma3KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsPuesto4.SetFocus
end;

procedure TfrmFirmas.tsFirma4KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsPuesto5.SetFocus 
end;

procedure TfrmFirmas.tsFirma5KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsPuesto6.SetFocus
end;

procedure TfrmFirmas.tsFirma6KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsPuesto8.SetFocus
end;

procedure TfrmFirmas.tsFirma7KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsFirma9.SetFocus
end;

procedure TfrmFirmas.tsFirma8KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsPuesto2.SetFocus
end;

procedure TfrmFirmas.tsFirma9KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsfirma10.SetFocus
end;

procedure TfrmFirmas.tsFirma10KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsFirma1.SetFocus
end;

procedure TfrmFirmas.tsFirma11Enter(Sender: TObject);
begin
    tsFirma11.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma11Exit(Sender: TObject);
begin
    tsFirma11.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma11KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsPuesto12.SetFocus
end;

procedure TfrmFirmas.tsFirma12Enter(Sender: TObject);
begin
    tsFirma12.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma12Exit(Sender: TObject);
begin
    tsFirma12.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma12KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsPuesto13.SetFocus
end;

procedure TfrmFirmas.tsFirma13Enter(Sender: TObject);
begin
    tsFirma13.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma13Exit(Sender: TObject);
begin
    tsFirma13.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma13KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsPuesto11.SetFocus
end;

procedure TfrmFirmas.tsFirma14Enter(Sender: TObject);
begin
    tsFirma14.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma14Exit(Sender: TObject);
begin
    tsFirma14.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma14KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsPuesto15.SetFocus
end;

procedure TfrmFirmas.tsFirma15Enter(Sender: TObject);
begin
    tsFirma15.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma15Exit(Sender: TObject);
begin
    tsFirma15.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma15KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsPuesto16.SetFocus
end;

procedure TfrmFirmas.tsFirma16Enter(Sender: TObject);
begin
    tsFirma16.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma16Exit(Sender: TObject);
begin
    tsFirma16.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma16KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsPuesto14.SetFocus
end;

procedure TfrmFirmas.tsFirma17Enter(Sender: TObject);
begin
    tsFirma17.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma17Exit(Sender: TObject);
begin
    tsFirma17.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma17KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsPuesto18.SetFocus
end;

procedure TfrmFirmas.tsFirma18Enter(Sender: TObject);
begin
    tsFirma18.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma18Exit(Sender: TObject);
begin
    tsFirma18.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma18KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsPuesto17.SetFocus
end;

procedure TfrmFirmas.tsFirma19Enter(Sender: TObject);
begin
    tsFirma19.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma19Exit(Sender: TObject);
begin
    tsFirma19.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma19KeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    tsPuesto20.SetFocus
end;

procedure TfrmFirmas.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure TfrmFirmas.Paste1Click(Sender: TObject);
begin
try
UtGrid.AddRowsFromClip;
except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_firmantes', 'Al pegar registro', 0);
  end;
end;
end;

procedure TfrmFirmas.Copy1Click(Sender: TObject);
begin
UtGrid.CopyRowsToClip;
end;

procedure TfrmFirmas.tsFirma7ClickBtn(Sender: TObject);
begin
  TraerNombre(tdbadveditbtn(Sender));
end;

procedure TfrmFirmas.TraerNombre(Comp:tdbadveditbtn);
var FormNombres:Tfrmcatnomfirmates;
begin
  FormNombres:= Tfrmcatnomfirmates.Create(nil);
  try
    if not(Comp.DataSource.DataSet.state in [dsEdit,dsInsert]) then
      raise Exception.Create('Actualmente no esta editando ni insertando registro.');

    FormNombres.btnSelect.Visible := True;
    FormNombres.ShowModal;
    if FormNombres.Seleccionar then
      Comp.DataSource.DataSet.FieldByName(Comp.DataField).AsString :=
      (FormNombres.Nombres.FieldByName('snombre').AsString+' '+FormNombres.Nombres.FieldByName('sapaterno').AsString+' '+FormNombres.Nombres.FieldByName('samaterno').AsString);
  finally
    FreeAndNil(FormNombres);
  end;
end;

procedure TfrmFirmas.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure TfrmFirmas.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click
end;

procedure TfrmFirmas.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmFirmas.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TfrmFirmas.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure TfrmFirmas.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

procedure TfrmFirmas.tsPuesto1KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 Then
      tsFirma1.SetFocus
end;

procedure TfrmFirmas.tsFirma1Enter(Sender: TObject);
begin
    tsFirma1.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma1Exit(Sender: TObject);
begin
    tsFirma1.Color := global_color_salida
end;

procedure TfrmFirmas.tsPuesto20KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma20.SetFocus
end;

procedure TfrmFirmas.tsPuesto21KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma21.SetFocus
end;

procedure TfrmFirmas.tsPuesto22KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma22.SetFocus
end;

procedure TfrmFirmas.tsPuesto2KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma2.SetFocus
end;

procedure TfrmFirmas.tsFirma20Enter(Sender: TObject);
begin
    tsFirma20.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma20Exit(Sender: TObject);
begin
    tsFirma20.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma20KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsPuesto21.SetFocus
end;

procedure TfrmFirmas.tsFirma21Enter(Sender: TObject);
begin
    tsFirma21.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma21Exit(Sender: TObject);
begin
    tsFirma21.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma21KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsPuesto22.SetFocus
end;

procedure TfrmFirmas.tsFirma22Enter(Sender: TObject);
begin
    tsFirma22.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma22Exit(Sender: TObject);
begin
    tsFirma22.Color := global_color_salida
end;

procedure TfrmFirmas.tsFirma22KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsPuesto19.SetFocus
end;

procedure TfrmFirmas.tsFirma2Enter(Sender: TObject);
begin
    tsFirma2.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma2Exit(Sender: TObject);
begin
    tsFirma2.Color := global_color_salida
end;

procedure TfrmFirmas.tsPuesto3KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma3.SetFocus
end;

procedure TfrmFirmas.tsFirma3Enter(Sender: TObject);
begin
    tsFirma3.Color := global_color_Entrada
end;

procedure TfrmFirmas.tsFirma3Exit(Sender: TObject);
begin
    tsFirma3.Color := global_color_salida
end;

procedure TfrmFirmas.tsPuesto4KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma4.SetFocus
end;

procedure TfrmFirmas.tsFirma4Enter(Sender: TObject);
begin
    tsFirma4.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma4Exit(Sender: TObject);
begin
    tsFirma4.Color := global_color_salida
end;

procedure TfrmFirmas.tsPuesto5KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma5.SetFocus
end;

procedure TfrmFirmas.tsFirma5Enter(Sender: TObject);
begin
    tsFirma5.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma5Exit(Sender: TObject);
begin
    tsFirma5.Color := global_color_salida
end;

procedure TfrmFirmas.tsPuesto6KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma6.SetFocus
end;

procedure TfrmFirmas.tsFirma6Enter(Sender: TObject);
begin
    tsFirma6.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma6Exit(Sender: TObject);
begin
    tsFirma6.Color := global_color_salida
end;

procedure TfrmFirmas.tsPuesto7KeyPress(Sender: TObject; var Key: Char);
begin
    If KEy = #13 Then
        tsFirma7.SetFocus
end;

procedure TfrmFirmas.tsFirma7Enter(Sender: TObject);
begin
    tsFirma7.Color := global_color_Entrada
end;

procedure TfrmFirmas.tsFirma7Exit(Sender: TObject);
begin
    tsFirma7.Color := global_color_salida
end;

procedure TfrmFirmas.tsPuesto8KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma8.SetFocus
end;

procedure TfrmFirmas.tsFirma8Enter(Sender: TObject);
begin
    tsFirma8.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma8Exit(Sender: TObject);
begin
    tsFirma8.Color := global_color_salida
end;

procedure TfrmFirmas.tsPuesto9KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma9.SetFocus
end;

procedure TfrmFirmas.tsFirma9Enter(Sender: TObject);
begin
    tsFirma9.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma9Exit(Sender: TObject);
begin
    tsFirma9.Color := global_color_salida
end;

procedure TfrmFirmas.tsPuesto10KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma10.SetFocus
end;

procedure TfrmFirmas.tsPuesto11KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma11.SetFocus
end;

procedure TfrmFirmas.tsPuesto12KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma12.SetFocus
end;

procedure TfrmFirmas.tsPuesto13KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma13.SetFocus
end;

procedure TfrmFirmas.tsPuesto14KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsFirma14.SetFocus
end;

procedure TfrmFirmas.tsPuesto15KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsFirma15.SetFocus
end;

procedure TfrmFirmas.tsPuesto16KeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
    tsFirma16.SetFocus
end;

procedure TfrmFirmas.tsPuesto17KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma17.SetFocus
end;

procedure TfrmFirmas.tsPuesto18KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma18.SetFocus
end;

procedure TfrmFirmas.tsPuesto19KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsFirma19.SetFocus
end;

procedure TfrmFirmas.tsFirma10Enter(Sender: TObject);
begin
    tsFirma10.Color := global_color_entrada
end;

procedure TfrmFirmas.tsFirma10Exit(Sender: TObject);
begin
    tsFirma10.Color := global_color_salida
end;

procedure TfrmFirmas.tdIdFechaEnter(Sender: TObject);
begin
    tdIdFecha.Color := global_color_entrada
end;

procedure TfrmFirmas.tdIdFechaExit(Sender: TObject);
begin
    tdIdFecha.Color := global_color_salida;
end;

procedure TfrmFirmas.tsNumeroOrdenExit(Sender: TObject);
begin
 try
    tsNumeroOrden.Color := global_color_Salida ;
    Firmas.Active := False ;
    Firmas.Params.ParamByName('Contrato').DataType := ftString ;
    Firmas.Params.ParamByName('Contrato').Value := tsNumeroOrden.Text ;
    Firmas.Open ;
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_firmantes', 'Al seleccionar frente de trabajo', 0);
  end;
 end;
end;

procedure TfrmFirmas.tsNumeroOrdenEnter(Sender: TObject);
begin
    frmBarra1.btnCancel.Click ;
    tsNumeroOrden.Color := global_color_entrada ;
end;

procedure TfrmFirmas.tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        Grid_Firmantes.SetFocus
end;

procedure TfrmFirmas.FirmasAfterPost(DataSet: TDataSet);
begin
if cbActualizarOrdenes.Checked then
begin
  ShowMessage('Se van a actualizar las firmas en todos los frentes registrados...');
  connection.QryBusca.Active:= false;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select sContrato, sNumeroOrden from ordenesdetrabajo ');
  connection.QryBusca.Open;

  while not connection.QryBusca.Eof  do
  begin
    connection.zCommand.Active:=False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add(' INSERT INTO firmas SET '+
                                '     sContrato=:Contrato, '+
                                '     sNumeroOrden=:Orden, '+
                                '     dIdFecha=:Fecha, '+
                                '     sFirmante1=:sFirmante1,'+
                                '     sFirmante2=:sFirmante2,'+
                                '     sFirmante3=:sFirmante3,'+
                                '     sFirmante4=:sFirmante4,'+
                                '     sFirmante5=:sFirmante5,'+
                                '     sFirmante6=:sFirmante6,'+
                                '     sFirmante7=:sFirmante7,'+
                                '     sFirmante8=:sFirmante8,'+
                                '     sFirmante9=:sFirmante9,'+
                                '     sFirmante10=:sFirmante10,'+
                                '     sFirmante11=:sFirmante11,'+
                                '     sFirmante12=:sFirmante12,'+
                                '     sFirmante13=:sFirmante13,'+
                                '     sFirmante14=:sFirmante14,'+
                                '     sFirmante15=:sFirmante15,'+
                                '     sFirmante16=:sFirmante16,'+
                                '     sFirmante17=:sFirmante17,'+
                                '     sFirmante18=:sFirmante18,'+
                                '     sPuesto1=:sPuesto1,'+
                                '     sPuesto2=:sPuesto2,'+
                                '     sPuesto3=:sPuesto3,'+
                                '     sPuesto4=:sPuesto4,'+
                                '     sPuesto5=:sPuesto5,'+
                                '     sPuesto6=:sPuesto6,'+
                                '     sPuesto7=:sPuesto7,'+
                                '     sPuesto8=:sPuesto8,'+
                                '     sPuesto9=:sPuesto9,'+
                                '     sPuesto10=:sPuesto10,' +
                                '     sPuesto11=:sPuesto11,' +
                                '     sPuesto12=:sPuesto12,' +
                                '     sPuesto13=:sPuesto13,' +
                                '     sPuesto14=:sPuesto14,' +
                                '     sPuesto15=:sPuesto15,' +
                                '     sPuesto16=:sPuesto16,' +
                                '     sPuesto17=:sPuesto17,' +
                                '     sPuesto18=:sPuesto18' +
                                ' ON DUPLICATE KEY UPDATE '+
                                '     sFirmante1=:sFirmante1,'+
                                '     sFirmante2=:sFirmante2,'+
                                '     sFirmante3=:sFirmante3,'+
                                '     sFirmante4=:sFirmante4,'+
                                '     sFirmante5=:sFirmante5,'+
                                '     sFirmante6=:sFirmante6,'+
                                '     sFirmante7=:sFirmante7,'+
                                '     sFirmante8=:sFirmante8,'+
                                '     sFirmante9=:sFirmante9,'+
                                '     sFirmante10=:sFirmante10,'+
                                '     sFirmante11=:sFirmante11,'+
                                '     sFirmante12=:sFirmante12,'+
                                '     sFirmante13=:sFirmante13,'+
                                '     sFirmante14=:sFirmante14,'+
                                '     sFirmante15=:sFirmante15,'+
                                '     sFirmante16=:sFirmante16,'+
                                '     sFirmante17=:sFirmante17,'+
                                '     sFirmante18=:sFirmante18,'+
                                '     sPuesto1=:sPuesto1,'+
                                '     sPuesto2=:sPuesto2,'+
                                '     sPuesto3=:sPuesto3,'+
                                '     sPuesto4=:sPuesto4,'+
                                '     sPuesto5=:sPuesto5,'+
                                '     sPuesto6=:sPuesto6,'+
                                '     sPuesto7=:sPuesto7,'+
                                '     sPuesto8=:sPuesto8,'+
                                '     sPuesto9=:sPuesto9,'+
                                '     sPuesto10=:sPuesto10,' +
                                '     sPuesto11=:sPuesto11,'+
                                '     sPuesto12=:sPuesto12,'+
                                '     sPuesto13=:sPuesto13,'+
                                '     sPuesto14=:sPuesto14,'+
                                '     sPuesto15=:sPuesto15,'+
                                '     sPuesto16=:sPuesto16,'+
                                '     sPuesto17=:sPuesto17,'+
                                '     sPuesto18=:sPuesto18' );

    connection.zCommand.Params.ParamByName('sFirmante1').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante1').Value := tsFirma1.Text;

    connection.zCommand.Params.ParamByName('sFirmante2').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante2').Value := tsFirma2.Text;

    connection.zCommand.Params.ParamByName('sFirmante3').DataType:=ftString ;
    connection.zCommand.Params.ParamByName('sFirmante3').Value := tsFirma3.Text;

    connection.zCommand.Params.ParamByName('sFirmante4').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante4').Value := tsFirma4.Text;

    connection.zCommand.Params.ParamByName('sFirmante5').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante5').Value := tsFirma5.Text;

    connection.zCommand.Params.ParamByName('sFirmante6').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante6').Value := tsFirma6.Text;

    connection.zCommand.Params.ParamByName('sFirmante7').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante7').Value := tsFirma7.Text;

    connection.zCommand.Params.ParamByName('sFirmante8').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante8').Value := tsFirma8.Text;

    connection.zCommand.Params.ParamByName('sFirmante9').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante9').Value := tsFirma9.Text;

    connection.zCommand.Params.ParamByName('sFirmante10').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante10').Value := tsFirma10.Text;

    connection.zCommand.Params.ParamByName('sFirmante11').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante11').Value := tsFirma11.Text;

    connection.zCommand.Params.ParamByName('sFirmante12').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante12').Value := tsFirma12.Text;

    connection.zCommand.Params.ParamByName('sFirmante13').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante13').Value := tsFirma13.Text;

    connection.zCommand.Params.ParamByName('sFirmante14').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante14').Value := tsFirma14.Text;

    connection.zCommand.Params.ParamByName('sFirmante15').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante15').Value := tsFirma15.Text;

    connection.zCommand.Params.ParamByName('sFirmante16').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante16').Value := tsFirma16.Text;

    connection.zCommand.Params.ParamByName('sFirmante17').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante17').Value := tsFirma17.Text;

    connection.zCommand.Params.ParamByName('sFirmante18').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sFirmante18').Value := tsFirma18.Text;

    connection.zCommand.Params.ParamByName('sPuesto1').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto1').Value := tsPuesto1.Text;

    connection.zCommand.Params.ParamByName('sPuesto2').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto2').Value := tsPuesto2.Text;

    connection.zCommand.Params.ParamByName('sPuesto3').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto3').Value := tsPuesto3.Text;

    connection.zCommand.Params.ParamByName('sPuesto4').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto4').Value := tsPuesto4.Text;

    connection.zCommand.Params.ParamByName('sPuesto5').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto5').Value := tsPuesto5.Text;

    connection.zCommand.Params.ParamByName('sPuesto6').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto6').Value := tsPuesto6.Text;

    connection.zCommand.Params.ParamByName('sPuesto7').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto7').Value := tsPuesto7.Text;

    connection.zCommand.Params.ParamByName('sPuesto8').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto8').Value := tsPuesto8.Text;

    connection.zCommand.Params.ParamByName('sPuesto9').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto9').Value := tsPuesto9.Text;

    connection.zCommand.Params.ParamByName('sPuesto10').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto10').Value := tsPuesto10.Text;

    connection.zCommand.Params.ParamByName('sPuesto11').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto11').Value := tsPuesto11.Text;

    connection.zCommand.Params.ParamByName('sPuesto12').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto12').Value := tsPuesto12.Text;

    connection.zCommand.Params.ParamByName('sPuesto13').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto13').Value := tsPuesto13.Text;

    connection.zCommand.Params.ParamByName('sPuesto14').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto14').Value := tsPuesto14.Text;

    connection.zCommand.Params.ParamByName('sPuesto15').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto15').Value := tsPuesto15.Text;

    connection.zCommand.Params.ParamByName('sPuesto16').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto16').Value := tsPuesto16.Text;

    connection.zCommand.Params.ParamByName('sPuesto17').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto17').Value := tsPuesto17.Text;

    connection.zCommand.Params.ParamByName('sPuesto18').DataType:=ftString;
    connection.zCommand.Params.ParamByName('sPuesto18').Value := tsPuesto18.Text;

    connection.zCommand.Params.ParamByName('Contrato').DataType:=ftString;
    connection.zCommand.Params.ParamByName('Contrato').Value   := Connection.QryBusca.FieldValues['sContrato'];

    connection.zCommand.Params.ParamByName('Orden').DataType:=ftString;
    connection.zCommand.Params.ParamByName('Orden').Value:=connection.QryBusca.FieldValues['sNumeroOrden'];

    connection.zCommand.Params.ParamByName('Fecha').DataType:=ftDate;
    connection.zCommand.Params.ParamByName('Fecha').Value:=tdIdFecha.Date;
//    try
      connection.zCommand.ExecSQL;
//    except
//      ShowMessage('No se pudo actualizar las firmas del dia ' + DateToStr(tdIdFecha.Date) + ' de la orden ' + connection.QryBusca.FieldValues['sNumeroOrden']);
//    end;
    connection.QryBusca.Next;
  end
end

end;

end.
