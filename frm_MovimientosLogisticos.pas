unit frm_MovimientosLogisticos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, frm_barra, StdCtrls, DBCtrls,
  Mask, ExtCtrls, DB, Global, Menus, frxClass, frxDBSet,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UdbGrid,
  UnitExcepciones, unittbotonespermisos, UnitValidaTexto, unitactivapop, UFunctionsGHH;

type
  TfrmMovimientosLogisticos = class(TForm)
    grid_plataformas: TDBGrid;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    frmBarra1: TfrmBarra;
    tsIdPlataforma: TDBEdit;
    tsDescripcion: TDBEdit;
    tlStatus: TDBComboBox;
    DBPlataformas: TfrxDBDataset;
    frxPlataformas: TfrxReport;
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
    ds_plataformas: TDataSource;
    Plataformas: TZQuery;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tsIdPlataformaKeyPress(Sender: TObject; var Key: Char);
    procedure tsDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure grid_plataformasCellClick(Column: TColumn);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure frmBarra1btnPrinterClick(Sender: TObject);
    procedure tlStatusEnter(Sender: TObject);
    procedure tlStatusExit(Sender: TObject);
    procedure tlStatusKeyPress(Sender: TObject; var Key: Char);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure tsIdPlataformaEnter(Sender: TObject);
    procedure tsIdPlataformaExit(Sender: TObject);
    procedure grid_plataformasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_plataformasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_plataformasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    function estaEnFrentes(sIdPlataforma: string): boolean;
    function estaEnJornadas(sIdPlataforma: string): boolean;
    function estaEnBitacoraDePersonal(sIdPlataforma: string): boolean;
  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmMovimientosLogisticos: TfrmMovimientosLogisticos;
   UtGrid:TicDbGrid;
   botonpermiso: tbotonespermisos;
   sOpcion, lStatusOrig: string;
implementation

uses frm_plataformas;

{$R *.dfm}

procedure TfrmMovimientosLogisticos.FormShow(Sender: TObject);
begin
  sMenuP:=stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cPlataformas', PopupPrincipal);
  OpcButton := '' ;
  frmbarra1.btnCancel.Click ;
  Plataformas.active := false ;
  Plataformas.Open;
  UtGrid:=TicdbGrid.create(grid_PLATAFORMAS);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmMovimientosLogisticos.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Plataformas.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmMovimientosLogisticos.tsIdPlataformaEnter(Sender: TObject);
begin
  tsIdPlataforma.Color := global_color_entrada

end;

procedure TfrmMovimientosLogisticos.tsIdPlataformaExit(Sender: TObject);
begin
  tsIdPlataforma.color := global_color_salida
end;

procedure TfrmMovimientosLogisticos.tsIdPlataformaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsDescripcion.SetFocus
end;

procedure TfrmMovimientosLogisticos.tsDescripcionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tlStatus.SetFocus 
end;

function TfrmMovimientosLogisticos.estaEnBitacoraDePersonal(
  sIdPlataforma: string): boolean;
begin
  result := false;
  with connection.QryBusca do
  begin
    Active := false;
    Filtered := false;
    SQL.Clear;
    SQL.Add('SELECT iIdDiario FROM bitacoradepersonal WHERE sIdPlataforma = :sIdPlataforma LIMIT 1');
    ParamByName('sIdPlataforma').Value := sIdPlataforma;
    Open;
    if RecordCount > 0 then
      result := true;
  end;
end;

function TfrmMovimientosLogisticos.estaEnFrentes(sIdPlataforma: string): boolean;
begin
  result := false;
  with connection.QryBusca do
  begin
    Active := false;
    Filtered := false;
    SQL.Clear;
    SQL.Add('SELECT sNumeroOrden FROM ordenesdetrabajo WHERE sIdPlataforma = :sIdPlataforma LIMIT 1');
    ParamByName('sIdPlataforma').Value := sIdPlataforma;
    Open;
    if RecordCount > 0 then
      result := true;
  end;
end;

function TfrmMovimientosLogisticos.estaEnJornadas(sIdPlataforma: string): boolean;
begin
  result := false;
  with connection.QryBusca do
  begin
    Active := false;
    Filtered := false;
    SQL.Clear;
    SQL.Add('SELECT dIdFecha FROM jornadasdiarias WHERE sIdPlataforma = :sIdPlataforma LIMIT 1');
    ParamByName('sIdPlataforma').Value := sIdPlataforma;
    Open;
    if RecordCount > 0 then
      result := true;
  end;
end;

procedure TfrmMovimientosLogisticos.frmBarra1btnAddClick(Sender: TObject);
begin
   activapop(frmPlataformas, popupprincipal);
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   Plataformas.Append ;
   Plataformas.FieldValues['sImagen'] := '' ;
   Plataformas.FieldValues['sIdDistrito'] := '' ;
   tsIdPlataforma.SetFocus ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;
end;

procedure TfrmMovimientosLogisticos.frmBarra1btnEditClick(Sender: TObject);
begin
   activapop(frmPlataformas, popupprincipal);
   frmBarra1.btnEditClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   sOpcion := 'Edit';
   lStatusOrig := Plataformas.FieldByName('lStatus').AsString;
   try
     Plataformas.Edit ;
   except
     on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al agregar registro', 0);
     frmbarra1.btnCancel.Click ;
     end;
   end ;
   tsIdPlataforma.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;

end;

procedure TfrmMovimientosLogisticos.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
    {Validaciones de campos}
    nombres:=TStringList.Create;cadenas:=TStringList.Create;
    nombres.Add('Identificacion');      nombres.Add('Descripcion');      nombres.Add('Status');
    cadenas.Add(tsIdPlataforma.Text); cadenas.Add(tsDescripcion.Text); cadenas.Add(tlStatus.Text);
    if not validaTexto(nombres, cadenas, 'Identificacion',tsIdPlataforma.Text) then
    begin
       MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
       exit;
    end;

    if (Plataformas.State = dsEdit) and (lStatusOrig <> tlStatus.Text) then
    begin
       if estaEnFrentes(Plataformas.FieldByName('sIdPlataforma').AsString) then
       begin
           MessageDlg('No es posible cambiar el status del registro porque ya ha sido ' + #10 +
           'usado en la ventana de registro de frentes de trabajo', mtInformation, [mbOk], 0);
           tlStatus.ItemIndex := 0;
       end
       else
       if estaEnJornadas(Plataformas.FieldByName('sIdPlataforma').AsString) then
       begin
           MessageDlg('No es posible cambiar el status del registro porque ya ha sido ' + #10 +
           'usado en la ventana de jornadas y tiempos', mtInformation, [mbOk], 0);
           tlStatus.ItemIndex := 0;
       end
       else
       if estaEnBitacoraDePersonal(Plataformas.FieldByName('sIdPlataforma').AsString) then
       begin
           MessageDlg('No es posible cambiar el status del registro porque ya ha sido ' + #10 +
           'usado en la ventana de registro de personal y equipo de construccion', mtInformation, [mbOk], 0);
           tlStatus.ItemIndex := 0;
       end;
    end;

    {Continua insercion de datos..}
  try
      desactivapop(popupprincipal);
      Plataformas.Post ;
      Insertar1.Enabled := True ;
      Editar1.Enabled := True ;
      Registrar1.Enabled := False ;
      Can1.Enabled := False ;
      Eliminar1.Enabled := True ;
      Refresh1.Enabled := True ;
      Salir1.Enabled := True ;
      frmBarra1.btnPostClick(Sender) ;
  except
      on e : exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al salvar registro', 0);
          frmbarra1.btnCancel.Click ;
      end;
  end;
  BotonPermiso.permisosBotones(frmBarra1);
  if sOpcion = 'Edit' then
  begin
      grid_plataformas.Enabled := True;
      sOpcion := '';
  end;

end;

procedure TfrmMovimientosLogisticos.frmBarra1btnCancelClick(Sender: TObject);
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
   Plataformas.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := True;
   sOpcion := '';
end;

procedure TfrmMovimientosLogisticos.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If Plataformas.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        Plataformas.Delete ;
      except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al eliminar registro', 0);
       end;
      end
    end
end;

procedure TfrmMovimientosLogisticos.frmBarra1btnRefreshClick(Sender: TObject);
begin
  Plataformas.refresh ;
end;

procedure TfrmMovimientosLogisticos.frmBarra1btnExitClick(Sender: TObject);
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


procedure TfrmMovimientosLogisticos.grid_plataformasCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
      frmbarra1.btnCancel.Click ;
end;

procedure TfrmMovimientosLogisticos.grid_plataformasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmMovimientosLogisticos.grid_plataformasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmMovimientosLogisticos.grid_plataformasTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmMovimientosLogisticos.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click;
end;

procedure TfrmMovimientosLogisticos.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click 
end;

procedure TfrmMovimientosLogisticos.Paste1Click(Sender: TObject);
begin
   UtGrid.AddRowsFromClip;
end;

procedure TfrmMovimientosLogisticos.Copy1Click(Sender: TObject);
begin
    UtGrid.CopyRowsToClip;
end;

procedure TfrmMovimientosLogisticos.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure TfrmMovimientosLogisticos.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click 
end;

procedure TfrmMovimientosLogisticos.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmMovimientosLogisticos.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TfrmMovimientosLogisticos.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure TfrmMovimientosLogisticos.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

procedure TfrmMovimientosLogisticos.frmBarra1btnPrinterClick(Sender: TObject);
begin
  If Plataformas.RecordCount > 0 Then
    frxPlataformas.ShowReport    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP))
  else
     messageDLG('No existen datos para imprimir!', mtInformation, [mbOk], 0);
end;

procedure TfrmMovimientosLogisticos.tlStatusEnter(Sender: TObject);
begin
    tlStatus.Color := global_color_entrada
end;

procedure TfrmMovimientosLogisticos.tlStatusExit(Sender: TObject);
begin
    tlStatus.Color := global_color_salida
end;

procedure TfrmMovimientosLogisticos.tlStatusKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsIdPlataforma.SetFocus 
end;

procedure TfrmMovimientosLogisticos.tsDescripcionEnter(Sender: TObject);
begin
    tsDescripcion.Color := global_color_entrada
end;

procedure TfrmMovimientosLogisticos.tsDescripcionExit(Sender: TObject);
begin
    tsDescripcion.Color := global_color_salida
end;

end.

