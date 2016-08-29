unit frm_proveedores;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_barra, global, db, StdCtrls,
  Mask, DBCtrls, Menus, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  udbgrid, unitexcepciones, unittbotonespermisos,UnitValidaTexto,
  unitactivapop,UnitValidacion;
type
  TfrmProveedores = class(TForm)
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    grid_proveedores: TDBGrid;
    tsIdProveedor: TDBEdit;
    tsRazon: TDBEdit;
    tsRfc: TDBEdit;
    tsDomicilio: TDBEdit;
    tsCiudad: TDBEdit;
    tsEstado: TDBEdit;
    tsTelefono: TDBEdit;
    tmComentarios: TDBMemo;
    frmBarra1: TfrmBarra;
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
    dsProveedores: TDataSource;
    Proveedores: TZQuery;
    Label9: TLabel;
    dbCuenta: TDBEdit;
    Label10: TLabel;
    dbSucursal: TDBEdit;
    Label11: TLabel;
    dbBanco: TDBEdit;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    dbClave: TDBEdit;
    dbNombre: TDBEdit;
    dbVendedor: TDBEdit;
    anexo_ocompras: TZReadOnlyQuery;
    ds_anexo_ocompras: TDataSource;
    ProveedoressIdProveedor: TStringField;
    ProveedoressRazon: TStringField;
    ProveedoressDomicilio: TStringField;
    ProveedoressCiudad: TStringField;
    ProveedoressEstado: TStringField;
    ProveedoressRfc: TStringField;
    ProveedoressTelefono: TStringField;
    ProveedoressCuenta: TStringField;
    ProveedoressSucursal: TStringField;
    ProveedoressBanco: TStringField;
    ProveedoresmComentarios: TMemoField;
    ProveedoressRepresentante: TStringField;
    ProveedoressVendedor: TStringField;
    ProveedoressEmail: TStringField;
    ProveedoressNombreCuenta: TStringField;
    ProveedoressClaveBan: TIntegerField;
    dbEmail: TDBEdit;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tsIdProveedorKeyPress(Sender: TObject; var Key: Char);
    procedure tsRazonKeyPress(Sender: TObject; var Key: Char);
    procedure tsRfcKeyPress(Sender: TObject; var Key: Char);
    procedure tsDomicilioKeyPress(Sender: TObject; var Key: Char);
    procedure tsCiudadKeyPress(Sender: TObject; var Key: Char);
    procedure tsEstadoKeyPress(Sender: TObject; var Key: Char);
    procedure grid_proveedoresCellClick(Column: TColumn);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure dbCuentaKeyPress(Sender: TObject; var Key: Char);
    procedure dbSucursalKeyPress(Sender: TObject; var Key: Char);
    procedure dbBancoKeyPress(Sender: TObject; var Key: Char);
    procedure tsTelefonoKeyPress(Sender: TObject; var Key: Char);
    procedure dbClaveKeyPress(Sender: TObject; var Key: Char);
    procedure dbNombreKeyPress(Sender: TObject; var Key: Char);
    procedure dbVendedorKeyPress(Sender: TObject; var Key: Char);
    procedure dbEmailKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdProveedorEnter(Sender: TObject);
    procedure tsIdProveedorExit(Sender: TObject);
    procedure tsRazonEnter(Sender: TObject);
    procedure tsRazonExit(Sender: TObject);
    procedure tsRfcEnter(Sender: TObject);
    procedure tsRfcExit(Sender: TObject);
    procedure tsDomicilioEnter(Sender: TObject);
    procedure tsDomicilioExit(Sender: TObject);
    procedure tsEstadoEnter(Sender: TObject);
    procedure tsEstadoExit(Sender: TObject);
    procedure tsCiudadEnter(Sender: TObject);
    procedure tsCiudadExit(Sender: TObject);
    procedure dbCuentaEnter(Sender: TObject);
    procedure dbCuentaExit(Sender: TObject);
    procedure tsTelefonoEnter(Sender: TObject);
    procedure tsTelefonoExit(Sender: TObject);
    procedure dbBancoEnter(Sender: TObject);
    procedure dbBancoExit(Sender: TObject);
    procedure dbSucursalEnter(Sender: TObject);
    procedure dbSucursalExit(Sender: TObject);
    procedure dbClaveEnter(Sender: TObject);
    procedure dbClaveExit(Sender: TObject);
    procedure dbNombreEnter(Sender: TObject);
    procedure dbNombreExit(Sender: TObject);
    procedure dbVendedorEnter(Sender: TObject);
    procedure dbVendedorExit(Sender: TObject);
    procedure dbEmailEnter(Sender: TObject);
    procedure dbEmailExit(Sender: TObject);
    procedure tmComentariosEnter(Sender: TObject);
    procedure tmComentariosExit(Sender: TObject);
    procedure grid_proveedoresMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_proveedoresMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_proveedoresTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure tsTelefonoChange(Sender: TObject);
    procedure dbCuentaChange(Sender: TObject);
    procedure dbClaveChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProveedores: TfrmProveedores;
  Opcion : String ;
  utgrid:ticdbgrid;
  botonpermiso:tbotonespermisos;
  banderaAgregar:Boolean;

implementation
uses frm_connection ;
{$R *.dfm}

procedure TfrmProveedores.FormShow(Sender: TObject);
begin
 try
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cProveedores', PopupPrincipal);
  OpcButton := '' ;
  frmBarra1.btnCancel.Click ;
  Proveedores.Active := False ;
  Proveedores.Open;
  anexo_ocompras.Active := False ;
  anexo_ocompras.Params.ParamByName('Contrato').DataType := ftString ;
  anexo_ocompras.Params.ParamByName('Contrato').Value    := global_contrato ;
  anexo_ocompras.Open ;
  UtGrid:=TicdbGrid.create(grid_proveedores);
  BotonPermiso.permisosBotones(frmBarra1);
  frmbarra1.btnPrinter.Enabled:=false;
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Proveedores', 'Al iniciar el formulario', 0);
  end;
 end;
end;

procedure TfrmProveedores.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  botonpermiso.Free;
  utgrid.Destroy;
  Proveedores.Cancel ;
  action := cafree ;
end;

procedure TfrmProveedores.tsIdProveedorEnter(Sender: TObject);
begin
tsidproveedor.color:= global_color_entrada;
end;

procedure TfrmProveedores.tsIdProveedorExit(Sender: TObject);
begin
tsidproveedor.Color:= global_color_salida
end;

procedure TfrmProveedores.tsIdProveedorKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsRazon.SetFocus
end;

procedure TfrmProveedores.tsRazonEnter(Sender: TObject);
begin
tsrazon.Color:= global_color_entrada
end;

procedure TfrmProveedores.tsRazonExit(Sender: TObject);
begin
tsrazon.Color:=global_color_salida
end;

procedure TfrmProveedores.tsRazonKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsRfc.SetFocus
end;

procedure TfrmProveedores.tsRfcEnter(Sender: TObject);
begin
tsrfc.Color:=global_color_entrada
end;

procedure TfrmProveedores.tsRfcExit(Sender: TObject);
begin
tsrfc.Color:=global_color_salida
end;

procedure TfrmProveedores.tsRfcKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsDomicilio.SetFocus
end;

procedure TfrmProveedores.tsTelefonoChange(Sender: TObject);
begin
//tdbeditchangef(tsTelefono,'Teléfono');
end;

procedure TfrmProveedores.tsTelefonoEnter(Sender: TObject);
begin

    tstelefono.Color:=global_color_entrada ;

end;

procedure TfrmProveedores.tsTelefonoExit(Sender: TObject);
begin
     tstelefono.Color:=global_color_salida;

end;
procedure TfrmProveedores.tsTelefonoKeyPress(Sender: TObject; var Key: Char);
begin
//  if not KeyFiltroTdbedit(tstelefono,key) then
//      key:=#0;
  if key = #13 then
    dbbanco.SetFocus;
end;

procedure TfrmProveedores.tsDomicilioEnter(Sender: TObject);
begin
     tsdomicilio.color:=global_color_entrada
end;

procedure TfrmProveedores.tsDomicilioExit(Sender: TObject);
begin
    tsdomicilio.Color:=global_color_salida
end;

procedure TfrmProveedores.tsDomicilioKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsestado.SetFocus
end;

procedure TfrmProveedores.tmComentariosEnter(Sender: TObject);
begin
tmcomentarios.Color:=global_color_entrada
end;

procedure TfrmProveedores.tmComentariosExit(Sender: TObject);
begin
tmcomentarios.Color:=global_color_salida
end;

procedure TfrmProveedores.tsCiudadEnter(Sender: TObject);
begin
tsciudad.Color:=global_color_entrada
end;

procedure TfrmProveedores.tsCiudadExit(Sender: TObject);
begin
tsciudad.Color:=global_color_salida
end;

procedure TfrmProveedores.tsCiudadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    dbcuenta.SetFocus
end;

procedure TfrmProveedores.tsEstadoEnter(Sender: TObject);
begin
tsestado.color:=global_color_entrada
end;

procedure TfrmProveedores.tsEstadoExit(Sender: TObject);
begin
tsestado.color:=global_color_salida
end;

procedure TfrmProveedores.tsEstadoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsciudad.SetFocus
end;

procedure TfrmProveedores.grid_proveedoresCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
    frmbarra1.btnCancel.Click
end;

procedure TfrmProveedores.grid_proveedoresMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmProveedores.grid_proveedoresMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmProveedores.grid_proveedoresTitleClick(Column: TColumn);
begin
 UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmProveedores.frmBarra1btnAddClick(Sender: TObject);
begin
  try
   //activapop(frmProveedores, popupprincipal);
   banderaAgregar:=true;
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   tsIdProveedor.SetFocus ;
   activapop(frmProveedores, popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   grid_Proveedores.Enabled:=false;
   Proveedores.Append ;
  except
   on e : exception do begin
   UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Proveedores', 'Al agregar registro ', 0);
   end;
  end;
    frmbarra1.btnPrinter.Enabled:=false;
end;

procedure TfrmProveedores.frmBarra1btnEditClick(Sender: TObject);
begin
   banderaAgregar:=false;
   frmBarra1.btnEditClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   try
      Proveedores.Edit ;
      activapop(frmProveedores, popupprincipal);
      grid_Proveedores.Enabled:=false;
   except
      on e : exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Proveedores', 'Al editar registro', 0);
          frmbarra1.btnCancel.Click ;
      end;
   end ;
   tsIdProveedor.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
     frmbarra1.btnPrinter.Enabled:=false;
end;

procedure TfrmProveedores.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
//empieza validacion
  nombres:=TStringList.Create;cadenas:=TStringList.Create;
  nombres.Add('Razon social');nombres.Add('Domicilio');
  nombres.Add('Ciudad');nombres.Add('Telefono');
  nombres.Add('Sucursal Banco');nombres.Add('RFC');
  nombres.Add('Estado');nombres.Add('Cuenta Banco');
  nombres.Add('Banco');nombres.Add('Clave del Banco');
  nombres.Add('Nombre de la Cuenta');nombres.Add('Vendedor');
  nombres.Add('E Mail');

  cadenas.Add(tsRazon.Text);cadenas.Add(tsDomicilio.Text);
  cadenas.Add(tsciudad.Text);cadenas.Add(tsTelefono.Text);
  cadenas.Add(Dbsucursal.Text);cadenas.Add(tsrfc.Text);
  cadenas.Add(tsestado.Text);cadenas.Add(dbcuenta.Text);
  cadenas.Add(dbbanco.Text);cadenas.Add(dbclave.Text);
  cadenas.Add(dbnombre.Text);cadenas.Add(dbvendedor.Text);
  cadenas.Add(dbemail.Text);

  if not validaTexto(nombres, cadenas, 'Proveedor', tsIdproveedor.Text) then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
//continuainserccion de datos
   try
       Proveedores.Post ;
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
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Proveedores', 'Al salvar registro', 0);
      frmBarra1.btnCancel.Click ;
     end;
   end;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  grid_proveedores.Enabled:=True;
  frmbarra1.btnCancel.Click;
  if banderaAgregar then
    frmbarra1.btnAdd.Click;
    frmbarra1.btnPrinter.Enabled:=false;
end;

procedure TfrmProveedores.frmBarra1btnCancelClick(Sender: TObject);
begin
 try
  desactivapop(popupprincipal);
  frmBarra1.btnCancelClick(Sender);
  Insertar1.Enabled := True ;
  Editar1.Enabled := True ;
  Registrar1.Enabled := False ;
  Can1.Enabled := False ;
  Eliminar1.Enabled := True ;
  Refresh1.Enabled := True ;
  Salir1.Enabled := True ;
  Proveedores.Cancel ;
  BotonPermiso.permisosBotones(frmBarra1);
  grid_Proveedores.Enabled:=True;
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Proveedores', 'Al cancelar', 0);
  end;
 end;
 frmbarra1.btnPrinter.Enabled:=false;
end;

procedure TfrmProveedores.frmBarra1btnDeleteClick(Sender: TObject);
begin
 // If Proveedores.RecordCount  > 0 then
 //    if anexo_ocompras.RecordCount > 0 then
    if grid_proveedores.DataSource.DataSet.IsEmpty=false then
       if grid_proveedores.DataSource.DataSet.RecordCount>0 then
       
              if MessageDlg('Desea eliminar el Registro Activo?',mtConfirmation, [mbYes, mbNo], 0) = mrYes then
                begin
                 if anexo_ocompras.FieldValues ['sIdProveedor']<> Proveedores.FieldValues ['sIdProveedor'] then
                   begin
                     try
                     Proveedores.Delete ;
                      except
                       on e : exception do begin
                       UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Proveedores', 'Al eliminar registro', 0);
                       end;
                   end
                end
                  else
                     ShowMessage('El Proveedor ya ha sido agregado a una Orden de Compra')
                end;
end;

procedure TfrmProveedores.frmBarra1btnRefreshClick(Sender: TObject);
begin
 try
  Proveedores.Active := False ;
  Proveedores.Open
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Proveedores', 'Al actualizar grid', 0);
  end;
 end;
end;

procedure TfrmProveedores.frmBarra1btnExitClick(Sender: TObject);
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

procedure TfrmProveedores.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click 
end;

procedure TfrmProveedores.Paste1Click(Sender: TObject);
begin
try
UtGrid.AddRowsFromClip;
except
on e : exception do begin
UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Proveedores', 'Al pegar registro', 0);
end;
end;
end;
procedure TfrmProveedores.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure TfrmProveedores.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click 
end;

procedure TfrmProveedores.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmProveedores.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TfrmProveedores.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure TfrmProveedores.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click 
end;

procedure TfrmProveedores.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

procedure TfrmProveedores.dbClaveChange(Sender: TObject);
begin
 tdbeditchangef(dbClave,'Clave Banco');
end;

procedure TfrmProveedores.dbClaveEnter(Sender: TObject);
begin
dbclave.Color:=global_color_entrada
end;

procedure TfrmProveedores.dbClaveExit(Sender: TObject);
begin
dbclave.Color:=global_color_salida
end;

procedure TfrmProveedores.dbClaveKeyPress(Sender: TObject; var Key: Char);
begin
if not keyFiltroTdbedit(dbClave,key) then
   key:=#0;
  if key = #13 then
    dbnombre.SetFocus
end;

procedure TfrmProveedores.dbCuentaChange(Sender: TObject);
begin
tdbeditchangef(dbCuenta,'Cuenta Banco');
end;

procedure TfrmProveedores.dbCuentaEnter(Sender: TObject);
begin
dbcuenta.Color:=global_color_entrada
end;

procedure TfrmProveedores.dbCuentaExit(Sender: TObject);
begin
dbcuenta.Color:=global_color_salida
end;

procedure TfrmProveedores.dbCuentaKeyPress(Sender: TObject; var Key: Char);
begin
  if not KeyFiltroTdbedit(dbCuenta,key) then
      key:=#0;
  If key = #13 Then
     tstelefono.SetFocus ;
end;

procedure TfrmProveedores.dbEmailEnter(Sender: TObject);
begin
dbemail.Color:=global_color_entrada
end;

procedure TfrmProveedores.dbEmailExit(Sender: TObject);
begin
dbemail.Color:=global_color_salida
end;

procedure TfrmProveedores.dbEmailKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tmcomentarios.SetFocus
end;

procedure TfrmProveedores.dbNombreEnter(Sender: TObject);
begin
dbnombre.Color:=global_color_entrada
end;

procedure TfrmProveedores.dbNombreExit(Sender: TObject);
begin
dbnombre.Color:=global_color_salida
end;

procedure TfrmProveedores.dbNombreKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    dbvendedor.SetFocus
end;

procedure TfrmProveedores.dbSucursalEnter(Sender: TObject);
begin
dbsucursal.Color:=global_color_entrada
end;

procedure TfrmProveedores.dbSucursalExit(Sender: TObject);
begin
dbsucursal.Color:=global_color_salida
end;

procedure TfrmProveedores.dbSucursalKeyPress(Sender: TObject; var Key: Char);
begin
      If key=#13 Then
         dbclave.SetFocus ;
end;

procedure TfrmProveedores.dbVendedorEnter(Sender: TObject);
begin
dbvendedor.Color:=global_color_entrada
end;

procedure TfrmProveedores.dbVendedorExit(Sender: TObject);
begin
dbvendedor.Color:=global_color_salida
end;

procedure TfrmProveedores.dbVendedorKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    dbemail.SetFocus
end;

procedure TfrmProveedores.Copy1Click(Sender: TObject);
begin
UtGrid.CopyRowsToClip;
end;

procedure TfrmProveedores.dbBancoEnter(Sender: TObject);
begin
dbbanco.Color:=global_color_entrada
end;

procedure TfrmProveedores.dbBancoExit(Sender: TObject);
begin
dbbanco.Color:=global_color_salida
end;

procedure TfrmProveedores.dbBancoKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    dbsucursal.SetFocus
end;

end.
