unit frm_Almacenes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, global, Grids, DBGrids, frm_barra, StdCtrls,
  Mask, DBCtrls, DB, Menus, ADODB, frxClass, frxDBSet, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, udbgrid, unitexcepciones, unittbotonespermisos,
  UnitValidaTexto,unitactivapop, UFunctionsGHH;

type
  tfrmAlmacenes = class(TForm)
    grid_Movtos: TDBGrid;
    Label1: TLabel;
    Label2: TLabel;
    tsIdAlmacen: TDBEdit;
    tsDescripcion: TDBEdit;
    frmBarra1: TfrmBarra;
    ds_almacenes: TDataSource;
    DBAlmacenes: TfrxDBDataset;
    frxAlmacenes: TfrxReport;
    Label3: TLabel;
    tsUbicacion: TDBEdit;
    Label4: TLabel;
    TSCOMENTARIOS: TDBMemo;
    Almacenes: TZQuery;
    AlmacenessIdAlmacen: TStringField;
    AlmacenessCiudad: TStringField;
    AlmacenessCp: TStringField;
    AlmacenessTelefono: TStringField;
    AlmacenessDescripcion: TStringField;
    AlmacenessDireccion: TStringField;
    AlmacenessFax: TStringField;
    AlmacenessComentarios: TStringField;
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
    Cut1: TMenuItem;
    Copy1: TMenuItem;
    N3: TMenuItem;
    Salir1: TMenuItem;
    Imprimir1: TMenuItem;
    procedure tsIdAlmacenKeyPress(Sender: TObject; var Key: Char);
    procedure tsDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure grid_MovtosCellClick(Column: TColumn);
    procedure FormShow(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tsIdAlmacenEnter(Sender: TObject);
    procedure tsIdAlmacenExit(Sender: TObject);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure frmBarra1btnPrinterClick(Sender: TObject);
    procedure tsUbicacionEnter(Sender: TObject);
    procedure tsUbicacionExit(Sender: TObject);
    procedure tsUbicacionKeyPress(Sender: TObject; var Key: Char);
    procedure TSCOMENTARIOSEnter(Sender: TObject);
    procedure TSCOMENTARIOSExit(Sender: TObject);
    procedure TSCOMENTARIOSKeyPress(Sender: TObject; var Key: Char);
    procedure grid_MovtosMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_MovtosMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_MovtosTitleClick(Column: TColumn);
    procedure Cut1Click(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);

  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmAlmacenes: TfrmAlmacenes;
  UtGrid:TicDbGrid;
  botonpermiso:tbotonespermisos;
  sOpcion : string;
implementation

{$R *.dfm}

procedure TfrmAlmacenes.tsIdAlmacenKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsDescripcion.SetFocus
end;

procedure TfrmAlmacenes.tsDescripcionKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsUbicacion.SetFocus 
end;


procedure TfrmAlmacenes.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Almacenes.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmAlmacenes.frmBarra1btnAddClick(Sender: TObject);
begin
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   tsIdAlmacen.SetFocus ;
   Almacenes.Append ;
   activapop(frmAlmacenes,popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   grid_movtos.Enabled := False;
end;

procedure TfrmAlmacenes.frmBarra1btnEditClick(Sender: TObject);
begin
   frmBarra1.btnEditClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   sOpcion := 'Edit';
   try
       Almacenes.Edit ;
       activapop(frmAlmacenes,popupprincipal);
   except
      on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Almacenes', 'Al agregar registro', 0);
       frmBarra1.btnCancel.Click ;
      end;
   end ;
   tsIdAlmacen.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmAlmacenes.frmBarra1btnPostClick(Sender: TObject);
var
   nombres, cadenas: TStringList;
begin
  {Validaciones de campos}
  nombres:=TStringList.Create;cadenas:=TStringList.Create;
  nombres.Add('Descripción');  nombres.Add('Ubicación');
  cadenas.Add(tsIdAlmacen.Text); cadenas.Add(tsDescripcion.Text); cadenas.Add(tsUbicacion.Text);

  nombres.Add('Comentarios');
  cadenas.Add(tsComentarios.Text);

  if not validaTexto(nombres, cadenas, 'Almacen',tsIdAlmacen.Text) then
  begin
     MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
     exit;
  end;

  {Continua insercion de datos..}

   try
       Almacenes.Post ;
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
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Almacenes', 'Al salvar registro', 0);
       frmBarra1.btnCancel.Click ;
       end;
   end;
   desactivapop(popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   if sOpcion = 'Edit' then
   begin
       grid_movtos.Enabled := True;
       sOpcion := '';
   end;
end;

procedure TfrmAlmacenes.frmBarra1btnCancelClick(Sender: TObject);
begin
   frmBarra1.btnCancelClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   Almacenes.Cancel ;
   desactivapop(popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   grid_movtos.Enabled := True;
   sOpcion := '';
end;

procedure TfrmAlmacenes.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If Almacenes.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
          Connection.QryBusca.Active := False ;
          Connection.QryBusca.SQL.Clear ;
          Connection.QryBusca.SQL.Add('Select sIdAlmacen from insumos Where sIdAlmacen =:almacen');
          Connection.QryBusca.Params.ParamByName('almacen').DataType := ftString ;
          Connection.QryBusca.Params.ParamByName('almacen').Value    := almacenes.FieldValues['sIdAlmacen'] ;
          Connection.QryBusca.Open ;
          If Connection.QryBusca.RecordCount > 0 Then
             MessageDlg('No se puede Borrar el Registro por que Existe en INSUMOS', mtInformation, [mbOk], 0)
          Else
             Almacenes.Delete ;
      except
        on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Almacenes', 'Al eliminar registro', 0);
        end;
      end
    end
end;

procedure TfrmAlmacenes.frmBarra1btnRefreshClick(Sender: TObject);
begin
  Almacenes.Refresh ;
end;

procedure TfrmAlmacenes.frmBarra1btnExitClick(Sender: TObject);
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

procedure TfrmAlmacenes.grid_MovtosCellClick(Column: TColumn);
begin
  if frmBarra1.btnCancel.Enabled = True then
       frmBarra1.btnCancel.Click ;
end;

procedure tfrmAlmacenes.grid_MovtosMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure tfrmAlmacenes.grid_MovtosMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure tfrmAlmacenes.grid_MovtosTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmAlmacenes.FormShow(Sender: TObject);
begin
  sMenuP:=stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'MnuCatalogodeMo', PopupPrincipal);
  OpcButton := '' ;
  Almacenes.Active := False ;
  Almacenes.Open ;  
  frmBarra1.btnCancel.Click ;
  UtGrid:=TicdbGrid.create(grid_movtos);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure tfrmAlmacenes.Imprimir1Click(Sender: TObject);
begin
    frmbarra1.btnPrinter.Click;
end;

procedure tfrmAlmacenes.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure tfrmAlmacenes.Copy1Click(Sender: TObject);
begin
  UtGrid.AddRowsFromClip;
end;

procedure tfrmAlmacenes.Cut1Click(Sender: TObject);
begin
  UtGrid.CopyRowsToClip;
end;

procedure tfrmAlmacenes.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure tfrmAlmacenes.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click 
end;

procedure tfrmAlmacenes.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click
end;

procedure tfrmAlmacenes.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure tfrmAlmacenes.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure tfrmAlmacenes.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

procedure tfrmAlmacenes.tsIdAlmacenEnter(Sender: TObject);
begin
    tsIdAlmacen.Color := global_color_entrada
end;

procedure tfrmAlmacenes.tsIdAlmacenExit(Sender: TObject);
begin
    tsIdAlmacen.Color := global_color_salida
end;

procedure tfrmAlmacenes.tsDescripcionEnter(Sender: TObject);
begin
    tsDescripcion.Color := global_color_entrada
end;

procedure tfrmAlmacenes.tsDescripcionExit(Sender: TObject);
begin
    tsDescripcion.Color := global_color_salida
end;

procedure tfrmAlmacenes.frmBarra1btnPrinterClick(Sender: TObject);
begin
   If Almacenes.RecordCount > 0 Then
      frxAlmacenes.ShowReport    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP))
   else
      messageDLG('No se encontro informacion a Imprimir.' , mtInformation, [mbOk], 0);
end;

procedure tfrmAlmacenes.tsUbicacionEnter(Sender: TObject);
begin
    tsUbicacion.Color := global_color_entrada
end;

procedure tfrmAlmacenes.tsUbicacionExit(Sender: TObject);
begin
    tsUbicacion.Color := global_color_salida
end;

procedure tfrmAlmacenes.tsUbicacionKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsComentarios.SetFocus ;
end;

procedure tfrmAlmacenes.TSCOMENTARIOSEnter(Sender: TObject);
begin
    tsComentarios.Color := global_color_entrada
end;

procedure tfrmAlmacenes.TSCOMENTARIOSExit(Sender: TObject);
begin
    tsComentarios.Color := global_color_salida
end;

procedure tfrmAlmacenes.TSCOMENTARIOSKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsidalmacen.SetFocus ;

end;

end.
