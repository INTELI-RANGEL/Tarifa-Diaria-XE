unit frm_unificadorequipos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, StdCtrls,global,
  Mask,UdbGrid, DBCtrls, Grids, DBGrids, frm_barra,UnitExcepciones, unittbotonespermisos, unitactivapop, 
  AdvSpin, DBAdvSp;

type
  TFrmUnificadorEquipos = class(TForm)
    frmBarra1: TfrmBarra;
    grid_plataformas: TDBGrid;
    tsIdPlataforma: TDBEdit;
    zq_UnificadorEquipos: TZQuery;
    ds_UnificadorEquipos: TDataSource;
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
    Label1: TLabel;
    Label2: TLabel;
    EdtId: TDBAdvSpinEdit;
    zq_UnificadorEquipossContrato: TStringField;
    zq_UnificadorEquiposiIdUnificador: TIntegerField;
    zq_UnificadorEquipossDescripcion: TStringField;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
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
    procedure DBEdit1Enter(Sender: TObject);
    procedure DBEdit1Exit(Sender: TObject);

  private
    { Private declarations }
      sMenuP: String;

  public
    { Public declarations }
  end;

var
  FrmUnificadorEquipos: TFrmUnificadorEquipos;
   UtGrid:TicDbGrid;
   botonpermiso: tbotonespermisos;
   sOpcion, lStatusOrig: string;
   
implementation

uses frm_connection;

{$R *.dfm}

procedure TFrmUnificadorEquipos.FormShow(Sender: TObject);
begin
  sMenuP:=stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cGruposdePersonal', PopupPrincipal);
  OpcButton := '' ;
  frmbarra1.btnCancel.Click ;
  zq_UnificadorEquipos.active := false ;
  zq_UnificadorEquipos.ParamByName('sContrato').AsString := global_Contrato_Barco;
  zq_UnificadorEquipos.Open;
  UtGrid:=TicdbGrid.create(grid_PLATAFORMAS);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TFrmUnificadorEquipos.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  zq_UnificadorEquipos.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TFrmUnificadorEquipos.tsIdPlataformaEnter(Sender: TObject);
begin
  tsIdPlataforma.Color := global_color_entrada
end;

procedure TFrmUnificadorEquipos.tsIdPlataformaExit(Sender: TObject);
begin
  tsIdPlataforma.color := global_color_salida
end;

procedure TFrmUnificadorEquipos.frmBarra1btnAddClick(Sender: TObject);
begin
   activapop(FrmUnificadorEquipos, popupprincipal);
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;

   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;
   zq_UnificadorEquipos.Append ;
   connection.QryBusca.Active := False;
   connection.QryBusca.SQL.Text := 'SELECT (MAX(iIdUnificador)+1) as Maximo FROM unificadorequipos WHERE sContrato = :sContrato';
   connection.QryBusca.ParamByName('sContrato').AsString := global_Contrato_Barco;
   connection.QryBusca.Open;

   if connection.QryBusca.RecordCount = 1 then
     zq_UnificadorEquipos.FieldByName('iIdUnificador').AsInteger := connection.QryBusca.FieldByName('Maximo').AsInteger
   else
     zq_UnificadorEquipos.FieldByName('iIdUnificador').AsInteger := 1;

   zq_UnificadorEquipos.FieldByName('sContrato').AsString := global_Contrato_Barco;
   EdtId.SetFocus ;   
end;

procedure TFrmUnificadorEquipos.frmBarra1btnEditClick(Sender: TObject);
begin
   activapop(FrmUnificadorEquipos, popupprincipal);
   frmBarra1.btnEditClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   sOpcion := 'Edit';
//   lStatusOrig := Plataformas.FieldByName('lStatus').AsString;
   try
     zq_UnificadorEquipos.Edit ;
   except
     on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Unificador de equipos', 'Al agregar registro', 0);
     frmbarra1.btnCancel.Click ;
     end;
   end ;
   tsIdPlataforma.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;

end;

procedure TFrmUnificadorEquipos.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
   frmBarra1.btnPost.SetFocus;
    {Validaciones de campos}
   if length(trim(zq_UnificadorEquipos.FieldByName('iIdUnificador').AsString)) = 0 then
     zq_UnificadorEquipos.FieldByName('iIdUnificador').asinteger := 1;

   if (zq_UnificadorEquipos.State = dsInsert) or (zq_UnificadorEquipos.FieldByName('iIdUnificador').OldValue <> zq_UnificadorEquipos.FieldByName('iIdUnificador').AsInteger) then
   begin
     connection.QryBusca.Active := False;
     connection.QryBusca.SQL.Text := 'SELECT iIdUnificador FROM unificadorequipos WHERE sContrato = :sContrato AND iIdUnificador = :iIdUnificador';
     connection.QryBusca.ParamByName('iIdUnificador').AsInteger := zq_UnificadorEquipos.FieldByName('iIdUnificador').AsInteger;
     connection.QryBusca.ParamByName('sContrato').AsString := global_Contrato_Barco;
     connection.QryBusca.Open;


     if connection.QryBusca.RecordCount <> 0 then
     begin
       ShowMessage('El identificador ya existe, intente con otro valor.');
       Exit;
     end;

     if length(trim(zq_UnificadorEquipos.FieldByName('sdescripcion').AsString)) = 0 then
     begin
       ShowMessage('La descripción no puede ir vacía, porfavor ingrese una deswscripción.');
       tsIdPlataforma.SetFocus;
       Exit;
     end;
   end;

    {Continua insercion de datos..}
  try
      desactivapop(popupprincipal);
      zq_UnificadorEquipos.Post ;
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
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Unificador de equipos', 'Al salvar registro', 0);
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

procedure TFrmUnificadorEquipos.frmBarra1btnCancelClick(Sender: TObject);
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
   zq_UnificadorEquipos.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := True;
   sOpcion := '';
end;

procedure TFrmUnificadorEquipos.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If zq_UnificadorEquipos.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        zq_UnificadorEquipos.Delete ;
      except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Unificador de equipos', 'Al eliminar registro', 0);
       end;
      end
    end
end;

procedure TFrmUnificadorEquipos.frmBarra1btnRefreshClick(Sender: TObject);
begin
  zq_UnificadorEquipos.refresh ;
end;

procedure TFrmUnificadorEquipos.frmBarra1btnExitClick(Sender: TObject);
begin
   frmBarra1.btnExitClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   close;
end;


procedure TFrmUnificadorEquipos.grid_plataformasCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
      frmbarra1.btnCancel.Click ;
end;

procedure TFrmUnificadorEquipos.grid_plataformasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TFrmUnificadorEquipos.grid_plataformasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TFrmUnificadorEquipos.grid_plataformasTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TFrmUnificadorEquipos.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click;
end;

procedure TFrmUnificadorEquipos.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure TFrmUnificadorEquipos.Paste1Click(Sender: TObject);
begin
   UtGrid.AddRowsFromClip;
end;

procedure TFrmUnificadorEquipos.Copy1Click(Sender: TObject);
begin
    UtGrid.CopyRowsToClip;
end;

procedure TFrmUnificadorEquipos.DBEdit1Enter(Sender: TObject);
begin
  EdtId.Color := global_color_entrada
end;

procedure TFrmUnificadorEquipos.DBEdit1Exit(Sender: TObject);
begin
  EdtId.Color := global_color_salida
end;

procedure TFrmUnificadorEquipos.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure TFrmUnificadorEquipos.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click 
end;

procedure TFrmUnificadorEquipos.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TFrmUnificadorEquipos.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TFrmUnificadorEquipos.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure TFrmUnificadorEquipos.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

procedure TFrmUnificadorEquipos.frmBarra1btnPrinterClick(Sender: TObject);
begin
//  If Plataformas.RecordCount > 0 Then
//    frxPlataformas.ShowReport(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP))
//  else
//     messageDLG('No existen datos para imprimir!', mtInformation, [mbOk], 0);
end;

end.
