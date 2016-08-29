unit frm_gruposdeequipo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, frm_barra, StdCtrls, DBCtrls,
  Mask, ExtCtrls, DB, Global, Menus, 
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UdbGrid,
  UnitExcepciones, unittbotonespermisos, unitactivapop;

type
  TfrmGruposdeEquipo = class(TForm)
    grid_plataformas: TDBGrid;
    Label1: TLabel;
    frmBarra1: TfrmBarra;
    tsIdPlataforma: TDBEdit;
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
    ds_gruposdepersonal: TDataSource;
    zq_gruposdepersonal: TZQuery;
    EdtIdentificador: TDBEdit;
    Label2: TLabel;
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
  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmGruposdeEquipo: TfrmGruposdeEquipo;
   UtGrid:TicDbGrid;
   botonpermiso: tbotonespermisos;
   sOpcion, lStatusOrig: string;
implementation

{$R *.dfm}

procedure TfrmGruposdeEquipo.FormShow(Sender: TObject);
begin
  sMenuP:=stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cGruposdePersonal', PopupPrincipal);
  OpcButton := '' ;
  frmbarra1.btnCancel.Click ;
  zq_gruposdepersonal.active := false ;
  zq_gruposdepersonal.Open;
  UtGrid:=TicdbGrid.create(grid_PLATAFORMAS);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmGruposdeEquipo.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  zq_gruposdepersonal.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmGruposdeEquipo.tsIdPlataformaEnter(Sender: TObject);
begin
  tsIdPlataforma.Color := global_color_entrada

end;

procedure TfrmGruposdeEquipo.tsIdPlataformaExit(Sender: TObject);
begin
  tsIdPlataforma.color := global_color_salida
end;

procedure TfrmGruposdeEquipo.frmBarra1btnAddClick(Sender: TObject);
begin
   activapop(frmGruposdeEquipo, popupprincipal);
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   zq_gruposdepersonal.Append ;
   tsIdPlataforma.SetFocus ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;
end;

procedure TfrmGruposdeEquipo.frmBarra1btnEditClick(Sender: TObject);
begin
   activapop(frmGruposdeEquipo, popupprincipal);
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
     zq_gruposdepersonal.Edit ;
   except
     on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Grupos de Personal', 'Al agregar registro', 0);
     frmbarra1.btnCancel.Click ;
     end;
   end ;
   tsIdPlataforma.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;

end;

procedure TfrmGruposdeEquipo.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
  frmBarra1.btnPost.SetFocus;
    {Validaciones de campos}
  if length(trim(EdtIdentificador.Text)) = 0 then
    raise Exception.Create('El campo identificador no puede ir vacío.');
    {Continua insercion de datos..}
  try
      desactivapop(popupprincipal);
      zq_gruposdepersonal.Post ;
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
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Agrupadores de Personal', 'Al salvar registro', 0);
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

procedure TfrmGruposdeEquipo.frmBarra1btnCancelClick(Sender: TObject);
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
   zq_gruposdepersonal.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := True;
   sOpcion := '';
end;

procedure TfrmGruposdeEquipo.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If zq_gruposdepersonal.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        zq_gruposdepersonal.Delete ;
      except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Agrupadores de Personal', 'Al eliminar registro', 0);
       end;
      end
    end
end;

procedure TfrmGruposdeEquipo.frmBarra1btnRefreshClick(Sender: TObject);
begin
  zq_gruposdepersonal.refresh ;
end;

procedure TfrmGruposdeEquipo.frmBarra1btnExitClick(Sender: TObject);
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


procedure TfrmGruposdeEquipo.grid_plataformasCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
      frmbarra1.btnCancel.Click ;
end;

procedure TfrmGruposdeEquipo.grid_plataformasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmGruposdeEquipo.grid_plataformasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmGruposdeEquipo.grid_plataformasTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmGruposdeEquipo.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click;
end;

procedure TfrmGruposdeEquipo.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click 
end;

procedure TfrmGruposdeEquipo.Paste1Click(Sender: TObject);
begin
   UtGrid.AddRowsFromClip;
end;

procedure TfrmGruposdeEquipo.Copy1Click(Sender: TObject);
begin
    UtGrid.CopyRowsToClip;
end;

procedure TfrmGruposdeEquipo.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure TfrmGruposdeEquipo.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click 
end;

procedure TfrmGruposdeEquipo.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmGruposdeEquipo.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TfrmGruposdeEquipo.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure TfrmGruposdeEquipo.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

procedure TfrmGruposdeEquipo.frmBarra1btnPrinterClick(Sender: TObject);
begin
//  If Plataformas.RecordCount > 0 Then
//    frxPlataformas.ShowReport(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP))
//  else
//     messageDLG('No existen datos para imprimir!', mtInformation, [mbOk], 0);
end;

end.

