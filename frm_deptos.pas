unit frm_deptos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, global, Grids, DBGrids, frm_barra, StdCtrls,unitactivapop,
  Mask, DBCtrls, DB, Menus, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  udbgrid, unitexcepciones, unittbotonespermisos, UnitValidaTexto;

type
  tfrmDeptos = class(TForm)
    grid_deptos: TDBGrid;
    Label1: TLabel;
    Label2: TLabel;
    tsIdDepartamento: TDBEdit;
    tsDescripcion: TDBEdit;
    frmBarra1: TfrmBarra;
    Label4: TLabel;
    tsJefatura: TDBEdit;
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
    ds_deptos: TDataSource;
    deptos: TZQuery;
    procedure tsIdDepartamentoKeyPress(Sender: TObject; var Key: Char);
    procedure tsDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure tsJefaturaKeyPress(Sender: TObject; var Key: Char);
    procedure grid_deptosCellClick(Column: TColumn);
    procedure FormShow(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tsIdDepartamentoEnter(Sender: TObject);
    procedure tsIdDepartamentoExit(Sender: TObject);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure tsJefaturaEnter(Sender: TObject);
    procedure tsJefaturaExit(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure grid_deptosMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_deptosMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_deptosTitleClick(Column: TColumn);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmDeptos: TfrmDeptos;
  UtGrid:TicDbGrid;
  botonpermiso:tbotonespermisos;
implementation

{$R *.dfm}

procedure TfrmDeptos.tsIdDepartamentoKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsDescripcion.SetFocus
end;

procedure TfrmDeptos.tsDescripcionKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsJefatura.SetFocus
end;


procedure TfrmDeptos.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  utgrid.Destroy;
  deptos.Cancel ;
  action := cafree ;
  botonpermiso.Free;
end;

procedure TfrmDeptos.frmBarra1btnAddClick(Sender: TObject);
begin
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   deptos.Append ;
   deptos.FieldValues ['sIdDepartamento'] := '' ;
   deptos.FieldValues ['sDescripcion'] := '' ;
   deptos.FieldValues ['sJefatura'] := '' ;
   tsIdDepartamento.SetFocus ;
   activapop(frmDeptos,popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
   grid_deptos.Enabled := False;
end;

procedure TfrmDeptos.frmBarra1btnEditClick(Sender: TObject);
begin
   frmBarra1.btnEditClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   try
       deptos.Edit ;
   except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Departamentos', 'Al editar registro', 0);
       frmBarra1.btnCancel.Click ;
       end;
   end ;
   tsIdDepartamento.SetFocus;
   activapop(frmDeptos,popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
   grid_deptos.Enabled := False;
end;

procedure TfrmDeptos.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
  {Validacion de campos}
  nombres:=TStringList.Create;cadenas:=TStringList.Create;
  nombres.Add('Descripcion');nombres.Add('Jefatura');cadenas.Add(tsDescripcion.Text);cadenas.Add(tsJefatura.Text);
  if not validaTexto(nombres, cadenas, 'Departamento', tsIdDepartamento.Text) then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
  {Continua insercion de datos}
  try
    deptos.Post ;
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
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Departamentos', 'Al salvar registro', 0);
      frmBarra1.btnCancel.Click ;
    end;
  end;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
  grid_deptos.Enabled := True;
end;

procedure TfrmDeptos.frmBarra1btnCancelClick(Sender: TObject);
begin
   frmBarra1.btnCancelClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   deptos.Cancel ;
   desactivapop(popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
   grid_deptos.Enabled := True;
end;

procedure TfrmDeptos.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If deptos.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        deptos.Delete ;
      except
        on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Departamentos', 'Al eliminar registro', 0);
        end;
      end
    end
end;

procedure TfrmDeptos.frmBarra1btnRefreshClick(Sender: TObject);
begin
  deptos.Refresh
end;

procedure TfrmDeptos.frmBarra1btnExitClick(Sender: TObject);
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

procedure TfrmDeptos.tsJefaturaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsIdDepartamento.SetFocus
end;

procedure TfrmDeptos.grid_deptosCellClick(Column: TColumn);
begin
  if frmBarra1.btnCancel.Enabled = True then
       frmBarra1.btnCancel.Click ;
end;

procedure tfrmDeptos.grid_deptosMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure tfrmDeptos.grid_deptosMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure tfrmDeptos.grid_deptosTitleClick(Column: TColumn);
begin
 UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmDeptos.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'adDeptos', PopupPrincipal);
  UtGrid:=TicdbGrid.create(grid_deptos);
  OpcButton := '' ;
  frmBarra1.btnCancel.Click ;
  deptos.Active := False ;
  deptos.SQL.Clear ;
  deptos.SQL.Add('select * from departamentos order by sIdDepartamento' ) ;
  deptos.Open ;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
end;

procedure tfrmDeptos.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure tfrmDeptos.Paste1Click(Sender: TObject);
begin
UtGrid.AddRowsFromClip;
end;

procedure tfrmDeptos.Copy1Click(Sender: TObject);
begin
UtGrid.CopyRowsToClip;
end;

procedure tfrmDeptos.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure tfrmDeptos.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click
end;

procedure tfrmDeptos.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure tfrmDeptos.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure tfrmDeptos.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure tfrmDeptos.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

procedure tfrmDeptos.tsIdDepartamentoEnter(Sender: TObject);
begin
    tsIdDepartamento.Color := global_color_entrada
end;

procedure tfrmDeptos.tsIdDepartamentoExit(Sender: TObject);
begin
    tsIdDepartamento.Color := global_color_salida
end;

procedure tfrmDeptos.tsDescripcionEnter(Sender: TObject);
begin
    tsDescripcion.Color := global_color_entrada
end;

procedure tfrmDeptos.tsDescripcionExit(Sender: TObject);
begin
    tsDescripcion.Color := global_color_salida
end;

procedure tfrmDeptos.tsJefaturaEnter(Sender: TObject);
begin
    tsJefatura.Color := global_color_entrada
end;

procedure tfrmDeptos.tsJefaturaExit(Sender: TObject);
begin
    tsJefatura.Color := global_color_salida
end;

end.
