unit frm_CatNomFirmantes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_barra, StdCtrls, Mask, DBCtrls, Grids, DBGrids, ZAbstractDataset,
  ZDataset, DB, ZAbstractRODataset, ExtCtrls,unitActivaPop, Menus,udbgrid,
  UnitExcepciones,UnitTBotonesPermisos,frm_connection,global, AdvGlowButton;

type
  Tfrmcatnomfirmates = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Nombres: TZQuery;
    ds_nombres: TDataSource;
    grid_fases: TDBGrid;
    Label1: TLabel;
    Panel3: TPanel;
    frmBarra1: TfrmBarra;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    Label2: TLabel;
    Label3: TLabel;
    NombresiIdCatNombreFirmante: TIntegerField;
    NombresdFechaAlta: TDateField;
    NombressNombre: TStringField;
    NombressAPaterno: TStringField;
    NombressAMaterno: TStringField;
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
    Copiar1: TMenuItem;
    Paste1: TMenuItem;
    N4: TMenuItem;
    Salir1: TMenuItem;
    btnSelect: TAdvGlowButton;
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure DBEdit1Enter(Sender: TObject);
    procedure DBEdit1Exit(Sender: TObject);
    procedure DBEdit2Exit(Sender: TObject);
    procedure DBEdit3Exit(Sender: TObject);
    procedure DBEdit2Enter(Sender: TObject);
    procedure DBEdit3Enter(Sender: TObject);
    procedure DBEdit1KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit2KeyPress(Sender: TObject; var Key: Char);
    procedure NombresBeforePost(DataSet: TDataSet);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Copiar1Click(Sender: TObject);
    procedure btnSelectClick(Sender: TObject);
  private

    { Private declarations }
  public
    Seleccionar:Boolean;
    { Public declarations }
  end;

var
  frmcatnomfirmates: Tfrmcatnomfirmates;
  UtGrid:TicDbGrid;
  botonpermiso: tbotonespermisos;
  banderaagregar:boolean;
implementation

{$R *.dfm}

procedure Tfrmcatnomfirmates.Copiar1Click(Sender: TObject);
begin
  UtGrid.CopyRowsToClip;
end;

procedure Tfrmcatnomfirmates.DBEdit1Enter(Sender: TObject);
begin
   DBEdit1.color := global_color_entrada
end;

procedure Tfrmcatnomfirmates.DBEdit1Exit(Sender: TObject);
begin
  DBEdit1.color := global_color_salida
end;

procedure Tfrmcatnomfirmates.DBEdit1KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    DBEdit2.SetFocus
end;

procedure Tfrmcatnomfirmates.DBEdit2Enter(Sender: TObject);
begin
  DBEdit2.color := global_color_entrada
end;

procedure Tfrmcatnomfirmates.DBEdit2Exit(Sender: TObject);
begin
  DBEdit2.color := global_color_salida
end;

procedure Tfrmcatnomfirmates.DBEdit2KeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    DBEdit3.SetFocus
end;

procedure Tfrmcatnomfirmates.DBEdit3Enter(Sender: TObject);
begin
  DBEdit3.color := global_color_entrada
end;

procedure Tfrmcatnomfirmates.DBEdit3Exit(Sender: TObject);
begin
  DBEdit3.color := global_color_salida
end;



procedure Tfrmcatnomfirmates.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  nombres.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure Tfrmcatnomfirmates.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'mnuAgrupacionP', PopupPrincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  Seleccionar := False;
  OpcButton := '' ;
  frmbarra1.btnCancel.Click ;
  nombres.Active := False ;
  nombres.Open ;

  frmbarra1.btnPrinter.Enabled:=false;
  UtGrid:=TicdbGrid.create(grid_fases);
end;

procedure Tfrmcatnomfirmates.frmBarra1btnAddClick(Sender: TObject);
begin
  try
    try
      activapop(frmcatnomfirmates, popupprincipal);
    except
      ;
    end;

    banderaagregar:=true;
    frmBarra1.btnAddClick(Sender);
    Insertar1.Enabled := False ;
    Editar1.Enabled := False ;
    Registrar1.Enabled := True ;
    Can1.Enabled := True ;
    Eliminar1.Enabled := False ;
    Refresh1.Enabled := False ;
    Salir1.Enabled := False ;
    Nombres.Append ;
    Nombres.FieldByName('dfechaalta').AsDateTime := now ;
  except
    on e : exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Nombres de Firmantes', 'Al agregar registro',0)    end;
  end;
  DBEdit1.SetFocus ;
  BotonPermiso.permisosBotones(frmBarra1);
  frmbarra1.btnPrinter.Enabled:=false;
  grid_fases.Enabled:=false;

end;

procedure Tfrmcatnomfirmates.frmBarra1btnCancelClick(Sender: TObject);
begin
  try
    desactivapop(popupprincipal);
  except
      ;
  end;
  frmBarra1.btnCancelClick(Sender);
  Insertar1.Enabled := True ;
  Editar1.Enabled := True ;
  Registrar1.Enabled := False ;
  Can1.Enabled := False ;
  Eliminar1.Enabled := True ;
  Refresh1.Enabled := True ;
  Salir1.Enabled := True ;
  Nombres.Cancel ;
  BotonPermiso.permisosBotones(frmBarra1);
  frmbarra1.btnPrinter.Enabled:=false;
  grid_fases.Enabled:=true;
end;

procedure Tfrmcatnomfirmates.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If Nombres.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
       Nombres.Delete ;
      except
        on e : exception do begin
         UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Nombres de firmantes', 'Al eliminar registro', 0);
        end;
      end
    end
end;

procedure Tfrmcatnomfirmates.frmBarra1btnEditClick(Sender: TObject);
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
    try
      activapop(frmcatnomfirmates, popupprincipal);
    except
      ;
    end;
    Nombres.Edit ;
  except
    on e : exception do
    begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Nombres de Firmantes', 'Al editar registro', 0);
      frmbarra1.btnCancel.Click ;
    end;
  end ;
  DBEdit1.SetFocus;
  BotonPermiso.permisosBotones(frmBarra1);
  frmbarra1.btnPrinter.Enabled:=false;
  grid_fases.Enabled:=false;
end;

procedure Tfrmcatnomfirmates.frmBarra1btnExitClick(Sender: TObject);
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
procedure Tfrmcatnomfirmates.frmBarra1btnPostClick(Sender: TObject);
begin
  try
    try
      desactivapop(popupprincipal);
    except
      ;
    end;
    Nombres.Post ;
    Insertar1.Enabled := True ;
    Editar1.Enabled := True ;
    Registrar1.Enabled := False ;
    Can1.Enabled := False ;
    Eliminar1.Enabled := True ;
    Refresh1.Enabled := True ;
    Salir1.Enabled := True ;
    frmBarra1.btnPostClick(Sender);
  except
    on e : exception do
    begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Nombres de Firmantes', 'Al salvar registro', 0);
      frmbarra1.btnCancel.Click ;
    end;
  end ;
  BotonPermiso.permisosBotones(frmBarra1);
  frmbarra1.btnPrinter.Enabled:=false;
  grid_fases.Enabled:=true;
  frmbarra1.btnCancel.Click;
  if banderaAgregar then
    frmbarra1.btnAdd.Click;
end;

procedure Tfrmcatnomfirmates.frmBarra1btnRefreshClick(Sender: TObject);
begin
  try
    nombres.Active := False ;
    nombres.Open
  except
    on e : exception do
    begin
       UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Nombres de Firmantes', 'Al actualizar Grid',0);    end;
  end;
end;

procedure Tfrmcatnomfirmates.Insertar1Click(Sender: TObject);
begin
  frmBarra1.btnAdd.Click;
end;

procedure Tfrmcatnomfirmates.NombresBeforePost(DataSet: TDataSet);
begin
  if nombres.state = dsinsert then
  begin
    connection.qrybusca.active := false;
    connection.qrybusca.sql.clear;
    connection.qrybusca.sql.text := 'SELECT max(iIdCatNombreFirmante)+1 as maximo from catnombrefirmantes';
    connection.qrybusca.open;
    nombres.fieldbyname('iIdCatNombreFirmante').asinteger :=  connection.qrybusca.fieldbyname('maximo').asinteger;
  end;
end;


procedure Tfrmcatnomfirmates.Paste1Click(Sender: TObject);
begin
  try
    UtGrid.AddRowsFromClip;
  except
    on e : exception do
    begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Categorias de Personal', 'Al pegar registro', 0);
    end;
  end;

end;

procedure Tfrmcatnomfirmates.Editar1Click(Sender: TObject);
begin
  frmBarra1.btnEdit.Click
end;

procedure Tfrmcatnomfirmates.Registrar1Click(Sender: TObject);
begin
  frmBarra1.btnPost.Click
end;

procedure Tfrmcatnomfirmates.btnSelectClick(Sender: TObject);
begin
  Seleccionar := True;
  close;
end;

procedure Tfrmcatnomfirmates.Can1Click(Sender: TObject);
begin
  frmBarra1.btnCancel.Click
end;

procedure Tfrmcatnomfirmates.Eliminar1Click(Sender: TObject);
begin
  frmBarra1.btnDelete.Click
end;

procedure Tfrmcatnomfirmates.Refresh1Click(Sender: TObject);
begin
  frmBarra1.btnRefresh.Click
end;

procedure Tfrmcatnomfirmates.Salir1Click(Sender: TObject);
begin
  frmBarra1.btnExit.Click
end;

end.
