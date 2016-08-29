unit frm_Condicionesclima;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, frm_barra, StdCtrls, DBCtrls,
  Mask, ExtCtrls, DB, Global, Menus, frxClass, frxDBSet,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UdbGrid,
  UnitExcepciones, unittbotonespermisos, unitactivapop, 
  ComCtrls, JvExComCtrls, JvDateTimePicker;

type
  TfrmCondicionesclima = class(TForm)
    grid_plataformas: TDBGrid;
    Label2: TLabel;
    frmBarra1: TfrmBarra;
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
    ds_condicionesdeclima: TDataSource;
    zq_Condicionesdeclima: TZQuery;
    tSeleccionarFecha: TJvDateTimePicker;
    zq_CondicionesdeclimaiId: TIntegerField;
    zq_CondicionesdeclimasContrato: TStringField;
    zq_CondicionesdeclimadFecha: TDateField;
    zq_CondicionesdeclimasOlas: TStringField;
    zq_CondicionesdeclimasVientos: TStringField;
    zq_CondicionesdeclimasPies: TStringField;
    zq_CondicionesdeclimasNudos: TStringField;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    Olas: TLabel;
    Label1: TLabel;
    Label3: TLabel;
    Button1: TButton;
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
    procedure grid_plataformasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_plataformasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_plataformasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure tSeleccionarFechaExit(Sender: TObject);
    procedure tsIdEquipoEnter(Sender: TObject);
    procedure tsIdEquipoExit(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCondicionesclima: TfrmCondicionesclima;
   UtGrid:TicDbGrid;
   botonpermiso: tbotonespermisos;
   sOpcion, lStatusOrig: string;
   FechaReporte: TDateTime;
implementation

{$R *.dfm}

procedure TfrmCondicionesclima.FormShow(Sender: TObject);
begin

  sMenuP:=stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cPlataformas', PopupPrincipal);
  OpcButton := '' ;
  frmbarra1.btnCancel.Click ;
  UtGrid:=TicdbGrid.create(grid_PLATAFORMAS);
  BotonPermiso.permisosBotones(frmBarra1);
  tSeleccionarFecha.DateTime := FechaReporte;
  zq_Condicionesdeclima.Active := False;
  zq_Condicionesdeclima.ParamByName('Fecha').AsDateTime := FechaReporte;
  zq_Condicionesdeclima.ParamByName('Contrato').AsString := Global_Contrato;
  zq_Condicionesdeclima.Open;
end;

procedure TfrmCondicionesclima.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  zq_Condicionesdeclima.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmCondicionesclima.tSeleccionarFechaExit(Sender: TObject);
begin
  zq_Condicionesdeclima.Active := False;
  zq_Condicionesdeclima.ParamByName('Fecha').AsDateTime := tSeleccionarFecha.DateTime;
  zq_Condicionesdeclima.ParamByName('Contrato').AsString := Global_Contrato;
  zq_Condicionesdeclima.Open;
end;

procedure TfrmCondicionesclima.tsIdEquipoEnter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmCondicionesclima.tsIdEquipoExit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmCondicionesclima.frmBarra1btnAddClick(Sender: TObject);
begin
   activapop(frmCondicionesclima, popupprincipal);
   zq_Condicionesdeclima.Append;
   zq_Condicionesdeclima.FieldByName('sContrato').AsString := Global_Contrato;
   zq_Condicionesdeclima.FieldByName('dFecha').AsDateTime := tSeleccionarFecha.DateTime;
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;

   DBEdit1.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;
end;

procedure TfrmCondicionesclima.frmBarra1btnEditClick(Sender: TObject);
begin
   activapop(frmCondicionesclima, popupprincipal);
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
     zq_Condicionesdeclima.Edit ;
   except
     on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Notas_Generales', 'Al agregar registro', 0);
     frmbarra1.btnCancel.Click ;
     end;
   end ;
   DBEdit1.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;

end;

procedure TfrmCondicionesclima.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin

    {Continua insercion de datos..}
  
  try
      desactivapop(popupprincipal);
      zq_Condicionesdeclima.Post ;
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

procedure TfrmCondicionesclima.frmBarra1btnCancelClick(Sender: TObject);
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
   zq_Condicionesdeclima.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := True;
   sOpcion := '';
end;

procedure TfrmCondicionesclima.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If zq_Condicionesdeclima.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        zq_Condicionesdeclima.Delete ;
      except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al eliminar registro', 0);
       end;
      end
    end
end;

procedure TfrmCondicionesclima.frmBarra1btnRefreshClick(Sender: TObject);
begin
  zq_Condicionesdeclima.refresh ;
end;

procedure TfrmCondicionesclima.frmBarra1btnExitClick(Sender: TObject);
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


procedure TfrmCondicionesclima.grid_plataformasCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
      frmbarra1.btnCancel.Click ;
end;

procedure TfrmCondicionesclima.grid_plataformasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmCondicionesclima.grid_plataformasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmCondicionesclima.grid_plataformasTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmCondicionesclima.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click;
end;

procedure TfrmCondicionesclima.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click 
end;

procedure TfrmCondicionesclima.Paste1Click(Sender: TObject);
begin
   UtGrid.AddRowsFromClip;
end;

procedure TfrmCondicionesclima.Copy1Click(Sender: TObject);
begin
    UtGrid.CopyRowsToClip;
end;

procedure TfrmCondicionesclima.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click
end;

procedure TfrmCondicionesclima.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click 
end;

procedure TfrmCondicionesclima.Button1Click(Sender: TObject);
begin
ShowMessage (Zq_Condicionesdeclima.Fieldbyname('sOlas').AsString)
end;

procedure TfrmCondicionesclima.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmCondicionesclima.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TfrmCondicionesclima.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure TfrmCondicionesclima.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

end.

