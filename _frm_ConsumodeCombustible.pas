unit frm_ConsumodeCombustible;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, frm_barra, StdCtrls, DBCtrls,
  Mask, ExtCtrls, jpeg, ExtDlgs, DB, Global, Menus, frxClass, frxDBSet,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UdbGrid,
  UnitExcepciones, unittbotonespermisos, UnitValidaTexto, unitactivapop, UFunctionsGHH,
  ComCtrls, JvExComCtrls, JvDateTimePicker;

type
  TfrmConsumodeCombustible = class(TForm)
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
    ds_NotasGenerales: TDataSource;
    zq_NotasGenerales: TZQuery;
    tSeleccionarFecha: TJvDateTimePicker;
    DBMemo1: TDBMemo;
    zq_NotasGeneralessNotaGeneral: TMemoField;
    zq_NotasGeneralessContrato: TStringField;
    zq_NotasGeneralesdIdFecha: TDateField;
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
    procedure DBMemo1Enter(Sender: TObject);
    procedure DBLookupComboBox2Enter(Sender: TObject);
    procedure DBEdit1Enter(Sender: TObject);
    procedure tsIdEquipoExit(Sender: TObject);
    procedure DBMemo1Exit(Sender: TObject);
    procedure DBLookupComboBox2Exit(Sender: TObject);
    procedure DBEdit1Exit(Sender: TObject);
    procedure tsIdEquipoKeyPress(Sender: TObject; var Key: Char);
    procedure DBLookupComboBox2KeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit1KeyPress(Sender: TObject; var Key: Char);
  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmConsumodeCombustible: TfrmConsumodeCombustible;
   UtGrid:TicDbGrid;
   botonpermiso: tbotonespermisos;
   sOpcion, lStatusOrig: string;
   FechaReporte: TDateTime;
implementation

{$R *.dfm}

procedure TfrmConsumodeCombustible.FormShow(Sender: TObject);
begin


  zQuery1.ParamByName('Contrato').AsString := Global_Contrato;
  zQuery1.Open;
  zQuery2.Open;


  zq_Equipos.Active := False;
  zq_Equipos.ParamByName('Contrato').AsString := Global_Contrato;
  zq_Equipos.Open;
  zq_recursos.Active := False;
  zq_recursos.Open;

  sMenuP:=stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cPlataformas', PopupPrincipal);
  OpcButton := '' ;
  frmbarra1.btnCancel.Click ;
  UtGrid:=TicdbGrid.create(grid_PLATAFORMAS);
  BotonPermiso.permisosBotones(frmBarra1);
  tSeleccionarFecha.DateTime := FechaReporte;
  zq_ConsumosPorEquipo.Active := False;
  zq_ConsumosPorEquipo.ParamByName('Fecha').AsDateTime := FechaReporte;
  zq_ConsumosPorEquipo.ParamByName('Contrato').AsString := Global_Contrato;
  zq_ConsumosPorEquipo.Open;
end;

procedure TfrmConsumodeCombustible.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  zq_ConsumosPorEquipo.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmConsumodeCombustible.tSeleccionarFechaExit(Sender: TObject);
begin
  zq_ConsumosPorEquipo.Active := False;
  zq_ConsumosPorEquipo.ParamByName('Fecha').AsDateTime := tSeleccionarFecha.DateTime;
  zq_ConsumosPorEquipo.ParamByName('Contrato').AsString := Global_Contrato;
  zq_ConsumosPorEquipo.Open;
end;

procedure TfrmConsumodeCombustible.tsIdEquipoEnter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmConsumodeCombustible.tsIdEquipoExit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmConsumodeCombustible.tsIdEquipoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    DBMemo1.SetFocus;
end;

procedure TfrmConsumodeCombustible.frmBarra1btnAddClick(Sender: TObject);
begin
   activapop(frmConsumodeCombustible, popupprincipal);
   zq_NotasGenerales.Append;
   zq_NotasGenerales.FieldByName('sContrato').AsString := Global_Contrato;
   zq_NotasGenerales.FieldByName('dIdFecha').AsDateTime := tSeleccionarFecha.DateTime;
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;

   DBMemo1.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;
end;

procedure TfrmConsumodeCombustible.frmBarra1btnEditClick(Sender: TObject);
begin
   activapop(frmConsumodeCombustible, popupprincipal);
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
     zq_NotasGenerales.Edit ;
   except
     on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al agregar registro', 0);
     frmbarra1.btnCancel.Click ;
     end;
   end ;
   DBMemo1.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;

end;

procedure TfrmConsumodeCombustible.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin

    {Continua insercion de datos..}
  
  try
      desactivapop(popupprincipal);
      zq_NotasGenerales.Post ;
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

procedure TfrmConsumodeCombustible.frmBarra1btnCancelClick(Sender: TObject);
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
   zq_ConsumosPorEquipo.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := True;
   sOpcion := '';
end;

procedure TfrmConsumodeCombustible.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If zq_ConsumosPorEquipo.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        zq_ConsumosPorEquipo.Delete ;
      except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al eliminar registro', 0);
       end;
      end
    end
end;

procedure TfrmConsumodeCombustible.frmBarra1btnRefreshClick(Sender: TObject);
begin
  zq_ConsumosPorEquipo.refresh ;
end;

procedure TfrmConsumodeCombustible.frmBarra1btnExitClick(Sender: TObject);
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


procedure TfrmConsumodeCombustible.grid_plataformasCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
      frmbarra1.btnCancel.Click ;
end;

procedure TfrmConsumodeCombustible.grid_plataformasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmConsumodeCombustible.grid_plataformasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmConsumodeCombustible.grid_plataformasTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmConsumodeCombustible.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click;
end;

procedure TfrmConsumodeCombustible.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click 
end;

procedure TfrmConsumodeCombustible.Paste1Click(Sender: TObject);
begin
   UtGrid.AddRowsFromClip;
end;

procedure TfrmConsumodeCombustible.Copy1Click(Sender: TObject);
begin
    UtGrid.CopyRowsToClip;
end;

procedure TfrmConsumodeCombustible.DBEdit1Enter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmConsumodeCombustible.DBEdit1Exit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmConsumodeCombustible.DBEdit1KeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsIdEquipo.SetFocus;
end;

procedure TfrmConsumodeCombustible.DBLookupComboBox2Enter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmConsumodeCombustible.DBLookupComboBox2Exit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmConsumodeCombustible.DBLookupComboBox2KeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    DBEdit1.SetFocus;
end;

procedure TfrmConsumodeCombustible.DBMemo1Enter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmConsumodeCombustible.DBMemo1Exit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmConsumodeCombustible.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure TfrmConsumodeCombustible.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click 
end;

procedure TfrmConsumodeCombustible.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmConsumodeCombustible.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TfrmConsumodeCombustible.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure TfrmConsumodeCombustible.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

end.

