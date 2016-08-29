unit frm_programas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, StdCtrls, Mask, DBCtrls, frm_barra, Grids,
  DBGrids, global, DB, Menus, unitactivapop,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UnitExcepciones, udbgrid, unittbotonespermisos, UnitValidaTexto ;

type
  TfrmProgramas = class(TForm)
    grid_programas: TDBGrid;
    Label1: TLabel;
    frmBarra1: TfrmBarra;
    ds_programas: TDataSource;
    Label2: TLabel;
    tsIdPrograma: TDBEdit;
    tsDescripcion: TDBEdit;
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
    Programas: TZQuery;
    ProgramassIdPrograma: TStringField;
    ProgramassDescripcion: TStringField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure grid_programasEnter(Sender: TObject);
    procedure grid_programasKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grid_programasKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure grid_programasCellClick(Column: TColumn);
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
    procedure Salir1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure tsIdProgramaEnter(Sender: TObject);
    procedure tsIdProgramaExit(Sender: TObject);
    procedure tsIdProgramaKeyPress(Sender: TObject; var Key: Char);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure tsDescripcionKeyPress(Sender: TObject; var Key: Char);
    function actualizarDependencias(sTabla: string): boolean;
    procedure grid_programasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_programasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_programasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmProgramas: TfrmProgramas;
  Opcion : String ;
  Registro_Actual : String ;
  sOldPrograma : string;  //indica el sIdPrograma anterior para casos de edicion
  UtGrid:TicDbGrid;
  botonpermiso:tbotonespermisos;
implementation

{$R *.dfm}

procedure TfrmProgramas.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Programas.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmProgramas.grid_programasEnter(Sender: TObject);
begin
  If frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;
end;

procedure TfrmProgramas.grid_programasKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  If frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;
end;

procedure TfrmProgramas.grid_programasKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;
end;

procedure TfrmProgramas.grid_programasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmProgramas.grid_programasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmProgramas.grid_programasTitleClick(Column: TColumn);
begin
 UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmProgramas.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'adProgramas', PopupPrincipal);
  sOldPrograma := '';
  OpcButton    := '' ;
  frmBarra1.btnCancel.Click ;
  frmBarra1.btnPrinter.Enabled := False;
  Programas.Active := False ;
  Programas.Open ;
  Grid_Programas.SetFocus ;
  UtGrid:=TicdbGrid.create(grid_programas);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmProgramas.grid_programasCellClick(Column: TColumn);
begin
  If frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;
end;

procedure TfrmProgramas.frmBarra1btnAddClick(Sender: TObject);
begin
  frmBarra1.btnAddClick(Sender);
  frmBarra1.btnPrinter.Enabled := False;
  Insertar1.Enabled := False ;
  Editar1.Enabled := False ;
  Registrar1.Enabled := True ;
  Can1.Enabled := True ;
  Eliminar1.Enabled := False ;
  Refresh1.Enabled := False ;
  Salir1.Enabled := False ;
  Programas.Append ;
  Programas.FieldValues['sIdPrograma']  := '' ;
  Programas.FieldValues['sDescripcion'] := '' ;
  tsIdPrograma.SetFocus ;
  activapop(frmProgramas,popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmProgramas.frmBarra1btnEditClick(Sender: TObject);
begin
  if Programas.RecordCount > 0 then begin
      sOldPrograma := Programas.FieldByName('sIdPrograma').AsString;
      frmBarra1.btnEditClick(Sender);
      frmBarra1.btnPrinter.Enabled := False;
      Insertar1.Enabled := False ;
      Editar1.Enabled := False ;
      Registrar1.Enabled := True ;
      Can1.Enabled := True ;
      Eliminar1.Enabled := False ;
      Refresh1.Enabled := False ;
      Salir1.Enabled := False ;
      try
          Programas.Edit ;
      except
          on e : exception do begin
              UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Programas', 'Al editar registro', 0);
              frmbarra1.btnCancel.Click ;
          end;
      end ;
      tsIdPrograma.SetFocus;
  end;
activapop(frmProgramas,popupprincipal);
BotonPermiso.permisosBotones(frmBarra1);
end;

function TfrmProgramas.actualizarDependencias(sTabla: string): boolean;
var
  sUpdate: string;
begin
  result := true;
  if sOldPrograma <> Programas.FieldByName('sIdPrograma').AsString  then begin//el ID cambio, actualizar la tabla gruposxprograma
    sUpdate :=
    'UPDATE '+sTabla+' '+
    'SET sIdPrograma = :programa '+
    'WHERE sIdPrograma = :oldPrograma';
    try
      connection.zCommand.Active := false;
      connection.zCommand.SQL.Clear;
      connection.zCommand.SQL.Add(sUpdate);
      connection.zCommand.ParamByName('programa').Value := Programas.FieldByName('sIdPrograma').AsString;
      connection.zCommand.ParamByName('oldPrograma').Value := sOldPrograma;
      connection.zCommand.ExecSQL;
    except
      result := false;
    end;
  end;
end;

procedure TfrmProgramas.frmBarra1btnPostClick(Sender: TObject);
var
  lContinua : boolean;//bandera para indicar seguir procediendo o no segun excepciones
  nombres, cadenas: TStringList;
begin
   {Validacion de campos}
   nombres:=TStringList.Create;cadenas:=TStringList.Create;
   nombres.Add('Descripcion');cadenas.Add(tsDescripcion.Text);
   if not validaTexto(nombres, cadenas, 'Programa', tsIdPrograma.Text) then
   begin
     MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
     exit;
   end;
   {Continua insercion de datos}
   lContinua := true;
   try
      Programas.Post ;
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
           lContinua := false;
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Programas', 'Al salvar registro', 0);
           frmbarra1.btnCancel.Click ;
       end;
   end;
   if lContinua then begin  //si todo salio bien, actualizar el id Programa en la tabla gruposxprograma
       if not actualizarDependencias('gruposxprograma') then begin
           lContinua := false;
           MessageDlg('El registro se actualizo correctamente pero sus dependencias en asignacion de Programas a Grupos no.' + #10 +
           'Esto puede generar problemas de relación. Informa al administrador del sistema de este error.', mtInformation, [mbOk], 0);
       end;
   end;
   if lContinua then begin  //si todo salio bien, actualizar el id Programa en la tabla gruposxprograma
       if not actualizarDependencias('usuariosxprograma') then begin
           lContinua := false;
           MessageDlg('El registro se actualizo correctamente pero sus dependencias en asignacion de Programas a Usuarios no.' + #10 +
           'Esto puede generar problemas de relación. Informa al administrador del sistema de este error.', mtInformation, [mbOk], 0);
       end;
   end;
   if lContinua then
       sOldPrograma := '';//resetear la variable
   desactivapop(popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmProgramas.frmBarra1btnCancelClick(Sender: TObject);
begin
  frmBarra1.btnCancelClick(Sender);
  frmBarra1.btnPrinter.Enabled := False;
  Insertar1.Enabled := True ;
  Editar1.Enabled := True ;
  Registrar1.Enabled := False ;
  Can1.Enabled := False ;
  Eliminar1.Enabled := True ;
  Refresh1.Enabled := True ;
  Salir1.Enabled := True ;
  Programas.Cancel ;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
end;

procedure TfrmProgramas.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If Programas.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
        Connection.QryBusca.Active := False ;
        Connection.QryBusca.SQL.Clear ;
        Connection.QryBusca.SQL.Add('Select sIdPrograma from gruposxprograma Where sIdPrograma =:programa');
        Connection.QryBusca.Params.ParamByName('Programa').DataType := ftString ;
        Connection.QryBusca.Params.ParamByName('Programa').Value    := programas.FieldValues['sIdPrograma'] ;
        Connection.QryBusca.Open ;
        if Connection.QryBusca.RecordCount > 0 Then
        Begin
           MessageDlg('No se puede Borrar el Registro por que esta ASIGNADO a uno o mas registros en GRUPOS DE USUARIOS.', mtInformation, [mbOk], 0)
        End
        Else Begin
          Connection.QryBusca.Active := False ;
          Connection.QryBusca.SQL.Clear ;
          Connection.QryBusca.SQL.Add('Select sIdPrograma from usuariosxprograma Where sIdPrograma =:programa');
          Connection.QryBusca.Params.ParamByName('Programa').DataType := ftString ;
          Connection.QryBusca.Params.ParamByName('Programa').Value    := programas.FieldValues['sIdPrograma'] ;
          Connection.QryBusca.Open ;
          if Connection.QryBusca.RecordCount > 0 Then
          Begin
            MessageDlg('No se puede Borrar el Registro por que esta ASIGNADO a uno o mas registros en USUARIOS.', mtInformation, [mbOk], 0)
          End
          Else Begin

            try
              Programas.Delete ;
            except
              on e : exception do begin
                UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Programas', 'Al eliminar registro', 0);
                frmbarra1.btnCancel.Click ;
              end;
            end;

          End;
        End;

    end
end;

procedure TfrmProgramas.frmBarra1btnRefreshClick(Sender: TObject);
begin
  Programas.Active := False ;
  Programas.Open ;
end;

procedure TfrmProgramas.frmBarra1btnExitClick(Sender: TObject);
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

procedure TfrmProgramas.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click;
end;

procedure TfrmProgramas.Paste1Click(Sender: TObject);
begin
 UtGrid.AddRowsFromClip;
end;

procedure TfrmProgramas.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click;
end;

procedure TfrmProgramas.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click;
end;

procedure TfrmProgramas.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click;
end;

procedure TfrmProgramas.Copy1Click(Sender: TObject);
begin
    UtGrid.CopyRowsToClip;
end;

procedure TfrmProgramas.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click;
end;

procedure TfrmProgramas.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click;
end;

procedure TfrmProgramas.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click;
end;


procedure TfrmProgramas.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click
end;


procedure TfrmProgramas.tsIdProgramaEnter(Sender: TObject);
begin
    tsIdPrograma.Color := global_color_Entrada
end;

procedure TfrmProgramas.tsIdProgramaExit(Sender: TObject);
begin
    tsIdPrograma.Color := global_color_salida
end;

procedure TfrmProgramas.tsIdProgramaKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsDescripcion.SetFocus 
end;

procedure TfrmProgramas.tsDescripcionEnter(Sender: TObject);
begin
    tsDescripcion.Color := global_color_entrada
end;

procedure TfrmProgramas.tsDescripcionExit(Sender: TObject);
begin
    tsDescripcion.Color := global_color_salida
end;

procedure TfrmProgramas.tsDescripcionKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key =  #13 Then
        tsIdPrograma.SetFocus
end;

End.
