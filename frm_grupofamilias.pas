unit frm_grupofamilias;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, global, 
  StdCtrls, ExtCtrls, DBCtrls, Mask, frm_barra, adodb, db, Menus, OleCtrls,
  ZAbstractRODataset, ZAbstractDataset, ZDataset,
  udbgrid, unitexcepciones, unittbotonespermisos, UnitValidaTexto,
  unitactivapop;

type
  TfrmGrupoFamilias = class(TForm)
    grid_GruposIsometrico: TDBGrid;
    Label2: TLabel;
    Label9: TLabel;
    frmBarra1: TfrmBarra;
    tsIdGrupo: TDBEdit;
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
    N4: TMenuItem;
    Cut1: TMenuItem;
    Copy1: TMenuItem;
    N3: TMenuItem;
    Salir1: TMenuItem;
    qryGruposFamilias: TZQuery;
    dsgfamilias: TDataSource;
    qryGruposFamiliassIdFamilia: TStringField;
    qryGruposFamiliassDescripcion: TStringField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure grid_GruposIsometricoCellClick(Column: TColumn);
    procedure tsIdPersonalKeyPress(Sender: TObject; var Key: Char);
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
    procedure grid_GruposIsometricoEnter(Sender: TObject);
    procedure grid_GruposIsometricoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grid_GruposIsometricoKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure tsIdGrupoEnter(Sender: TObject);
    procedure tsIdGrupoExit(Sender: TObject);
    procedure tsIdGrupoKeyPress(Sender: TObject; var Key: Char);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure tsDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure tlFaseKeyPress(Sender: TObject; var Key: Char);
    procedure grid_GruposIsometricoMouseMove(Sender: TObject;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_GruposIsometricoMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure grid_GruposIsometricoTitleClick(Column: TColumn);
    procedure Cut1Click(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmGrupoFamilias : TfrmGrupoFamilias;
  utgrid:ticdbgrid;
  sOldId: string;
  botonpermiso:tbotonespermisos;
  sOpcion : string;
implementation

{$R *.dfm}

procedure TfrmGrupoFamilias.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
  end;

procedure TfrmGrupoFamilias.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'MnuFamiliadePro', PopupPrincipal);
  OpcButton := '' ;
  frmbarra1.btnCancel.Click ;

  qryGruposFamilias.Active := False ;
  qryGruposFamilias.Open ;
  Grid_GruposIsometrico.SetFocus;
  UtGrid:=TicdbGrid.create(grid_gruposisometrico);
  BotonPermiso.permisosBotones(frmBarra1);
  frmbarra1.btnPrinter.Enabled := False;
end;
procedure TfrmGrupoFamilias.grid_GruposIsometricoCellClick(Column: TColumn);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure TfrmGrupoFamilias.tsIdPersonalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsDescripcion.SetFocus ;
end;

procedure TfrmGrupoFamilias.frmBarra1btnAddClick(Sender: TObject);
begin
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   qryGruposFamilias.Append ;
   qryGruposFamilias.FieldValues['sIdFamilia']      := '' ;
   qryGruposFamilias.FieldValues['sDescripcion']    := '' ;
   tsIdGrupo.SetFocus ;
   activapop(frmGrupoFamilias,popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   frmbarra1.btnPrinter.Enabled := False;
   grid_gruposisometrico.Enabled := False;
end;

procedure TfrmGrupoFamilias.frmBarra1btnEditClick(Sender: TObject);
begin
    If qryGruposFamilias.RecordCount > 0 Then
    Begin
        try
       frmBarra1.btnEditClick(Sender);
       Insertar1.Enabled := False ;
       Editar1.Enabled := False ;
       Registrar1.Enabled := True ;
       Can1.Enabled := True ;
       Eliminar1.Enabled := False ;
       Refresh1.Enabled := False ;
       Salir1.Enabled := False ;
       sOpcion := 'Edit';
       sOldId := qryGruposFamilias.FieldValues['sIdFamilia'];
       qryGruposFamilias.Edit ;
       tsIdGrupo.SetFocus;
       activapop(frmGrupoFamilias,popupprincipal)
        except
           on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Clasificacion de Familia de Materiales', 'Al agregar registro', 0);
           end;
        end;
    End;
   BotonPermiso.permisosBotones(frmBarra1);
   frmbarra1.btnPrinter.Enabled := False;
end;

procedure TfrmGrupoFamilias.frmBarra1btnPostClick(Sender: TObject);
var
  lEdicion: boolean;
  nombres, cadenas: TStringList;
begin
  {Validaciones de campos}
  nombres:=TStringList.Create;cadenas:=TStringList.Create;
  nombres.Add('Grupo');         nombres.Add('Descripcion');
  cadenas.Add(tsIdGrupo.Text); cadenas.Add(tsDescripcion.Text);

  if not validaTexto(nombres, cadenas, 'Grupo',tsIdGrupo.Text) then
  begin
     MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
     exit;
  end;

  {Continua insercion de datos..}

  lEdicion := qryGruposFamilias.state = dsEdit;//capturar la bandera para usarla luego del post
  Try
     qryGruposFamilias.Post ;
     Insertar1.Enabled := True ;
     Editar1.Enabled := True ;
     Registrar1.Enabled := False ;
     Can1.Enabled := False ;
     Eliminar1.Enabled := True ;
     Refresh1.Enabled := True ;
     Salir1.Enabled := True ;
     frmBarra1.btnPostClick(Sender);
     desactivapop(popupprincipal);
  except
     on e : exception do begin
       UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Clasificacion de Familia de Materiales', 'Al salvar registro', 0);
       frmBarra1.btnCancel.Click ;
       lEdicion := false;
     end;
  end;
  if (lEdicion) and (sOldId <> qryGruposFamilias.FieldValues['sIdFamilia']) then begin
    //El registro fue editado y su ID cambio, es necesario actualizar este ID en tablas dependientes
    Connection.zCommand.Active := False ;
    Connection.zCommand.SQL.Clear ;
    Connection.zCommand.SQL.Add('UPDATE insumos SET sIdGrupo = :nuevo WHERE sIdGrupo = :viejo');
    Connection.zCommand.Params.ParamByName('nuevo').value := qryGruposFamilias.FieldValues['sIdFamilia'];
    Connection.zCommand.Params.ParamByName('viejo').value := sOldId;
    try
      Connection.zCommand.ExecSQL;
    except
      MessageDlg('Ocurrio un error al actualizar los registros de la tabla dependiente "insumos".', mtInformation, [mbOk], 0);
    end;
  end;
  BotonPermiso.permisosBotones(frmBarra1);
  frmbarra1.btnPrinter.Enabled := False;
  if sOpcion = 'Edit' then
  begin
       grid_gruposisometrico.Enabled := True;
       sOpcion := '';
  end;
end;

procedure TfrmGrupoFamilias.frmBarra1btnCancelClick(Sender: TObject);
begin
   frmBarra1.btnCancelClick(Sender);
   qryGruposFamilias.Cancel ;
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   desactivapop(popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   frmbarra1.btnPrinter.Enabled := False;
   grid_gruposisometrico.Enabled := True;
   sOpcion := '';
end;

procedure TfrmGrupoFamilias.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If   qryGruposFamilias.RecordCount > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
          Connection.QryBusca.Active := False ;
          Connection.QryBusca.SQL.Clear ;
          Connection.QryBusca.SQL.Add('Select sIdGrupo from insumos Where sIdGrupo =:Grupo');
          Connection.QryBusca.Params.ParamByName('Grupo').DataType := ftString ;
          Connection.QryBusca.Params.ParamByName('Grupo').Value    := qryGruposFamilias.FieldValues['sIdFamilia'] ;
          Connection.QryBusca.Open ;
          If Connection.QryBusca.RecordCount > 0 Then
             MessageDlg('No se puede Borrar el Registro por que Existe en INSUMOS', mtInformation, [mbOk], 0)
          Else
             qryGruposFamilias.Delete ;
      except
         on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Clasificacion de Familia de Materiales', 'Al eliminar registro', 0);
         end;
      end
    end
end;

procedure TfrmGrupoFamilias.frmBarra1btnRefreshClick(Sender: TObject);
begin
    qryGruposFamilias.refresh ;
end;

procedure TfrmGrupoFamilias.frmBarra1btnExitClick(Sender: TObject);
begin
   frmBarra1.btnExitClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   Close
end;

procedure TfrmGrupoFamilias.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure TfrmGrupoFamilias.Copy1Click(Sender: TObject);
begin
UtGrid.AddRowsFromClip;
end;

procedure TfrmGrupoFamilias.Cut1Click(Sender: TObject);
begin
  UtGrid.CopyRowsToClip;
end;

procedure TfrmGrupoFamilias.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click
end;

procedure TfrmGrupoFamilias.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click
end;

procedure TfrmGrupoFamilias.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click
end;

procedure TfrmGrupoFamilias.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click
end;

procedure TfrmGrupoFamilias.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmGrupoFamilias.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

procedure TfrmGrupoFamilias.grid_GruposIsometricoEnter(Sender: TObject);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure TfrmGrupoFamilias.grid_GruposIsometricoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure TfrmGrupoFamilias.grid_GruposIsometricoKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure TfrmGrupoFamilias.grid_GruposIsometricoMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmGrupoFamilias.grid_GruposIsometricoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmGrupoFamilias.grid_GruposIsometricoTitleClick(Column: TColumn);
begin
   UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmGrupoFamilias.tsIdGrupoEnter(Sender: TObject);
begin
    tsIdGrupo.Color := global_color_entrada
end;

procedure TfrmGrupoFamilias.tsIdGrupoExit(Sender: TObject);
begin
    tsIdGrupo.Color := global_color_salida
end;

procedure TfrmGrupoFamilias.tsIdGrupoKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsDescripcion.SetFocus
end;

procedure TfrmGrupoFamilias.tsDescripcionEnter(Sender: TObject);
begin
    tsDescripcion.Color := global_color_entrada
end;

procedure TfrmGrupoFamilias.tsDescripcionExit(Sender: TObject);
begin
    tsDescripcion.Color := global_color_salida
end;

procedure TfrmGrupoFamilias.tsDescripcionKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 then
        tsIdGrupo.SetFocus
end;

procedure TfrmGrupoFamilias.tlFaseKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 then
        tsIdGrupo.SetFocus
end;

end.
