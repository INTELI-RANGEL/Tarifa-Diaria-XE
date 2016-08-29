unit frm_Kardex;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, RXDBCtrl, frm_connection, DBCtrls,UnitTBotonesPermisos,
  StdCtrls, Buttons, global, ZAbstractRODataset, ZDataset, udbgrid,
  Menus ;

type
  TfrmKardex = class(TForm)
    Grid_Kardex: TRxDBGrid;
    ds_usuarios: TDataSource;
    Filtro: TGroupBox;
    chkTodos: TCheckBox;
    tsIdUsuario: TDBLookupComboBox;
    Label1: TLabel;
    btnVisualizar: TBitBtn;
    ds_KardexSistema: TDataSource;
    btnSalir: TBitBtn;
    Usuarios: TZReadOnlyQuery;
    Kardex: TZReadOnlyQuery;
    PopupPrincipal: TPopupMenu;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    N4: TMenuItem;
    Salir1: TMenuItem;
    procedure tsIdUsuarioEnter(Sender: TObject);
    procedure tsIdUsuarioExit(Sender: TObject);
    procedure btnVisualizarClick(Sender: TObject);
    procedure Grid_KardexTitleBtnClick(Sender: TObject; ACol: Integer;
      Field: TField);
    procedure tsIdUsuarioKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnSalirClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Grid_KardexMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_KardexMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_KardexTitleClick(Column: TColumn);
    procedure Salir1Click(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmKardex: TfrmKardex;
  utgrid:ticdbgrid;
  BotonPermiso: TBotonesPermisos;
implementation

{$R *.dfm}

procedure TfrmKardex.tsIdUsuarioEnter(Sender: TObject);
begin
    tsIdUsuario.Color := global_color_entrada
end;

procedure TfrmKardex.tsIdUsuarioExit(Sender: TObject);
begin
    tsIdUsuario.Color := global_color_salida
end;

procedure TfrmKardex.btnVisualizarClick(Sender: TObject);
begin
    Kardex.Active := False ;
    Kardex.SQL.Clear ;
    If chkTodos.Checked Then
    Begin
        Kardex.SQL.Add('Select * From kardex_sistema Where sContrato = :Contrato Order By dIdFecha, sHora DESC') ;
        Kardex.Params.ParamByName('Contrato').DataType := ftString ;
        Kardex.Params.ParamByName('Contrato').Value := global_contrato ;
        Kardex.Open ;
    End
    Else
    Begin
        Kardex.SQL.Add('Select * From kardex_sistema Where sContrato = :Contrato And sIdUsuario = :Usuario Order By dIdFecha, sHora DESC') ;
        Kardex.Params.ParamByName('Contrato').DataType := ftString ;
        Kardex.Params.ParamByName('Contrato').Value := global_contrato ;
        Kardex.Params.ParamByName('Usuario').DataType := ftString ;
        Kardex.Params.ParamByName('Usuario').Value := tsIdUsuario.KeyValue ;
        Kardex.Open ;
    End;

end;

procedure TfrmKardex.Copy1Click(Sender: TObject);
begin
     UtGrid.CopyRowsToClip;
end;

procedure TfrmKardex.Grid_KardexMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
if grid_kardex.datasource.DataSet.IsEmpty=false  then

if grid_kardex.DataSource.DataSet.RecordCount>0 then
  UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmKardex.Grid_KardexMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
if grid_kardex.datasource.DataSet.IsEmpty=false  then
if grid_kardex.DataSource.DataSet.RecordCount>0  then

  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmKardex.Grid_KardexTitleBtnClick(Sender: TObject;
  ACol: Integer; Field: TField);
Var
  sCampo : String ;
begin
  sCampo := Field.FieldName ;

  Kardex.Active := False ;
  Kardex.SQL.Clear ;
  If sTipoOrden = 'ASC' Then
  Begin
      If chkTodos.Checked Then
      Begin
          Kardex.SQL.Add('Select * From kardex_sistema Where sContrato = :Contrato Order By :Ordenado ASC') ;
          Kardex.Params.ParamByName('Contrato').DataType := ftString ;
          Kardex.Params.ParamByName('Contrato').Value := global_contrato ;
          Kardex.Params.ParamByName('Ordenado').DataType := ftString ;
          Kardex.Params.ParamByName('Ordenado').Value := sCampo ;
          Kardex.Open ;
      End
      Else
      Begin
          Kardex.SQL.Add('Select * From kardex_sistema Where sContrato = :Contrato And sIdUsuario = :Usuario Order By :Ordenado ASC') ;
          Kardex.Params.ParamByName('Contrato').DataType := ftString ;
          Kardex.Params.ParamByName('Contrato').Value := global_contrato ;
          Kardex.Params.ParamByName('Usuario').DataType := ftString ;
          Kardex.Params.ParamByName('Usuario').Value := tsIdUsuario.KeyValue ;
          Kardex.Params.ParamByName('Ordenado').DataType := ftString ;
          Kardex.Params.ParamByName('Ordenado').Value := sCampo ;
          Kardex.Open ;
      End ;
      sTipoOrden := 'DESC'
  End
  Else
  Begin
      If chkTodos.Checked Then
      Begin
          Kardex.SQL.Add('Select * From kardex_sistema Where sContrato = :Contrato Order By :Ordenado DESC') ;
          Kardex.Params.ParamByName('Contrato').DataType := ftString ;
          Kardex.Params.ParamByName('Contrato').Value := global_contrato ;
          Kardex.Params.ParamByName('Ordenado').DataType := ftString ;
          Kardex.Params.ParamByName('Ordenado').Value := sCampo ;
          Kardex.Open ;
      End
      Else
      Begin
          Kardex.SQL.Add('Select * From kardex_sistema Where sContrato = :Contrato And sIdUsuario = :Usuario Order By :Ordenado DESC') ;
          Kardex.Params.ParamByName('Contrato').DataType := ftString ;
          Kardex.Params.ParamByName('Contrato').Value := global_contrato ;
          Kardex.Params.ParamByName('Usuario').DataType := ftString ;
          Kardex.Params.ParamByName('Usuario').Value := tsIdUsuario.KeyValue ;
          Kardex.Params.ParamByName('Ordenado').DataType := ftString ;
          Kardex.Params.ParamByName('Ordenado').Value := sCampo ;
          Kardex.Open ;
      End ;
      sTipoOrden := 'ASC'
  End ;
end;

procedure TfrmKardex.Grid_KardexTitleClick(Column: TColumn);
begin
if grid_kardex.datasource.DataSet.IsEmpty=false  then
if grid_kardex.DataSource.DataSet.RecordCount>0  then

   UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmKardex.Paste1Click(Sender: TObject);
begin
     UtGrid.AddRowsFromClip;
end;

procedure TfrmKardex.Salir1Click(Sender: TObject);
begin
    Close
end;

procedure TfrmKardex.tsIdUsuarioKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        btnVisualizar.SetFocus;




end;

procedure TfrmKardex.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  botonpermiso.free;
  action := cafree ;
  utgrid.Destroy;
end;

procedure TfrmKardex.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'mnuKardex',popupPrincipal);
  BotonPermiso.permisosBotones(nil);
  usuarios.Open ;
  UtGrid:=TicdbGrid.create(grid_kardex);
end;

procedure TfrmKardex.btnSalirClick(Sender: TObject);
begin
    Close
end;

end.
