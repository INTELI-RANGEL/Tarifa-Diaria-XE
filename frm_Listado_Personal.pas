unit frm_Listado_Personal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, dblookup, StdCtrls, Mask, DBCtrls, Grids, DBGrids, ExtCtrls,
  frm_barra, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, UnitExcepciones, global,
  AdvCombo, rxToolEdit, RXDBCtrl;

type
  TfrmListado_Personal = class(TForm)
    frmBarra1: TfrmBarra;
    Panel1: TPanel;
    Label1: TLabel;
    edtFicha: TDBEdit;
    edtNombre: TDBEdit;
    edtApellidoP: TDBEdit;
    edtApellidoM: TDBEdit;
    edtRfc: TDBEdit;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    zq_listadoper: TZQuery;
    zq_compania: TZQuery;
    zq_categoria: TZQuery;
    ds_listadoper: TDataSource;
    ds_compania: TDataSource;
    ds_categoria: TDataSource;
    lkcbCompania: TDBLookupComboBox;
    GridListPer: TDBGrid;
    Label8: TLabel;
    lkcbEsp: TDBLookupComboBox;
    zq_Esp: TZQuery;
    ds_Esp: TDataSource;
    lkcbCategoria: TDBLookupComboBox;
    Panel2: TPanel;
    edtBuscar: TEdit;
    cbbBuscar: TAdvComboBox;
    btnBuscar: TButton;
    edtLibreta: TDBEdit;
    Label9: TLabel;
    Label10: TLabel;
    dedtVigencia: TDBDateEdit;
    procedure cambio_estado;
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure ds_categoriaDataChange(Sender: TObject; Field: TField);
    procedure edtFichaKeyPress(Sender: TObject; var Key: Char);
    procedure edtNombreKeyPress(Sender: TObject; var Key: Char);
    procedure edtApellidoPKeyPress(Sender: TObject; var Key: Char);
    procedure edtApellidoMKeyPress(Sender: TObject; var Key: Char);
    procedure edtRfcKeyPress(Sender: TObject; var Key: Char);
    procedure lkcbCategoriaKeyPress(Sender: TObject; var Key: Char);
    procedure lkcbEspKeyPress(Sender: TObject; var Key: Char);
    procedure edtFichaEnter(Sender: TObject);
    procedure edtNombreEnter(Sender: TObject);
    procedure edtApellidoPEnter(Sender: TObject);
    procedure edtApellidoMEnter(Sender: TObject);
    procedure edtRfcEnter(Sender: TObject);
    procedure lkcbCategoriaEnter(Sender: TObject);
    procedure lkcbEspEnter(Sender: TObject);
    procedure lkcbCompaniaEnter(Sender: TObject);
    procedure edtFichaExit(Sender: TObject);
    procedure edtNombreExit(Sender: TObject);
    procedure edtApellidoPExit(Sender: TObject);
    procedure edtApellidoMExit(Sender: TObject);
    procedure edtRfcExit(Sender: TObject);
    procedure lkcbCategoriaExit(Sender: TObject);
    procedure lkcbEspExit(Sender: TObject);
    procedure lkcbCompaniaExit(Sender: TObject);
    procedure zq_listadoperAfterScroll(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnBuscarClick(Sender: TObject);
    procedure cbbBuscarKeyPress(Sender: TObject; var Key: Char);
    procedure edtBuscarKeyPress(Sender: TObject; var Key: Char);
    procedure lkcbCompaniaKeyPress(Sender: TObject; var Key: Char);
    procedure edtLibretaKeyPress(Sender: TObject; var Key: Char);
    procedure dedtVigenciaEnter(Sender: TObject);
    procedure edtLibretaEnter(Sender: TObject);
    procedure edtLibretaExit(Sender: TObject);
    procedure dedtVigenciaExit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  IsOpen: Boolean;
  frmListado_Personal: TfrmListado_Personal;

implementation

uses frm_connection;

{$R *.dfm}

procedure TfrmListado_Personal.btnBuscarClick(Sender: TObject);
begin
  if zq_listadoper.RecordCount>0 then
    if cbbBuscar.ItemIndex=0 then
    begin
      if not zq_listadoper.Locate('sIdTripulacion',trim(edtBuscar.Text),[LoCaseInsensitive]) then
        MessageDlg('No se Encontro la Ficha Buscada',mtInformation,[MbOk],0);
    end
    else
    begin
      if not zq_listadoper.Locate('sNombre',trim(edtBuscar.Text),[LoCaseInsensitive]) then
        MessageDlg('No se Encontro el Nombre Buscado',mtInformation,[MbOk],0);
    end;
end;

procedure TfrmListado_Personal.cambio_estado;
begin
  if zq_listadoper.State in [dsInsert,dsEdit] then
  begin
    frmBarra1.btnAdd.Enabled      :=False;
    frmBarra1.btnEdit.Enabled        :=False;
    frmBarra1.btnPost.Enabled       :=True;
    frmBarra1.btnCancel.Enabled      :=True;
    frmBarra1.btnDelete.Enabled      :=False;
    frmBarra1.btnPrinter.Enabled      :=False;
    frmBarra1.btnRefresh.Enabled     :=False;
    frmBarra1.btnExit.Enabled         :=False;
    GridListPer.Enabled    :=False;
  end
  else
  if zq_listadoper.State in [dsBrowse] then
  begin
    frmBarra1.btnAdd.Enabled      :=True;
    frmBarra1.btnEdit.Enabled        :=True;
    frmBarra1.btnPost.Enabled       :=False;
    frmBarra1.btnCancel.Enabled      :=False;
    frmBarra1.btnDelete.Enabled      :=True;
    frmBarra1.btnPrinter.Enabled      :=True;
    frmBarra1.btnRefresh.Enabled     :=True;
    frmBarra1.btnExit.Enabled         :=True;
    GridListPer.Enabled    :=True;
  end;
end;

procedure TfrmListado_Personal.cbbBuscarKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    edtbuscar.SetFocus
end;

procedure TfrmListado_Personal.dedtVigenciaEnter(Sender: TObject);
begin
  dedtVigencia.color := global_color_entrada
end;

procedure TfrmListado_Personal.dedtVigenciaExit(Sender: TObject);
begin
  dedtVigencia.color := global_color_salida
end;

procedure TfrmListado_Personal.ds_categoriaDataChange(Sender: TObject;
  Field: TField);
begin
  zq_Esp.Active:=False;
  zq_Esp.ParamByName('Contrato').AsString:=global_contrato_barco;
  zq_Esp.ParamByName('TipoPer').AsString:=zq_categoria.FieldByName('sIdTipoPersonal').AsString;
  zq_Esp.ParamByName('Per').AsString:='-1';
  zq_Esp.Open;
end;

procedure TfrmListado_Personal.edtApellidoMEnter(Sender: TObject);
begin
  edtApellidoM.color := global_color_entrada
end;

procedure TfrmListado_Personal.edtApellidoMExit(Sender: TObject);
begin
  edtApellidoM.color := global_color_salida
end;

procedure TfrmListado_Personal.edtApellidoMKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    edtRfc.SetFocus
end;

procedure TfrmListado_Personal.edtApellidoPEnter(Sender: TObject);
begin
  edtApellidoP.color := global_color_entrada
end;

procedure TfrmListado_Personal.edtApellidoPExit(Sender: TObject);
begin
  edtApellidoP.color := global_color_salida
end;

procedure TfrmListado_Personal.edtApellidoPKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    edtApellidoM.SetFocus
end;

procedure TfrmListado_Personal.edtBuscarKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
  begin
    btnBuscar.SetFocus;
    btnBuscar.Click;
  end;
end;

procedure TfrmListado_Personal.edtFichaEnter(Sender: TObject);
begin
  edtFicha.color := global_color_entrada
end;

procedure TfrmListado_Personal.edtFichaExit(Sender: TObject);
begin
  edtFicha.color := global_color_salida
end;

procedure TfrmListado_Personal.edtFichaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    edtNombre.SetFocus
end;

procedure TfrmListado_Personal.edtLibretaEnter(Sender: TObject);
begin
  edtLibreta.color := global_color_entrada
end;

procedure TfrmListado_Personal.edtLibretaExit(Sender: TObject);
begin
  edtLibreta.color := global_color_salida
end;

procedure TfrmListado_Personal.edtLibretaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    dedtVigencia.SetFocus
end;

procedure TfrmListado_Personal.edtNombreEnter(Sender: TObject);
begin
  edtNombre.color := global_color_entrada
end;

procedure TfrmListado_Personal.edtNombreExit(Sender: TObject);
begin
  edtNombre.color := global_color_salida
end;

procedure TfrmListado_Personal.edtNombreKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    edtApellidoP.SetFocus
end;

procedure TfrmListado_Personal.edtRfcEnter(Sender: TObject);
begin
  edtRfc.color := global_color_entrada
end;

procedure TfrmListado_Personal.edtRfcExit(Sender: TObject);
begin
  edtRfc.color := global_color_salida
end;

procedure TfrmListado_Personal.edtRfcKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    lkcbCategoria.SetFocus
end;

procedure TfrmListado_Personal.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  action := cafree;
end;

procedure TfrmListado_Personal.FormShow(Sender: TObject);
begin
  try
    IsOpen:=False;
    zq_listadoper.Active:=False;
    zq_listadoper.Open;

    zq_compania.Active:=False;
    zq_compania.Open;

    zq_categoria.Active:=False;
    zq_categoria.Open;
  finally
    IsOpen:=True;
    zq_listadoper.First;
  end;   
end;

procedure TfrmListado_Personal.frmBarra1btnAddClick(Sender: TObject);
begin
  zq_listadoper.Append;
  zq_listadoper.FieldByName('sContrato').AsString:=global_contrato_barco;
  zq_listadoper.FieldByName('sLibretadeMar').AsString:='';

  cambio_estado;
  frmBarra1.btnAddClick(Sender);
  edtFicha.SetFocus;
end;

procedure TfrmListado_Personal.frmBarra1btnCancelClick(Sender: TObject);
begin
  zq_listadoper.Cancel;
  cambio_estado;
  frmBarra1.btnCancelClick(Sender);

end;

procedure TfrmListado_Personal.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If zq_listadoper.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Seleccionado?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        zq_listadoper.Delete ;
      except
        on e : exception do
        begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Categorias de Personal', 'Al eliminar registro', 0);
        end;
      end
    end
end;

procedure TfrmListado_Personal.frmBarra1btnEditClick(Sender: TObject);
begin
  if zq_listadoper.RecordCount>0 then
  begin
    zq_listadoper.Edit;
    cambio_estado;
    frmBarra1.btnEditClick(Sender);
  end;

end;

procedure TfrmListado_Personal.frmBarra1btnExitClick(Sender: TObject);
begin
  frmBarra1.btnExitClick(Sender);
  close;
end;

procedure TfrmListado_Personal.frmBarra1btnPostClick(Sender: TObject);
const
  NombreFields: Array[1..6] of string = ('sidTripulacion', 'sNombre', 'sApellidoP', 'sApellidoM', 'sIdPersonal', 'sRfc');
var
  i: Integer;
begin
  try
    if Length(Trim(edtFicha.Text)) = 0  then
    begin
      MessageDlg('El campo Ficha debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtFicha.CanFocus then
        edtFicha.SetFocus;
      Exit;
    end;
    if Length(Trim(edtNombre.Text)) = 0  then
    begin
      MessageDlg('El campo Nombre debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtNombre.CanFocus then
        edtNombre.SetFocus;
      Exit;
    end;
    if Length(Trim(edtApellidoP.Text)) = 0  then
    begin
      MessageDlg('El campo Apellido Paterno debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtApellidoP.CanFocus then
        edtApellidoP.SetFocus;
      Exit;
    end;
    if Length(Trim(edtApellidoM.Text)) = 0  then
    begin
      MessageDlg('El campo Apellido Materno debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtApellidoM.CanFocus then
        edtApellidoM.SetFocus;
      Exit;
    end;
    if Length(Trim(edtRfc.Text)) = 0  then
    begin
      MessageDlg('El campo Rfc debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if edtRfc.CanFocus then
        edtRfc.SetFocus;
      Exit;
    end;
    if Length(Trim(lkcbEsp.Text)) = 0  then
    begin
      MessageDlg('El campo Categoria debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if lkcbEsp.CanFocus then
        lkcbEsp.SetFocus;
      Exit;
    end;
    if Length(Trim(lkcbCompania.Text)) = 0  then
    begin
      MessageDlg('El campo Compañia debe ser llenado correctamente antes de proceder a grabar el registro.', mtInformation, [mbOK], 0);
      if lkcbCompania.CanFocus then
        lkcbCompania.SetFocus;
      Exit;
    end;

    if zq_listadoper.state=dsEdit then
    begin
      if (MessageDlg('Desea Actualizar los Datos en el Listado de Tripulació Diaria', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
      begin
        for i := 1 to 6 do
        begin
          if zq_listadoper.FieldByName(NombreFields[i]).OldValue <> zq_listadoper.FieldByName(NombreFields[i]).Value then
          begin
            connection.zCommand.Active := False;
            Connection.zCommand.SQL.Clear;
            if (i>1) and (i<5) then
              Connection.zCommand.SQL.Add('UPDATE tripulaciondiaria_listado set sNombre='+zq_listadoper.FieldByName('sNombre').AsString+' '+zq_listadoper.FieldByName('sApellidoP').AsString+' '+zq_listadoper.FieldByName('sApellidoM').AsString+' where sIdTripulacion='+zq_listadoper.FieldByName('sIdTripulacion').OldValue)
            else
              Connection.zCommand.SQL.Add('UPDATE tripulaciondiaria_listado set '+NombreFields[i]+'='+zq_listadoper.FieldByName(NombreFields[i]).AsString+' where sIdTripulacion='+zq_listadoper.FieldByName('sIdTripulacion').OldValue);
            Connection.zCommand.ExecSQL;
          end;  
        end;
      end;
    end;

    zq_listadoper.Post;

    cambio_estado;
    //frmBarra1.btnPostClick(Sender);
  except
    on e:Exception do
    begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Categorias de Personal', 'Al guardar registro', 0);

      frmbarra1.btnCancel.Click ;
    end; 
  end;

end;

procedure TfrmListado_Personal.frmBarra1btnRefreshClick(Sender: TObject);
begin
  try
    zq_listadoper.Refresh;
  except
    on e : exception do
    begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Categorias de Personal', 'Al actualizar los Datos',0);    end;
  end;
end;

procedure TfrmListado_Personal.lkcbCategoriaEnter(Sender: TObject);
begin
  lkcbCategoria.color := global_color_entrada
end;

procedure TfrmListado_Personal.lkcbCategoriaExit(Sender: TObject);
begin
  lkcbCategoria.color := global_color_salida
end;

procedure TfrmListado_Personal.lkcbCategoriaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    lkcbEsp.SetFocus
end;

procedure TfrmListado_Personal.lkcbCompaniaEnter(Sender: TObject);
begin
  lkcbCompania.color := global_color_entrada
end;

procedure TfrmListado_Personal.lkcbCompaniaExit(Sender: TObject);
begin
  lkcbCompania.color := global_color_salida
end;

procedure TfrmListado_Personal.lkcbCompaniaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    edtLibreta.SetFocus
end;

procedure TfrmListado_Personal.lkcbEspEnter(Sender: TObject);
begin
  lkcbEsp.color := global_color_entrada
end;

procedure TfrmListado_Personal.lkcbEspExit(Sender: TObject);
begin
  lkcbEsp.color := global_color_salida
end;

procedure TfrmListado_Personal.lkcbEspKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    lkcbCompania.SetFocus
end;

procedure TfrmListado_Personal.zq_listadoperAfterScroll(DataSet: TDataSet);
begin
  if IsOpen then
  begin
    zq_Esp.Active:=False;
    zq_Esp.ParamByName('Contrato').AsString:=global_contrato_barco;
    zq_Esp.ParamByName('TipoPer').AsString:='-1';
    zq_Esp.ParamByName('Per').AsString:=zq_listadoper.FieldByName('sIdPersonal').AsString;
    zq_Esp.Open;

    lkcbCategoria.keyvalue:=zq_Esp.FieldByName('sIdTipoPersonal').AsString;
  end;

end;

end.
