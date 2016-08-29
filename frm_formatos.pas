unit frm_formatos;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvSmoothPanel, AdvOfficeButtons, StdCtrls, Buttons, frm_connection,
  global, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, NxScrollControl,
  NxCustomGridControl, NxCustomGrid, NxDBGrid, AdvAppStyler, frm_barra,
  NxDBColumns, NxColumns, Grids, DBGrids, Mask, DBCtrls, StrUtils, cxGraphics,
  cxLookAndFeels, cxLookAndFeelPainters, Menus, cxButtons, ShellApi,
  dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel,
  dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven,
  dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus, dxSkinSilver,
  dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008, dxSkinTheAsphaltWorld,
  dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue, RxToolEdit;

type
  TfrmFormatos = class(TForm)
    qrFormatos: TZQuery;
    pnlCarga: TAdvSmoothPanel;
    frmBarra1: TfrmBarra;
    dsFormatos: TDataSource;
    DBGrid1: TDBGrid;
    Label1: TLabel;
    dbModulo: TDBEdit;
    Label2: TLabel;
    Label3: TLabel;
    txtArchivo: TFilenameEdit;
    btnSubir: TcxButton;
    cxButton5: TcxButton;
    chkAbrirFinalizar: TAdvOfficeCheckBox;
    dbDescripcion: TDBEdit;
    dbTitulo: TDBEdit;
    procedure FormCreate(Sender: TObject);
    procedure btnSubirClick(Sender: TObject);
    procedure cxButton5Click(Sender: TObject);
    procedure txtArchivoAfterDialog(Sender: TObject; var Name: string;
      var Action: Boolean);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure dbTituloKeyPress(Sender: TObject; var Key: Char);
    procedure dbModuloKeyPress(Sender: TObject; var Key: Char);
    procedure dbDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure dbTituloEnter(Sender: TObject);
    procedure dbTituloExit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure SubirFormato(sArchivo, formato : string);
    procedure GuardarFormato(reporte,extencion : string);

    function DescargaReporte(Dir, Reporte: String): Boolean;
  end;

var
  frmFormatos: TfrmFormatos;

implementation

{$R *.dfm}

procedure TfrmFormatos.btnSubirClick(Sender: TObject);
begin
  try
    if qrFormatos.State in [dsEdit, dsInsert] then
    raise exception.Create('Amtes tiene que guardar los cambios al registro, guarde el registro e intente nuevamente.');

    SubirFormato(txtArchivo.Text, qrFormatos.FieldByName('stitulo').AsString);
  except
    on e:exception do
      MessageDlg(e.Message, mtInformation, [mbOk], 0);
  end;
end;

procedure TfrmFormatos.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmFormatos.FormCreate(Sender: TObject);
begin
  qrFormatos.Active := False;
  qrFormatos.ParamByName('contrato').AsString := global_contrato;
  QrFormatos.Open;
end;

procedure TfrmFormatos.frmBarra1btnAddClick(Sender: TObject);
begin
  qrFormatos.Append;
  qrFormatos.FieldByName('sContrato').AsString := global_contrato;
  frmBarra1.btnAddClick(Sender);
  pnlCarga.enabled := frmBarra1.btnAdd.Enabled;
  dbtitulo.setfocus;
end;

procedure TfrmFormatos.frmBarra1btnCancelClick(Sender: TObject);
begin
  qrFormatos.Cancel;
  frmBarra1.btnCancelClick(Sender);
  pnlCarga.Enabled := frmBarra1.btnAdd.Enabled;
end;

procedure TfrmFormatos.frmBarra1btnDeleteClick(Sender: TObject);
begin
  if MessageDlg('¿Desea eliminar el registro activo?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    qrFormatos.Delete;
end;

procedure TfrmFormatos.frmBarra1btnEditClick(Sender: TObject);
begin
  qrFormatos.Edit;
  frmBarra1.btnEditClick(Sender);
  pnlCarga.Enabled := frmBarra1.btnAdd.Enabled;
end;

procedure TfrmFormatos.frmBarra1btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmFormatos.frmBarra1btnPostClick(Sender: TObject);
begin
  try
    connection.qrybusca.Active := false;
    connection.qrybusca.sql.text := 'select stitulo from tarifa_formatos where sContrato = :contrato and stitulo = :reporte';
    connection.qrybusca.ParamByName('contrato').asstring := global_contrato;
    connection.qrybusca.parambyname('reporte').asstring := Trim(qrFormatos.FieldByName('stitulo').asstring);
    connection.QryBusca.open;

    if connection.QryBusca.RecordCount > 0 then
      raise exception.Create('El formato ya esta dado de alta en la base de datos');

    qrFormatos.Post;
    frmBarra1.btnPostClick(Sender);
    pnlCarga.enabled := frmBarra1.btnAdd.Enabled;
  except
    on e:exception do
      MessageDlg(e.message, mtInformation, [mbOk], 0);
  end;
end;

procedure TfrmFormatos.frmBarra1btnRefreshClick(Sender: TObject);
begin
  qrFormatos.Refresh;
end;

procedure TfrmFormatos.SubirFormato(sArchivo: string; formato: string);
var
  BlobStream : TStream;
  FileStream : TFileStream;
begin
  if not FileExists(sArchivo) then
    raise exception.Create('La ruta especificada no es valida');

  connection.zCommand.Active := False;
  connection.zCommand.SQL.Text := 'select bArchivo from tarifa_formatos where sContrato = :contrato and sTitulo = :reporte';
  connection.zcommand.parambyname('contrato').asstring := global_contrato;
  connection.zcommand.ParamByName('reporte').AsString := formato;
  connection.zCommand.Open;
  if connection.zCommand.RecordCount > 0 then
  begin
    try
      connection.zCommand.Edit;
      BlobStream := connection.zCommand.CreateBlobStream(connection.zCommand.FieldByName('bArchivo')
      ,bmWrite);
      try
        FileStream := TFileStream.Create(sArchivo, fmOpenRead);
        try
          BlobStream.CopyFrom(FileStream, FileStream.Size);
        finally
          FileStream.Destroy;
        end;
      finally
        BlobStream.Destroy;
      end;
    finally
      connection.zCommand.Post;
      MessageDlg('Formato cargado correctamente', mtInformation,[mbOk],0);
    end;
  end;
end;


procedure TfrmFormatos.txtArchivoAfterDialog(Sender: TObject; var Name: string;
  var Action: Boolean);
begin
  if Trim(name) <> '' then
    btnSubir.Enabled := true
  else
    btnSubir.Enabled := False;

  if Trim(name) <> '' then
  begin
    qrFormatos.Edit;
    qrFormatos.FieldByName('sTipo').AsString := ExtractFileExt(Name);
    qrFormatos.Post;
  end;
end;

procedure TFrmFormatos.GuardarFormato(reporte,extencion : string);
var
  sArchivo : string;

function GuardarComo:Boolean;
var
  Correcto,Cancelar:Boolean;
  Guardar : TSaveDialog;
begin
  Guardar := TSaveDialog.Create(nil);
  Guardar.Filter := 'Excel(*'+extencion+')|*'+extencion;
  Guardar.DefaultExt := extencion;
  Guardar.FileName := Reporte;


  Correcto := False;
  Cancelar := False;
  repeat
    if Guardar.Execute then
    begin
      SArchivo := Guardar.FileName;
      if (AnsiEndsText(extencion,sArchivo)) then
      begin
        if FileExists(SArchivo) then
        begin
          if MessageDlg('¿El archivo ya existe, desea sobreescribirlo?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            Correcto := True;
          end
          else
            ShowMessage('Intente con otro nombre porfavor.');
        end
        else
          Correcto := True;
      end;

    end
    else
      Cancelar := True;
  until (Correcto) or Cancelar;
  Result := Correcto;
end;


begin
  if not GuardarComo then
    raise Exception.Create('Proceso cancelado por el usuario');

  if not DescargaReporte(sArchivo, reporte) then
    raise Exception.Create('No se encontró el reporte en la base de datos');

  if chkAbrirFinalizar.Checked then
    ShellExecute(Handle,'open',pwidechar(SArchivo), nil, nil,  SW_SHOWNORMAL);
end;

procedure TfrmFormatos.cxButton5Click(Sender: TObject);
begin
  try
    if qrFormatos.RecordCount = 0 then
      raise exception.create('No hay información que exportar');

    GuardarFormato(qrFormatos.FieldbyName('sTitulo').AsString, qrFormatos.FieldByName('sTipo').AsString);
  except
    on e:exception do
      MessageDlg(e.Message, mtInformation, [mbOk], 0);
  end;
end;

procedure TfrmFormatos.dbDescripcionKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    dbtitulo.setfocus;
end;

procedure TfrmFormatos.dbModuloKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    dbDescripcion.Setfocus;
end;

procedure TfrmFormatos.dbTituloEnter(Sender: TObject);
begin
  (Sender as Tdbedit).Color := global_color_entrada;
end;

procedure TfrmFormatos.dbTituloExit(Sender: TObject);
begin
    (Sender as tdbedit).Color := global_color_salida;
end;

procedure TfrmFormatos.dbTituloKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    dbModulo.SetFocus;
end;

function TFrmFormatos.DescargaReporte(Dir,Reporte:String):Boolean;
var ZqReporte:TZReadOnlyQuery;
RDescarga:Boolean;
begin
  RDescarga := False;
  ZqReporte := TZReadOnlyQuery.Create(nil);
  try
    ZqReporte.Connection := connection.zConnection;
    ZqReporte.Active := False;
    ZqReporte.SQL.Text := 'Select bArchivo from tarifa_formatos where sContrato = :contrato and sTitulo = :TRep ';
    zqReporte.parambyname('contrato').Asstring := global_contrato;
    ZqReporte.ParamByName('TRep').AsString := Reporte;
    ZqReporte.Open;
    if ZqReporte.RecordCount = 1 then
    begin
      if FileExists(Dir) then
        DeleteFile(Dir);
      if FileExists(Dir) then
        raise Exception.Create('No se puede eliminar el archivo existente: '+Dir);

      TBlobField(ZqReporte.FieldByName('bArchivo')).SaveToFile(Dir);
      Sleep(200); //por seguridad
      RDescarga :=  FileExists(Dir);
    end
    else
      raise exception.create('El formato de reporte "'+Reporte+'" no se encuentra en la base de datos.');

  finally
    ZqReporte.free;
    Result := RDescarga;
  end;
end;


end.
