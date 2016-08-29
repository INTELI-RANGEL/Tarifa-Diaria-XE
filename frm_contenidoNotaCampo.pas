unit frm_contenidoNotaCampo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxStyles, dxSkinsCore, dxSkinOffice2010Silver, dxSkinSharp,
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit,
  cxNavigator, DB, cxDBData, ExtCtrls, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid, frm_connection,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, frm_barra, cxContainer,
  cxLabel, cxTextEdit, cxMemo, cxDBEdit,StdCtrls, ComCtrls, jpeg, AdvGlowButton, ExtDlgs, axCtrls,
  Grids, DBGrids, cxCheckBox, DBCtrls, AdvEdit, DBAdvEd, AdvCombo, AdvDBComboBox,
  dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee,
  dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2013DarkGray, dxSkinOffice2013LightGray,
  dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic,
  dxSkinSharpPlus, dxSkinSilver, dxSkinSpringTime, dxSkinStardust,
  dxSkinSummer2008, dxSkinTheAsphaltWorld, dxSkinsDefaultPainters,
  dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint, dxSkinXmas2008Blue, Mask;

type
  TfrmContenidoNotacampo = class(TForm)
    Panel1: TPanel;
    dsContenido: TDataSource;
    zqContenido: TZQuery;
    frmBarra1: TfrmBarra;
    LbContenido: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    TabSheet5: TTabSheet;
    TabSheet6: TTabSheet;
    TabSheet7: TTabSheet;
    TabSheet8: TTabSheet;
    TabSheet9: TTabSheet;
    TabSheet10: TTabSheet;
    TabSheet11: TTabSheet;
    PnlSuperior: TPanel;
    Splitter1: TSplitter;
    OpenPicture: TOpenPictureDialog;
    DBGrid1: TDBGrid;
    PnlPrincipal: TPanel;
    CmbTipo: TAdvDBComboBox;
    DBAdvEdit1: TDBAdvEdit;
    DBAdvEdit2: TDBAdvEdit;
    cxLabel1: TcxLabel;
    CbxIndice: TcxCheckBox;
    PnlPortada: TPanel;
    PnlIndice: TPanel;
    ImgPortada: TImage;
    btnAdd: TAdvGlowButton;
    PnlPresentacion: TPanel;
    PnlOficio1: TPanel;
    PnlOficio2: TPanel;
    PnlOficio3: TPanel;
    PnlOficio4: TPanel;
    PnlOficio5: TPanel;
    PnlOficio6: TPanel;
    PnlOficio7: TPanel;
    DBMemo2: TDBMemo;
    cxLabel2: TcxLabel;
    GroupBox1: TGroupBox;
    DBAdvEdit3: TDBAdvEdit;
    DBAdvEdit4: TDBAdvEdit;
    GroupBox2: TGroupBox;
    DBAdvEdit5: TDBAdvEdit;
    DBAdvEdit6: TDBAdvEdit;
    DBAdvEdit7: TDBAdvEdit;
    DBMemo3: TDBMemo;
    cxLabel3: TcxLabel;
    DBAdvEdit8: TDBAdvEdit;
    GroupBox3: TGroupBox;
    DBAdvEdit9: TDBAdvEdit;
    DBAdvEdit10: TDBAdvEdit;
    GroupBox4: TGroupBox;
    DBAdvEdit11: TDBAdvEdit;
    DBAdvEdit12: TDBAdvEdit;
    GroupBox5: TGroupBox;
    DBAdvEdit13: TDBAdvEdit;
    DBAdvEdit14: TDBAdvEdit;
    DBAdvEdit15: TDBAdvEdit;
    DBMemo4: TDBMemo;
    cxLabel4: TcxLabel;
    GroupBox6: TGroupBox;
    DBAdvEdit16: TDBAdvEdit;
    DBAdvEdit17: TDBAdvEdit;
    GroupBox7: TGroupBox;
    DBAdvEdit18: TDBAdvEdit;
    DBAdvEdit19: TDBAdvEdit;
    DBAdvEdit20: TDBAdvEdit;
    DBMemo5: TDBMemo;
    cxLabel5: TcxLabel;
    DBMemo1: TDBMemo;
    procedure FormShow(Sender: TObject);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure zqContenidoAfterScroll(DataSet: TDataSet);
    procedure edtDescripcionEnter(Sender: TObject);
    procedure CmbTipoChange(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure zqContenidoAfterInsert(DataSet: TDataSet);
    procedure zqContenidoAfterOpen(DataSet: TDataSet);
    procedure zqContenidoAfterCancel(DataSet: TDataSet);
    procedure zqContenidoAfterClose(DataSet: TDataSet);
    procedure zqContenidoAfterEdit(DataSet: TDataSet);
    procedure zqContenidoAfterPost(DataSet: TDataSet);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    procedure ControlBarra(Zqr: Tzquery);
    procedure HabilitarApartado(indice: integer);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmContenidoNotacampo: TfrmContenidoNotacampo;
  max : integer;
  id : string;
  estado : string;


implementation

{$R *.dfm}



procedure TfrmContenidoNotacampo.btnAddClick(Sender: TObject);

{$REGION 'generar tbitmap'}
Function LoadGraphicsFile(Const Filename: String):  TBitmap;
Var
  Picture: TPicture;
  f : TFileStream;
  graphic : TOleGraphic;
Begin
  Result := NIL;
  If FileExists(Filename) Then
  Begin
    Result := TBitmap.Create;
    Try
      Picture := TPicture.Create;
      graphic := TOleGraphic.Create;
      Try
        f := TFileStream.Create (Filename,fmOpenRead or fmShareDenyNone);
        try
          try
            Graphic.LoadFromStream(f);
            Picture.Assign(Graphic);
          except
            on e:exception do
            begin
              if e.ClassName='EOleSysError' then
              begin
                try
                  Picture.LoadFromFile(Filename);
                except
                  Result:=nil;
                end;
              end;
            end;
          end;
          Try
            Result.Assign(Picture.Graphic);
          Except
            Result.Width  := Picture.Graphic.Width;
            Result.Height := Picture.Graphic.Height;
            Result.PixelFormat := pf24bit;
            Result.Canvas.Draw(0, 0, Picture.Graphic);
          End;
        finally
          freeandnil(f);
        end;
      Finally
        Picture.Free;
        freeandnil(graphic);
      End;
    Except
      Result.Free;
      Raise;
    End;
  End;
End;
{$ENDREGION}

begin
  OpenPicture.Title := 'Cargar imagen (*.jpg)';
  if OpenPicture.Execute then
  begin
    try
      ImgPortada.Picture.LoadFromFile(OpenPicture.FileName);
    except
      ShowMessage('El tipo de imagen seleccionada no esta permitida.');
      ImgPortada.Picture.LoadFromFile('');
    end
  end
end;

procedure TfrmContenidoNotacampo.CmbTipoChange(Sender: TObject);
begin
  HabilitarApartado(CmbTipo.Items.IndexOf(CmbTipo.Text)+1);
end;

procedure TfrmContenidoNotacampo.edtDescripcionEnter(Sender: TObject);
begin
 // edtDescripcion.SelStart := Length(edtDescripcion.Text);
end;

procedure TfrmContenidoNotacampo.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  action := cafree ;
end;

procedure TfrmContenidoNotacampo.FormShow(Sender: TObject);
begin
  zqContenido.Active := False;
  zqContenido.Open;
  LbContenido.TabIndex := 0;
end;

procedure TfrmContenidoNotacampo.frmBarra1btnAddClick(Sender: TObject);
begin
  zqContenido.Append;
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select ifnull(max(iorden)+1,1) as m from contenidonotacampo ');
  connection.QryBusca.Open;
  if connection.QryBusca.RecordCount > 0 then begin
    connection.QryBusca.Last;
    max := connection.QryBusca.FieldByName('m').AsInteger;
  end
  else begin
    max := 1;
  end;
  zqContenido.FieldByName('iorden').AsInteger := max;



end;

procedure TfrmContenidoNotacampo.frmBarra1btnCancelClick(Sender: TObject);
begin
  zqContenido.Cancel;
end;

procedure TfrmContenidoNotacampo.frmBarra1btnDeleteClick(Sender: TObject);
begin
  zqContenido.Delete;
end;

procedure TfrmContenidoNotacampo.frmBarra1btnEditClick(Sender: TObject);
begin
  zqContenido.Edit;
end;

procedure TfrmContenidoNotacampo.frmBarra1btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmContenidoNotacampo.frmBarra1btnPostClick(Sender: TObject);
var
  query : TZQuery;
  des, nom : string;
  continuar : boolean;

  StImagen:TStream;
   {
  procedure GuardarImagen(qr:tzquery;Cargar:Boolean);
  var
    bS: TStream;
    Pic: TJpegImage;
    BlobField: tField;
  begin
    try
      BlobField := qr.FieldByName('bImagen1');
      BS := qr.CreateBlobStream(BlobField, bmWrite);
      try
        Pic := TJpegImage.Create;
        try
          if cargar then
          begin
         //   Pic.LoadFromFile(OpenPictureAux2.FileName);
            Pic.SaveToStream(Bs);
          end
          else
            ImgPortada.Picture.Graphic.SaveToStream(Bs);
        finally
          Pic.Free;
        end;
      finally
        bS.Free
      end
    except
      ;
    end
  end;
      }

begin
  try
    if length(trim(zqContenido.FieldByName('snombreportada').asstring)) = 0 then
      raise Exception.Create('El nombre de la portada no puede ir vacío.');

    if length(trim(DBMemo1.Text)) = 0 then
      raise Exception.Create('La descripción no puede ir vacía.');

    if zqContenido.State = dsEdit then
    begin
      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select * from contenidonotacampo where snombreportada = :nombrep and sid <> :id order by iOrden');
      connection.QryBusca.ParamByName('nombrep').AsString := zqContenido.FieldByName('snombreportada').asstring;
      connection.QryBusca.ParamByName('id').AsString := zqContenido.FieldByName('sid').asstring;
      connection.QryBusca.Open;
      if connection.QryBusca.RecordCount > 0 then
        raise Exception.Create('Ya existe una hoja con este nombre, cambielo e intente otra vez.');

    end;

    if zqContenido.State = dsInsert then
    begin
      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select * from contenidonotacampo where snombreportada = :nombrep  order by iOrden');
      connection.QryBusca.ParamByName('nombrep').AsString := zqContenido.FieldByName('snombreportada').asstring;
      connection.QryBusca.Open;
      if connection.QryBusca.RecordCount > 0 then
        raise Exception.Create('Ya existe una hoja con este nombre, cambielo e intente otra vez.');

      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select ifnull(max(sid)+1,1) as m from contenidonotacampo ');
      connection.QryBusca.Open;
      if connection.QryBusca.RecordCount > 0 then begin
        connection.QryBusca.Last;
        max := connection.QryBusca.FieldByName('m').AsInteger;
      end
      else begin
        max := 1;
      end;
      zqContenido.FieldByName('sid').AsInteger := max;
    end;

    try
      try
        StImagen:= zqContenido.CreateBlobStream(TBlobField(zqContenido.FieldByName('bimagen1')), bmWrite);
        ImgPortada.Picture.Graphic.SaveToStream(StImagen);
        TBlobField(zqContenido.FieldByName('bimagen1')).LoadFromStream(stimagen);
      finally
        StImagen.Free;
      end;
    except
      ;
    end;

    if CbxIndice.Checked then
      zqContenido.FieldByName('lincluirindice').AsString := 'Si'
    else
      zqContenido.FieldByName('lincluirindice').AsString := 'No';
    zqContenido.Post;

  finally

  end;
 { try
    CmbTipo.Enabled := False;
    des := edtDescripcion.Text;
    nom := edtNombre.Text;

    if Length(Trim(nom)) = 0 then
      raise Exception.Create('El nombre de la hoja no puede ir vacío');

    if Length(Trim(des)) = 0 then
      raise Exception.Create('La descripción de la hoja no puede ir vacío');

    query := TZQuery.Create(nil);
    query.Connection := connection.zConnection;

    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;

    if estado = 'insert' then
    begin
      connection.QryBusca.SQL.Add('select * from contenidonotacampo where snombreportada = :nombrep order by iOrden');
      connection.QryBusca.ParamByName('nombrep').AsString := edtNombre.Text;
      connection.QryBusca.Open;
      if connection.QryBusca.RecordCount > 0 then
        raise Exception.Create('Ya existe una hoja con ese nombre.');

      query.Active := False;
      query.SQL.Clear;
      query.SQL.Add('insert into contenidonotacampo (sId, iOrden, sDescripcion,sNombrePortada,lTipo,lincluirindice) '+
                          'values (:id, :orden, :descripcion, :nombre,:tipo,indice)');
      query.Params.ParamByName('id').DataType := ftString;
      query.Params.ParamByName('orden').DataType := ftInteger;
      query.Params.ParamByName('descripcion').DataType := ftString;
      query.Params.ParamByName('nombre').DataType := ftString;
      query.Params.ParamByName('tipo').DataType := ftString;
      
      query.Params.ParamByName('id').Value := edtID.Text;
      query.Params.ParamByName('orden').Value := StrToInt(edtOrden.Text);
      query.Params.ParamByName('descripcion').Value := edtDescripcion.Text;
      query.Params.ParamByName('nombre').Value := edtNombre.Text;
      query.Params.ParamByName('tipo').Value := CmbTipo.text;

      if CbxIndice.Checked then
        query.ParamByName('indice').AsString := 'Si'
      else
        query.ParamByName('indice').AsString := 'No';
        
      query.ExecSQL;
    end
    else
    if estado = 'edit' then
    begin
      connection.QryBusca.SQL.Add('select * from contenidonotacampo where snombreportada = :nombrep and sid <> :id order by iOrden');
      connection.QryBusca.ParamByName('nombrep').AsString := edtNombre.Text;
      connection.QryBusca.ParamByName('id').AsString := edtID.Text;
      connection.QryBusca.Open;
      if connection.QryBusca.RecordCount > 0 then
        raise Exception.Create('Ya existe una hoja con ese nombre.');

      query.Active := False;
      query.SQL.Clear;
      query.SQL.Add('update contenidonotacampo set sId = :id, iOrden = :orden, sDescripcion = :descripcion, sNombrePortada = :nombre, lincluirindice = :indice, bimagen1 = :bimagen  '+
                    'where sId = :id');
      query.Params.ParamByName('id').DataType := ftString;
      query.Params.ParamByName('orden').DataType := ftInteger;
      query.Params.ParamByName('descripcion').DataType := ftString;
      query.Params.ParamByName('nombre').DataType := ftString;
      query.Params.ParamByName('nombre').DataType := ftBlob;

      query.Params.ParamByName('id').Value := edtID.Text;
      query.Params.ParamByName('orden').Value := StrToInt(edtOrden.Text);
      query.Params.ParamByName('descripcion').Value := edtDescripcion.Text;
      query.Params.ParamByName('nombre').Value := edtNombre.Text;

      StImagen := TStream.Create;
      ImgPortada.Picture.Graphic.SaveToStream(StImagen);
      TBlobfield(query.Params.ParamByName('imagen').AsBlob).loadfromstream(StImagen);

      if CbxIndice.Checked then
        query.ParamByName('indice').AsString := 'Si'
      else
        query.ParamByName('indice').AsString := 'No';      
      query.ExecSQL;
    end;

    desbloquearControles;
    frmBarra1.btnPrinter.Enabled := False;
    zqContenido.Refresh;
    zqContenido.Last;
    CbxIndice.Enabled := False;
  finally
    if Assigned(query) then
      query.Destroy;
  end; }
end;

procedure TfrmContenidoNotacampo.frmBarra1btnRefreshClick(Sender: TObject);
begin
  zqContenido.Refresh;
end;

procedure TfrmContenidoNotacampo.zqContenidoAfterCancel(DataSet: TDataSet);
begin
  ControlBarra(zqcontenido);
end;

procedure TfrmContenidoNotacampo.zqContenidoAfterClose(DataSet: TDataSet);
begin
  ControlBarra(zqcontenido);
end;

procedure TfrmContenidoNotacampo.zqContenidoAfterEdit(DataSet: TDataSet);
begin
  ControlBarra(zqcontenido);
end;

procedure TfrmContenidoNotacampo.zqContenidoAfterInsert(DataSet: TDataSet);
begin
  ControlBarra(zqcontenido);
end;

procedure TfrmContenidoNotacampo.zqContenidoAfterOpen(DataSet: TDataSet);
begin
  ControlBarra(zqcontenido);
end;

procedure TfrmContenidoNotacampo.zqContenidoAfterPost(DataSet: TDataSet);
begin
  ControlBarra(zqcontenido);
end;

procedure TfrmContenidoNotacampo.zqContenidoAfterScroll(DataSet: TDataSet);
var Bs:tstream;
 pic:TJpegImage;
 BlobField: tField;
begin
  CbxIndice.Checked := zqContenido.FieldByName('lincluirindice').AsString = 'Si';

  try
    HabilitarApartado(CmbTipo.Items.IndexOf(CmbTipo.Text)+1);
  except
    ;
  end;

  try
    try
      BlobField := zqContenido.FieldByName('bImagen1');
      BS := zqContenido.CreateBlobStream(BlobField, bmRead);
      if bs.Size > 1 then
      begin
        Pic := TJpegImage.Create;
        try
          Pic.LoadFromStream(bS);
          ImgPortada.Picture.Graphic := Pic;
        finally
          Pic.Free;
        end;
      end
      else
        ImgPortada.Picture.LoadFromFile('');

    except
      ImgPortada.Picture.LoadFromFile('');
    end;
  finally
    bS.Free
  end;
end;

{$Region 'Procedimientos varios'}
procedure TfrmContenidoNotacampo.HabilitarApartado(indice:integer);
var y:Integer;
begin
  for y := 1 to LbContenido.PageCount-1 do
  begin
    LbContenido.Pages[y].TabVisible := False;
  end;
  LbContenido.Pages[indice].TabVisible := True;
end;

procedure TfrmContenidoNotacampo.ControlBarra(Zqr:Tzquery);
var estado:TDataSetState;
begin
  estado := zqr.state;
  with frmBarra1 do
  begin
    //Botones
    btnadd.Enabled := estado in [dsBrowse];
    btnEdit.Enabled := estado in [dsBrowse];
    btnPost.Enabled := estado in [dsEdit,dsInsert];
    btncancel.Enabled := estado in [dsEdit,dsInsert];
    btnDelete.Enabled := estado in [dsBrowse];
    btnRefresh.Enabled := estado in [dsBrowse];
    btnPrinter.Enabled := False;
    btnExit.Enabled := estado in [dsBrowse];

    //Paneles
    PnlPrincipal.Enabled := estado in [dsEdit,dsInsert];
    PnlPortada.Enabled := estado in [dsEdit,dsInsert];
    PnlIndice.Enabled := estado in [dsEdit,dsInsert];
    PnlPresentacion.Enabled := estado in [dsEdit,dsInsert];
    PnlOficio1.Enabled := estado in [dsEdit,dsInsert];
    PnlOficio2.Enabled := estado in [dsEdit,dsInsert];
    PnlOficio3.Enabled := estado in [dsEdit,dsInsert];
    PnlOficio4.Enabled := estado in [dsEdit,dsInsert];
    PnlOficio5.Enabled := estado in [dsEdit,dsInsert];
    PnlOficio6.Enabled := estado in [dsEdit,dsInsert];
    PnlOficio7.Enabled := estado in [dsEdit,dsInsert];
  end;
end;

{$ENDREGION}

end.
