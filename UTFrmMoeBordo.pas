unit UTFrmMoeBordo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ZDataset, DB, ZAbstractRODataset, ZAbstractDataset, ImgList,
  RXDBCtrl, StdCtrls, Newpanel, DBCtrls, ComCtrls, AdvDateTimePicker,
  AdvDBDateTimePicker, NxPageControl, Grids, DBGrids, JvExDBGrids, JvDBGrid,
  JvDBUltimGrid, ExtCtrls, frm_barra, cxGraphics, cxControls,
  cxLookAndFeels, cxLookAndFeelPainters, cxStyles, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom,
  dxSkinDarkSide, dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle, dxSkinFoggy,
  dxSkinGlassOceans, dxSkinHighContrast, dxSkiniMaginary, dxSkinLilian,
  dxSkinLiquidSky, dxSkinLondonLiquidSky, dxSkinMcSkin, dxSkinMetropolis,
  dxSkinMetropolisDark, dxSkinMoneyTwins, dxSkinOffice2007Black,
  dxSkinOffice2007Blue, dxSkinOffice2007Green, dxSkinOffice2007Pink,
  dxSkinOffice2007Silver, dxSkinOffice2010Black, dxSkinOffice2010Blue,
  dxSkinOffice2010Silver, dxSkinOffice2013DarkGray, dxSkinOffice2013LightGray,
  dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven, dxSkinSevenClassic,
  dxSkinSharp, dxSkinSharpPlus, dxSkinSilver, dxSkinSpringTime, dxSkinStardust,
  dxSkinSummer2008, dxSkinTheAsphaltWorld, dxSkinsDefaultPainters,
  dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint, dxSkinXmas2008Blue,
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit,
  cxNavigator, cxDBData, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGridLevel, cxClasses, cxGridCustomView, cxGrid,
  cxContainer, cxButtons, cxTextEdit;

type
  TFrmMoeBordo = class(TForm)
    pnl2: TPanel;
    pnl3: TPanel;
    pnl5: TPanel;
    pnl6: TPanel;
    pnl4: TPanel;
    JDbUG1: TJvDBUltimGrid;
    pnl1: TPanel;
    pgDatos: TNxPageControl;
    NxTshDatos: TNxTabSheet;
    lbl1: TLabel;
    AvDbDtpFecha: TAdvDBDateTimePicker;
    dbmmoComentario: TDBMemo;
    pgPersonal: TNxTabSheet;
    pgEquipo: TNxTabSheet;
    Panel: tNewGroupBox;
    ListaObjeto: TRxDBGrid;
    ImgBtns: TImageList;
    QMoe: TZQuery;
    dsMoe: TDataSource;
    QMoePersonal: TZQuery;
    QMoePersonaliIdMoe: TIntegerField;
    QMoePersonalsIdRecurso: TStringField;
    QMoePersonaleTipoRecurso: TStringField;
    QMoePersonalsDescripcion: TStringField;
    QMoePersonaldCantidad: TFloatField;
    QMoeEquipo: TZQuery;
    QMoeEquipoiIdMoe: TIntegerField;
    QMoeEquiposIdRecurso: TStringField;
    QMoeEquipoeTipoRecurso: TStringField;
    QMoeEquiposDescripcion: TStringField;
    QMoeEquipodCantidad: TFloatField;
    dsMoePersonal: TDataSource;
    dsMoeEquipo: TDataSource;
    BuscaObjeto: TZReadOnlyQuery;
    ds_buscaobjeto: TDataSource;
    mnuVigencia: TPopupMenu;
    mnuCarga: TMenuItem;
    Contratos: TZQuery;
    DsContratos: TDataSource;
    frmBarra1: TfrmBarra;
    fltfldQMoePersonaldCantR: TFloatField;
    CxGridMoePersonal: TcxGridDBTableView;
    CxLevel1: TcxGridLevel;
    CxGridMoeRecursos: TcxGrid;
    CxColumnCxGridMoePersonalColumn1: TcxGridDBColumn;
    CxColumnCxGridMoePersonalColumn2: TcxGridDBColumn;
    CxColumnCxGridMoePersonalColumn3: TcxGridDBColumn;
    CxColumnCxGridMoePersonalColumn5: TcxGridDBColumn;
    cxGrid1: TcxGrid;
    cxGridDBTableView1: TcxGridDBTableView;
    cxGridDBColumn1: TcxGridDBColumn;
    cxGridDBColumn2: TcxGridDBColumn;
    cxGridDBColumn3: TcxGridDBColumn;
    cxGridDBColumn4: TcxGridDBColumn;
    cxGridLevel1: TcxGridLevel;
    QMoeEquipodCantR: TFloatField;
    GboxCantidad: tNewGroupBox;
    CxTextEdtCantidad: TcxTextEdit;
    CxBtnok: TcxButton;
    cxStyleReposCellColor: TcxStyleRepository;
    cxstylRed: TcxStyle;
    cxstylGreen: TcxStyle;
    cxstyleYellow: TcxStyle;
    cxstylGray: TcxStyle;
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure QMoeEquipoAfterScroll(DataSet: TDataSet);
    procedure QMoeEquiposIdRecursoChange(Sender: TField);
    procedure CxBtnokClick(Sender: TObject);
    procedure CxGridMoePersonalCellDblClick(Sender: TcxCustomGridTableView;
      ACellViewInfo: TcxGridTableDataCellViewInfo; AButton: TMouseButton;
      AShift: TShiftState; var AHandled: Boolean);
    procedure QMoeAfterScroll(DataSet: TDataSet);
    procedure pgDatosChange(Sender: TObject);
    procedure CxTextEdtCantidadKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure CxColumnCxGridMoePersonalColumn3StylesGetContentStyle(
      Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
      AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
    procedure cxGridDBColumn3StylesGetContentStyle(
      Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
      AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
  private
    mReSult: Integer;
    parIdMoe: Integer;
    parsIdRecurso: string;
    gForma: TForm;
    procedure ScrollMoe;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmMoeBordo: TFrmMoeBordo;

implementation

{$R *.dfm}

uses
  global, frm_connection;

procedure TFrmMoeBordo.CxBtnokClick(Sender: TObject);
var
  zMoeaBordoUpt: TZQuery;
  TipoMoe: string;
begin
  try
    zMoeaBordoUpt := TZQuery.Create(Self);
    if pgDatos.ActivePage = pgEquipo then
    begin
      TipoMoe := 'Equipo';
      parIdMoe := QMoeEquipo.FieldByName('iIdMoe').AsInteger;
      parsIdRecurso := QMoeEquipo.FieldByName('sIdRecurso').AsString;
    end
    else if pgDatos.ActivePage = pgPersonal then
    begin
      TipoMoe := 'Personal';
      parIdMoe := QMoePersonal.FieldByName('iIdMoe').AsInteger;
      parsIdRecurso := QMoePersonal.FieldByName('sIdRecurso').AsString;
    end;

    try
      zMoeaBordoUpt.Active := False;
      zMoeaBordoUpt.Connection := connection.zConnection;
      zMoeaBordoUpt.SQL.Clear;
      zMoeaBordoUpt.SQL.Text := 'Select * from Moerecursos_aBordo where iidMoe = :IdMoe and sIdRecurso = :idRecurso and ETipoRecurso ="' + TipoMoe + '"';
      zMoeaBordoUpt.Params.ParamByName('idMoe').AsInteger := parIdMoe ;
      zMoeaBordoUpt.Params.ParamByName('IdRecurso').AsString := parsIdRecurso;
      zMoeaBordoUpt.Open;
      zMoeaBordoUpt.Edit;
      zMoeaBordoUpt.FieldByName('dCantidad').AsFloat := StrToFloatDef(CxTextEdtCantidad.Text, 0);
      zMoeaBordoUpt.Post;
      ScrollMoe;
    finally
      if Assigned(zMoeaBordoUpt) then
        zMoeaBordoUpt.Destroy;


      GboxCantidad.Visible := False;
      GboxCantidad.Parent := Self;
      GboxCantidad.Align := alNone;
    end;
  except
    on e: Exception do
      MessageDlg('Ha ocurrido un error inesperado, informar al administrador del sistema del siguiente error:' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TFrmMoeBordo.CxColumnCxGridMoePersonalColumn3StylesGetContentStyle(
  Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
  AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
begin
  if ARecord.Values[2] > ARecord.Values[3]  then
    AStyle := cxstylRed;

  if ARecord.Values[2] = ARecord.Values[3]  then
    AStyle := cxstylGreen;

  if ARecord.Values[2] < ARecord.Values[3]  then
    AStyle := cxstyleYellow;
end;

procedure TFrmMoeBordo.cxGridDBColumn3StylesGetContentStyle(
  Sender: TcxCustomGridTableView; ARecord: TcxCustomGridRecord;
  AItem: TcxCustomGridTableItem; var AStyle: TcxStyle);
begin
  if ARecord.Values[2] > ARecord.Values[3]  then
    AStyle := cxstylRed;

  if ARecord.Values[2] = ARecord.Values[3]  then
    AStyle := cxstylGreen;

  if ARecord.Values[2] < ARecord.Values[3]  then
    AStyle := cxstyleYellow;
end;

procedure TFrmMoeBordo.CxGridMoePersonalCellDblClick(
  Sender: TcxCustomGridTableView; ACellViewInfo: TcxGridTableDataCellViewInfo;
  AButton: TMouseButton; AShift: TShiftState; var AHandled: Boolean);
begin
  try
    mReSult := mrNone;
    if pgDatos.ActivePage = pgEquipo then
      CxTextEdtCantidad.text := QMoeEquipo.FieldByName('dCantidad').AsString
    else if pgDatos.ActivePage = pgPersonal then
      CxTextEdtCantidad.Text := QMoePersonal.FieldByName('dCantidad').AsString;

    if Assigned(FindComponent('FrmEdit')) then
      TForm(FindComponent('FrmEdit')).Destroy;
      
    gForma := TForm.Create(Self);
    gForma.Name := 'FrmEdit';
    gForma.Caption := 'Editar valor';
    gForma.Width := 195;
    gForma.Height := 100;
    gForma.BorderIcons := [];
    gForma.BorderStyle := bsDialog;
    gForma.Position := poScreenCenter;
    GboxCantidad.Parent := gForma;
    GboxCantidad.Visible := True;
    GboxCantidad.Align := alClient;
    gForma.Showmodal;
    if CxTextEdtCantidad.canFocus then
    begin
      CxTextEdtCantidad.SetFocus;
    end;
  finally
    gForma.Free;
  end;
end;

procedure TFrmMoeBordo.CxTextEdtCantidadKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = 13 then
    CxBtnok.Click;
end;

procedure TFrmMoeBordo.FormShow(Sender: TObject);
Var
  Cursor: TCursor;
begin
  try
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    try
      QMoe.Active := False;
      QMoe.SQL.Clear;
      QMoe.SQL.Text := 'select * from moe where scontrato=:Contrato order by dIdFecha';
      QMoe.Params.ParamByName('Contrato').AsString := Global_Contrato;
      QMoe.Open;
    finally
      Screen.Cursor := Cursor;
    end;
  Except
    on e: Exception do
      MessageDlg('Ha Ocurrido un error inesperado, informar al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

procedure TFrmMoeBordo.frmBarra1btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmMoeBordo.pgDatosChange(Sender: TObject);
begin
  ScrollMoe;
end;

procedure TFrmMoeBordo.QMoeAfterScroll(DataSet: TDataSet);
begin
  ScrollMoe;
end;

procedure TFrmMoeBordo.QMoeEquipoAfterScroll(DataSet: TDataSet);
Var
  Cursor: TCursor;
begin
//  try
//    Cursor := Screen.Cursor;
//    Screen.Cursor := crAppStart;
//    try
//      QMoeEquipo.Active := False;
//      QMoeEquipo.SQL.Clear;
//      QMoeEquipo.SQL.Text := 'select * from moerecursos where iIdMoe=:Id and eTipoRecurso="Equipo"';
//      QMoeEquipo.Open;
//    finally
//      Screen.Cursor := Cursor;
//    end;
//  Except
//    on e: Exception do
//      MessageDlg('Ha Ocurrido un error inesperado, informar al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
//  end;
end;
procedure TFrmMoeBordo.QMoeEquiposIdRecursoChange(Sender: TField);
var
  sDescripcion:string;
begin
  Connection.qryBusca.Active := False;
  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('select iItemOrden, sDescripcion, dCostoDLL, dCostoMN, sAgrupaPersonal from personal where sContrato = :Contrato And sIdPersonal = :Personal');
  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value := Global_Contrato_Barco;
  Connection.qryBusca.Params.ParamByName('Personal').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Personal').Value := QMoePersonalsIdRecurso.Text;
  Connection.qryBusca.Open;
  if Connection.qryBusca.RecordCount > 0 then
  begin
      QMoePersonal.FieldByName('sDescripcion').AsString := Connection.qryBusca.FieldByName('sDescripcion').AsString;
  end
  else
    if not QMoePersonal.FieldByName('sIdRecurso').IsNull then
      if Trim(QMoePersonal.FieldValues['sIdRecurso']) <> '' then
      begin
        sDescripcion := '%' + Trim(QMoePersonal.FieldValues['sIdRecurso']) + '%';
        BuscaObjeto.Active := False;
        ListaObjeto.Columns.Clear;
        ListaObjeto.Columns.Add;
        ListaObjeto.Columns[0].FieldName := 'sNumeroActividad';
        ListaObjeto.Columns.Add;
        ListaObjeto.Columns[1].FieldName := 'sDescripcion';
        BuscaObjeto.SQL.Clear;
        BuscaObjeto.SQL.Add('Select iItemOrden, sIdPersonal as sNumeroActividad, sDescripcion, dCostoDLL, dCostoMN  from personal Where ' +
          'sContrato = :Contrato And sDescripcion Like :Descripcion Order by sDescripcion');
        BuscaObjeto.Params.ParamByName('Contrato').DataType := ftString;
        BuscaObjeto.Params.ParamByName('Contrato').Value := Global_Contrato_Barco;
        BuscaObjeto.Params.ParamByName('Descripcion').DataType := ftString;
        BuscaObjeto.Params.ParamByName('Descripcion').Value := sDescripcion;
        BuscaObjeto.Open;
        Panel.Visible := True;
        Panel.Height  := 358;
        Panel.Width   := 590;
        ListaObjeto.Columns[0].Width := 50;
        ListaObjeto.Columns[1].Width := 680;
        ListaObjeto.SetFocus;
      end;
end;

procedure TFrmMoeBordo.ScrollMoe;
Var
  Cursor: TCursor;
  TipoMoe: string;
begin
  try
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    try
      if (pgDatos.ActivePage = pgEquipo) or (pgDatos.ActivePage = pgPersonal) then
      begin
        if pgDatos.ActivePage = pgEquipo then
        begin
          TipoMoe := 'Equipo';
          QMoeEquipo.Active := False;
          QMoeEquipo.SQL.Clear;
          QMoeEquipo.SQL.Text := 'select MoeE.*, mr.dCantidad as dCantR from moerecursos_aBordo as MoeE' +
                                 ' inner join moerecursos as mr on (mr.iIdMoe = MoeE.iIdMoe and mr.sIdRecurso = MoeE.sIdRecurso and MoeE.eTipoRecurso = "' + TipoMoe + '") ' +
                                 ' where MoeE.iIdMoe=:Id and MoeE.eTipoRecurso="' + TipoMoe + '"';
          QMoeEquipo.ParamByName('Id').AsInteger := QMoe.FieldByName('iIdMoe').AsInteger;
          QMoeEquipo.Open;
        end;
        if pgDatos.ActivePage = pgPersonal then
        begin
          TipoMoe := 'Personal';
          QMoePersonal.Active := False;
          QMoePersonal.SQL.Clear;
          QMoePersonal.SQL.Text := 'select MoeP.*, mr.dCantidad as dCantR from moerecursos_aBordo as MoeP, moerecursos as mr ' +
                                   ' where MoeP.iIdMoe=:Id and MoeP.eTipoRecurso="' + TipoMoe + '" and mr.iIdMoe = MoeP.iIdMoe and mr.sIdRecurso = MoeP.sIdRecurso and MoeP.eTipoRecurso = "' + TipoMoe + '"';
          QMoePersonal.ParamByName('Id').AsInteger := QMoe.FieldByName('iIdMoe').AsInteger;
          QMoePersonal.Open;
        end;
      end;
    finally
      Screen.Cursor := Cursor;
    end;
  except
    on e: Exception do
      MessageDlg('Ha Ocurrido un error inesperado, informar al administrador del sistema del siguiente error: ' + e.Message, mtError, [mbOK], 0);
  end;
end;

end.
