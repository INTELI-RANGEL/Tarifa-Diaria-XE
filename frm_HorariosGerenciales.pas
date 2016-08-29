unit frm_HorariosGerenciales;

interface

uses
  frm_connection, UnitMetodos, global,

  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxStyles, dxSkinsCore, dxSkinDevExpressStyle, dxSkinFoggy, dxSkinscxPCPainter,
  cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit, cxNavigator, DB,
  cxDBData, frm_barra, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxClasses, cxGridCustomView, cxGrid, cxContainer,
  cxGroupBox, cxLabel, cxMaskEdit, cxTextEdit, cxDropDownEdit, cxCalc, cxDBEdit,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, cxCheckBox;

type
  TfrmHorariosGerenciales = class(TForm)
    gridHorarios: TcxGrid;
    gridDbHorarios: TcxGridDBTableView;
    gridDbHorariosColumn3: TcxGridDBColumn;
    gridDbHorariosColumn2: TcxGridDBColumn;
    cxFoliosLvl: TcxGridLevel;
    TfrmBarra1: TfrmBarra;
    grpCaptura: TcxGroupBox;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    dsHorarios: TDataSource;
    ZHorarios: TZQuery;
    dbHorario: TcxDBMaskEdit;
    dbGerencial: TcxDBTextEdit;
    gridDbHorariosColumn1: TcxGridDBColumn;
    dbPrincipal: TcxDBCheckBox;
    dbSecundario: TcxDBCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure TfrmBarra1btnAddClick(Sender: TObject);
    procedure TfrmBarra1btnEditClick(Sender: TObject);
    procedure TfrmBarra1btnPostClick(Sender: TObject);
    procedure TfrmBarra1btnCancelClick(Sender: TObject);
    procedure TfrmBarra1btnDeleteClick(Sender: TObject);
    procedure TfrmBarra1btnRefreshClick(Sender: TObject);
    procedure TfrmBarra1btnExitClick(Sender: TObject);
    procedure GlobalKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmHorariosGerenciales: TfrmHorariosGerenciales;

implementation

{$R *.dfm}

procedure TfrmHorariosGerenciales.GlobalKeyUp(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    Perform( CM_DIALOGKEY, VK_TAB, 0 );
    Key := 0
  end;
end;

procedure TfrmHorariosGerenciales.FormCreate(Sender: TObject);
begin
  ZHorarios.Active := False;
  ZHorarios.SQL.Text := ObtenerSentencia( 'horarios_gerenciales', 'sql_horarios_gerenciales', ftCatalogo );
  ZHorarios.ParamByName( 'todos' ).AsInteger := Integer( True );
  ZHorarios.ParamByName( 'principales' ).AsInteger := Integer( False );
  ZHorarios.Open;
end;

procedure TfrmHorariosGerenciales.TfrmBarra1btnAddClick(Sender: TObject);
begin
  TfrmBarra1.btnAddClick(Sender);
  ZHorarios.Append;
  grpCaptura.Enabled := True;
  gridHorarios.Enabled := False;
  dbGerencial.SetFocus;
end;

procedure TfrmHorariosGerenciales.TfrmBarra1btnCancelClick(Sender: TObject);
begin
  ZHorarios.Cancel;
  grpCaptura.Enabled := False;
  gridHorarios.Enabled := True;
  TfrmBarra1.btnCancelClick(Sender);
end;

procedure TfrmHorariosGerenciales.TfrmBarra1btnDeleteClick(Sender: TObject);
begin
  if ( ZHorarios.RecordCount > 0 ) and ( TaskMessageDlg( 'Confirmación', '¿Desea eliminar el registro activo?', mtConfirmation, [ mbYes, mbNo ], 0 ) = mrYes ) then
    ZHorarios.Delete;
end;

procedure TfrmHorariosGerenciales.TfrmBarra1btnEditClick(Sender: TObject);
begin
  TfrmBarra1.btnEditClick(Sender);
  if ZHorarios.RecordCount > 0 then
  begin
    ZHorarios.Edit;
    grpCaptura.Enabled := True;
    gridHorarios.Enabled := False;
    dbGerencial.SetFocus;
  end;
end;

procedure TfrmHorariosGerenciales.TfrmBarra1btnExitClick(Sender: TObject);
begin  
  Close;
end;

procedure TfrmHorariosGerenciales.TfrmBarra1btnPostClick(Sender: TObject);
begin
  OpcButton := EmptyStr;
  if ( ZHorarios.State in [ dsEdit, dsInsert ] ) then
  begin
    ZHorarios.FieldByName( 'NumeroGerencial' ).AsInteger := StrToInt( dbGerencial.Text );
    ZHorarios.FieldByName( 'Horario' ).AsString := dbHorario.Text;
    ZHorarios.Post;
    grpCaptura.Enabled := False;
    gridHorarios.Enabled := True;
  end;
  TfrmBarra1.btnPostClick(Sender);
end;

procedure TfrmHorariosGerenciales.TfrmBarra1btnRefreshClick(Sender: TObject);
begin
  ZHorarios.Active := False;
  ZHorarios.Open;
end;

end.
