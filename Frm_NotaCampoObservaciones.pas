unit Frm_NotaCampoObservaciones;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, AdvPageControl, ComCtrls, StdCtrls,
  AdvDateTimePicker, AdvDBDateTimePicker, Grids, DBGrids, JvExDBGrids, JvDBGrid,
  JvDBUltimGrid, frm_barra, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  Mask, DBCtrls, AdvEdit, AdvEdBtn, DBAdvEdBtn, JvExStdCtrls, JvCombobox,
  JvDBCombobox;

type
  TFrmNotaCampoObservaciones = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    AdvPageControl1: TAdvPageControl;
    AdvTabSheet1: TAdvTabSheet;
    AdvTabSheet2: TAdvTabSheet;
    Panel3: TPanel;
    Label1: TLabel;
    lblFolio: TLabel;
    Label2: TLabel;
    AdvDBDateTimePicker1: TAdvDBDateTimePicker;
    Panel4: TPanel;
    frmBarra1: TfrmBarra;
    JvDBUltimGrid1: TJvDBUltimGrid;
    QNotaCampo: TZQuery;
    dsNotaCampo: TDataSource;
    AvDbDtpFecha: TAdvDBDateTimePicker;
    Label3: TLabel;
    tsFirma1: TDBAdvEditBtn;
    Label4: TLabel;
    Label45: TLabel;
    tsPuesto1: TDBEdit;
    Label5: TLabel;
    tsFirma2: TDBAdvEditBtn;
    Label6: TLabel;
    Label7: TLabel;
    tsPuesto2: TDBEdit;
    Label8: TLabel;
    tsFirma3: TDBAdvEditBtn;
    Label9: TLabel;
    Label10: TLabel;
    tsPuesto3: TDBEdit;
    Label11: TLabel;
    JDbCmbFirmantes: TJvDBComboBox;
    QNotaObservaciones: TZQuery;
    dsNotaObservaciones: TDataSource;
    Label12: TLabel;
    dbmmoObs: TDBMemo;
    Label13: TLabel;
    DBMemo1: TDBMemo;
    procedure JDbCmbFirmantesEnter(Sender: TObject);
    procedure JDbCmbFirmantesExit(Sender: TObject);
    procedure JDbCmbFirmantesKeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma1KeyPress(Sender: TObject; var Key: Char);
    procedure tsFirma2KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto1KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto2KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto3KeyPress(Sender: TObject; var Key: Char);
    procedure tsPuesto1Enter(Sender: TObject);
    procedure tsPuesto1Exit(Sender: TObject);
    procedure tsFirma1Enter(Sender: TObject);
    procedure tsFirma1Exit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure AdvDBDateTimePicker1Enter(Sender: TObject);
    procedure AdvDBDateTimePicker1Exit(Sender: TObject);
    procedure AvDbDtpFechaEnter(Sender: TObject);
    procedure AvDbDtpFechaExit(Sender: TObject);
    procedure DBMemo1Enter(Sender: TObject);
    procedure DBMemo1Exit(Sender: TObject);
    procedure dbmmoObsEnter(Sender: TObject);
    procedure dbmmoObsExit(Sender: TObject);
    procedure AdvDBDateTimePicker1KeyPress(Sender: TObject; var Key: Char);
    procedure AvDbDtpFechaKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
    ParamContrato,
    ParamFolio:String;
  end;

var
  FrmNotaCampoObservaciones: TFrmNotaCampoObservaciones;

implementation

uses frm_connection, global;

{$R *.dfm}

procedure TFrmNotaCampoObservaciones.AdvDBDateTimePicker1Enter(Sender: TObject);
begin
  AdvDBDateTimePicker1.Color:=Global_Color_Entrada;
end;

procedure TFrmNotaCampoObservaciones.AdvDBDateTimePicker1Exit(Sender: TObject);
begin
  AdvDBDateTimePicker1.Color:=Global_Color_Salida;
end;

procedure TFrmNotaCampoObservaciones.AdvDBDateTimePicker1KeyPress(
  Sender: TObject; var Key: Char);
begin
   If Key = #13 Then
      dbMemo1.SetFocus
end;

procedure TFrmNotaCampoObservaciones.AvDbDtpFechaEnter(Sender: TObject);
begin
  AvDbDtpFecha.Color:=Global_Color_Entrada;
end;

procedure TFrmNotaCampoObservaciones.AvDbDtpFechaExit(Sender: TObject);
begin
  AvDbDtpFecha.Color:=Global_Color_Salida;
end;

procedure TFrmNotaCampoObservaciones.AvDbDtpFechaKeyPress(Sender: TObject;
  var Key: Char);
begin
   If Key = #13 Then
      dbmmoObs.SetFocus
end;

procedure TFrmNotaCampoObservaciones.DBMemo1Enter(Sender: TObject);
begin
  DBMemo1.Color:=Global_color_entrada;
end;

procedure TFrmNotaCampoObservaciones.DBMemo1Exit(Sender: TObject);
begin
  DBMemo1.Color:=Global_color_salida;
end;

procedure TFrmNotaCampoObservaciones.dbmmoObsEnter(Sender: TObject);
begin
  dbmmoObs.Color:=Global_color_entrada;
end;

procedure TFrmNotaCampoObservaciones.dbmmoObsExit(Sender: TObject);
begin
  dbmmoObs.Color:=Global_color_Salida;
end;

procedure TFrmNotaCampoObservaciones.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  if QNotaCampo.State in [dsinsert,dsedit] then
    QNotaCampo.Post;
end;

procedure TFrmNotaCampoObservaciones.FormCreate(Sender: TObject);
begin
  AdvPageControl1.ActivePageIndex:=0;
end;

procedure TFrmNotaCampoObservaciones.FormShow(Sender: TObject);
begin
  lblFolio.Caption:= ParamFolio;
  QNotaCampo.Active:=false;
  QNotaCampo.ParamByName('Contrato').AsString:=ParamContrato;
  QNotaCampo.ParamByName('Folio').AsString:=ParamFolio;
  QNotaCampo.Open;
  if QNotaCampo.recordcount=0 then
  begin
    connection.QryBusca.Active:=false;
    connection.QryBusca.SQL.Text:='select * from notacampo_general where dfecha='+
                                   '(select max(dfecha) from notacampo_general)' ;
    connection.QryBusca.Open;

    QNotaCampo.Append;
    QNotaCampo.FieldByName('sContrato').AsString:=ParamContrato;
    QNotaCampo.FieldByName('sNumeroOrden').AsString:=ParamFolio;
    QNotaCampo.FieldByName('dFecha').AsDateTime:=now;
    QNotaCampo.FieldByName('iNumFirmante').AsInteger:=2;
    QNotaCampo.FieldByName('sPeriodo').AsString:='';

    if connection.QryBusca.RecordCount=0 then
    begin
      QNotaCampo.FieldByName('sFirmante1').AsString:='';
      QNotaCampo.FieldByName('sFirmante2').AsString:='';
      QNotaCampo.FieldByName('sFirmante3').AsString:='';
      QNotaCampo.FieldByName('sPuesto1').AsString:='';
      QNotaCampo.FieldByName('sPuesto2').AsString:='';
      QNotaCampo.FieldByName('sPuesto3').AsString:='';
    end
    else
    begin
      QNotaCampo.FieldByName('sFirmante1').AsString:=connection.QryBusca.FieldByName('sFirmante1').AsString;
      QNotaCampo.FieldByName('sFirmante2').AsString:=connection.QryBusca.FieldByName('sFirmante2').AsString;
      QNotaCampo.FieldByName('sFirmante3').AsString:=connection.QryBusca.FieldByName('sFirmante3').AsString;
      QNotaCampo.FieldByName('sPuesto1').AsString:=connection.QryBusca.FieldByName('sPuesto1').AsString;
      QNotaCampo.FieldByName('sPuesto2').AsString:=connection.QryBusca.FieldByName('sPuesto2').AsString;
      QNotaCampo.FieldByName('sPuesto3').AsString:=connection.QryBusca.FieldByName('sPuesto3').AsString;
    end;

    connection.QryBusca.Active:=false;
    connection.QryBusca.SQL.Text:='select (ifnull(max(iIdNota),0)+1) as Next from notacampo_general' ;
    connection.QryBusca.Open;

    if connection.QryBusca.RecordCount=0 then
      QNotaCampo.FieldByName('iIdNota').AsInteger:=1
    else
      QNotaCampo.FieldByName('iIdNota').AsInteger:= connection.QryBusca.FieldByName('next').AsInteger;
    QNotaCampo.Post;
  end;

  QNotaObservaciones.Active:=false;
  QNotaObservaciones.ParamByName('Nota').AsInteger:=QNotaCampo.FieldByName('iIdNota').AsInteger;
  QNotaObservaciones.Open;
  QNotaCampo.edit;

end;

procedure TFrmNotaCampoObservaciones.frmBarra1btnAddClick(Sender: TObject);
begin
  QNotaObservaciones.Append;
  QNotaObservaciones.FieldByName('iIdNota').AsInteger:=QNotaCampo.FieldByName('iIdNota').AsInteger;
  QNotaObservaciones.FieldByName('dFecha').AsDateTime:=now;
  frmBarra1.btnAddClick(Sender);

end;

procedure TFrmNotaCampoObservaciones.frmBarra1btnCancelClick(Sender: TObject);
begin
  QNotaObservaciones.Cancel;
  frmBarra1.btnCancelClick(Sender);

end;

procedure TFrmNotaCampoObservaciones.frmBarra1btnDeleteClick(Sender: TObject);
begin
  if MessageDlg('La Observacion del dia: ' + QNotaObservaciones.fieldByname('DFecha').AsString + ' ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    QNotaObservaciones.Delete;
end;

procedure TFrmNotaCampoObservaciones.frmBarra1btnEditClick(Sender: TObject);
begin
  QNotaObservaciones.Edit;
  frmBarra1.btnEditClick(Sender);

end;

procedure TFrmNotaCampoObservaciones.frmBarra1btnExitClick(Sender: TObject);
begin
  close;
  frmBarra1.btnExitClick(Sender);

end;

procedure TFrmNotaCampoObservaciones.frmBarra1btnPostClick(Sender: TObject);
begin
  if Length(Trim(dbmmoObs.Text))=0 then
  begin
    messagedlg('Debe agregar una Observacion.',mtinformation,[MbOK],0);
    dbmmoObs.setfocus;
    exit;
  end;
  QNotaObservaciones.Post;
  frmBarra1.btnPostClick(Sender);

end;

procedure TFrmNotaCampoObservaciones.frmBarra1btnRefreshClick(Sender: TObject);
begin
  QNotaObservaciones.Refresh;
end;

procedure TFrmNotaCampoObservaciones.JDbCmbFirmantesEnter(Sender: TObject);
begin
  JDbCmbFirmantes.Color:=global_Color_Entrada;
end;

procedure TFrmNotaCampoObservaciones.JDbCmbFirmantesExit(Sender: TObject);
begin
  JDbCmbFirmantes.Color:=global_Color_Salida;
end;

procedure TFrmNotaCampoObservaciones.JDbCmbFirmantesKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
      TsPuesto1.SetFocus
end;

procedure TFrmNotaCampoObservaciones.tsFirma1Enter(Sender: TObject);
begin
  TdbADvEditBtn(Sender).Color:=global_Color_Entrada;
end;

procedure TFrmNotaCampoObservaciones.tsFirma1Exit(Sender: TObject);
begin
  TdbADvEditBtn(Sender).Color:=global_Color_Salida;
end;

procedure TFrmNotaCampoObservaciones.tsFirma1KeyPress(Sender: TObject;
  var Key: Char);
begin
  If Key = #13 Then
    TsPuesto2.SetFocus
end;

procedure TFrmNotaCampoObservaciones.tsFirma2KeyPress(Sender: TObject;
  var Key: Char);
begin
  If Key = #13 Then
    TsPuesto3.SetFocus
end;

procedure TFrmNotaCampoObservaciones.tsPuesto1Enter(Sender: TObject);
begin
  TdbEdit(Sender).Color:=global_Color_Entrada;
end;

procedure TFrmNotaCampoObservaciones.tsPuesto1Exit(Sender: TObject);
begin
  TdbEdit(Sender).Color:=global_Color_Salida;
end;

procedure TFrmNotaCampoObservaciones.tsPuesto1KeyPress(Sender: TObject;
  var Key: Char);
begin
  If Key = #13 Then
    TsFirma1.SetFocus
end;

procedure TFrmNotaCampoObservaciones.tsPuesto2KeyPress(Sender: TObject;
  var Key: Char);
begin
  If Key = #13 Then
    TsFirma2.SetFocus
end;

procedure TFrmNotaCampoObservaciones.tsPuesto3KeyPress(Sender: TObject;
  var Key: Char);
begin
  If Key = #13 Then
    TsFirma3.SetFocus
end;

end.
