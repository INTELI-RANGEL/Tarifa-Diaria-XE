unit Frm_Materiales;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxContainer, cxEdit, dxSkinsCore, dxSkinDevExpressStyle, dxSkinFoggy,
  cxGroupBox, DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, cxStyles,
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxNavigator, cxDBData, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  cxDBLookupComboBox, cxMemo;

type
  TFrmAltaMAterial = class(TForm)
    GBx1: TcxGroupBox;
    GBx2: TcxGroupBox;
    QInsumos: TZQuery;
    dsInsumos: TDataSource;
    QrAlmacen: TZReadOnlyQuery;
    dsAlmacenes: TDataSource;
    CxGrdDbTblVMateriales: TcxGridDBTableView;
    CxGLvlGrid1Level1: TcxGridLevel;
    CxGrd1: TcxGrid;
    CxGrdDbTblVMaterialesColumn1: TcxGridDBColumn;
    CxGrdDbTblVMaterialesColumn2: TcxGridDBColumn;
    CxGrdDbTblVMaterialesColumn3: TcxGridDBColumn;
    CxGrdDbTblVMaterialesColumn4: TcxGridDBColumn;
    CxGrdDbTblVMaterialesColumn5: TcxGridDBColumn;
    procedure FormShow(Sender: TObject);
    procedure QInsumosAfterInsert(DataSet: TDataSet);
    procedure FormCreate(Sender: TObject);
    procedure QInsumosBeforePost(DataSet: TDataSet);
  private
    { Private declarations }
    function GenerarCodigo:string;
  public
    { Public declarations }
    SeInserto:Boolean;
  end;

var
  FrmAltaMAterial: TFrmAltaMAterial;

implementation

uses frm_connection, global;

{$R *.dfm}

function TFrmAltaMAterial.GenerarCodigo:String;
var
  QRCode:TZReadOnlyQuery;
  MaxCode:Integer;
  Code:string;
begin
  MaxCode:=0;

  QRCode:=TZReadOnlyQuery.Create(nil);
  try
    QRCode.Connection:=connection.zConnection;
    QRCode.SQL.Text:='select count(*) as Numt from insumos where sContrato=:Contrato';
    QRCode.ParamByName('Contrato').AsString:=global_Contrato;
    QRCode.Open;
    if QRCode.RecordCount=0 then
      MaxCode:=1
    else
      MaxCode:= QRCode.FieldByName('numt').AsInteger + 1;


    Code:=IntToStr(MaxCode);
    while Length(Code)<5 do
      Code:='0'+ Code;

  finally
    QRCode.Destroy;
  end;
  Result:='MAT' + Code;
end;
procedure TFrmAltaMAterial.QInsumosAfterInsert(DataSet: TDataSet);
begin
  QInsumos.FieldByName('sContrato').AsString:= global_contrato;
  QInsumos.FieldByName('sIdInsumo').AsString:=GenerarCodigo;
  QInsumos.FieldByName('sTrazabilidad').AsString:='S/T' ;
  QInsumos.FieldByName('smedida').AsString:='' ;
  QInsumos.FieldByName('sColumnaAux').AsString:='*';

  if QrAlmacen.RecordCount>0 then
  begin
    QrAlmacen.First;
    QInsumos.FieldByName('sIdAlmacen').AsString:=QrAlmacen.FieldByName('sIdAlmacen').AsString;
  end;

end;

procedure TFrmAltaMAterial.QInsumosBeforePost(DataSet: TDataSet);
begin
  if QInsumos.FieldByName('mDescripcion').IsNull then
  begin
    QInsumos.Cancel;
    Abort;
  end;
  SeInserto:=True;
end;

procedure TFrmAltaMAterial.FormCreate(Sender: TObject);
begin
  SeInserto:=False;
end;

procedure TFrmAltaMAterial.FormShow(Sender: TObject);
begin
  QrAlmacen.Open;
  QInsumos.Active:=False;
  QInsumos.ParamByName('Contrato').AsString:=global_Contrato;
  QInsumos.Open;
  QInsumos.Append;

end;

end.
