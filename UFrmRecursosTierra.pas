unit UFrmRecursosTierra;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxContainer, cxEdit, dxSkinsCore, dxSkinDevExpressStyle, dxSkinFoggy,
  cxSplitter, cxGroupBox, cxStyles, dxSkinscxPCPainter, cxCustomData, cxFilter,
  cxData, cxDataStorage, cxNavigator, DB, cxDBData, cxGridLevel, cxClasses,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGrid, dxBarBuiltInMenu, cxPC, ZAbstractRODataset, ZDataset, FramedPanel,
  ZAbstractDataset, cxTextEdit;

type
  TFrmRecursosTierra = class(TForm)
    GBxClient: TcxGroupBox;
    GBxBottom: TcxGroupBox;
    SplPrincipal: TcxSplitter;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    CxPageRecursos: TcxPageControl;
    cTsPersonal: TcxTabSheet;
    cTsEquipo: TcxTabSheet;
    QrNotasCortes: TZReadOnlyQuery;
    dsNotasCortes: TDataSource;
    cxGrid1DBTableView1Column1: TcxGridDBColumn;
    cxGrid1DBTableView1Column2: TcxGridDBColumn;
    cxGrid1DBTableView1Column3: TcxGridDBColumn;
    cxGrid1DBTableView1Column4: TcxGridDBColumn;
    cxGrid1DBTableView1Column6: TcxGridDBColumn;
    cxGrid1DBTableView1Column7: TcxGridDBColumn;
    QrConceptos: TZReadOnlyQuery;
    dsConceptos: TDataSource;
    CxGLvlGrid1Level2: TcxGridLevel;
    CxGrdDbTblVGrid1DBTableView2: TcxGridDBTableView;
    CxGrdDbTblVGrid1DBTableView2Column1: TcxGridDBColumn;
    CxGrdDbTblVGrid1DBTableView2Column2: TcxGridDBColumn;
    CxGrdDbTblVGrid1DBTableView2Column3: TcxGridDBColumn;
    CxGrdDbTblVGrid1DBTableView2Column4: TcxGridDBColumn;
    CxGrdDbTblVGrid1DBTableView2Column5: TcxGridDBColumn;
    CxGrdDbTblVGrid1DBTableView2Column6: TcxGridDBColumn;
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    FramedPanel1: TFramedPanel;
    GBx1: TcxGroupBox;
    cxGrid2DBTableView1: TcxGridDBTableView;
    cxGrid2Level1: TcxGridLevel;
    cxGrid2: TcxGrid;
    cxGrid3DBTableView1: TcxGridDBTableView;
    cxGrid3Level1: TcxGridLevel;
    cxGrid3: TcxGrid;
    QPersonal: TZQuery;
    dsPersonal: TDataSource;
    QEquipos: TZQuery;
    dsEquipos: TDataSource;
    cxGrid3DBTableView1Column1: TcxGridDBColumn;
    cxGrid3DBTableView1Column2: TcxGridDBColumn;
    cxGrid3DBTableView1Column3: TcxGridDBColumn;
    cxGrid3DBTableView1Column4: TcxGridDBColumn;
    cxGrid4: TcxGrid;
    cxGridDBTableView1: TcxGridDBTableView;
    cxGridDBColumn1: TcxGridDBColumn;
    cxGridDBColumn2: TcxGridDBColumn;
    cxGridDBColumn3: TcxGridDBColumn;
    cxGridDBColumn4: TcxGridDBColumn;
    cxGridDBColumn5: TcxGridDBColumn;
    cxGridDBColumn6: TcxGridDBColumn;
    cxGridLevel1: TcxGridLevel;
    cxStyle2: TcxStyle;
    procedure FormShow(Sender: TObject);
    procedure CxGrdDbTblVGrid1DBTableView2SelectionChanged(
      Sender: TcxCustomGridTableView);
  private
    { Private declarations }
    paramContrato,
    ParamConvenio,
    ParamFolio    :string;
    ParamFecha:TDate;
  public
    { Public declarations }
    constructor MiCreate(AOwner: TComponent;pContrato,pConvenio,pFolio:string;pFecha:TDate);
  end;

var
  FrmRecursosTierra: TFrmRecursosTierra;

implementation

uses frm_connection, global;

{$R *.dfm}

constructor TFrmRecursosTierra.MiCreate(AOwner: TComponent; pContrato: string; pConvenio: string; pFolio: string; pFecha: TDate);
begin
  paramContrato:=pContrato;
  ParamConvenio:=pConvenio;
  ParamFolio:=pFolio;
  ParamFecha:=pFecha;
  inherited Create(AOwner);
end;

procedure TFrmRecursosTierra.CxGrdDbTblVGrid1DBTableView2SelectionChanged(
  Sender: TcxCustomGridTableView);
begin
  QPersonal.close;
  QPersonal.ParamByName('contrato').AsString    := paramContrato ;
  QPersonal.ParamByName('fecha').AsDate    := ParamFecha ;
  QPersonal.ParamByName('Diario').AsInteger      := StrToIntDef(sender.DataController.GetDisplayText(sender.Controller.FocusedRecordIndex,0),-1);
  QPersonal.ParamByName('Actividad').AsInteger         :=StrToIntDef(sender.DataController.GetDisplayText(sender.Controller.FocusedRecordIndex,5),-1);
  QPersonal.Open;
end;

procedure TFrmRecursosTierra.FormShow(Sender: TObject);
begin
  QrConceptos.Active := false;
  QrConceptos.ParamByName('Contrato').AsString  := paramContrato;
  QrConceptos.ParamByName('Convenio').AsString  := ParamConvenio;
  QrConceptos.ParamByName('Fecha').AsDate       := ParamFecha;
  QrConceptos.ParamByName('folio').asstring     := ParamFolio;
  QrConceptos.Open;

  QrNotasCortes.Active := false;
  QrNotasCortes.ParamByName('barco').AsString  :=global_Contrato_Barco;
  QrNotasCortes.ParamByName('Contrato').AsString  := paramContrato;
  QrNotasCortes.ParamByName('Convenio').AsString  := ParamConvenio;
  QrNotasCortes.ParamByName('Fecha').AsDate       := ParamFecha;
  QrNotasCortes.ParamByName('folio').asstring     := ParamFolio;
  QrNotasCortes.Open;

  cxGrid1DBTableView1.ViewData.Expand(true);
  //ShowMessage(IntToStr(QrNotasCortes.RecordCount));
end;

end.
