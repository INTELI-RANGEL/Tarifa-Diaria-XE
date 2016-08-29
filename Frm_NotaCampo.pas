unit Frm_NotaCampo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxContainer, cxEdit, dxSkinsCore, dxSkinDevExpressStyle, dxSkinFoggy,
  cxStyles, dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxNavigator, DB, cxDBData, frm_barra, cxGridLevel, cxClasses,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGrid, cxSplitter, cxGroupBox, dxLayoutContainer, dxLayoutControl,
  dxLayoutcxEditAdapters, cxTextEdit, cxDBEdit, cxMaskEdit, cxDropDownEdit,
  cxLookupEdit, cxDBLookupEdit, cxDBLookupComboBox, cxCalendar, cxRadioGroup,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UnitTarifa, frxClass,
  dxBarBuiltInMenu, cxPC, cxSSheet, StdCtrls, Menus,cxSSTypes, JvMemoryDataset,
  cxMemo, frxDBSet, cxCheckGroup, cxDBCheckGroup, dxmdaset, cxCheckBox,
  cxButtonEdit, ZSqlUpdate, cxCalc, cxButtons,
  dxLayoutControlAdapters, dxSkinBlack, dxSkinBlue, dxSkinBlueprint,
  dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide,
  dxSkinDevExpressDarkStyle, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven,
  dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus, dxSkinSilver,
  dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008, dxSkinTheAsphaltWorld,
  dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue;

type TypeDatos=(ctNew,ctLoad);
type TypeFormat=(FrTitle,FrSubTitle,FrContent,FrNone);

type
  TFrmNotaCampo = class(TForm)
    QActa: TZQuery;
    dsActa: TDataSource;
    QrFolios: TZReadOnlyQuery;
    dsFolios: TDataSource;
    RptActa: TfrxReport;
    QrImprimir: TZReadOnlyQuery;
    CxPage1: TcxPageControl;
    cTs1: TcxTabSheet;
    cTs2: TcxTabSheet;
    Spl1: TcxSplitter;
    GBx1: TcxGroupBox;
    cxGrid1: TcxGrid;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1DBTableView1Column1: TcxGridDBColumn;
    cxGrid1DBTableView1Column2: TcxGridDBColumn;
    cxGrid1DBTableView1Column3: TcxGridDBColumn;
    cxGrid1Level1: TcxGridLevel;
    BrPrincipal: TfrmBarra;
    GBx3: TcxGroupBox;
    GBx4: TcxGroupBox;
    SprShBkDatos: TcxSpreadSheetBook;
    pmDatos: TPopupMenu;
    mniAdd: TMenuItem;
    mniDelete: TMenuItem;
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    Md1: TdxMemData;
    JMry1: TJvMemoryData;
    CxPageDetalle: TcxPageControl;
    cTsCaratula: TcxTabSheet;
    cTsMateriales: TcxTabSheet;
    GBx2: TcxGroupBox;
    dxLayoutControl1: TdxLayoutControl;
    DbMmObservacion: TcxDBMemo;
    DbRdGrpTipo: TcxDBRadioGroup;
    DbTxtEdtActa: TcxDBTextEdit;
    DbLkpCmbFolio: TcxDBLookupComboBox;
    DbTxtEdtEspecialidad: TcxDBTextEdit;
    DbDtEdtFecha: TcxDBDateEdit;
    GBx5: TcxGroupBox;
    dxLayoutControl2: TdxLayoutControl;
    DbChkBxPernocta: TcxDBCheckBox;
    DbChkBxMaterial: TcxDBCheckBox;
    DbChkBxPaginas: TcxDBCheckBox;
    DbChkBxPdas: TcxDBCheckBox;
    dxLayoutControl2Group_Root: TdxLayoutGroup;
    dxLayoutItem1: TdxLayoutItem;
    dxLayoutControl2Item2: TdxLayoutItem;
    dxLayoutControl2Item3: TdxLayoutItem;
    dxLayoutControl2Item4: TdxLayoutItem;
    cxDBMemo1: TcxDBMemo;
    cxDBMemo2: TcxDBMemo;
    DbTxtEdtActivo: TcxDBTextEdit;
    dxLayoutControl1Group_Root: TdxLayoutGroup;
    dxLayoutItem2: TdxLayoutItem;
    dxLayoutControl1Item5: TdxLayoutItem;
    dxLayoutControl1Group2: TdxLayoutAutoCreatedGroup;
    dxLayoutItem3: TdxLayoutItem;
    dxLayoutControl1Item2: TdxLayoutItem;
    dxLayoutControl1Group4: TdxLayoutAutoCreatedGroup;
    dxLayoutControl1Item3: TdxLayoutItem;
    dxLayoutControl1Item4: TdxLayoutItem;
    dxLayoutControl1Group3: TdxLayoutAutoCreatedGroup;
    dxLayoutControl1Item6: TdxLayoutItem;
    dxLayoutAutoCreatedGroup1: TdxLayoutAutoCreatedGroup;
    dxLayoutControl1Item7: TdxLayoutItem;
    dxLayoutControl1Item8: TdxLayoutItem;
    dxLayoutControl1Item9: TdxLayoutItem;
    CxGrdDbTblVMateriales: TcxGridDBTableView;
    CxGLvlGrid2Level1: TcxGridLevel;
    CxGrd1: TcxGrid;
    QMateriales: TZQuery;
    dsMateriales: TDataSource;
    CxGrdDbTblVMaterialesColumn1: TcxGridDBColumn;
    CxGrdDbTblVMaterialesColumn2: TcxGridDBColumn;
    CxGrdDbTblVMaterialesColumn3: TcxGridDBColumn;
    CxGrdDbTblVMaterialesColumn4: TcxGridDBColumn;
    CxGrdDbTblVMaterialesColumn5: TcxGridDBColumn;
    QrActividades: TZReadOnlyQuery;
    QrInsumos: TZReadOnlyQuery;
    dsActividades: TDataSource;
    dsInsumos: TDataSource;
    CxGrdDbTblVMaterialesColumn6: TcxGridDBColumn;
    CxGrdDbTblVMaterialesColumn7: TcxGridDBColumn;
    UdSqlMateriales: TZUpdateSQL;
    dlgSaveGuardar: TSaveDialog;
    btnAjustar: TcxButton;
    dxLayoutControl2Item1: TdxLayoutItem;
    dxLayoutControl3Group_Root: TdxLayoutGroup;
    dxLayoutControl3: TdxLayoutControl;
    dxLayoutControl3Item1: TdxLayoutItem;
    btnExcel: TcxButton;
    btnGuardar: TcxButton;
    dxLayoutControl3Item2: TdxLayoutItem;
    btnCancelar: TcxButton;
    dxLayoutControl3Item3: TdxLayoutItem;
    cMmCelda: TcxMemo;
    dxLayoutControl3Item4: TdxLayoutItem;
    pmActa: TPopupMenu;
    mniRecalcular: TMenuItem;
    DbTxtEdtCentro: TcxDBTextEdit;
    dxLayoutControl1Item1: TdxLayoutItem;
    dxLayoutControl1Group1: TdxLayoutAutoCreatedGroup;
    procedure BrPrincipalbtnExitClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure BrPrincipalbtnAddClick(Sender: TObject);
    procedure BrPrincipalbtnEditClick(Sender: TObject);
    procedure BrPrincipalbtnPostClick(Sender: TObject);
    procedure BrPrincipalbtnCancelClick(Sender: TObject);
    procedure BrPrincipalbtnDeleteClick(Sender: TObject);
    procedure BrPrincipalbtnRefreshClick(Sender: TObject);
    procedure pmDatosPopup(Sender: TObject);
    procedure mniAddClick(Sender: TObject);
    procedure BrPrincipalbtnPrinterClick(Sender: TObject);
    procedure RptActaGetValue(const VarName: string; var Value: Variant);
    procedure DbLkpCmbFolioKeyPress(Sender: TObject; var Key: Char);
    procedure DbTxtEdtActaKeyPress(Sender: TObject; var Key: Char);
    procedure DbDtEdtFechaKeyPress(Sender: TObject; var Key: Char);
    procedure DbTxtEdtEspecialidadKeyPress(Sender: TObject; var Key: Char);
    procedure DbLkpCmbFolioEnter(Sender: TObject);
    procedure DbLkpCmbFolioExit(Sender: TObject);
    procedure DbTxtEdtActaEnter(Sender: TObject);
    procedure DbTxtEdtActaExit(Sender: TObject);
    procedure DbDtEdtFechaEnter(Sender: TObject);
    procedure DbDtEdtFechaExit(Sender: TObject);
    procedure DbTxtEdtEspecialidadEnter(Sender: TObject);
    procedure DbTxtEdtEspecialidadExit(Sender: TObject);
    procedure DbMmObservacionEnter(Sender: TObject);
    procedure DbMmObservacionExit(Sender: TObject);
    procedure DbLkpCmbFolioPropertiesCloseUp(Sender: TObject);
    procedure DbTxtEdtActivoEnter(Sender: TObject);
    procedure DbTxtEdtActivoExit(Sender: TObject);
    procedure DbTxtEdtActivoKeyPress(Sender: TObject; var Key: Char);
    procedure CxPageDetalleChange(Sender: TObject);
    procedure CxGrdDbTblVMaterialesColumn6PropertiesButtonClick(Sender: TObject;
      AButtonIndex: Integer);
    procedure CxGrdDbTblVMaterialesColumn2PropertiesCloseUp(Sender: TObject);
    procedure QMaterialesAfterInsert(DataSet: TDataSet);
    procedure QMaterialesBeforePost(DataSet: TDataSet);
    procedure QActaAfterScroll(DataSet: TDataSet);
    procedure CxPageDetallePageChanging(Sender: TObject; NewPage: TcxTabSheet;
      var AllowChange: Boolean);
    procedure dlgSaveGuardarTypeChange(Sender: TObject);
    procedure btnExcelClick(Sender: TObject);
    procedure btnAjustarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnGuardarClick(Sender: TObject);
    procedure mniRecalcularClick(Sender: TObject);
    procedure SprShBkDatosActiveCellChanging(Sender: TcxSSBookSheet;
      const ActiveCell: TPoint; var CanSelect: Boolean);
    procedure cMmCeldaKeyPress(Sender: TObject; var Key: Char);
    procedure cMmCeldaEnter(Sender: TObject);
    procedure cMmCeldaExit(Sender: TObject);
    procedure DbTxtEdtCentroEnter(Sender: TObject);
    procedure DbTxtEdtCentroExit(Sender: TObject);
    procedure DbTxtEdtCentroKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
    Procedure GeneraActaEntrega_PDF(RTipo:FtTipo;RImpresion:FtSeccion;sAplica:string='');
    Procedure GeneraActaEntrega_Ex(RTipo:FtTipo;RImpresion:FtSeccion;sAplica:string='');
    procedure LoadFromDataset(Hoja:TcxssBookSheet;Datos:TzQuery;sPda:String;Modo:TypeDatos);
    procedure ActaExLoad(Libro:TcxSpreadSheetBook;Datos:TzReadOnlyQuery;sPda:String='');
    procedure ActaExSave(Libro:TcxSpreadSheetBook;Datos:TzReadOnlyQuery;sPda:String='');
    procedure LockUnLock(Hoja:TcxssBookSheet;ColI,ColT,RowI,NumRows:Integer;Lock:Boolean=true);
    procedure RangeMerge(Hoja:TcxssBookSheet;ColI,ColT,RowI,NumRows:Integer);
    procedure CrearFormato(Hoja:TcxssBookSheet);
    procedure FormatoCelda(ACellObj: TcxSSCellObject;Formato:TypeFormat=FrNone);
    procedure BordeCelda(ACellObj: TcxSSCellObject);   overload;
    procedure BordeCelda(Hoja:TcxssBookSheet;ColI,ColT,RowI,NumRows:Integer);   overload;
    procedure AlineacionCelda(Hoja:TcxssBookSheet;AlHorz:TcxHorzTextAlign;AlVert:TcxVertTextAlign;ColI,ColT,RowI,NumRows:Integer);
  public
    { Public declarations }
    var
      tmpCelda:string;
  end;

var
  FrmNotaCampo: TFrmNotaCampo;

implementation

uses frm_connection, UnitExcepciones, global, Utilerias, UnitValidaTexto,
  Frm_Materiales, UnitExcel, masUtilerias;

{$R *.dfm}

procedure TFrmNotaCampo.BordeCelda(ACellObj: TcxSSCellObject);
begin
  ACellObj.Style.Borders.Left.Style:=lsMedium;
  ACellObj.Style.Borders.Right.Style:=lsMedium;
  ACellObj.Style.Borders.Top.Style:=lsMedium;
  ACellObj.Style.Borders.Bottom.Style:=lsMedium;
end;

procedure TFrmNotaCampo.BordeCelda(Hoja:TcxssBookSheet;ColI,ColT,RowI,NumRows:Integer);
var
  ACellObj: TcxSSCellObject;
  iCol,iRow:Integer;
begin
  for iCol := ColI to ColT do
    for IRow := RowI to RowI + NumRows  do
    begin
      ACellObj:=Hoja.GetCellObject(iCol,IRow);
      ACellObj.Style.Borders.Left.Style:=lsThin;
      ACellObj.Style.Borders.Right.Style:=lsThin;
      ACellObj.Style.Borders.Top.Style:=lsThin;
      ACellObj.Style.Borders.Bottom.Style:=lsThin;
    end;

end;

procedure TFrmNotaCampo.AlineacionCelda(Hoja:TcxssBookSheet;AlHorz:TcxHorzTextAlign;AlVert:TcxVertTextAlign;ColI,ColT,RowI,NumRows:Integer);
var
  ACellObj: TcxSSCellObject;
  iCol,iRow:Integer;
begin
{TcxHorzTextAlign = (haGENERAL, haLEFT, haCENTER, haRIGHT, haFILL, haJUSTIFY);
  TcxVertTextAlign = (vaTOP, vaCENTER, vaBOTTOM, vaJUSTIFY);
}
  for iCol := ColI to ColT do
    for IRow := RowI to RowI + NumRows  do
    begin
      ACellObj:=Hoja.GetCellObject(iCol,IRow);
      ACellObj.Style.HorzTextAlign := AlHorz;
      ACellObj.Style.VertTextAlign := AlVert;
    end;

end;

procedure TFrmNotaCampo.FormatoCelda(ACellObj: TcxSSCellObject;Formato:TypeFormat=FrNone);
begin
  //(FrTitle,FrSubTitle,FrContent,FrNone);
  //ACellObj.Style.Brush.BackgroundColor := 5;
  ACellObj.Style.Font.Name := 'Arial';
  ACellObj.Style.Font.Size := 8;
  //ACellObj.Style.Font.FontColor := 2;


  case Formato of
    FrTitle:    begin
                  ACellObj.Style.HorzTextAlign := haCENTER;
                  ACellObj.Style.Font.Style:=[fsBold];
                  ACellObj.Style.Font.Size := 8;
                end;
    FrSubTitle: begin
                  ACellObj.Style.HorzTextAlign := haCENTER;
                //ACellObj.Style.Font.Style:=[fsBold];
                end;
    FrContent:  begin
                  ACellObj.Style.HorzTextAlign := haCENTER;
                  ACellObj.Style.Font.Size := 8;
                //ACellObj.Style.Font.Style:=[fsBold];
                end;
  end;

end;
procedure TFrmNotaCampo.CrearFormato(Hoja:TcxssBookSheet);
var
  ACellObj: TcxSSCellObject;
  IRow:Byte;
  NPdas:Integer;
begin
  Hoja.ClearAll;
  NPdas:=0;
  with Hoja do
  begin
    Cols.Size[7]:=Cols.Size[7] - 5;
    IRow:=NPdas;
    ACellObj := GetCellObject(0, IRow);
    FormatoCelda(ACellObj,FrTitle);
    ACellObj.SetCellText('PARTIDA');
    ACellObj := GetCellObject(2, IRow);
    FormatoCelda(ACellObj,FrTitle);
    ACellObj.SetCellText('ACTIVIDAD');

    GetCellObject(2, IRow+1).Style.WordBreak:=True;
    Rows.Size[IRow+1] := 40;

    inc(IRow,3);
    ACellObj := GetCellObject(1, IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('PERIODOS DE EJECUCION DE LA ACTIVIDAD');
    //IRow:=5;
    Inc(IRow);
    ACellObj := GetCellObject(1, IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('FECHA');
    ACellObj := GetCellObject(2,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('INICIO');
    ACellObj := GetCellObject(3, IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('TERMINO');
    ACellObj := GetCellObject(4, IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('AFECTACIÓN');
    ACellObj := GetCellObject(5, IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.Style.WordBreak:=True;
    ACellObj.SetCellText('INTERVALO TIEMPO');
    ACellObj := GetCellObject(6, IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.Style.WordBreak:=True;
    ACellObj.SetCellText('AVANCE ANTERIOR');
    ACellObj := GetCellObject(7, IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.Style.WordBreak:=True;
    ACellObj.SetCellText('AVANCE ACTUAL');
    ACellObj := GetCellObject(8,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.Style.WordBreak:=True;
    ACellObj.SetCellText('AVANCE ACUMULADO');
    Rows.Size[IRow] :=30;

    inc(IRow,2);
    //IRow:=7;
    ACellObj := GetCellObject(1, IRow);
    FormatoCelda(ACellObj,FrTitle);
    ACellObj.SetCellText('DURACION TIEMPO EFECTIVO (HRS):');


    inc(iRow,2);
    //IRow:=9;
    ACellObj := GetCellObject(0,IRow);
    FormatoCelda(ACellObj,FrTitle);
    ACellObj.SetCellText('PERSONAL');
    inc(IRow);
    //IRow:=10;
    ACellObj := GetCellObject(0,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('PARTIDA');
    ACellObj := GetCellObject(1,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('DESCRIPCIÓN');
    ACellObj := GetCellObject(5,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('UNIDAD');
    ACellObj := GetCellObject(6,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('CANTIDAD');
    ACellObj := GetCellObject(7,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('PU MN');
    ACellObj := GetCellObject(8,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('PU USD');
    ACellObj := GetCellObject(9,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('IMP MN');
    ACellObj := GetCellObject(10,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('IMP USD');
    inc(iRow,2);
    //IRow:=12;
    ACellObj := GetCellObject(6, IRow);
    FormatoCelda(ACellObj,FrTitle);
    ACellObj.SetCellText('IMPORTE PERSONAL:');

    inc(IRow,2);
    //IRow:=14;
    ACellObj := GetCellObject(0,IRow);
    FormatoCelda(ACellObj,FrTitle);
    ACellObj.SetCellText('EQUIPO');
    inc(IRow);
    //IRow:=15;
    ACellObj := GetCellObject(0,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('PARTIDA');
    ACellObj := GetCellObject(1,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('DESCRIPCIÓN');
    ACellObj := GetCellObject(5,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('UNIDAD');
    ACellObj := GetCellObject(6,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('CANTIDAD');
    ACellObj := GetCellObject(7,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('PU MN');
    ACellObj := GetCellObject(8,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('PU USD');
    ACellObj := GetCellObject(9,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('IMP MN');
    ACellObj := GetCellObject(10,IRow);
    FormatoCelda(ACellObj,FrSubTitle);
    ACellObj.SetCellText('IMP USD');
    inc(IRow,2);
    //IRow:=17;
    ACellObj := GetCellObject(6, IRow);
    FormatoCelda(ACellObj,FrTitle);
    ACellObj.SetCellText('IMPORTE EQUIPO:');

    IRow:=NPdas;
    LockUnLock(Hoja,0,1,IRow,0,False);
    RangeMerge(Hoja,0,1,IRow,0);
    LockUnLock(Hoja,0,1,IRow,0);

    LockUnLock(Hoja,2,10,IRow,0,False);
    RangeMerge(Hoja,2,10,IRow,0);
    LockUnLock(Hoja,2,10,IRow,0);

    Inc(IRow);
    LockUnLock(Hoja,0,1,IRow,0,False);
    RangeMerge(Hoja,0,1,IRow,0);
    LockUnLock(Hoja,0,1,IRow,0);

    LockUnLock(Hoja,2,10,IRow,0,False);
    RangeMerge(Hoja,2,10,IRow,0);
    LockUnLock(Hoja,2,10,IRow,0);

    Inc(IRow,2);
    LockUnLock(Hoja,1,8,IRow,0,False);
    RangeMerge(Hoja,1,8,IRow,0);
    LockUnLock(Hoja,1,8,IRow,0);


    Inc(IRow,3);
    LockUnLock(Hoja,1,4,IRow,0,False);
    RangeMerge(Hoja,1,4,IRow,0);
    LockUnLock(Hoja,1,4,IRow,0);

    Inc(IRow,2);
    LockUnLock(Hoja,0,10,IRow,0,False);
    RangeMerge(Hoja,0,10,IRow,0);
    LockUnLock(Hoja,0,10,IRow,0);

    Inc(IRow);
    LockUnLock(Hoja,1,4,IRow,0,False);
    RangeMerge(Hoja,1,4,IRow,0);

    Inc(IRow);
    LockUnLock(Hoja,1,4,IRow,0,False);
    RangeMerge(Hoja,1,4,IRow,0);

    Inc(IRow);
    LockUnLock(Hoja,6,8,IRow,0,False);
    RangeMerge(Hoja,6,8,IRow,0);
    LockUnLock(Hoja,6,8,IRow,0);

    IRow:=NPdas;
    BordeCelda(Hoja,0,10,IRow,1);

   // IRow:=3;
    Inc(IRow,3);
    BordeCelda(Hoja,1,8,IRow,2);



    //IRow:=5;
    Inc(IRow,3);
    BordeCelda(Hoja,1,5,IRow,0);

    //IRow:=7;
    Inc(IRow,2);
    BordeCelda(Hoja,0,10,IRow,2);

   // IRow:=10;
    Inc(IRow,3);
    BordeCelda(Hoja,6,10,IRow,0);

    {IIRow:=12;
    BordeCelda(Hoja,0,10,IRow,2);
    IRow:=15;
    BordeCelda(Hoja,6,10,IRow,0);

    nc(IRow,3);
    LockUnLock(Hoja,1,4,IRow,0,False);
    RangeMerge(Hoja,1,4,IRow,0);
    LockUnLock(Hoja,1,4,IRow,0);






    LockUnLock(Hoja,0,10,12,0,False);
    RangeMerge(Hoja,0,10,12,0);
    LockUnLock(Hoja,0,10,12,0);

    LockUnLock(Hoja,1,4,13,0,False);
    RangeMerge(Hoja,1,4,13,0);
    LockUnLock(Hoja,1,4,13,0);

    LockUnLock(Hoja,1,4,14,0,False);
    RangeMerge(Hoja,1,4,14,0);

    LockUnLock(Hoja,6,8,15,0,False);
    RangeMerge(Hoja,6,8,15,0);
    LockUnLock(Hoja,6,8,15,0);

    }

  end;


end;

procedure TFrmNotaCampo.CxGrdDbTblVMaterialesColumn2PropertiesCloseUp(
  Sender: TObject);
begin
   if QMateriales.State in [dsInsert,dsEdit] then
  begin
    QMateriales.FieldByName('sIdInsumo').AsString:=QrInsumos.FieldByName('sIdInsumo').AsString;
    QMateriales.FieldByName('sMedida').AsString:=QrInsumos.FieldByName('sMedida').AsString;
    QMateriales.FieldByName('sTrazabilidad').AsString:=QrInsumos.FieldByName('sTrazabilidad').AsString;
    QMateriales.FieldByName('mDescripcion').AsString:=QrInsumos.FieldByName('mDescripcion').AsString;
  end;
end;

procedure TFrmNotaCampo.CxGrdDbTblVMaterialesColumn6PropertiesButtonClick(
  Sender: TObject; AButtonIndex: Integer);
begin
  //ShowMessage('Aparece Pantalla de Material');
  if QMateriales.State in [dsInsert,dsEdit] then
  begin
    Application.CreateForm(TFrmAltaMAterial,FrmAltaMAterial);
    try
      FrmAltaMAterial.ShowModal;
      if FrmAltaMAterial.SeInserto then
      begin
        QrInsumos.Refresh;
        if QMateriales.State=dsInsert then
        begin
          QMateriales.FieldByName('sIdInsumo').AsString:=FrmAltaMAterial.QInsumos.FieldByName('sIdInsumo').AsString;
          QMateriales.FieldByName('sMedida').AsString:=FrmAltaMAterial.QInsumos.FieldByName('sMedida').AsString;
          QMateriales.FieldByName('sTrazabilidad').AsString:=FrmAltaMAterial.QInsumos.FieldByName('sTrazabilidad').AsString;
          QMateriales.FieldByName('mDescripcion').AsString:=FrmAltaMAterial.QInsumos.FieldByName('mDescripcion').AsString;
        end;
      end;

    finally
      FrmAltaMAterial.Destroy;
    end;
  end;

end;

procedure TFrmNotaCampo.CxPageDetalleChange(Sender: TObject);
begin
  QrActividades.Active:=False;
  QrInsumos.Active:=False;
  QMateriales.Active:=False;
  if (CxPageDetalle.ActivePageIndex=1) and (QActa.RecordCount>0) then
  begin
    with QrActividades do
    begin
      ParamByName('Contrato').AsString:=QActa.FieldByName('sContrato').AsString;
      ParamByName('Orden').AsString:=QActa.FieldByName('sNumeroOrden').AsString;
      Open;
    end;

    with QrInsumos do
    begin
      ParamByName('Contrato').AsString:=QActa.FieldByName('sContrato').AsString;
      Open;
    end;

    with QMateriales do
    begin
      ParamByName('Contrato').AsString:=QActa.FieldByName('sContrato').AsString;
      ParamByName('Acta').AsInteger:=QActa.FieldByName('iIdActa').AsInteger;
      Open;
    end;

  end;
end;

procedure TFrmNotaCampo.CxPageDetallePageChanging(Sender: TObject;
  NewPage: TcxTabSheet; var AllowChange: Boolean);
begin
  AllowChange:=True;
  if QActa.State in [dsInsert,dsEdit] then
    AllowChange:=False;
end;

procedure TFrmNotaCampo.DbDtEdtFechaEnter(Sender: TObject);
begin
  DbDtEdtFecha.Style.Color:=global_color_Entrada;
end;

procedure TFrmNotaCampo.DbDtEdtFechaExit(Sender: TObject);
begin
   DbDtEdtFecha.Style.Color:=global_color_Salida;
end;

procedure TFrmNotaCampo.DbDtEdtFechaKeyPress(Sender: TObject; var Key: Char);
begin
  If key = #13 then
    DbTxtEdtEspecialidad.SetFocus;
end;

procedure TFrmNotaCampo.DbLkpCmbFolioEnter(Sender: TObject);
begin

  DbLkpCmbFolio.Style.Color:=global_color_entrada;
end;

procedure TFrmNotaCampo.DbLkpCmbFolioExit(Sender: TObject);
begin
  DbLkpCmbFolio.Style.Color:=global_color_Salida;
end;

procedure TFrmNotaCampo.DbLkpCmbFolioKeyPress(Sender: TObject; var Key: Char);
begin
  If key = #13 then
    DbTxtEdtActa.SetFocus
end;

procedure TFrmNotaCampo.DbLkpCmbFolioPropertiesCloseUp(Sender: TObject);
begin
  if QActa.State=dsInsert then
  begin
    QActa.FieldByName('sActivo').AsString:=QrFolios.FieldByName('sUbicacion').AsString;
   

  end;
end;

procedure TFrmNotaCampo.DbMmObservacionEnter(Sender: TObject);
begin
  DbMmObservacion.Style.Color:=global_color_Entrada;
end;

procedure TFrmNotaCampo.DbMmObservacionExit(Sender: TObject);
begin
  DbMmObservacion.Style.Color:=global_color_Salida;
end;

procedure TFrmNotaCampo.DbTxtEdtActaEnter(Sender: TObject);
begin
  DbTxtEdtActa.Style.Color:=global_color_Entrada;
end;

procedure TFrmNotaCampo.DbTxtEdtActaExit(Sender: TObject);
begin
  DbTxtEdtActa.Style.Color:=global_color_Salida;
end;

procedure TFrmNotaCampo.DbTxtEdtActaKeyPress(Sender: TObject; var Key: Char);
begin
  If key = #13 then
    DbDtEdtFecha.SetFocus;
end;

procedure TFrmNotaCampo.DbTxtEdtActivoEnter(Sender: TObject);
begin
  DbTxtEdtActivo.Style.Color:=global_color_Entrada;
end;

procedure TFrmNotaCampo.DbTxtEdtActivoExit(Sender: TObject);
begin
  DbTxtEdtActivo.Style.Color:=global_color_Salida;
end;

procedure TFrmNotaCampo.DbTxtEdtActivoKeyPress(Sender: TObject; var Key: Char);
begin
  If key = #13 then
    DbMmObservacion.SetFocus;
end;

procedure TFrmNotaCampo.DbTxtEdtCentroEnter(Sender: TObject);
begin
  DbTxtEdtCentro.Style.Color:=global_color_Entrada;
end;

procedure TFrmNotaCampo.DbTxtEdtCentroExit(Sender: TObject);
begin
    DbTxtEdtCentro.Style.Color:=global_color_Salida;
end;

procedure TFrmNotaCampo.DbTxtEdtCentroKeyPress(Sender: TObject; var Key: Char);
begin
  If key = #13 then
    DbTxtEdtActivo.SetFocus;
end;

procedure TFrmNotaCampo.DbTxtEdtEspecialidadEnter(Sender: TObject);
begin
   DbTxtEdtEspecialidad.Style.Color:=global_color_Entrada;
end;

procedure TFrmNotaCampo.DbTxtEdtEspecialidadExit(Sender: TObject);
begin
  DbTxtEdtEspecialidad.Style.Color:=global_color_Salida;
end;

procedure TFrmNotaCampo.DbTxtEdtEspecialidadKeyPress(Sender: TObject;
  var Key: Char);
begin
    If key = #13 then
    DbTxtEdtCentro.SetFocus;
end;

procedure TFrmNotaCampo.dlgSaveGuardarTypeChange(Sender: TObject);
begin
  case dlgSaveGuardar.FilterIndex of
    1 : dlgSaveGuardar.DefaultExt := 'xls';
  else
    dlgSaveGuardar.DefaultExt := '';
  end;
end;

procedure TFrmNotaCampo.RangeMerge(Hoja:TcxssBookSheet;ColI,ColT,RowI,NumRows:Integer);
var
  Rango:TRect;
  iCol,iRow,i:Integer;
begin
  rango.Left:=ColI;
  Rango.Right:=ColT;
  Rango.Top:= RowI;
  Rango.Bottom:=RowI + NumRows;
  Hoja.SetMergedState(Rango,true);
end;

procedure TFrmNotaCampo.RptActaGetValue(const VarName: string;
  var Value: Variant);
begin

  if CompareText(VarName, 'PERSONALMN') = 0 then
    Value :=Montos[1,1];

  if CompareText(VarName, 'PERSONALDLL') = 0 then
    Value :=Montos[1,2];

  if CompareText(VarName, 'EQUIPOMN') = 0 then
    Value :=Montos[2,1];

  if CompareText(VarName, 'EQUIPODLL') = 0 then
    Value :=Montos[2,2];

  if CompareText(VarName, 'BARCOMN') = 0 then
    Value :=Montos[3,1];

  if CompareText(VarName, 'BARCODLL') = 0 then
    Value :=Montos[3,2];

  if CompareText(VarName, 'PERNOCTAMN') = 0 then
    Value :=Montos[4,1];

  if CompareText(VarName, 'PERNOCTADLL') = 0 then
    Value :=Montos[4,2];


  {[PERSONALMN + EQUIPOMN]
        Montos[i,1]:=Montos[i,1] + dImporteMn;
        Montos[i,2]:=Montos[i,1] + dImporteDll; }
end;

procedure TFrmNotaCampo.SprShBkDatosActiveCellChanging(Sender: TcxSSBookSheet;
  const ActiveCell: TPoint; var CanSelect: Boolean);
begin

    cMmCelda.Text:=SprShBkDatos.ActiveSheet.GetCellObject(ActiveCell.X,ActiveCell.Y).Text;
    cMmCelda.Properties.ReadOnly:=SprShBkDatos.ActiveSheet.GetCellObject(ActiveCell.X,ActiveCell.Y).Style.Locked;
end;

procedure TFrmNotaCampo.LockUnLock(Hoja:TcxssBookSheet;ColI,ColT,RowI,NumRows:Integer;Lock:Boolean=true);
var
  iCol,iRow:Integer;
begin
  for iCol := ColI to ColT do
    for IRow := RowI to RowI + NumRows  do
      Hoja.GetCellObject(iCol,IRow).Style.Locked:=Lock ;
      //SprShBkDatos.ActiveSheet.GetCellObject(i,SprShBkDatos.ActiveSheet.row).Style.Borders:=
end;

procedure TFrmNotaCampo.BrPrincipalbtnAddClick(Sender: TObject);
var
  Firma1,Firma2:string;
begin
  try
    CxPageDetalle.ActivePageIndex:=0;
    if QActa.RecordCount>0 then
    begin
      Firma1:=QActa.FieldByName('sFirma1').AsString;
      Firma2:=QActa.FieldByName('sFirma2').AsString;
    end;


    QActa.Append;
    QActa.FieldByName('sContrato').AsString:=global_contrato;
    QActa.FieldByName('iIdActa').AsInteger:=0;
    QActa.FieldByName('dFecha').AsDateTime:=Now;
    QActa.FieldByName('eTipo').AsString:='Parcial';
    if connection.contrato.FieldByName('eLugarOt').AsString='Tierra' then
    begin
      QActa.FieldByName('lPernocta').AsString:='No';
      QActa.FieldByName('lPaginado').AsString:='No';
    end
    else
    begin
      QActa.FieldByName('lPernocta').AsString:='Si';
      QActa.FieldByName('lPaginado').AsString:='Si';
      QActa.FieldByName('eTipo').AsString:='Total';
    end;

    QActa.FieldByName('lMaterial').AsString:='No';
    QActa.FieldByName('lPartidas').AsString:='No';
    QActa.FieldByName('sFirma1').AsString:=Firma1;
    QActa.FieldByName('sFirma2').AsString:=Firma2;
    QActa.FieldByName('sCentroProceso').AsString:='';
    DbLkpCmbFolio.SetFocus;
    BrPrincipal.btnAddClick(Sender);
    btnAjustar.Enabled:=False;
  except
    on e : exception do
      UnitExcepciones.manejarExcep(E.Message, E.ClassName,Self.Caption, 'Al agregar nuevo registro', 0);
  end;
end;


//TypeDatos=(ctNew,ctLoad);
procedure TFrmNotaCampo.LoadFromDataset(Hoja:TcxssBookSheet;Datos:TzQuery;sPda:String;Modo:TypeDatos);
const
  SQlRef: array[1..2,1..3] of string=(('bitacoradepersonal','personal','sIdPersonal'),('bitacoradeequipos','equipos','sIdEquipo'));
var
  Row,Col,i:Integer;
  CampoName:string;
  CampoOrigen:TField;
  QrActividad,
  QrAvance,
  QrRecursos:TZReadOnlyQuery;
  TotalHrs:string;
begin
  if Datos.RecordCount>0 then
  begin
    QrActividad:=TZReadOnlyQuery.Create(nil);
    QrAvance:=TZReadOnlyQuery.Create(nil);
    QrRecursos:=TZReadOnlyQuery.Create(nil);
    try
      QrActividad.Connection:=connection.zConnection;
      QrAvance.Connection:=connection.zConnection;
      QrRecursos.Connection:=connection.zConnection;

      QrActividad.SQL.Text:='select * from actividadesxorden where scontrato=:Contrato and sNUmeroOrden=:Orden and sNumeroActividad=:Actividad';
      QrActividad.ParamByName('Contrato').AsString:=Datos.FieldByName('sContrato').AsString;
      QrActividad.ParamByName('Orden').AsString:=Datos.FieldByName('sNumeroOrden').AsString;
      QrActividad.ParamByName('Actividad').AsString:=sPda;
      QrActividad.Open;
      if QrActividad.RecordCount=1 then
      begin
        Hoja.GetCellObject(0,1).SetCellText(QrActividad.FieldByName('sNumeroActividad').AsString);
        Hoja.GetCellObject(2,1).SetCellText(QrActividad.FieldByName('mDescripcion').AsString);
       

        Row:=4;
        QrAvance.Connection:=connection.zConnection;
        QrAvance.SQL.Text:= 'select b.*, ' +
                            '( SELECT (ifnull(sum(ba.dAvance), 0)) ' +
                                              '		FROM ' +
                                              '			bitacoradeactividades AS ba ' +
                                              '		WHERE ' +
                                              '			ba.sContrato = b.sContrato ' +
                                              '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                              '		AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento ' +
                                              '		AND ba.swbs = b.swbs ' +
                                              '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                              '		AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
                                              '   < cast(b.sHoraInicio AS Time))  )	) AS AvanceAnterior ' +
                            ' from bitacoradeactividades b' + #13#10 +
                            'where b.sContrato=:Contrato and b.snumeroorden=:Orden and b.sNumeroActividad=:Actividad' + #13#10 +
                            'and b.sIdTipoMovimiento=:Tipo' + #13#10 +
                            'order by b.didfecha,time(b.sHoraInicio)' ;
        QrAvance.ParamByName('contrato').AsString:=Datos.FieldByName('sContrato').AsString;
        QrAvance.ParamByName('Orden').AsString:=Datos.FieldByName('sNumeroOrden').AsString;
        QrAvance.ParamByName('Actividad').AsString:=QrActividad.FieldByName('sNumeroActividad').AsString;
        QrAvance.ParamByName('tipo').AsString:='ED';
        QrAvance.Open;
        TotalHrs:='00:00';
        while not QrAvance.Eof do
        begin
          if QrAvance.RecordCount<>QrAvance.RecNo then
          begin
            LockUnLock(Hoja,1,8,Row,0,false);
            Hoja.SelectCell(1,Row);
            Hoja.InsertCells(Hoja.SelectionRect,msallRow );
          end
          else
            LockUnLock(Hoja,1,8,Row,0,true);
          //Hoja.GetCellObject(1,Row).SetCellText(QrAvance.FieldByName('dIdFecha').AsString);
          Hoja.GetCellObject(1,Row).DateTime:= QrAvance.FieldByName('dIdFecha').AsDateTime;
          Hoja.GetCellObject(2,Row).SetCellText(QrAvance.FieldByName('sHoraInicio').AsString);
          Hoja.GetCellObject(3,Row).SetCellText(QrAvance.FieldByName('sHoraFinal').AsString);
          Hoja.GetCellObject(4,Row).SetCellText(QrAvance.FieldByName('sIdTipoMovimiento').AsString);
          Hoja.GetCellObject(5,Row).SetCellText(sfnRestaHoras(QrAvance.FieldValues['sHoraFinal'], QrAvance.FieldValues['sHoraInicio']));
          Hoja.GetCellObject(6,Row).SetCellText(FormatFloat( '0.00',QrAvance.FieldByName('AvanceAnterior').asfloat)  + '%');
          Hoja.GetCellObject(7,Row).SetCellText(FormatFloat( '0.00',QrAvance.FieldByName('dAvance').AsFloat) + '%');
          Hoja.GetCellObject(8,Row).SetCellText(FormatFloat( '0.00',QrAvance.FieldByName('dAvance').AsFloat + QrAvance.FieldByName('AvanceAnterior').AsFloat) + '%');

          TotalHrs:=sfnSumaHoras(TotalHrs,sfnRestaHoras(QrAvance.FieldByName('sHoraFinal').AsString, QrAvance.FieldByName('sHoraInicio').AsString));
          QrAvance.Next;
          BordeCelda(Hoja,1,8,Row,0);
          inc(row);
        end;

        if QrAvance.RecordCount>0 then
        begin
         // inc(row);
          LockUnLock(Hoja,5,5,Row,0,false);
          FormatoCelda(Hoja.GetCellObject(5,Row),FrTitle);
          Hoja.GetCellObject(5,Row).SetCellText(TotalHrs);
          AlineacionCelda(Hoja,haCENTER,vaCENTER,1,8,4,Row-4);
          LockUnLock(Hoja,5,5,Row,0);
          Inc(Row,4)
        end
        else
          Inc(Row,5);

        for I := 1 to Length(SQlRef) do
        begin
          QrRecursos.Active:=False;
          QrRecursos.SQL.Text:= 'select br.' +SQlRef[i,3] + ' as sIdRecurso,br.sDescripcion,r.sMedida,br.dCanthh,r.dVentaMn,r.dVentaDll ' +
                                'from '+ SQlRef[i,1]  + ' br ' +
                                'left join ' + SQlRef[i,2] + ' r ' +
                                'on(r.sContrato=:Contrato and br.'+SQlRef[i,3]+'=r.'+ SQlRef[i,3] +') ' +
                                'where br.sContrato=:Orden and br.sNumeroOrden=:Folio and br.sNumeroActividad=:Actividad' ;
          QrRecursos.ParamByName('contrato').AsString:=global_Contrato_Barco;
          QrRecursos.ParamByName('Orden').AsString:=Datos.FieldByName('sContrato').AsString;
          QrRecursos.ParamByName('Folio').AsString:=Datos.FieldByName('sNumeroOrden').AsString;
          QrRecursos.ParamByName('Actividad').AsString:=QrActividad.FieldByName('sNumeroActividad').AsString;
          QrRecursos.Open;

          while not QrRecursos.Eof do
          begin
           // if QrRecursos.RecordCount<>QrRecursos.RecNo then
           // begin
              //LockUnLock(Hoja,1,8,Row,0,false);
              LockUnLock(Hoja,0,10,Row,0,false);
              Hoja.SelectCell(0,Row);
              Hoja.InsertCells(Hoja.SelectionRect,msallRow );
              RangeMerge(Hoja,1,4,Row,0);
              LockUnLock(Hoja,0,10,Row,0,false);
           // end;
            //else
            //  LockUnLock(Hoja,1,8,Row,0,true);
            Hoja.GetCellObject(0,Row).SetCellText(QrRecursos.FieldByName('sIdRecurso').AsString);
            Hoja.GetCellObject(1,Row).SetCellText(QrRecursos.FieldByName('sDescripcion').AsString);
            Hoja.GetCellObject(5,Row).SetCellText(QrRecursos.FieldByName('sMedida').AsString);
            Hoja.GetCellObject(6,Row).SetCellText(FormatFloat( '0.00',QrRecursos.FieldByName('dCanthh').Asfloat));
            Hoja.GetCellObject(7,Row).SetCellText(FormatFloat( '#,##0.00',QrRecursos.FieldByName('dVentaMn').AsFloat));
            AlineacionCelda(Hoja,haRIGHT,vaCENTER,7,7,Row,0);
            Hoja.GetCellObject(8,Row).SetCellText(FormatFloat( '#,##0.00',QrRecursos.FieldByName('dVentaDll').AsFloat));
            Hoja.GetCellObject(9,Row).SetCellText(FormatFloat( '#,##0.00',QrRecursos.FieldByName('dCanthh').AsFloat *  QrRecursos.FieldByName('dVentaMn').AsFloat));
            Hoja.GetCellObject(10,Row).SetCellText(FormatFloat( '#,##0.00',QrRecursos.FieldByName('dCanthh').AsFloat * QrRecursos.FieldByName('dVentaDll').AsFloat));
            {Hoja.GetCellObject(4,Row).SetCellText(QrAvance.FieldByName('sIdTipoMovimiento').AsString);
            Hoja.GetCellObject(5,Row).SetCellText(sfnRestaHoras(QrAvance.FieldValues['sHoraFinal'], QrAvance.FieldValues['sHoraInicio']));
            Hoja.GetCellObject(6,Row).SetCellText(QrAvance.FieldByName('AvanceAnterior').AsString);
            Hoja.GetCellObject(7,Row).SetCellText(QrAvance.FieldByName('dAvance').AsString);
            Hoja.GetCellObject(8,Row).SetCellText(FloatToStr(QrAvance.FieldByName('dAvance').AsFloat + QrAvance.FieldByName('AvanceAnterior').AsFloat));
                }
            QrRecursos.Next;
            BordeCelda(Hoja,0,10,Row,0);
            inc(row);
          end;

          if QrRecursos.RecordCount>0 then
            Inc(Row,5)
          else
            Inc(Row,6);


        end;
      end;

     //('bitacoradepersonal','personal','sIdPersonal'),('bitacoradeequipos','equipos','sIdEquipo'));


    finally
      QrActividad.Destroy;
      QrAvance.Destroy;
      QrRecursos.Destroy;
    end;

  end;

  {  while not Datos.Eof do
  begin
    for Col := 0 to HojaColumns.CountCols-1 do
    begin
      CampoOrigen:=nil;
      if HojaColumns.Cols[Col].CampoPos<>-1 then
        CampoOrigen := Datos.FindField(HojaColumns.Cols[Col].CampoName);    

      if Assigned(CampoOrigen) then
      begin
        if (Hoja.Cells[Row,HojaColumns.Cols[Col].CampoPos] = nil) then
          Hoja.CreateCell(Row,HojaColumns.Cols[Col].CampoPos);

        
        Hoja.Cells[Row,HojaColumns.Cols[Col].CampoPos].AsString:=CampoOrigen.AsString;
      end;
    end;

    Inc(Row);
    Datos.Next;
  end;    }
end;


//LoadFromDataset(Hoja:TcxssBookSheet;Datos:TzQuery;sPda:String;Modo:TypeDatos);
procedure TFrmNotaCampo.ActaExSave(Libro:TcxSpreadSheetBook;Datos:TzReadOnlyQuery;sPda:String='');
var
  Hoja:TcxssBookSheet;
  QImportes,
  QRecursos:TZQuery;
  QrActividades:TZReadOnlyQuery;
  i:Byte;
  ACellObj: TcxSSCellObject;
begin
  QImportes:=TZQuery.Create(nil);
  QRecursos:=TZQuery.Create(nil);
  QrActividades:=TZReadOnlyQuery.Create(nil);
  try
    QImportes.Connection:=connection.zConnection;
    QRecursos.Connection:=connection.zConnection;
    QrActividades.Connection:=connection.zConnection;

    QrActividades.SQL.Text:='select ao.* from actividadesxorden ao inner join acta_campo ac '+
                        'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                        'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                        'where ac.iIdActa=:Acta and ao.sTipoActividad=:Tipo ' +
                        'group by ao.swbs order by ao.iItemOrden';


    QImportes.SQL.Text:='select * from acta_campo where iIdActa=:Acta and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                        'sIdRecurso like "$IMPORTE%" order by iOrdenTipo';

    QRecursos.SQL.Text:='select * from acta_campo where iIdActa=:Acta and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                        'eTipo=:Tipo and sIdRecurso not like "$IMPORTE%" order by iOrdenRecurso';


    for I := 3 downto 2 do
    begin
      Hoja:=Libro.Pages[i];
      with Hoja do
      begin
        QrActividades.Active:=False;
        QrActividades.ParamByName('Acta').AsInteger:=Datos.FieldByName('iIdActa').AsInteger;
        if (I=2) then
          QrActividades.ParamByName('Tipo').AsString:='Actividad'
        else
          QrActividades.ParamByName('Tipo').AsString:='Paquete';
        QrActividades.Open;

        while not QrActividades.Eof do
        begin
          QImportes.Active:=False;
          QImportes.ParamByName('Acta').AsInteger:=Datos.FieldByName('iIdActa').AsInteger;
          QImportes.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
          QImportes.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
          QImportes.Open;
          while not QImportes.Eof do
          begin
            QRecursos.Active:=False;
            QRecursos.ParamByName('Acta').AsInteger:=Datos.FieldByName('iIdActa').AsInteger;
            QRecursos.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
            QRecursos.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
            QRecursos.ParamByName('Tipo').AsString:=QImportes.FieldByName('eTipo').AsString;
            QRecursos.Open;
            while not QRecursos.Eof do
            begin
              QRecursos.Edit;
              try
                ACellObj := GetCellObject(6,QRecursos.FieldByName('xRow').AsInteger-1);
                QRecursos.FieldByName('dCantidad').AsFloat:=ACellObj.CellValue;
                ACellObj := GetCellObject(9,QRecursos.FieldByName('xRow').AsInteger-1);
                QRecursos.FieldByName('sFormulaMn').AsString:=ACellObj.Text;
                QRecursos.FieldByName('dImporteMn').AsFloat:=ACellObj.CellValue;
                ACellObj := GetCellObject(10,QRecursos.FieldByName('xRow').AsInteger-1);
                QRecursos.FieldByName('sFormulaDll').AsString:=ACellObj.Text;
                QRecursos.FieldByName('dImporteDll').AsFloat:=ACellObj.CellValue;
                QRecursos.Post;
              except
                QRecursos.Cancel;
              end;
              QRecursos.Next;
            end;

            QImportes.Edit;
            try
              ACellObj := GetCellObject(9,QImportes.FieldByName('xRow').AsInteger-1);
              QImportes.FieldByName('sFormulaMn').AsString:=ACellObj.Text;
              QImportes.FieldByName('dImporteMn').AsFloat:=ACellObj.CellValue;
              ACellObj := GetCellObject(10,QImportes.FieldByName('xRow').AsInteger-1);
              QImportes.FieldByName('sFormulaDll').AsString:=ACellObj.Text;
              QImportes.FieldByName('dImporteDll').AsFloat:=ACellObj.CellValue;
              QImportes.Post;
            except
              QImportes.Cancel;
            end;
            QImportes.Next;
          end;
          QrActividades.Next;
        end;
      end;

    end;
  finally
    QImportes.Destroy;
    QRecursos.Destroy;
  end;

///Aqui se guardan los datos
end;

procedure TFrmNotaCampo.ActaExLoad(Libro:TcxSpreadSheetBook;Datos:TzReadOnlyQuery;sPda:String='');
var
  ACellObj: TcxSSCellObject;
  IRow,pIRow:Integer;
  NPdas:Integer;
  Hoja:TcxssBookSheet;
  i:Byte;
  QrActividades,QrAvance:TZReadOnlyQuery;
  QImportes,
  QRecursos:TZQuery;
  TotalHrs,sFormulaSumMn,sFormulaSumDll:string;
  posI:Integer;
  sItem,DesgloseMn,DesgloseDll:string;
begin

  QrActividades:=TZReadOnlyQuery.Create(nil);
  QrAvance:=TZReadOnlyQuery.Create(nil);
  QImportes:=TZQuery.Create(nil);
  QRecursos:=TZQuery.Create(nil);
  try
    QrActividades.Connection:=connection.zConnection;
    QImportes.Connection:=connection.zConnection;
    QRecursos.Connection:=connection.zConnection;
    QrAvance.Connection:=connection.zConnection;

    QrActividades.SQL.Text:='select ao.* from actividadesxorden ao inner join acta_campo ac '+
                            'on (ac.sContrato=ao.sContrato and ac.sNumeroOrden=ao.sNumeroOrden and ' +
                            'ac.swbs=ao.swbs and ac.sNumeroActividad=ao.sNumeroActividad) ' +
                            'where ac.iIdActa=:Acta and ao.sTipoActividad=:Tipo ' +
                            'group by ao.swbs order by ao.iItemOrden';

    QrAvance.Connection:=connection.zConnection;
    QrAvance.SQL.Text:= 'select b.*, ' +
                        '( SELECT (ifnull(sum(ba.dAvance), 0)) ' +
                                          '		FROM ' +
                                          '			bitacoradeactividades AS ba ' +
                                          '		WHERE ' +
                                          '			ba.sContrato = b.sContrato ' +
                                          '		AND ba.sNumeroOrden = b.sNumeroOrden ' +
                                          '		AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento ' +
                                          '		AND ba.swbs = b.swbs ' +
                                          '		AND ba.sNumeroActividad = b.sNumeroActividad ' +
                                          '		AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
                                          '   < cast(b.sHoraInicio AS Time))  )	) AS AvanceAnterior ' +
                        ' from bitacoradeactividades b' + #13#10 +
                        'where b.sContrato=:Contrato and b.snumeroorden=:Orden and b.sNumeroActividad=:Actividad' + #13#10 +
                        'and b.sIdTipoMovimiento=:Tipo' + #13#10 +
                        'order by b.didfecha,time(b.sHoraInicio)' ;



    QRecursos.SQL.Text:='select * from acta_campo where iIdActa=:Acta and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                        'eTipo=:Tipo and sIdRecurso not like "$IMPORTE%" and sAnexo=:Anexo order by iOrdenRecurso';

    for I := 3 downto 0 do
    begin
      Hoja:=Libro.Pages[i];

      if i=0 then
        Hoja.Caption:='ACTA DE ACTIVIDADES';

      if i=1 then
        Hoja.Caption:='COSTO POR ACTIVIDAD';

      if i=2 then
        Hoja.Caption:='CAMPO';

      if i=3 then
        Hoja.Caption:='DESGLOSE DE COSTOS';

      if i=0 then
        QImportes.SQL.Text:='select *,GROUP_CONCAT(xRow) as xRows from acta_campo where iIdActa=:Acta and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                        'sIdRecurso like "$IMPORTE%" group by etipo order by iOrdenTipo,sAnexo'

      else
        QImportes.SQL.Text:='select * from acta_campo where iIdActa=:Acta and swbs=:wbs and sNumeroActividad=:Actividad and '   +
                        'sIdRecurso like "$IMPORTE%" order by iOrdenTipo,sAnexo';


      QrActividades.Active:=False;
      QrActividades.ParamByName('Acta').AsInteger:=Datos.FieldByName('iIdActa').AsInteger;
      if (I=2) or (I=1) then
        QrActividades.ParamByName('Tipo').AsString:='Actividad'
      else
        QrActividades.ParamByName('Tipo').AsString:='Paquete';
      QrActividades.Open;

      Hoja.ClearAll;
      //NPdas:=0;

      IRow:=0;
      while not QrActividades.Eof do
      begin
        sFormulaSumMn:='';
        sFormulaSumDll:='';
        QImportes.Active:=False;
        QImportes.ParamByName('Acta').AsInteger:=Datos.FieldByName('iIdActa').AsInteger;
        QImportes.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
        QImportes.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
        QImportes.Open;

        with Hoja do
        begin
          //IRow:=NPdas;
          if (Hoja.Caption='CAMPO') or (Hoja.Caption='COSTO POR ACTIVIDAD') then
          begin
            Hoja.Rows.ResetDefault(IRow);
            Hoja.Rows.ResetDefault(IRow+1);
       //     Cols.Size[7]:=Cols.Size[7] - 5;

            ACellObj := GetCellObject(0, IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('PARTIDA');
            ACellObj := GetCellObject(2, IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('ACTIVIDAD');

            GetCellObject(2, IRow+1).Style.WordBreak:=True;

           //
            Rows.Size[IRow+1] := 70;

            LockUnLock(Hoja,0,1,IRow,0,False);
            RangeMerge(Hoja,0,1,IRow,0);
            LockUnLock(Hoja,0,1,IRow,0);

            LockUnLock(Hoja,2,10,IRow,0,False);
            RangeMerge(Hoja,2,10,IRow,0);
            LockUnLock(Hoja,2,10,IRow,0);

            LockUnLock(Hoja,0,1,IRow+1,0,False);
            RangeMerge(Hoja,0,1,IRow+1,0);
            LockUnLock(Hoja,0,1,IRow+1,0);

            LockUnLock(Hoja,2,10,IRow+1,0,False);
            RangeMerge(Hoja,2,10,IRow+1,0);
            LockUnLock(Hoja,2,10,IRow+1,0);

            BordeCelda(Hoja,0,10,IRow,1);
            Hoja.GetCellObject(0,IRow+1).SetCellText(QrActividades.FieldByName('sNumeroActividad').AsString);
            Hoja.GetCellObject(2,IRow+1).SetCellText(Trim(QrActividades.FieldByName('mDescripcion').AsString));
            AlineacionCelda(Hoja,haCENTER,vaCENTER,0,0,IRow+1,0);


            if (Hoja.Caption='CAMPO') then
            begin
              inc(IRow,3);
              Hoja.Rows.ResetDefault(IRow);
              ACellObj := GetCellObject(1, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('PERIODOS DE EJECUCION DE LA ACTIVIDAD');
              LockUnLock(Hoja,1,8,IRow,0,False);
              RangeMerge(Hoja,1,8,IRow,0);
              LockUnLock(Hoja,1,8,IRow,0);
              BordeCelda(Hoja,1,8,IRow,2);
                        //IRow:=5;
              Inc(IRow);
              Hoja.Rows.ResetDefault(IRow);
              ACellObj := GetCellObject(1, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('FECHA');
              ACellObj := GetCellObject(2,IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('INICIO');
              ACellObj := GetCellObject(3, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('TERMINO');
              ACellObj := GetCellObject(4, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.SetCellText('AFECTACIÓN');
              ACellObj := GetCellObject(5, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.Style.WordBreak:=True;
              ACellObj.SetCellText('INTERVALO TIEMPO');
              ACellObj := GetCellObject(6, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.Style.WordBreak:=True;
              ACellObj.SetCellText('AVANCE ANTERIOR');
              ACellObj := GetCellObject(7, IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.Style.WordBreak:=True;
              ACellObj.SetCellText('AVANCE ACTUAL');
              ACellObj := GetCellObject(8,IRow);
              FormatoCelda(ACellObj,FrSubTitle);
              ACellObj.Style.WordBreak:=True;
              ACellObj.SetCellText('AVANCE ACUMULADO');
              Rows.Size[IRow] :=30;

              inc(IRow,2);
              //IRow:=7;
              Hoja.Rows.ResetDefault(IRow);
              ACellObj := GetCellObject(1, IRow);
              FormatoCelda(ACellObj,FrTitle);
              ACellObj.SetCellText('DURACION TIEMPO EFECTIVO (HRS):');

                           //Inc(IRow,3);
              LockUnLock(Hoja,1,4,IRow,0,False);
              RangeMerge(Hoja,1,4,IRow,0);
              LockUnLock(Hoja,1,4,IRow,0);
              BordeCelda(Hoja,1,5,IRow,0);
              IRow:=IRow-1;

              QrAvance.Active:=False;
              QrAvance.ParamByName('contrato').AsString:=Datos.FieldByName('sContrato').AsString;
              QrAvance.ParamByName('Orden').AsString:=Datos.FieldByName('sNumeroOrden').AsString;
              QrAvance.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
              QrAvance.ParamByName('tipo').AsString:='ED';
              QrAvance.Open;
              TotalHrs:='00:00';
              while not QrAvance.Eof do
              begin
                if QrAvance.RecordCount<>QrAvance.RecNo then
                begin
                  LockUnLock(Hoja,1,8,IRow,0,false);
                  Hoja.SelectCell(1,IRow);
                  Hoja.InsertCells(Hoja.SelectionRect,msallRow );
                end
                else
                  LockUnLock(Hoja,1,8,IRow,0,true);
                //Hoja.GetCellObject(1,Row).SetCellText(QrAvance.FieldByName('dIdFecha').AsString);
                Hoja.GetCellObject(1,IRow).DateTime:= QrAvance.FieldByName('dIdFecha').AsDateTime;
                Hoja.GetCellObject(2,IRow).SetCellText(QrAvance.FieldByName('sHoraInicio').AsString);
                Hoja.GetCellObject(3,IRow).SetCellText(QrAvance.FieldByName('sHoraFinal').AsString);
                Hoja.GetCellObject(4,IRow).SetCellText(QrAvance.FieldByName('sIdTipoMovimiento').AsString);
                Hoja.GetCellObject(5,IRow).SetCellText(sfnRestaHoras(QrAvance.FieldValues['sHoraFinal'], QrAvance.FieldValues['sHoraInicio']));
                Hoja.GetCellObject(6,IRow).SetCellText(FormatFloat( '0.00',QrAvance.FieldByName('AvanceAnterior').asfloat)  + '%');
                Hoja.GetCellObject(7,IRow).SetCellText(FormatFloat( '0.00',QrAvance.FieldByName('dAvance').AsFloat) + '%');
                Hoja.GetCellObject(8,IRow).SetCellText(FormatFloat( '0.00',QrAvance.FieldByName('dAvance').AsFloat + QrAvance.FieldByName('AvanceAnterior').AsFloat) + '%');

                TotalHrs:=sfnSumaHoras(TotalHrs,sfnRestaHoras(QrAvance.FieldByName('sHoraFinal').AsString, QrAvance.FieldByName('sHoraInicio').AsString));
                QrAvance.Next;
                BordeCelda(Hoja,1,8,IRow,0);
                inc(IRow);
                Hoja.Rows.ResetDefault(IRow);
              end;

              if QrAvance.RecordCount>0 then
              begin
               // inc(row);
                LockUnLock(Hoja,5,5,IRow,0,false);
                FormatoCelda(Hoja.GetCellObject(5,IRow),FrTitle);
                Hoja.GetCellObject(5,IRow).SetCellText(TotalHrs);
                AlineacionCelda(Hoja,haCENTER,vaCENTER,1,8,4,IRow-4);
                LockUnLock(Hoja,5,5,IRow,0);
               // Inc(IRow,4)
              end;
            end;
          end;
         // else
          //  Inc(IRow,5);

          IF (Hoja.Caption='COSTO POR ACTIVIDAD') then
          begin
            inc(IRow,3);
            Hoja.Rows.ResetDefault(IRow);
            ACellObj := GetCellObject(9,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('IMP MN');

            ACellObj := GetCellObject(10,IRow);
            FormatoCelda(ACellObj,Frtitle);
            ACellObj.SetCellText('IMP USD');

            AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
            BordeCelda(Hoja,9,10,IRow,0);
            //Inc(IRow);
            pIRow:=IRow+2;
          end;

          IF (Hoja.Caption='ACTA DE ACTIVIDADES') then
          begin
            inc(IRow,1);
            Hoja.Rows.ResetDefault(IRow);
            ACellObj := GetCellObject(0,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('RESUMEN DE COSTOS:');
            RangeMerge(Hoja,0,9,IRow,0);
            Rows.Size[IRow] :=40;

            inc(IRow,1);
            Hoja.Rows.ResetDefault(IRow);
            ACellObj := GetCellObject(0,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('ANEXOS "C"');
            RangeMerge(Hoja,0,9,IRow,0);
            Rows.Size[IRow] :=30;

            inc(IRow,1);
            Hoja.Rows.ResetDefault(IRow);
            ACellObj := GetCellObject(0,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('ANEXO C');
            RangeMerge(Hoja,0,2,IRow,0);

            ACellObj := GetCellObject(3,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('DESCRIPCIÓN');
            RangeMerge(Hoja,3,7,IRow,0);

            ACellObj := GetCellObject(8,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('IMPORTE TOTAL');
            RangeMerge(Hoja,8,9,IRow,0);

            Rows.Size[IRow] :=25;

            inc(IRow,1);
            Hoja.Rows.ResetDefault(IRow);
            {ACellObj := GetCellObject(0,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('ANEXO C');   }
            RangeMerge(Hoja,0,2,IRow,0);

            ACellObj := GetCellObject(3,IRow);
            {FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('DESCRIPCIÓN'); }
            RangeMerge(Hoja,3,7,IRow,0);

            ACellObj := GetCellObject(8,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('M.N.');

            ACellObj := GetCellObject(9,IRow);
            FormatoCelda(ACellObj,FrTitle);
            ACellObj.SetCellText('U.S.D.');


           // AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
            //BordeCelda(Hoja,9,10,IRow,0);
            //Inc(IRow);
            pIRow:=IRow+1;
          end;





          while not QImportes.Eof do
          begin

              if QImportes.FieldByName('xRow').AsInteger=-1 then
                QImportes.Edit;

              IF (Hoja.Caption='COSTO POR ACTIVIDAD') or
                  (Hoja.Caption='ACTA DE ACTIVIDADES') then
                inc(iRow)
              else
                inc(iRow,2);

              Hoja.Rows.ResetDefault(IRow);
              if QImportes.fieldByName('eTipo').AsString<>'ACTIVIDAD' then
              begin
                if (Hoja.Caption='CAMPO') or (Hoja.Caption='DESGLOSE DE COSTOS') then
                begin
                  QRecursos.Active:=False;
                  QRecursos.ParamByName('Acta').AsInteger:=Datos.FieldByName('iIdActa').AsInteger;
                  QRecursos.ParamByName('wbs').AsString:=QrActividades.FieldByName('swbs').AsString;
                  QRecursos.ParamByName('Actividad').AsString:=QrActividades.FieldByName('sNumeroActividad').AsString;
                  QRecursos.ParamByName('Tipo').AsString:=QImportes.FieldByName('eTipo').AsString;
                  QRecursos.ParamByName('anexo').AsString:=QImportes.FieldByName('sAnexo').AsString;
                  QRecursos.Open;


                  if QImportes.FieldByName('eTipo').AsString='BARCO' then
                  begin
                    ACellObj := GetCellObject(0,IRow);
                    FormatoCelda(ACellObj,FrTitle);
                    ACellObj.SetCellText('MOVIMIENTO DE EMBARCACION');

                  end
                  else
                  begin
                    if QImportes.FieldByName('eTipo').AsString='MATERIAL' then
                    begin
                      ACellObj := GetCellObject(0,IRow);
                      FormatoCelda(ACellObj,FrTitle);
                      ACellObj.SetCellText('ANEXO ' + QImportes.FieldByName('SaNEXO').AsString );
                    end
                    else
                    if (i=3) and (QImportes.FieldByName('sAnexo').AsString<>'')  then
                    begin
                      ACellObj := GetCellObject(0,IRow);
                      FormatoCelda(ACellObj,FrTitle);
                      if (QImportes.FieldByName('eTipo').AsString='PERSONAL') then
                        ACellObj.SetCellText('PERSONAL TIEMPO EXTRA');

                    end
                    else
                    begin
                      ACellObj := GetCellObject(0,IRow);
                      FormatoCelda(ACellObj,FrTitle);
                      ACellObj.SetCellText(QImportes.FieldByName('eTipo').AsString);
                    end;
                  end;

                  LockUnLock(Hoja,0,10,IRow,0,False);
                  RangeMerge(Hoja,0,10,IRow,0);
                  LockUnLock(Hoja,0,10,IRow,0);
                  pIRow:=IRow;
                  inc(IRow);
                  Hoja.Rows.ResetDefault(IRow);
                  ACellObj := GetCellObject(0,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('PARTIDA');
                  ACellObj := GetCellObject(1,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('DESCRIPCIÓN');

                  if QImportes.FieldByName('eTipo').AsString='BARCO' then
                  begin
                    ACellObj := GetCellObject(5,IRow);
                    FormatoCelda(ACellObj,FrSubTitle);
                    ACellObj.SetCellText('CLAS.');
                  end
                  else
                  begin
                    ACellObj := GetCellObject(5,IRow);
                    FormatoCelda(ACellObj,FrSubTitle);
                    ACellObj.SetCellText('UNIDAD');
                  end;

                  ACellObj := GetCellObject(6,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('CANTIDAD');
                  ACellObj := GetCellObject(7,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('PU MN');
                  ACellObj := GetCellObject(8,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('PU USD');
                  ACellObj := GetCellObject(9,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('IMP MN');
                  ACellObj := GetCellObject(10,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('IMP USD');

                  LockUnLock(Hoja,1,4,IRow,0,False);
                  RangeMerge(Hoja,1,4,IRow,0);

                  while not QRecursos.Eof do
                  begin
                    Inc(IRow);
                    Hoja.Rows.ResetDefault(IRow);
                    if QRecursos.FieldByName('xRow').AsInteger=-1 then
                      QRecursos.Edit;
                    Hoja.GetCellObject(0,IRow).SetCellText(QRecursos.FieldByName('sIdRecurso').AsString);
                    AlineacionCelda(Hoja,haCENTER,vaCENTER,0,0,IRow,0);
                    Hoja.GetCellObject(1,IRow).SetCellText(QRecursos.FieldByName('mDescripcion').AsString);
                    Hoja.GetCellObject(5,IRow).SetCellText(QRecursos.FieldByName('sMedida').AsString);
                    AlineacionCelda(Hoja,haCENTER,vaCENTER,5,5,IRow,0);
                    //"#,##0.00"
                    //"$#,##0.00"
                    if QImportes.FieldByName('eTipo').AsString='BARCO' then
                     // Hoja.GetCellObject(6,IRow).SetCellText(FormatFloat( '0.000000',QRecursos.FieldByName('dCantidad').Asfloat))
                      Hoja.GetCellObject(6,IRow).Text:=FormatFloat( '0.000000',QRecursos.FieldByName('dCantidad').Asfloat)
                    else
                    begin
                     // Hoja.GetCellObject(6,IRow).Style.
                      Hoja.GetCellObject(6,IRow).Style.Format:=$2; //$2
                      Hoja.GetCellObject(6,IRow).Text:=FormatFloat( '0.00',QRecursos.FieldByName('dCantidad').AsFloat);
                    end;

                   //  Hoja.GetCellObject(6,IRow).Style.Format:=$11; //$2

                   //  Hoja.GetCellObject(6,IRow).Text:=QRecursos.FieldByName('dCantidad').AsString;
                    LockUnLock(Hoja,6,6,IRow,0,False);
                   // Hoja.GetCellObject(7,IRow).SetCellText(FormatFloat( '#,##0.00',QRecursos.FieldByName('dCostoMn').AsFloat));
                   Hoja.GetCellObject(7,IRow).Style.Format:=$4;
                   Hoja.GetCellObject(7,IRow).Text:=QRecursos.FieldByName('dCostoMn').AsString;
                   // Hoja.GetCellObject(8,IRow).SetCellText(FormatFloat( '#,##0.00',QRecursos.FieldByName('dCostoDll').AsFloat));
                   Hoja.GetCellObject(8,IRow).Style.Format:=$4;
                   Hoja.GetCellObject(8,IRow).Text:=QRecursos.FieldByName('dCostoDll').AsString;
                    Hoja.GetCellObject(9,IRow).Style.Format:=$4;
                    if QRecursos.FieldByName('sFormulaMN').AsString='##' then
                    begin

                      Hoja.GetCellObject(9,IRow).Text:= '=Round((G' + IntToStr(IRow+1)+ ' * ' + 'H' + IntToStr(IRow+1)+ '),2)';
                      if QRecursos.State=dsEdit then
                        QRecursos.FieldByName('sFormulaMN').AsString:='=Round((G' + IntToStr(IRow+1)+ ' * ' + 'H' + IntToStr(IRow+1)+ '),2)';
                    end
                    else
                      Hoja.GetCellObject(9,IRow).Text:=QRecursos.FieldByName('sFormulaMN').AsString;


                    Hoja.GetCellObject(10,IRow).Style.Format:=$4;
                    if QRecursos.FieldByName('sFormulaDll').AsString='##' then
                    begin
                      Hoja.GetCellObject(10,IRow).Text:= '=Round((G' + IntToStr(IRow+1)+ ' * ' + 'I' + IntToStr(IRow+1)+ '),2)' ;
                      if QRecursos.State=dsEdit then
                        QRecursos.FieldByName('sFormulaDll').AsString:='=Round((G' + IntToStr(IRow+1)+ ' * ' + 'I' + IntToStr(IRow+1)+ '),2)' ;
                    end
                    else
                      Hoja.GetCellObject(10,IRow).Text:=QRecursos.FieldByName('sFormulaDll').AsString;
                    RangeMerge(Hoja,1,4,IRow,0);

                    LockUnLock(Hoja,9,10,IRow,0,False);

                    if Length(QRecursos.FieldByName('mDescripcion').AsString)>200 then
                      Rows.Size[IRow] := Trunc((Length(QRecursos.FieldByName('mDescripcion').AsString)/60) * 20)
                    else
                      if Length(QRecursos.FieldByName('mDescripcion').AsString)>55 then
                        Rows.Size[IRow] := Trunc((Length(QRecursos.FieldByName('mDescripcion').AsString)/60) * 30);

                    if QRecursos.State=dsEdit then
                    begin
                      QRecursos.FieldByName('xRow').AsInteger:=IRow+1;
                      QRecursos.Post;
                    end;
                    QRecursos.Next;
                  end;


                  inc(iRow);
                  Hoja.Rows.ResetDefault(IRow);
                //IRow:=12;
                  ACellObj := GetCellObject(6, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  if QImportes.FieldByName('eTipo').AsString='MATERIAL' then
                    ACellObj.SetCellText('IMPORTE ANEXO ' + QImportes.FieldByName('sAnexo').AsString + ':')
                  else
                  ACellObj.SetCellText(QImportes.FieldByName('mDescripcion').AsString);
                  LockUnLock(Hoja,6,8,IRow,0,False);
                  RangeMerge(Hoja,6,8,IRow,0);
                  LockUnLock(Hoja,6,8,IRow,0);
                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Style.Format:=$4;
                  if QImportes.FieldByName('sFormulaMn').AsString='##' then
                  begin
                    ACellObj.Text:= '=Round(sum(J' + IntToStr(pIRow+3)+ ':' + 'J' + IntToStr(IRow)+ '),2)' ;

                    if QImportes.State=dsEdit then
                      QImportes.FieldByName('sFormulaMn').AsString:='=Round(sum(J' + IntToStr(pIRow+3)+ ':' + 'J' + IntToStr(IRow)+ '),2)' ;
                  end
                  else
                    ACellObj.Text:=QImportes.FieldByName('sFormulaMn').AsString;

                  ACellObj := GetCellObject(10, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Style.Format:=$4;
                  if QImportes.FieldByName('sFormulaDll').AsString='##' then
                  begin
                    ACellObj.Text:= '=Round(sum(K' + IntToStr(pIRow+3)+ ':' + 'K' + IntToStr(IRow)+ '),2)' ;
                    if QImportes.State=dsEdit then
                      QImportes.FieldByName('sFormulaDll').AsString:='=Round(sum(K' + IntToStr(pIRow+3)+ ':' + 'K' + IntToStr(IRow)+ '),2)' ;
                  end
                  else
                    ACellObj.Text:=QImportes.FieldByName('sFormulaDll').AsString;

                  if QImportes.State=dsEdit then
                    QImportes.FieldByName('xRow').AsInteger:=IRow+1;
                  BordeCelda(Hoja,0,10,pIRow,IRow-pIRow-1);
                  BordeCelda(Hoja,6,10,IRow,0);
                  if sFormulaSumMn='' then
                     sFormulaSumMn:= 'J' + IntToStr(IRow+1)
                  else
                    sFormulaSumMn:= sFormulaSumMn +  '+J' + IntToStr(IRow+1);

                  if sFormulaSumDll='' then
                    sFormulaSumDll:='K' + IntToStr(IRow+1)
                  else
                    sFormulaSumDll:=sFormulaSumDll + '+K' + IntToStr(IRow+1);
                  AlineacionCelda(Hoja,haCENTER,vaCENTER,6,10,IRow,0);
                  LockUnLock(Hoja,9,10,IRow,0,False);
                  if  (QImportes.FieldByName('sLeyendaAnexo').AsString='') and
                      (QImportes.State = dsEdit) then
                    if Datos.FieldByName('eLugarOT').AsString='Tierra' then
                    begin
                      if QImportes.FieldByName('eTipo').AsString='BARCO' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C 1.1';

                      if QImportes.FieldByName('eTipo').AsString='PERSONAL' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-5 "PERSONAL OPT."';

                      if QImportes.FieldByName('eTipo').AsString='EQUIPO' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-5 "EQUIPO OPT."';

                      if QImportes.FieldByName('eTipo').AsString='PERNOCTA' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-4 "SERV. DE HOTEL."';
                    end
                    else
                    begin
                      if QImportes.FieldByName('eTipo').AsString='BARCO' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C 1.1';

                      if QImportes.FieldByName('eTipo').AsString='PERSONAL' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-2 "PERSONAL OPT."';

                      if QImportes.FieldByName('eTipo').AsString='EQUIPO' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-3 "EQUIPO OPT."';

                      if QImportes.FieldByName('eTipo').AsString='PERNOCTA' then
                        QImportes.FieldByName('sLeyendaAnexo').AsString:='ANEXO C-4 "SERV. DE HOTEL."';
                    end;

                end;

                if (Hoja.Caption='COSTO POR ACTIVIDAD')  then
                begin
                  ACellObj := GetCellObject(0, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  if QImportes.FieldByName('eTipo').AsString='BARCO' then
                    ACellObj.SetCellText('MOVIMIENTO DE EMBARCACIÓN')
                  else
                    ACellObj.SetCellText(QImportes.FieldByName('eTipo').AsString);
                  RangeMerge(Hoja,0,5,IRow,0);

                  ACellObj := GetCellObject(6, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.SetCellText(QImportes.FieldByName('mDescripcion').AsString);
                  LockUnLock(Hoja,6,8,IRow,0,False);
                  RangeMerge(Hoja,6,8,IRow,0);
                  LockUnLock(Hoja,6,8,IRow,0);
                  //AlineacionCelda(Hoja,haLeft,vaCENTER,5,5,IRow,0);


                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.Style.Format:=$4;
                  ACellObj.Text:='=CAMPO!J' + QImportes.FieldByName('xRow').AsString ;

                  ACellObj := GetCellObject(10, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.Style.Format:=$4;
                  ACellObj.Text:='=CAMPO!K' + QImportes.FieldByName('xRow').AsString ;

                  AlineacionCelda(Hoja,haLeft,vaCENTER,0,5,IRow,0);
                  AlineacionCelda(Hoja,haCENTER,vaCENTER,6,8,IRow,0);
                  BordeCelda(Hoja,0,10,IRow,0);
                  //Rows.Size[IRow] :=30;

                end;

                if (Hoja.Caption='ACTA DE ACTIVIDADES')  then
                begin
                  ACellObj := GetCellObject(0, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.SetCellText(QImportes.FieldByName('sLeyendaAnexo').AsString);
                  RangeMerge(Hoja,0,2,IRow,0);

                  ACellObj := GetCellObject(3,IRow);
                  FormatoCelda(ACellObj,FrContent);
                  if QImportes.FieldByName('eTipo').AsString='BARCO' then
                    ACellObj.SetCellText('UTILIZANDO POSICIONAMIENTO DINÁMICO')
                  else
                    if QImportes.FieldByName('eTipo').AsString='PERNOCTA' then
                      ACellObj.SetCellText('SERVICIOS DE HOTELERIA')
                    else
                      if QImportes.FieldByName('eTipo').AsString='MATERIAL' then
                        ACellObj.SetCellText('')
                      ELSE
                        ACellObj.SetCellText(QImportes.FieldByName('eTipo').AsString);
                  RangeMerge(Hoja,3,7,IRow,0);


                 // if i=0 then
                 // QImportes

                  //  posI:Integer;
                  sItem:='';
                  desgloseMn:='';
                  DesgloseDll:='';
                  for PosI := 1 to NumItems(QImportes.FieldByName('xRows').AsString,',') do
                  begin
                    sItem:=TraerItem(QImportes.FieldByName('xRows').AsString,',',PosI);
                    if sItem<>'' then
                    begin
                      if desgloseMn='' then
                        desgloseMn:='='+quotedstr('DESGLOSE DE COSTOS')+'!J' + sItem
                      else
                        desgloseMn:=desgloseMn + '+'+quotedstr('DESGLOSE DE COSTOS')+'!J' + sItem;

                      if DesgloseDll='' then
                        DesgloseDll:='='+quotedstr('DESGLOSE DE COSTOS')+'!K' + sItem
                      else
                        DesgloseDll:=DesgloseDll + '+'+quotedstr('DESGLOSE DE COSTOS')+'!K' + sItem ;

                    end;
                  end;

                  ACellObj := GetCellObject(8, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.Style.Format:=$4;
                  ACellObj.Text:=desgloseMn;//'='+quotedstr('DESGLOSE DE COSTOS')+'!J' + QImportes.FieldByName('xRow').AsString ;

                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrContent);
                  ACellObj.Style.Format:=$4;
                  ACellObj.Text:=DesgloseDll;//'='+quotedstr('DESGLOSE DE COSTOS')+'!K' + QImportes.FieldByName('xRow').AsString ;

                  AlineacionCelda(Hoja,haLeft,vaCENTER,0,7,IRow,0);
                  AlineacionCelda(Hoja,haRIGHT,vaCENTER,8,9,IRow,0);
                  //BordeCelda(Hoja,0,10,IRow,0);
                  //Rows.Size[IRow] :=30;

                end;



              end
              else
              begin
                if (Hoja.Caption='CAMPO') or (Hoja.Caption='DESGLOSE DE COSTOS') then
                begin
                  ACellObj := GetCellObject(5, IRow);
                  ACellObj.SetCellText('COSTO TOTAL DE LA ACTIVIDAD:');
                  LockUnLock(Hoja,5,8,IRow,0,False);
                  RangeMerge(Hoja,5,8,IRow,0);
                  LockUnLock(Hoja,5,8,IRow,0);
                  AlineacionCelda(Hoja,haLeft,vaCENTER,5,5,IRow,0);
                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Style.Format:=$4;
                  if QImportes.FieldByName('sFormulaMN').AsString='##' then
                  begin
                    ACellObj.Text:= '=' + sFormulaSumMn;
                    if QImportes.State=dsEdit then
                      QImportes.FieldByName('sFormulaMN').AsString:= '=' + sFormulaSumMn;
                  end
                  else
                    ACellObj.Text:=QImportes.FieldByName('sFormulaMN').AsString;

                  ACellObj := GetCellObject(10, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Style.Format:=$4;
                  if QImportes.FieldByName('sFormulaDll').AsString='##' then
                  begin
                    ACellObj.Text:= '='+  sFormulaSumDll;
                    if QImportes.State=dsEdit then
                      QImportes.FieldByName('sFormulaDll').AsString:= '='+  sFormulaSumDll;
                  end
                  else
                    ACellObj.Text:=QImportes.FieldByName('sFormulaDll').AsString;

                  AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
                  BordeCelda(Hoja,5,10,IRow,0);
                  LockUnLock(Hoja,9,10,IRow,0,False);
                  Rows.Size[IRow] :=30;
                  if QImportes.State=dsEdit then
                    QImportes.FieldByName('xRow').AsInteger:=IRow+1;
                end;
                
                if (Hoja.Caption='COSTO POR ACTIVIDAD') then
                begin
                  inc(iRow);
                  Hoja.Rows.ResetDefault(IRow);
                  ACellObj := GetCellObject(9,IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.SetCellText('IMP MN');

                  ACellObj := GetCellObject(10,IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.SetCellText('IMP USD');

                  AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
                  BordeCelda(Hoja,9,10,IRow,0);

                  Inc(iRow);
                  Hoja.Rows.ResetDefault(IRow);
                  ACellObj := GetCellObject(5, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.SetCellText('COSTO TOTAL DE LA ACTIVIDAD:');
                  LockUnLock(Hoja,5,8,IRow,0,False);
                  RangeMerge(Hoja,5,8,IRow,0);
                  LockUnLock(Hoja,5,8,IRow,0);
                  AlineacionCelda(Hoja,haRight,vaCENTER,5,10,IRow,0);
                  ACellObj := GetCellObject(9, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Style.Format:=$4;
                  ACellObj.Text:='=sum(J' + IntToStr(pIRow) + ':J' + IntToStr(IRow-2)+ ')';
                 // ACellObj.Style.Format:=0;
                  //"#,##0.00"
                  ACellObj := GetCellObject(10, IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Style.Format:=$4;
                  ACellObj.Text:='=sum(K' + IntToStr(pIRow) + ':K' + IntToStr(IRow-2)+ ')';
                  AlineacionCelda(Hoja,haCENTER,vaCENTER,9,10,IRow,0);
                  BordeCelda(Hoja,5,10,IRow,0);
                  Rows.Size[IRow] :=30;
                end;

                if (Hoja.Caption='ACTA DE ACTIVIDADES')  then
                begin
                  RangeMerge(Hoja,0,2,IRow,0);

                  ACellObj := GetCellObject(3,IRow);
                  FormatoCelda(ACellObj,FrSubTitle);
                  ACellObj.SetCellText('TOTALES');
                  RangeMerge(Hoja,3,7,IRow,0);

                  ACellObj := GetCellObject(8,IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Style.Format:=$4;
                  ACellObj.Text:='=sum(I' + IntToStr(pIRow+1) + ':I' + IntToStr(IRow)+ ')';

                  ACellObj := GetCellObject(9,IRow);
                  FormatoCelda(ACellObj,FrTitle);
                  ACellObj.Style.Format:=$4;
                  ACellObj.Text:='=sum(J' + IntToStr(pIRow+1) + ':J' + IntToStr(IRow)+ ')';

                  BordeCelda(Hoja,0,9,pIRow-4,IRow-(pIRow-4));
                end;
              end;
            //IRow:=NPdas;
            {LockUnLock(Hoja,0,1,IRow,0,False);
            RangeMerge(Hoja,0,1,IRow,0);
            LockUnLock(Hoja,0,1,IRow,0);

            LockUnLock(Hoja,2,10,IRow,0,False);
            RangeMerge(Hoja,2,10,IRow,0);
            LockUnLock(Hoja,2,10,IRow,0);

            Inc(IRow);
            LockUnLock(Hoja,0,1,IRow,0,False);
            RangeMerge(Hoja,0,1,IRow,0);
            LockUnLock(Hoja,0,1,IRow,0);

            LockUnLock(Hoja,2,10,IRow,0,False);
            RangeMerge(Hoja,2,10,IRow,0);
            LockUnLock(Hoja,2,10,IRow,0);

            Inc(IRow,2);
            LockUnLock(Hoja,1,8,IRow,0,False);
            RangeMerge(Hoja,1,8,IRow,0);
            LockUnLock(Hoja,1,8,IRow,0);


            Inc(IRow,3);
            LockUnLock(Hoja,1,4,IRow,0,False);
            RangeMerge(Hoja,1,4,IRow,0);
            LockUnLock(Hoja,1,4,IRow,0);

            Inc(IRow,2);
            LockUnLock(Hoja,0,10,IRow,0,False);
            RangeMerge(Hoja,0,10,IRow,0);
            LockUnLock(Hoja,0,10,IRow,0);

            Inc(IRow);
            LockUnLock(Hoja,1,4,IRow,0,False);
            RangeMerge(Hoja,1,4,IRow,0);

            Inc(IRow);
            LockUnLock(Hoja,1,4,IRow,0,False);
            RangeMerge(Hoja,1,4,IRow,0);

            Inc(IRow);
            LockUnLock(Hoja,6,8,IRow,0,False);
            RangeMerge(Hoja,6,8,IRow,0);
            LockUnLock(Hoja,6,8,IRow,0);

            IRow:=NPdas;
            BordeCelda(Hoja,0,10,IRow,1);

           // IRow:=3;
            Inc(IRow,3);
            BordeCelda(Hoja,1,8,IRow,2);



            //IRow:=5;
            Inc(IRow,3);
            BordeCelda(Hoja,1,5,IRow,0);  

            //IRow:=7;
            Inc(IRow,2);
            BordeCelda(Hoja,0,10,IRow,2);

           // IRow:=10;
            Inc(IRow,3);
            BordeCelda(Hoja,6,10,IRow,0);

            IIRow:=12;
            BordeCelda(Hoja,0,10,IRow,2);
            IRow:=15;
            BordeCelda(Hoja,6,10,IRow,0);

            nc(IRow,3);
            LockUnLock(Hoja,1,4,IRow,0,False);
            RangeMerge(Hoja,1,4,IRow,0);
            LockUnLock(Hoja,1,4,IRow,0);






            LockUnLock(Hoja,0,10,12,0,False);
            RangeMerge(Hoja,0,10,12,0);
            LockUnLock(Hoja,0,10,12,0);

            LockUnLock(Hoja,1,4,13,0,False);
            RangeMerge(Hoja,1,4,13,0);
            LockUnLock(Hoja,1,4,13,0);

            LockUnLock(Hoja,1,4,14,0,False);
            RangeMerge(Hoja,1,4,14,0);

            LockUnLock(Hoja,6,8,15,0,False);
            RangeMerge(Hoja,6,8,15,0);
            LockUnLock(Hoja,6,8,15,0);

            }

            if QImportes.State=dsEdit then
              QImportes.Post;
            QImportes.Next;
          end;
        end;
        Inc(iRow,2);
        QrActividades.Next;
      end;




    end;

  finally
    QrActividades.Destroy;
    QImportes.Destroy;
    QRecursos.Destroy;
    QrAvance.Destroy;
  end;
end;


Procedure TFrmNotaCampo.GeneraActaEntrega_Ex(RTipo:FtTipo;RImpresion:FtSeccion;sAplica:string='');
var
   sSeccion: string;
   i:Integer;
   sDatasets:string;
begin
    //RptActa.DataSets.Clear;
    QrImprimir.Active:=False;
    QrImprimir.ParamByName('Acta').AsInteger:=QActa.FieldByName('iIdActa').AsInteger;
    QrImprimir.Open;
    ActaPdf_Informacion(QrImprimir,RptActa);
    EncabezadoPDF_Horizontal(QrImprimir,RptActa,RTipo);
    ActaPdf_Actividades(QrImprimir,RptActa);
    ActaEx_CostoActividad(QrImprimir,RptActa,RTipo);
    ActaEx_NotaCampo(QrImprimir,RptActa,RTipo);
    ActaEx_DesgloseCostos(QrImprimir,RptActa,RTipo);
    ActaPdf_CostoInterferencia(QrImprimir,RptActa,RTipo);

    for I := 0 to RptActa.DataSets.Count - 1 do
    begin
      if not Assigned(RptActa.DataSets.Items[i].DataSet) then
        sDatasets:=sDatasets + RptActa.DataSets.Items[i].DataSetName + '= NILL' + #13 + #10
      else
        sDatasets:=sDatasets + RptActa.DataSets.Items[i].DataSetName + '= '  + RptActa.DataSets.Items[i].DataSet.Name+ #13 + #10;
    end;

    if RTipo=FtTierra then
      RptActa.LoadFromFile(global_files + global_Mireporte + '_TDActaEntrega.fr3')
    else
      RptActa.LoadFromFile(global_files + global_Mireporte + '_TDActaEntregaAbordo.fr3') ;

    sDatasets:='';
    for I := 0 to RptActa.DataSets.Count - 1 do
    begin
      if not Assigned(RptActa.DataSets.Items[i].DataSet) then
        sDatasets:=sDatasets + RptActa.DataSets.Items[i].DataSetName + '= NILL' + #13 + #10
      else
        sDatasets:=sDatasets + RptActa.DataSets.Items[i].DataSetName + '= '  + TfrxDBDataset(RptActa.DataSets.Items[i].DataSet).DataSet.Name + #13 + #10;
    end;

    RptActa.ShowReport(true);
    ReportePDF_ClearDataset(RptActa);

end;

Procedure TFrmNotaCampo.GeneraActaEntrega_PDF(RTipo:FtTipo;RImpresion:FtSeccion;sAplica:string='');
var
   sSeccion: string;
   i:Integer;
   sDatasets:string;
begin
    //RptActa.DataSets.Clear;
    QrImprimir.Active:=False;
    QrImprimir.ParamByName('Acta').AsInteger:=QActa.FieldByName('iIdActa').AsInteger;
    QrImprimir.Open;
    ActaPdf_Informacion(QrImprimir,RptActa);
    EncabezadoPDF_Horizontal(QrImprimir,RptActa,RTipo);
    ActaPdf_Actividades(QrImprimir,RptActa);
    ActaPdf_CostoActividad(QrImprimir,RptActa,RTipo);
    ActaPdf_NotaCampo(QrImprimir,RptActa,RTipo);
    ActaPdf_DesgloseCostos(QrImprimir,RptActa,RTipo);
    ActaPdf_CostoInterferencia(QrImprimir,RptActa,RTipo);

    for I := 0 to RptActa.DataSets.Count - 1 do
    begin
      if not Assigned(RptActa.DataSets.Items[i].DataSet) then
        sDatasets:=sDatasets + RptActa.DataSets.Items[i].DataSetName + '= NILL' + #13 + #10
      else
        sDatasets:=sDatasets + RptActa.DataSets.Items[i].DataSetName + '= '  + RptActa.DataSets.Items[i].DataSet.Name+ #13 + #10;
    end;

    if RTipo=FtTierra then
      RptActa.LoadFromFile(global_files + global_Mireporte + '_TDActaEntrega.fr3')
    else
      RptActa.LoadFromFile(global_files + global_Mireporte + '_TDActaEntregaAbordo.fr3') ;

    sDatasets:='';
    for I := 0 to RptActa.DataSets.Count - 1 do
    begin
      if not Assigned(RptActa.DataSets.Items[i].DataSet) then
        sDatasets:=sDatasets + RptActa.DataSets.Items[i].DataSetName + '= NILL' + #13 + #10
      else
        sDatasets:=sDatasets + RptActa.DataSets.Items[i].DataSetName + '= '  + TfrxDBDataset(RptActa.DataSets.Items[i].DataSet).DataSet.Name + #13 + #10;
    end;

    RptActa.ShowReport(true);

    if Assigned(MemoryTmp) then
      MemoryTmp:=nil;

    if Assigned(MemoryDetalleTmp) then
      MemoryDetalleTmp:=nil;

    ReportePDF_ClearDataset(RptActa);

  (* FirmasPDF_Generales(ReporteDiario,     rDiario,FtAbordo);
    sSeccion := connection.configuracion.FieldByName('sSeccionImprime').AsString;
    {Clasificacion de secciones a Imprimir..}

        if pos('Movimientos de Embarcacion', sSeccion) > 0 then
           ReportePDF_MovimientosLogisticos2(ReporteDiario,rDiario,RTipo,RImpresion)
        else
           ReportePDF_MovimientosLogisticos2(ReporteDiario,rDiario,RTipo,ftsNone);

        if pos('Actividades', sSeccion) > 0 then
           ReportePDF_ActividadesPorFolio(ReporteDiario,   rDiario,RTipo,RImpresion)
        else
           ReportePDF_ActividadesPorFolio(ReporteDiario,   rDiario,RTipo,ftsNone);

        if pos('Avances', sSeccion) > 0 then
           ReportePDF_AvancesCortes(ReporteDiario,   rDiario,RTipo,RImpresion)
        else
           ReportePDF_AvancesCortes(ReporteDiario,   rDiario,RTipo,ftsNone);

        if pos('Resumen de Material', sSeccion) > 0 then
           ReportePDF_ResumenMaterial2(ReporteDiario,      rDiario,RTipo,RImpresion)
        else
           ReportePDF_ResumenMaterial2(ReporteDiario,      rDiario,RTipo,ftsNone);

        if pos('Balance de Embarcacion', sSeccion) > 0 then
           ReportePDF_BalanceGeneral(ReporteDiario,        rDiario,RTipo,RImpresion)
        else
           ReportePDF_BalanceGeneral(ReporteDiario,        rDiario,RTipo,ftsNone);

        if pos('Notas Generales', sSeccion) > 0 then
           ReportePDF_NotasGenerales(ReporteDiario,        rDiario,RTipo,RImpresion)
        else
           ReportePDF_NotasGenerales(ReporteDiario,        rDiario,RTipo,ftsNone);

        if pos('Consumos de Diesel EQ', sSeccion) > 0 then
           ReportePDF_ConsumosDiesel(ReporteDiario,        rDiario,RTipo,RImpresion)
        else
           ReportePDF_ConsumosDiesel(ReporteDiario,        rDiario,RTipo,ftsNone);

        if pos('Concentrado de Personal', sSeccion) > 0 then
           ReportePDF_ConcentradoDePersonal2(ReporteDiario,rDiario,RTipo,RImpresion)
        else
           ReportePDF_ConcentradoDePersonal2(ReporteDiario,rDiario,RTipo,ftsNone);

        if pos('Concentrado de Equipos', sSeccion) > 0 then
           ReportePDF_DistribucionDeEquipos2(ReporteDiario,rDiario,RTipo,RImpresion)
        else
           ReportePDF_DistribucionDeEquipos2(ReporteDiario,rDiario,RTipo,ftsNone);

        if pos('Concentrado de Pernoctas', sSeccion) > 0 then
           ReportePDF_ConcentradoDePernoctas2(ReporteDiario,rDiario,RTipo,RImpresion)
        else
           ReportePDF_ConcentradoDePernoctas2(ReporteDiario,rDiario,RTipo,ftsNone);

        if pos('Lista de Personal', sSeccion) > 0 then
           ReportePDF_Listadepersonal(ReporteDiario,rDiario,RTipo,RImpresion)
        else
           ReportePDF_Listadepersonal(ReporteDiario,rDiario,RTipo,ftsNone);

        if pos('Reporte Fotografico', sSeccion) > 0 then
           ReportePDF_ReporteFotografico(ReporteDiario,    rDiario,RTipo,RImpresion)
        else
           ReportePDF_ReporteFotografico(ReporteDiario,    rDiario,RTipo,ftsNone);

        if pos('Resumen de Personal', sSeccion) > 0 then
           ReportePDF_TotalDePersonal2(ReporteDiario,      rDiario,RTipo,RImpresion)
        else
           ReportePDF_TotalDePersonal2(ReporteDiario,      rDiario,RTipo,ftsNone);

        if pos('Diario de Cobro', sSeccion) > 0 then
           ReportePDF_PartidasAnexoC(ReporteDiario,        rDiario,RTipo,RImpresion)
        else
           ReportePDF_PartidasAnexoC(ReporteDiario,        rDiario,RTipo,ftsNone);

        if pos('Partidas de Anexo C', sSeccion) > 0 then
          ReportePDF_PartidasAnexoC_detalle(ReporteDiario,rDiario,RTipo,ftsNone)
        else
          ReportePDF_PartidasAnexoC_detalle(ReporteDiario,rDiario,RTipo,ftsNone);

        if pos('Detalle de Anexo C', sSeccion) > 0 then
        begin
          ReportePDF_PartidasAnexoC_detalle(ReporteDiario,rDiario,RTipo,ftsNone);
          ReportePDF_PartidasAnexoC_detalleV2(ReporteDiario,rDiario,RTipo,RImpresion)
        end
        else
        begin
          ReportePDF_PartidasAnexoC_detalle(ReporteDiario,rDiario,RTipo,ftsNone);
          ReportePDF_PartidasAnexoC_detalleV2(ReporteDiario,rDiario,RTipo,ftsNone);
        end;

  sDatasets:='';
  for I := 0 to rDiario.DataSets.Count - 1 do
  begin
    if not Assigned(rDiario.DataSets.Items[i].DataSet) then
      sDatasets:=sDatasets + rDiario.DataSets.Items[i].DataSetName + '= NILL' + #13 + #10
    else
      sDatasets:=sDatasets + rDiario.DataSets.Items[i].DataSetName + '= '  + rDiario.DataSets.Items[i].DataSet.Name+ #13 + #10;

  end;

  if RTipo=FtTierra then
  begin
    if global_contrato='BLP-OTt-002' then
      rDiario.LoadFromFile(global_files + global_Mireporte + '_TDReporteDiarioTierraB.fr3')
    else
      rDiario.LoadFromFile(global_files + global_Mireporte + '_TDReporteDiarioTierra.fr3');
  end
  else
    rDiario.LoadFromFile(global_files + global_Mireporte + '_TDReporteDiario.fr3') ;

  sDatasets:='';
  for I := 0 to rDiario.DataSets.Count - 1 do
  begin
    if not Assigned(rDiario.DataSets.Items[i].DataSet) then
      sDatasets:=sDatasets + rDiario.DataSets.Items[i].DataSetName + '= NILL' + #13 + #10
    else
      sDatasets:=sDatasets + rDiario.DataSets.Items[i].DataSetName + '= '  + TfrxDBDataset(rDiario.DataSets.Items[i].DataSet).DataSet.Name + #13 + #10;

  end;

    rDiario.ShowReport();
    ReportePDF_ClearDataset(rDiario);  *)
end;
procedure TFrmNotaCampo.mniAddClick(Sender: TObject);
var
  Rango:TRect;
  i:Integer;
begin
  SprShBkDatos.ActiveSheet.InsertCells(SprShBkDatos.ActiveSheet.SelectionRect,msallRow );
  // SprShBkDatos.ActiveSheet.SelectCell();
  //SprShBkDatos.ActiveSheet.
  //SprShBkDatos.ActiveSheet.FormatCel;
  rango.Left:=SprShBkDatos.ActiveSheet.SelectionRect.left;
  //SprShBkDatos.act
  Rango.Right:=SprShBkDatos.ActiveSheet.SelectionRect.left + 3;
  Rango.Top:= SprShBkDatos.ActiveSheet.SelectionRect.Top;
  Rango.Bottom:=SprShBkDatos.ActiveSheet.SelectionRect.Bottom;
  //SprShBkDatos.ActiveSheet.f
  for I := SprShBkDatos.ActiveSheet.col to SprShBkDatos.ActiveSheet.col +4 do
  begin
    SprShBkDatos.ActiveSheet.GetCellObject(i,SprShBkDatos.ActiveSheet.row).Style.Locked:=false ;
  //SprShBkDatos.ActiveSheet.GetCellObject(i,SprShBkDatos.ActiveSheet.row).Style.Borders:=
  end;

  SprShBkDatos.ActiveSheet.SetMergedState(Rango,true);
  //InsertCells(ActiveSheet.SelectionRect, msAllCol)
end;

procedure TFrmNotaCampo.mniRecalcularClick(Sender: TObject);
var
  Opciones:TSetRtRecurso;
begin
  if QActa.State=dsBrowse then
  begin
    QrImprimir.Active:=False;
    QrImprimir.ParamByName('Acta').AsInteger:=QActa.FieldByName('iIdActa').AsInteger;
    QrImprimir.Open;
    if QrImprimir.FieldByName('eLugarOT').AsString='Tierra' then
    begin
      //(RtPersonal=1,RtEquipo=2,RtPernocta=3,RtBarco=4,RtExtraordinaria=5);

      Include(Opciones,RtPersonal);
      Include(Opciones,RtEquipo);
      Include(Opciones,RtExtraordinaria);

      Load_NotaCampo(QrImprimir,Opciones);
      Load_DesgloseCostos(QrImprimir,Opciones);
      Load_Ajuste(QrImprimir,Opciones);

    end
    else
    begin
      Include(Opciones,RtPersonal);
      Include(Opciones,RtEquipo);
      Include(Opciones,RtBarco);
      Include(Opciones,RtExtraordinaria);
      Include(Opciones,RtMAterial);
      if QrImprimir.FieldByName('lPernocta').AsString='Si' then
        Include(Opciones,RtPernocta);
      Load_NotaCampo(QrImprimir,Opciones);
      Load_DesgloseCostos(QrImprimir,Opciones);
      Load_Ajuste(QrImprimir,Opciones);
    end;
  end;
end;

procedure TFrmNotaCampo.pmDatosPopup(Sender: TObject);
begin
  mniAdd.Enabled:=False;
  mniDelete.Enabled:=False;
  if  SprShBkDatos.ActiveSheet.GetCellObject(SprShBkDatos.ActiveSheet.Col,SprShBkDatos.ActiveSheet.Row).Style.Locked=false
  then
  begin
    mniAdd.Enabled:=true;
    mniDelete.Enabled:=true;


  end;
end;

procedure TFrmNotaCampo.QActaAfterScroll(DataSet: TDataSet);
begin
  if QActa.FieldByName('lMaterial').AsString='Si' then
    cTsMateriales.TabVisible:=True
  else
    cTsMateriales.TabVisible:=False;

  if QActa.State=dsBrowse then
    CxPageDetalleChange(nil);


end;

procedure TFrmNotaCampo.QMaterialesAfterInsert(DataSet: TDataSet);
begin
  QMateriales.FieldByName('sContrato').AsString:=QActa.FieldByName('sContrato').AsString;
  QMateriales.FieldByName('iIdActa').AsInteger:=QActa.FieldByName('iIdActa').AsInteger;
   QMateriales.FieldByName('dCantidad').AsFloat:=  0;
end;

procedure TFrmNotaCampo.QMaterialesBeforePost(DataSet: TDataSet);
begin
  if QMateriales.FieldByName('sIdInsumo').IsNull then
  begin
    QMateriales.Cancel;
    Abort;
  end;
end;

procedure TFrmNotaCampo.BrPrincipalbtnCancelClick(Sender: TObject);
begin
  BrPrincipal.btnCancelClick(Sender);
  QActa.Cancel;
  DbLkpCmbFolio.Enabled:=true;
  btnAjustar.Enabled:=true;
end;

procedure TFrmNotaCampo.BrPrincipalbtnDeleteClick(Sender: TObject);
begin
  if MessageDlg('Desea eliminar el registro Seleccionado?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
        QActa.Delete;
end;

procedure TFrmNotaCampo.BrPrincipalbtnEditClick(Sender: TObject);
begin
  try
    CxPageDetalle.ActivePageIndex:=0;
    DbLkpCmbFolio.Enabled:=False;
    btnAjustar.Enabled:=False;
    QActa.Edit;
    BrPrincipal.btnEditClick(Sender);

  except
    on e : exception do
      UnitExcepciones.manejarExcep(E.Message, E.ClassName,Self.Caption, 'Al Editar un registro', 0);
  end;
end;

procedure TFrmNotaCampo.BrPrincipalbtnExitClick(Sender: TObject);
begin
  BrPrincipal.btnExitClick(Sender);
  Close;
end;

procedure TFrmNotaCampo.BrPrincipalbtnPostClick(Sender: TObject);
var
  nombres,cadenas:TStringList;
  esNuevo:Boolean;
  Opciones:TSetRtRecurso;
begin
  nombres:=TStringList.Create;
  cadenas:=TStringList.Create;
  try


    try

      nombres.Add('Folio');  cadenas.Add(DbLkpCmbFolio.Text);
      nombres.Add('No. de Acta');  cadenas.Add(DbTxtEdtActa.Text);
      nombres.Add('Especialidad');  cadenas.Add(DbTxtEdtEspecialidad.Text);
      nombres.Add('No. de Acta');  cadenas.Add(DbTxtEdtActa.Text);

      if not validaTexto(nombres, cadenas, '', '') then
      begin
          MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
          exit;
      end;



      esNuevo:=(QActa.State=dsInsert);
      QActa.Post;
      DbLkpCmbFolio.Enabled:=true;
      btnAjustar.Enabled:=true;
      if esNuevo then
      begin
        QrImprimir.Active:=False;
        QrImprimir.ParamByName('Acta').AsInteger:=QActa.FieldByName('iIdActa').AsInteger;
        QrImprimir.Open;
        if QrImprimir.FieldByName('eLugarOT').AsString='Tierra' then
        begin
          //(RtPersonal=1,RtEquipo=2,RtPernocta=3,RtBarco=4,RtExtraordinaria=5);

          Include(Opciones,RtPersonal);
          Include(Opciones,RtEquipo);
          Include(Opciones,RtExtraordinaria);

          Load_NotaCampo(QrImprimir,Opciones);
          Load_DesgloseCostos(QrImprimir,Opciones);
          Load_Ajuste(QrImprimir,Opciones);

        end
        else
        begin
          Include(Opciones,RtPersonal);
          Include(Opciones,RtEquipo);
          Include(Opciones,RtBarco);
          Include(Opciones,RtExtraordinaria);
          Include(Opciones,RtMAterial);
          if QrImprimir.FieldByName('lPernocta').AsString='Si' then
            Include(Opciones,RtPernocta);
          Load_NotaCampo(QrImprimir,Opciones);
          Load_DesgloseCostos(QrImprimir,Opciones);
          Load_Ajuste(QrImprimir,Opciones);
        end;
      end;
      BrPrincipal.btnPostClick(Sender);

    except
      on e : exception do
        UnitExcepciones.manejarExcep(E.Message, E.ClassName,Self.Caption, 'Al Salvar el registro', 0);
    end;
  finally
    nombres.Destroy;
    cadenas.Destroy;
  end;
end;

procedure TFrmNotaCampo.BrPrincipalbtnPrinterClick(Sender: TObject);
begin
   if connection.contrato.FieldByName('eLugarOt').AsString='Tierra' then
    {GeneraActaEntrega_PDF(FtTierra,FtsAll)//}GeneraActaEntrega_Ex(FtTierra,FtsAll)//
   else
    {GeneraActaEntrega_PDF(FtAbordo,FtsAll) ;//}GeneraActaEntrega_Ex(FtAbordo,FtsAll) ;//







end;

procedure TFrmNotaCampo.BrPrincipalbtnRefreshClick(Sender: TObject);
var
  PosReg:TBookmark;
begin
  try
    QActa.DisableControls;
    if QActa.RecordCount>0 then
      PosReg:=QActa.GetBookmark;
    try
      QActa.Refresh;
    finally
      try
        QActa.GotoBookmark(PosReg);
      except
        QActa.FreeBookmark(PosReg);
      end;
      QActa.EnableControls;
    end;
  except
    on e : exception do
      UnitExcepciones.manejarExcep(E.Message, E.ClassName,Self.Caption, 'Al Actualizar los Datos', 0);
  end;
end;

procedure TFrmNotaCampo.btnAjustarClick(Sender: TObject);
begin
  if QActa.State=dsBrowse then
  begin
    QrImprimir.Active:=False;
    QrImprimir.ParamByName('Acta').AsInteger:=QActa.FieldByName('iIdActa').AsInteger;
    QrImprimir.Open;
     CxPage1.ActivePageIndex:=1;
    SprShBkDatos.BeginUpdate;
    SprShBkDatos.Protected:=False;
    ActaExLoad(SprShBkDatos,QrImprimir) ;
    SprShBkDatos.Protected:=true;

    SprShBkDatos.EndUpdate;
   // SprShBkDatos.UpdateControl;
   // SprShBkDatos.Repaint;
    SprShBkDatos.ActivePage:=0;

  end;
end;

procedure TFrmNotaCampo.btnCancelarClick(Sender: TObject);
begin
  CxPage1.ActivePageIndex:=0;
end;

procedure TFrmNotaCampo.btnExcelClick(Sender: TObject);
begin
  if dlgSaveGuardar.Execute then
  begin
    dlgSaveGuardarTypeChange(Self) ;
    SprShBkDatos.SaveToFile(dlgSaveGuardar.FileName);
  end;
end;

procedure TFrmNotaCampo.btnGuardarClick(Sender: TObject);
begin
  QrImprimir.Active:=False;
  QrImprimir.ParamByName('Acta').AsInteger:=QActa.FieldByName('iIdActa').AsInteger;
  QrImprimir.Open;

  ActaExSave(SprShBkDatos,QrImprimir) ;
  CxPage1.ActivePageIndex:=0;
end;

procedure TFrmNotaCampo.cMmCeldaEnter(Sender: TObject);
begin
  tmpCelda:=cMmCelda.Text;
end;

procedure TFrmNotaCampo.cMmCeldaExit(Sender: TObject);
begin
  if tmpCelda<>cMmCelda.Text then
    SprShBkDatos.ActiveSheet.GetCellObject(SprShBkDatos.ActiveSheet.ActiveCell.X,SprShBkDatos.ActiveSheet.ActiveCell.Y).Text:=cMmCelda.Text ;
  tmpCelda:='';
end;

procedure TFrmNotaCampo.cMmCeldaKeyPress(Sender: TObject; var Key: Char);
begin
  If key = #13 then
    SprShBkDatos.SetFocus;
end;

procedure TFrmNotaCampo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:=caFree;
end;

procedure TFrmNotaCampo.FormCreate(Sender: TObject);
begin
  CxPage1.ActivePageIndex:=0;
  CxPage1.HideTabs:=True;
end;

procedure TFrmNotaCampo.FormShow(Sender: TObject);
begin
  QActa.Close;
  QActa.ParamByName('Contrato').AsString:=global_contrato;
  QActa.Open;

  QrFolios.Close;
  QrFolios.ParamByName('Contrato').AsString:=global_contrato;
  QrFolios.Open;

  CxPageDetalle.ActivePageIndex:=0;
end;

end.
