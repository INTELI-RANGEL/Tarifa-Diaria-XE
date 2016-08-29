unit Frm_generadores;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore,
  dxSkinDevExpressStyle, dxSkinFoggy, cxGroupBox, dxSkinscxPCPainter,
  dxLayoutContainer, dxLayoutControl, ComCtrls, dxCore, cxDateUtils,
  dxLayoutcxEditAdapters, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxCalendar,
  cxDBEdit, cxRadioGroup, dxLayoutControlAdapters, Menus, StdCtrls, cxButtons,
  DB, ZAbstractRODataset, ZDataset, cxLookupEdit, cxDBLookupEdit,
  cxDBLookupComboBox, frxClass, frxDBSet, dxSkinBlack, dxSkinBlue,
  dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide,
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

type
  TFrmGeneradores = class(TForm)
    GBx1: TcxGroupBox;
    GBx2: TcxGroupBox;
    dxLayoutControl1Group_Root: TdxLayoutGroup;
    dxLayoutControl1: TdxLayoutControl;
    dxLayoutControl2Group_Root: TdxLayoutGroup;
    dxLayoutControl2: TdxLayoutControl;
    DtEdtFechaInicio: TcxDateEdit;
    dxLayoutControl1Item1: TdxLayoutItem;
    DtEdtFechaFin: TcxDateEdit;
    dxLayoutControl1Item2: TdxLayoutItem;
    RdGpGeneradores: TcxRadioGroup;
    dxLayoutControl2Item1: TdxLayoutItem;
    btnIMprimir: TcxButton;
    dxLayoutControl2Item2: TdxLayoutItem;
    QrOrdenes: TZReadOnlyQuery;
    QrFolios: TZReadOnlyQuery;
    dsOrdenes: TDataSource;
    dsFolios: TDataSource;
    dxLayoutControl1Item3: TdxLayoutItem;
    LCmbOrdenes: TcxLookupComboBox;
    dxLayoutControl1Group1: TdxLayoutAutoCreatedGroup;
    LCmbFolios: TcxLookupComboBox;
    dxLayoutControl1Item4: TdxLayoutItem;
    dxLayoutControl1Group2: TdxLayoutAutoCreatedGroup;
    FrReporte: TfrxReport;
    fDbDts1: TfrxDBDataset;
    procedure FormCreate(Sender: TObject);
    procedure DtEdtFechaFinPropertiesCloseUp(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnIMprimirClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FrReporteGetValue(const VarName: string; var Value: Variant);
    procedure DtEdtFechaInicioPropertiesEditValueChanged(Sender: TObject);
    procedure DtEdtFechaFinPropertiesEditValueChanged(Sender: TObject);
  private
    { Private declarations }
    TituloGenerador:string;
    Procedure InicializeFilters;
  public
    { Public declarations }
  end;

var
  FrmGeneradores: TFrmGeneradores;

implementation

uses frm_connection,DateUtils, global, UnitTarifa, Utilerias, UFunctionsGHH;

{$R *.dfm}

procedure TFrmGeneradores.InicializeFilters;
begin
  QrOrdenes.Close;
  QrOrdenes.ParamByName('fechaI').AsDateTime:=DtEdtFechaInicio.Date;
  QrOrdenes.ParamByName('fechaF').AsDateTime:=DtEdtFechaFin.Date;
  QrOrdenes.Open;

  if QrFolios.Active then
    QrFolios.Refresh
  else
    QrFolios.Open;
  //:fechaI and :fechaF
end;

procedure TFrmGeneradores.btnIMprimirClick(Sender: TObject);
var
  LstParams:TstringList;
  lParamContrato,lParamFolio:String;
  Tipo:FtGenerador;
  sSeccion: string;
  i:Integer;
  sDatasets:string;
  QrAnexo:TZReadOnlyQuery;
begin

  QrAnexo:=TZReadOnlyQuery.Create(nil);
  QrAnexo.Connection:=connection.zConnection;
  QrAnexo.SQL.Text:='select * from anexos where sTipo=:Tipo and sTierra=:Tierra';

  if connection.contrato.FieldByName('eLugarOt').AsString='Tierra' then
    QrAnexo.ParamByNAme('Tierra').AsString:='Si'
  else
    QrAnexo.ParamByNAme('Tierra').AsString:='No';


  case RdGpGeneradores.ItemIndex of
    0:  begin
          QrAnexo.ParamByNAme('Tipo').AsString:='PERSONAL';
          QrAnexo.Open;
          Tipo:=FtGPersonal;
          TituloGenerador:='NUMEROS GENERADORES DE PERSONAL';
          if QrAnexo.RecordCount>0 then
            TituloGenerador:= TituloGenerador + ' (' +  QrAnexo.FieldByName('sTitulo').AsString + ')';


        end;
    1:  begin
          Tipo:=FtTiempoExtra;
          QrAnexo.ParamByNAme('Tipo').AsString:='PERSONAL';
          QrAnexo.Open;
          TituloGenerador:='NUMEROS GENERADORES DE PERSONAL';
          if QrAnexo.RecordCount>0 then
            TituloGenerador:= TituloGenerador + ' (' +  QrAnexo.FieldByName('sTitulo').AsString + ')';
        end;
    2:  begin
          Tipo:=FtGEquipo;
          QrAnexo.ParamByNAme('Tipo').AsString:='EQUIPO';
          QrAnexo.Open;
          TituloGenerador:='NUMEROS GENERADORES DE EQUIPO';
          if QrAnexo.RecordCount>0 then
            TituloGenerador:= TituloGenerador + ' (' +  QrAnexo.FieldByName('sTitulo').AsString + ')';
        end;
  end;
  sDiarioPeriodo := UpperCase(FormatDateTime('dd', DtEdtFechaInicio.Date))+ ' AL ' + UpperCase(FormatDateTime('dd''-''mmmm''-''yyyy', DtEdtFechaFin.Date));

  LstParams:=TstringList.Create;
  try
    lParamContrato:=QrOrdenes.FieldByName('sContrato').AsString;
    LstParams.Add('CONTRATO='+lParamContrato);
    lParamFolio:=QrFolios.FieldByName('sNumeroOrden').AsString;
    LstParams.Add('FOLIO='+lParamFolio);
    LstParams.Add('COBRO=Si') ;

    LstParams.Add('CONTRATO_BARCO='+global_Contrato_Barco);
    LstParams.Add('INICIO='+DateToStr(DtEdtFechaInicio.Date));
    LstParams.Add('TERMINO='+DateToStr(DtEdtFechaFin.Date));
    LstParams.Add('TIPO=FOLIO');

    TdConfiguracionGenerador(lParamContrato,lParamFolio,FrReporte,DtEdtFechaInicio.Date,DtEdtFechaFin.Date);
    TdGenerador(LstParams,Tipo,FrReporte);

    for I := 0 to FrReporte.DataSets.Count - 1 do
    begin
      if not Assigned(FrReporte.DataSets.Items[i].DataSet) then
        sDatasets:=sDatasets + FrReporte.DataSets.Items[i].DataSetName + '= NILL' + #13 + #10
      else
        sDatasets:=sDatasets + FrReporte.DataSets.Items[i].DataSetName + '= '  + FrReporte.DataSets.Items[i].DataSet.Name+ #13 + #10;
    end;


    rDiarioFirmas(lParamContrato, '', 'A',DtEdtFechaFin.Date , self) ;
    FrReporte.LoadFromFile(Global_Files + global_miReporte + '_TDGeneradorPorFolioH.fr3');
    sDatasets:='';
    for I := 0 to FrReporte.DataSets.Count - 1 do
    begin
      if not Assigned(FrReporte.DataSets.Items[i].DataSet) then
        sDatasets:=sDatasets + FrReporte.DataSets.Items[i].DataSetName + '= NILL' + #13 + #10
      else
        sDatasets:=sDatasets + FrReporte.DataSets.Items[i].DataSetName + '= '  + TfrxDBDataset(FrReporte.DataSets.Items[i].DataSet).DataSet.Name + #13 + #10;
    end;
    FrReporte.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString);
    
  finally
    QrAnexo.Destroy;
    ReportePDF_ClearDataset(FrReporte);
    LstParams.destroy;
  end;

end;

procedure TFrmGeneradores.DtEdtFechaFinPropertiesCloseUp(Sender: TObject);
begin
  if (Sender=nil) then
    InicializeFilters;

  if (Sender<>nil) then
  begin
    if (QrOrdenes.ParamByName('fechaI').AsDateTime<>DtEdtFechaInicio.Date) or
       (QrOrdenes.ParamByName('fechaF').AsDateTime<>DtEdtFechaFin.Date)  then
       InicializeFilters;


  end;
end;

procedure TFrmGeneradores.DtEdtFechaFinPropertiesEditValueChanged(
  Sender: TObject);
begin
  DtEdtFechaFinPropertiesCloseUp(Sender);
end;

procedure TFrmGeneradores.DtEdtFechaInicioPropertiesEditValueChanged(
  Sender: TObject);
begin
  DtEdtFechaFin.Date:=dxGetEndDateOfMonth(DtEdtFechaInicio.Date,true);
end;

procedure TFrmGeneradores.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:=caFree;
end;

procedure TFrmGeneradores.FormCreate(Sender: TObject);
begin
 // GetMonthNumber
 //dxGetStartDateOfMonth
  DtEdtFechaInicio.Date:=dxGetStartDateOfMonth(Now);//EncodeDate(CurrentYear,dxGetMonthNumber(Now),1);
  DtEdtFechaFin.Date:=dxGetEndDateOfMonth(Now,true);//EncodeDate(CurrentYear,dxGetMonthNumber(Now),DayOfTheMonth(EndOfAMonth(CurrentYear,dxGetMonthNumber(Now))));
    
end;

procedure TFrmGeneradores.FormShow(Sender: TObject);
begin
  DtEdtFechaFinPropertiesCloseUp(nil)
end;

procedure TFrmGeneradores.FrReporteGetValue(const VarName: string;
  var Value: Variant);
begin
  if CompareText( VarName,'SUPERINTENDENTE' ) = 0 then
     Value := sSuperIntendente ;
  if CompareText( VarName,'SUPERVISOR' ) = 0 then
    Value := sSupervisorGenerador;
  if CompareText( VarName,'SUPERVISOR_TIERRA' ) = 0 then
    Value := sSupervisorTierra ;

  if CompareText( VarName,'PUESTO_SUPERINTENDENTE' ) = 0 then
       if pos('#', sPuestoSuperIntendente) > 0 then
             Value := copy(sPuestoSuperIntendente,0, pos('#', sPuestoSuperIntendente)-1) +#13+ copy(sPuestoSuperIntendente,pos('#', sPuestoSuperIntendente)+1, length(sPuestoSuperIntendente))
          else
             Value := sPuestoSuperIntendente;

  //  Value := sPuestoSuperIntendente ;





  if CompareText( VarName,'PUESTO_SUPERVISOR' ) = 0 then
    if pos('#', sPuestoSupervisorGenerador) > 0 then
             Value := copy(sPuestoSupervisorGenerador,0, pos('#', sPuestoSupervisorGenerador)-1) +#13+ copy(sPuestoSupervisorGenerador,pos('#', sPuestoSupervisorGenerador)+1, length(sPuestoSupervisorGenerador))
          else
             Value := sPuestoSupervisorGenerador;


  //  Value := sPuestoSupervisorGenerador  ;
  if CompareText( VarName,'PUESTO_SUPERVISOR_TIERRA' ) = 0 then
    if pos('#', sPuestoSupervisorTierra) > 0 then
             Value := copy(sPuestoSupervisorTierra,0, pos('#', sPuestoSupervisorTierra)-1) +#13+ copy(sPuestoSupervisorTierra,pos('#', sPuestoSupervisorTierra)+1, length(sPuestoSupervisorTierra))
          else
             Value := sPuestoSupervisorTierra;
 //   Value := sPuestoSupervisorTierra  ;




  if CompareText( VarName,'PERGENOPT' ) = 0 then
    Value := sDiarioPeriodo  ;
  if CompareText( VarName,'TITULOGENERADOR' ) = 0 then
    Value := TituloGenerador  ;


end;

end.
