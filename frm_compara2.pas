unit frm_compara2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, DBCtrls, StdCtrls, Grids, DBGrids, DB, global,
  Buttons, Mask, ExtCtrls, frxClass, frxDBSet, RXCtrls, frxDMPExport,
  frxCross, ComCtrls, TeEngine, Series, TeeProcs, Chart, DbChart, Newpanel,
  RxMemDS, ZAbstractRODataset, ZDataset, Menus, DateUtils, udbgrid,
  unitexcepciones, UFunctionsGHH,UnitTBotonesPermisos, RxLookup, dxSkinsCore,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinscxPCPainter, cxPCdxBarPopupMenu,
  cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxStyles,
  cxCustomData, cxFilter, cxData, cxDataStorage, cxEdit, cxNavigator, cxDBData,
  cxGridLevel, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxClasses, cxGridCustomView, cxGrid, cxPC, ZAbstractDataset, cxCalc,
  cxTextEdit, dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel,
  dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
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
  dxBarBuiltInMenu;

type
  TfrmComparativo2 = class(TForm)
    ds_avancesglobales: TDataSource;
    GroupBox1: TGroupBox;
    btnExit: TBitBtn;
    btnPrinter: TBitBtn;
    Bevel2: TBevel;
    Label6: TLabel;
    Avances: TfrxDBDataset;
    Catalogo_001: TfrxReport;
    rxGraficaProgramado: TRxMemoryData;
    StringField16: TStringField;
    rxGraficaProgramadodFecha: TDateField;
    FloatField4: TFloatField;
    rxGraficaFisico: TRxMemoryData;
    StringField14: TStringField;
    rxGraficaFisicodFecha: TDateField;
    FloatField7: TFloatField;
    rxGraficaFinanciero: TRxMemoryData;
    StringField2: TStringField;
    DateField2: TDateField;
    FloatField3: TFloatField;
    chkSeries: TGroupBox;
    chkProgramado: TCheckBox;
    chkFisico: TCheckBox;
    chkFinanciero: TCheckBox;
    chkParametros: TGroupBox;
    Label2: TLabel;
    chk3D: TCheckBox;
    chkLeyendas: TCheckBox;
    up3D: TUpDown;
    ti3D: TMaskEdit;
    chkEjes: TCheckBox;
    rxAvancesContrato: TRxMemoryData;
    rxAvancesContratodIdFecha: TDateField;
    rxAvancesContratodProgramadoDia: TFloatField;
    rxAvancesContratodProgramadoAcum: TFloatField;
    rxAvancesContratodFisicoDia: TFloatField;
    rxAvancesContratodFisicoAcumulado: TFloatField;
    SaveSql: TSaveDialog;
    popGraphics: TPopupMenu;
    Exportar1: TMenuItem;
    Label1: TLabel;
    ds_ordenesdetrabajo: TDataSource;
    ordenesdetrabajo: TZReadOnlyQuery;
    tsNumeroOrden: TComboBox;
    chkTurnos: TCheckBox;
    Label9: TLabel;
    tsPlataforma: TRxDBLookupCombo;
    ds_plataformas: TDataSource;
    zqPlataformas: TZReadOnlyQuery;
    chkGeneral: TCheckBox;
    rxAvancesContratosHora: TStringField;
    cxPageAv_Grafica: TcxPageControl;
    cxTabSheet1: TcxTabSheet;
    grGeneral: tNewGroupBox;
    dbGraphicsRespaldo: TDBChart;
    Series1: TFastLineSeries;
    Series2: TFastLineSeries;
    Series3: TFastLineSeries;
    dbGraphics: TDBChart;
    FastLineSeries1: TFastLineSeries;
    FastLineSeries2: TFastLineSeries;
    FastLineSeries3: TFastLineSeries;
    cxTabSheet2: TcxTabSheet;
    Grid_Bitacora: TcxGrid;
    BView_Actividades: TcxGridDBTableView;
    sNumeroActividad: TcxGridDBColumn;
    Horario: TcxGridDBColumn;
    dAvance: TcxGridDBColumn;
    Grid_BitacoraLevel1: TcxGridLevel;
    ds_programados_partidas: TDataSource;
    Fecha: TcxGridDBColumn;
    cxDistribuye: TBitBtn;
    zqAvProgramados: TZQuery;
    zqAvProgramadossContrato: TStringField;
    zqAvProgramadossIdConvenio: TStringField;
    zqAvProgramadosdIdFecha: TDateField;
    zqAvProgramadossNumeroOrden: TStringField;
    zqAvProgramadossPaquete: TStringField;
    zqAvProgramadossWbs: TStringField;
    zqAvProgramadossNumeroActividad: TStringField;
    zqAvProgramadosiNumeroGerencial: TIntegerField;
    zqAvProgramadosdCantidad: TFloatField;
    zqAvProgramadossHorario: TStringField;
    chkTodos: TCheckBox;
    grid_avances: TcxGrid;
    BView_Folios: TcxGridDBTableView;
    cxGridLevel1: TcxGridLevel;
    BView_dIdFecha: TcxGridDBColumn;
    BView_sIdConvenio: TcxGridDBColumn;
    BView_dProgramadoDia: TcxGridDBColumn;
    BView_sHorario: TcxGridDBColumn;
    BView_dProgramadoAcum: TcxGridDBColumn;
    BView_dFisicoDia: TcxGridDBColumn;
    BView_dFisicoAcumulado: TcxGridDBColumn;
    rxAvancesContratosIdConvenio: TStringField;
    cxStyleRepository1: TcxStyleRepository;
    cxFecha: TcxStyle;
    cxEstado: TcxStyle;
    cxConvenio: TcxStyle;
    rxAvancesContratoiNumeroGerencial: TIntegerField;
    zqAvProgramadosdAcumulado: TFloatField;
    dAcumulado: TcxGridDBColumn;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnExitClick(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure btnPrinterClick(Sender: TObject);
    procedure Exportar1Click(Sender: TObject);
    procedure up3DChanging(Sender: TObject; var AllowChange: Boolean);
    procedure chk3DClick(Sender: TObject);
    procedure chkLeyendasClick(Sender: TObject);
    procedure chkEjesClick(Sender: TObject);
    procedure chkProgramadoClick(Sender: TObject);
    procedure chkFisicoClick(Sender: TObject);
    procedure chkFinancieroClick(Sender: TObject);
    procedure tsNumeroOrdenChange(Sender: TObject);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure frxAvancesTotalesGetValue(const VarName: String;
      var Value: Variant);
    procedure frxAvancesGetValue(const VarName: String;
      var Value: Variant);
    procedure Catalogo_001GetValue(const VarName: String;
      var Value: Variant);
    procedure chkTurnosClick(Sender: TObject);
    procedure tsPlataformaChange(Sender: TObject);
    procedure chkGeneralClick(Sender: TObject);
    procedure zqAvProgramadosCalcFields(DataSet: TDataSet);
    procedure rxAvancesContratoAfterScroll(DataSet: TDataSet);
    procedure cxDistribuyeClick(Sender: TObject);
    function  GeneraProgramadosActividad : integer;
  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmComparativo2: TfrmComparativo2;
  BotonPermiso: TBotonesPermisos;
  BotonPermiso2: TBotonesPermisos;
  isOpen : boolean;
  QryGerenciales  : TZReadOnlyQuery;
implementation

{$R *.dfm}

procedure TfrmComparativo2.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cAvOrden');
  BotonPermiso2 := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cAvOrden',popGraphics);
  BotonPermiso.permisosBotones2(nil, nil, nil, btnPrinter);
  BotonPermiso2.permisosBotones(nil);

  QryGerenciales := TZReadOnlyQuery.Create(self);
  QryGerenciales.Connection := connection.zConnection;

  QryGerenciales.Active := False;
  QryGerenciales.SQL.Add('select NumeroGerencial from horarios_gerenciales where Principal = "Si" ');
  QryGerenciales.Open;

  try
    sMenuP:=stMenu;
    OrdenesdeTrabajo.Active := False ;
    OrdenesdeTrabajo.Params.ParamByName('contrato').DataType := ftString ;
    OrdenesdeTrabajo.Params.ParamByName('contrato').Value := global_contrato ;
    OrdenesdeTrabajo.Open ;
    While NOT OrdenesdeTrabajo.Eof Do
    Begin
        tsNumeroOrden.Items.Add(OrdenesdeTrabajo.FieldValues['sNumeroOrden'] ) ;
        OrdenesdeTrabajo.Next ;
    End ;

    if Ordenesdetrabajo.RecordCount > 0 then
    begin
        tsNumeroOrden.ItemIndex := 0;
        tsNumeroOrden.OnChange(sender);

        zqPlataformas.Active := False;
        zqPlataformas.ParamByName('Contrato').AsString := global_contrato;
        zqPlataformas.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
        zqPlataformas.ParamByName('Convenio').AsString := global_convenio;
        zqPlataformas.Open;

        if zQplataformas.RecordCount > 0 then
           tsPlataforma.KeyValue := zqPlataformas.FieldValues['sIdPlataforma'];
    end;

    tsNumeroOrden.Text := 'SELECCIONE ORDEN DE TRABAJO' ;
    tsNumeroOrden.SetFocus
  except
   on e : exception do begin
   UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Comparativo de Avances', 'Al iniciar el formulario', 0);
   end;
  end;
end;

procedure TfrmComparativo2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  BotonPermiso.Free;
  BotonPermiso2.Free;
  action := cafree ;

end;

procedure TfrmComparativo2.btnExitClick(Sender: TObject);
begin
      close
end;

procedure TfrmComparativo2.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        grid_avances.SetFocus
end;

procedure TfrmComparativo2.tsPlataformaChange(Sender: TObject);
begin
    tsNumeroOrden.OnChange(sender);
end;

procedure TfrmComparativo2.btnPrinterClick(Sender: TObject);
begin
   if rxAvancesContrato.RecordCount > 0 then
   begin
       catalogo_001.PreviewOptions.MDIChild := False ;
       catalogo_001.PreviewOptions.Modal := True ;
       catalogo_001.PreviewOptions.ShowCaptions := False ;
       catalogo_001.Previewoptions.ZoomMode := zmPageWidth ;
       catalogo_001.LoadFromFile(Global_Files+global_Mireporte+'_AvancesProgramados.fr3') ;
       catalogo_001.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
   end
   else
       showmessage('No existen registros para imprimir');
end;

procedure TfrmComparativo2.Exportar1Click(Sender: TObject);
begin
  try
  if rxAvancesContrato.RecordCount > 0 then
   begin
    SaveSql.Title := 'Guardar Grafica';
    If SaveSql.Execute Then
         dbGraphics.SaveToBitmapFile(SaveSql.FileName);
   end
   else
   showmessage('No existen datos para exportar');
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Comparativo de Avances', 'Al exportar gráfica', 0);
    end;
  end;
end;

procedure TfrmComparativo2.up3DChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
    dbGraphics.Chart3DPercent := Up3D.Position ;

end;

procedure TfrmComparativo2.zqAvProgramadosCalcFields(DataSet: TDataSet);
var
   dAcumuladoAnt : double;
begin
    if zqAvProgramados.RecordCount > 0 then
    begin
        if zqAvProgramados.FieldByName('iNumeroGerencial').AsInteger  = 1 then
           zqAvProgramados.FieldValues['sHorario'] := '05:00';

        if zqAvProgramados.FieldByName('iNumeroGerencial').AsInteger  = 2 then
           zqAvProgramados.FieldValues['sHorario'] := '17:00';

        if zqAvProgramados.FieldByName('iNumeroGerencial').AsInteger  = 3 then
           zqAvProgramados.FieldValues['sHorario'] := '24:00';

        dAcumuladoAnt  := 0;
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select iNumeroGerencial, sum(dCantidad) as Cantidad '+
                                'from distribuciondeactividades '+
                                'where sContrato =:Contrato and sIdConvenio =:convenio '+
                                'and sNumeroOrden =:Orden and dIdFecha <:Fecha and sWbs = :Wbs '+
                                'group by sWbs '+
                                'order by dIdFecha, sWbs, iNumeroGerencial');
        connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
        connection.zCommand.ParamByName('Convenio').AsString := rxAvancesContrato.FieldByName('sIdConvenio').AsString;
        connection.zCommand.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
        connection.zCommand.ParamByName('Fecha').AsDate      := rxAvancesContrato.FieldByName('dIdFecha').AsDateTime;
        connection.zCommand.ParamByName('Wbs').AsString      := zqAvProgramados.FieldByName('sWbs').AsString;
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
           dAcumuladoAnt := dAcumuladoAnt + connection.zCommand.FieldByName('Cantidad').AsFloat ;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select iNumeroGerencial, sum(dCantidad) as Cantidad '+
                                'from distribuciondeactividades '+
                                'where sContrato =:Contrato and sIdConvenio =:convenio '+
                                'and sNumeroOrden =:Orden and dIdFecha =:Fecha and sWbs = :Wbs and iNumeroGerencial <= :Num '+
                                'group by sWbs '+
                                'order by dIdFecha, sWbs, iNumeroGerencial');
        connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
        connection.zCommand.ParamByName('Convenio').AsString := rxAvancesContrato.FieldByName('sIdConvenio').AsString;
        connection.zCommand.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
        connection.zCommand.ParamByName('Fecha').AsDate      := rxAvancesContrato.FieldByName('dIdFecha').AsDateTime;
        connection.zCommand.ParamByName('Wbs').AsString      := zqAvProgramados.FieldByName('sWbs').AsString;
        connection.zCommand.ParamByName('Num').AsInteger     := zqAvProgramados.FieldByName('iNumeroGerencial').AsInteger;
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
           dAcumuladoAnt := dAcumuladoAnt + connection.zCommand.FieldByName('Cantidad').AsFloat ;

        zqAvProgramados.FieldByName('dAcumulado').AsFloat :=  dAcumuladoAnt;

    end;
end;

procedure TfrmComparativo2.chk3DClick(Sender: TObject);
begin
    dbGraphics.View3D := chk3d.Checked ;
end;

procedure TfrmComparativo2.chkLeyendasClick(Sender: TObject);
begin
    dbGraphics.Series[0].Marks.Visible := chkLeyendas.Checked ;
    dbGraphics.Series[1].Marks.Visible := chkLeyendas.Checked ;
    dbGraphics.Series[2].Marks.Visible := chkLeyendas.Checked ;
end;

procedure TfrmComparativo2.chkEjesClick(Sender: TObject);
begin
    dbGraphics.LeftAxis.Visible := chkEjes.Checked
end;

procedure TfrmComparativo2.chkProgramadoClick(Sender: TObject);
begin
    dbGraphics.Series[0].Active := chkProgramado.Checked ;
end;

procedure TfrmComparativo2.chkTurnosClick(Sender: TObject);
begin
    tsNumeroOrden.OnChange(sender);
end;

procedure TfrmComparativo2.cxDistribuyeClick(Sender: TObject);
begin
    if chkTodos.Checked then
    begin
        rxAvancesContrato.first;
        while not rxAvancesContrato.Eof do
        begin
            GeneraProgramadosActividad;
            rxavancesContrato.Next;
        end;
    end
    else
       GeneraProgramadosActividad;
    zqAvProgramados.Refresh;
end;

procedure TfrmComparativo2.chkFisicoClick(Sender: TObject);
begin
    dbGraphics.Series[1].Active := chkFisico.Checked ;
end;

procedure TfrmComparativo2.chkGeneralClick(Sender: TObject);
begin
   tsNumeroOrden.OnChange(sender);
end;

procedure TfrmComparativo2.chkFinancieroClick(Sender: TObject);
begin
    dbGraphics.Series[2].Active := chkFinanciero.Checked ;
end;

procedure TfrmComparativo2.tsNumeroOrdenChange(Sender: TObject);
var
    sFecha     : String ;
    iMiMes     : Byte ;
    dAcumulado : Currency ;
    dAcumuladoFisico : Currency ;
    dAvanceFisico    : Currency ;
    QryBuscarTurnos  : TZReadOnlyQuery;

    dProgramadoActual,
    dProgramadoAnterior,
    dProgramadoAcumulado,
    dProgramadoAcumulado_Aux : Currency;
    sTurno : string;
    dFecha : tdate;
begin

try
      IsOpen:=false;
      rxAvancesContrato.DisableControls;

      try
      // Primero Genera la Grafica ....

        Caption := tsNumeroOrden.Text  + '-' + connection.configuracion.FieldValues['sNombre'] +']' ;

        SaveSql.FileName := global_contrato ;
        If rxGraficaProgramado.RecordCount > 0 then
          rxGraficaProgramado.EmptyTable   ;

        If rxGraficaFisico.RecordCount > 0 then
          rxGraficaFisico.EmptyTable  ;

        If rxGraficaFinanciero.RecordCount > 0 then
            rxGraficaFinanciero.EmptyTable  ;

        If rxAvancesContrato.RecordCount > 0 then
            rxAvancesContrato.EmptyTable  ;

        dbGraphics.RefreshData ;

        QryBuscarTurnos := TZReadOnlyQuery.Create(self);
        QryBuscarTurnos.Connection := connection.zConnection;     

        Connection.qryBusca.Active := False ;
        Connection.qryBusca.SQL.Clear ;
        if chkTurnos.Checked then
           Connection.qryBusca.SQL.Add('Select a.dIdFecha, a.iNumeroGerencial, a.dAvancePonderadoDia, a.dAvancePonderadoGlobal, af.dAvance, r.sIdTurno, a.sIdConvenio ' +
                                       'From avancesglobales a ' +
                                       'left join reportediario r on (r.sContrato = a.sContrato and r.sNumeroOrden = a.sNumeroOrden and r.dIdFecha = a.dIdfecha ) '+
                                       'left join avancesglobalesxorden af on (a.sContrato = af.sContrato and a.sNumeroOrden = af.sNumeroOrden and a.dIdFecha = af.dIdFecha and af.sIdTurno = r.sIdTurno) '+
                                       'Where a.sContrato = :Contrato And a.sIdConvenio = :Convenio And a.sNumeroOrden = :Orden order by a.dIdFecha' )
        else
            Connection.qryBusca.SQL.Add('Select a.dIdFecha,  a.iNumeroGerencial, a.dAvancePonderadoDia, a.dAvancePonderadoGlobal, sum(af.dAvance) as dAvance, a.sIdConvenio ' +
                                       'From avancesglobales a ' +
                                       'left join avancesglobalesxorden af on (a.sContrato = af.sContrato and a.sIdConvenio = af.sIdConvenio and a.sNumeroOrden = af.sNumeroOrden and a.dIdFecha = af.dIdFecha ) '+
                                       'Where a.sContrato = :Contrato And a.sNumeroOrden = :Orden group by a.dIdFecha, a.iNumeroGerencial order by a.dIdFecha,a.iNumeroGerencial ');
        Connection.qryBusca.params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca.params.ParamByName('Contrato').Value    := global_contrato ;
        Connection.qryBusca.params.ParamByName('Orden').DataType    := ftString ;
        Connection.qryBusca.params.ParamByName('Orden').Value       := tsNumeroOrden.Text ;
        Connection.qryBusca.Open ;

        If connection.QryBusca.RecordCount > 0 Then
            iMiMes := MonthOf(Connection.QryBusca.FieldValues['dIdFecha'])
        Else
            iMiMes := 0 ;
        dAcumuladoFisico := 0 ;
        dProgramadoAnterior := 0;
        While NOT Connection.qryBusca.Eof Do
        Begin
            If iMiMes <> MonthOf(Connection.QryBusca.FieldValues['dIdFecha']) Then
            Begin
                iMiMes := MonthOf(Connection.QryBusca.FieldValues['dIdFecha']) ;
                rxGraficaProgramado.Append ;
                rxGraficaProgramado.FieldValues['sDescripcion'] := global_contrato ;
                rxGraficaProgramado.FieldValues['dFecha']       := Connection.QryBusca.FieldValues['dIdFecha'] - 1 ;
                rxGraficaProgramado.FieldValues['dProgramado']  := dAcumulado ;
                rxGraficaProgramado.Post ;
                dAcumulado := 0 ;
            End ;
            dAvanceFisico := 0 ;
            If Connection.QryBusca.FieldValues['dAvance'] <> Null Then
            begin
                dAcumuladoFisico := dAcumuladoFisico + Connection.QryBusca.FieldValues['dAvance'] ;
                dAvanceFisico    := Connection.QryBusca.FieldValues['dAvance']
            End ;

            dProgramadoActual    := Connection.QryBusca.FieldValues['dAvancePonderadoDia'];
            dProgramadoAcumulado := Connection.QryBusca.FieldValues['dAvancePonderadoGlobal'];

            rxAvancesContrato.Append ;
            rxAvancesContrato.FieldValues['dIdFecha']    := Connection.QryBusca.FieldValues['dIdFecha'] ;
            rxAvancesContrato.FieldValues['sIdConvenio'] := Connection.QryBusca.FieldByName('sIdConvenio').AsString;
            rxAvancesContrato.FieldValues['sHora']       := '00:00';
            rxAvancesContrato.FieldValues['iNumeroGerencial'] := Connection.QryBusca.FieldByName('iNumeroGerencial').AsInteger;

            QryGerenciales.First;
            while not QryGerenciales.Eof do
            begin
               if Connection.QryBusca.FieldByName('iNumeroGerencial').AsInteger  = QryGerenciales.FieldByName('NumeroGerencial').AsInteger then
                  rxAvancesContrato.FieldValues['sHora'] := QryGerenciales.FieldByName('NumeroGerencial').AsString;
               QryGerenciales.Next;
            end;

            rxAvancesContrato.FieldValues['dProgramadoDia']   := dProgramadoActual;
            rxAvancesContrato.FieldValues['dProgramadoAcum']  := dProgramadoAcumulado;
            rxAvancesContrato.FieldValues['dFisicoDia']       := dAvanceFisico;
            rxAvancesContrato.FieldValues['dFisicoAcumulado'] := dAcumuladoFisico;
            rxAvancesContrato.Post ;

            dAcumulado := Connection.QryBusca.FieldValues['dAvancePonderadoGlobal'] ;

            Connection.qryBusca.Next
        End ;
        QryBuscarTurnos.Destroy;


        If dAcumulado <> 0 Then
            With Connection.qryBusca DO
            begin
                If MonthOf(FieldValues['dIdFecha']) <= 8 Then
                   sFecha := '01/0' + Trim(IntToStr(MonthOf(FieldValues['dIdFecha']) + 1))  + '/' + Trim(IntToStr(YearOf(FieldValues['dIdFecha'])))
                Else
                   If MonthOf(FieldValues['dIdFecha']) <= 11 Then
                       sFecha := '01/' + Trim(IntToStr(MonthOf(FieldValues['dIdFecha']) + 1)) + '/' + Trim(IntToStr(YearOf(FieldValues['dIdFecha'])))
                    Else
                       sFecha := '01/01/' + Trim(IntToStr(YearOf(FieldValues['dIdFecha']) + 1 )) ;
                sFecha := DateToStr(StrToDate(sFecha) - 1) ;
                rxGraficaProgramado.Append ;
                rxGraficaProgramado.FieldValues['sDescripcion'] := global_contrato ;
                rxGraficaProgramado.FieldValues['dFecha']       := sFecha ;
                rxGraficaProgramado.FieldValues['dProgramado']  := dAcumulado ;
                rxGraficaProgramado.Post ;

          End ;
         // El resto del avance Fisico .....

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select a.dIdFecha, sum((b.dPonderado / 100)* a.dCantidad) as dAvance, a.iNumeroGerencial '+
                   ' From actividadesxorden b '+
                   ' inner JOIN bitacoradeactividades a '+
                   ' ON (a.sContrato=b.sContrato and b.sIdConvenio = a.sIdConvenio And a.sWbs=b.sWbs and b.sNumeroOrden=a.sNumeroOrden ) '+
                   ' left JOIN tiposdemovimiento t '+
                   ' ON (b.sContrato=t.sContrato And a.sIdTipoMovimiento=t.sIdTipoMovimiento And t.sClasificacion="Tiempo en Operacion") '+
                   ' Where b.sContrato=:Contrato And b.sNumeroOrden =:Orden '+
                   ' Group By a.dIdFecha, a.iNumeroGerencial Order By a.dIdFecha' ) ;
        Connection.qryBusca2.params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca2.params.ParamByName('Contrato').Value    := global_contrato ;
        Connection.qryBusca2.params.ParamByName('Orden').DataType    := ftString ;
        Connection.qryBusca2.params.ParamByName('Orden').Value       := tsNumeroOrden.Text ;
        Connection.qryBusca2.Open ;

        dFecha := 0;
        rxAvancesContrato.First;
        while not rxAvancesContrato.Eof do
        begin
            dAvanceFisico := 0 ;
            Connection.QryBusca2.First;
            While NOT Connection.qryBusca2.Eof Do
            Begin
                if Connection.QryBusca2.FieldValues['dIdFecha'] = rxAvancesContrato.FieldValues['dIdFecha'] then
                begin
                    if Connection.QryBusca2.FieldValues['iNumeroGerencial'] = rxAvancesContrato.FieldValues['iNumeroGerencial'] then
                    begin
                        If Connection.QryBusca2.FieldValues['dAvance'] <> Null Then
                        begin
                            dAcumuladoFisico := dAcumuladoFisico + Connection.QryBusca2.FieldValues['dAvance'] ;
                            dAvanceFisico := Connection.QryBusca2.FieldValues['dAvance']
                        End;
                       // dFecha := rxAvancesContrato.FieldValues['dIdFecha'];
                    end;
                end;
                rxAvancesContrato.Edit;
                rxAvancesContrato.FieldValues['dFisicoDia']       := dAvanceFisico ;
                rxAvancesContrato.FieldValues['dFisicoAcumulado'] := dAcumuladoFisico ;
                rxAvancesContrato.Post;
                Connection.qryBusca2.Next
            End ;
            rxAvancesContrato.Next;
        end;


        // Real ...
        Connection.qryBusca.Active := False ;
        Connection.qryBusca.SQL.Clear ;
        Connection.qryBusca.SQL.Add('Select year(dIdFecha) as dAnno , month(dIdFecha) as dMes From bitacoradeactividades Where ' +
                                    'sContrato = :Contrato And sNumeroOrden = :Orden Group By Year(dIdFecha), month(dIdFecha)' ) ;
        Connection.qryBusca.params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca.params.ParamByName('Contrato').Value    := global_contrato ;
        Connection.qryBusca.params.ParamByName('Orden').DataType    := ftString ;
        Connection.qryBusca.params.ParamByName('Orden').Value       := tsNumeroOrden.Text ;
        Connection.qryBusca.Open ;
        While NOT Connection.qryBusca.Eof Do
        Begin
           If Connection.qryBusca.FieldValues['dMes'] <= 8 Then
              sFecha := '01/0' + Trim(IntToStr(Connection.qryBusca.FieldValues['dMes'] + 1))  + '/' + Connection.qryBusca.fieldByName('dAnno').AsString
           Else
              If Connection.qryBusca.FieldValues['dMes'] <= 11 Then
                  sFecha := '01/' + Trim(IntToStr(Connection.qryBusca.FieldValues['dMes'] + 1)) + '/' + Connection.qryBusca.fieldByName('dAnno').AsString
              Else
                  sFecha := '01/01/' + Trim(IntToStr(Connection.qryBusca.FieldValues['dAnno'] + 1)) ;

              sFecha := DateToStr(StrToDate(sFecha) - 1) ;
              Connection.qryBusca2.Active := False ;
              Connection.qryBusca2.SQL.Clear ;
              Connection.qryBusca2.SQL.Add('Select sum((b.dPonderado / 100)* a.dCantidad) as dMensual '+
                   ' From actividadesxorden b '+
                   ' inner JOIN bitacoradeactividades a ');
              Connection.qryBusca2.SQL.Add(' ON (a.sContrato=b.sContrato And a.sWbs=b.sWbs and b.sNumeroOrden=a.sNumeroOrden and a.dIdFecha <=:fecha) ');
              Connection.qryBusca2.SQL.Add(' left JOIN tiposdemovimiento t '+
                   ' ON (b.sContrato=t.sContrato And a.sIdTipoMovimiento=t.sIdTipoMovimiento And t.sClasificacion="Tiempo en Operacion") '+
                   ' Where b.sContrato=:Contrato And b.sNumeroOrden =:Orden Group By a.sContrato'  ) ;
              Connection.qryBusca2.params.ParamByName('Contrato').DataType := ftString ;
              Connection.qryBusca2.params.ParamByName('Contrato').Value    := global_contrato ;
              Connection.qryBusca2.params.ParamByName('Orden').DataType    := ftString ;
              Connection.qryBusca2.params.ParamByName('Orden').Value       := tsNumeroOrden.Text ;
              Connection.qryBusca2.params.ParamByName('Fecha').DataType    := ftDate ;
              Connection.qryBusca2.params.ParamByName('Fecha').Value       := StrToDate(sFecha) ;
              Connection.qryBusca2.Open ;
              If Connection.qryBusca.RecordCount > 0 Then
              Begin
                 rxGraficaFisico.Append ;
                 rxGraficaFisico.FieldValues['sDescripcion'] := global_contrato ;
                 rxGraficaFisico.FieldValues['dFecha']       := sFecha ;
                 rxGraficaFisico.FieldValues['dFisico']      := Connection.qryBusca2.FieldValues['dMensual'] ;
                 rxGraficaFisico.Post ;
              End ;
              Connection.qryBusca.Next
         End ;

         // Financiero ....
         Connection.qryBusca.Active := False ;
         Connection.qryBusca.SQL.Clear ;
         Connection.qryBusca.SQL.Add('Select year(dFechaFinal) as dAnno , month(dFechaFinal) as dMes From estimaciones Where ' +
                                     'sContrato = :Contrato And sNumeroOrden = :orden Group By Year(dFechaFinal), month(dFechaFinal)' ) ;
         Connection.qryBusca.params.ParamByName('Contrato').DataType := ftString ;
         Connection.qryBusca.params.ParamByName('Contrato').Value := global_contrato ;
         Connection.qryBusca.params.ParamByName('orden').DataType := ftString ;
         Connection.qryBusca.params.ParamByName('orden').Value := tsNumeroOrden.Text ;
         Connection.qryBusca.Open ;
         While NOT Connection.qryBusca.Eof Do
         Begin
             If Connection.qryBusca.FieldValues['dMes'] <= 8 Then
                 sFecha := '01/0' + Trim(IntToStr(Connection.qryBusca.FieldValues['dMes'] + 1))  + '/' + Connection.qryBusca.fieldByName('dAnno').AsString
             Else
                 If Connection.qryBusca.FieldValues['dMes'] <= 11 Then
                     sFecha := '01/' + Trim(IntToStr(Connection.qryBusca.FieldValues['dMes'] + 1)) + '/' + Connection.qryBusca.fieldByName('dAnno').AsString
                 Else
                     sFecha := '01/01/' + Trim(IntToStr(Connection.qryBusca.FieldValues['dAnno'] + 1)) ;
             sFecha := DateToStr(StrToDate(sFecha) - 1) ;

             Connection.qryBusca2.Active := False ;
             Connection.qryBusca2.SQL.Clear ;
             Connection.qryBusca2.SQL.Add('Select Sum(dMontoMN) as dReal From estimaciones ' +
                                          'Where sContrato = :Contrato And dFechaFinal <= :Fecha And sNumeroGenerador NOT Like "%A%" and sNumeroOrden = :orden Group By sContrato' ) ;

             Connection.qryBusca2.params.ParamByName('Contrato').DataType := ftString ;
             Connection.qryBusca2.params.ParamByName('Contrato').Value := global_contrato ;
             Connection.qryBusca2.params.ParamByName('Fecha').DataType := ftDate ;
             Connection.qryBusca2.params.ParamByName('Fecha').Value := strToDate(sFecha) ;
             Connection.qryBusca2.params.ParamByName('orden').DataType := ftString ;
             Connection.qryBusca2.params.ParamByName('orden').Value := tsNumeroOrden.Text ;
             Connection.qryBusca2.Open ;
             If Connection.qryBusca2.RecordCount > 0 Then
             Begin
                try
                   rxGraficaFinanciero.Append ;
                   rxGraficaFinanciero.FieldValues['sDescripcion'] := global_contrato ;
                   rxGraficaFinanciero.FieldValues['dFecha'] := sFecha ;
                   rxGraficaFinanciero.FieldValues['dFinanciero'] := (Connection.qryBusca2.FieldValues['dReal'] / dMontoContrato) * 100 ;
                   rxGraficaFinanciero.Post ;
                Except
                    //No hace nada, con esto se evita el mensaje de error, division por cero.. 17 Febrero de 2011
                end;
             End ;
             Connection.qryBusca.Next
         End ;

         dbGraphics.Title.Text.Clear ;
         dbGraphics.Title.Text.Add ('Avances Programado/Fisico/Financiero') ;
         dbGraphics.Title.Text.Add (Caption) ;
         dbGraphics.Title.Text.Add (connection.contrato.FieldValues['mDescripcion']) ;

         rxAvancesContrato.Locate('dIdFecha', Date() , [loPartialKey]) ;
        except
        on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Comparativo de Avances', 'Al generar la gráfica', 0);
        end;
     end;
finally
   rxAvancesContrato.EnableControls;
   IsOpen:=true;
   rxAvancesContratoAfterScroll(rxAvancesContrato);
end;

end;

procedure TfrmComparativo2.tsNumeroOrdenEnter(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmComparativo2.tsNumeroOrdenExit(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_salida;
end;

procedure TfrmComparativo2.frxAvancesTotalesGetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'ORDEN_TRABAJO') = 0 then
      Value := tsNumeroOrden.Text ;
end;

procedure TfrmComparativo2.rxAvancesContratoAfterScroll(DataSet: TDataSet);
begin
    if isOpen then
    begin
        if rxAvancesContrato.RecordCount > 0 then
        begin
            zqAvProgramados.Active := False;
            zqAvProgramados.SQL.Clear;
            zqAvProgramados.SQL.Add('select b.*, 0 as dAcumulado from distribuciondeactividades b where b.sContrato =:Contrato and b.sIdConvenio =:convenio and b.sNumeroOrden =:Orden and b.dIdFecha =:Fecha order by b.dIdFecha, b.sWbs, b.iNumeroGerencial ');
            zqAvProgramados.ParamByName('Contrato').AsString := global_contrato;
            zqAvProgramados.ParamByName('Convenio').AsString := rxAvancesContrato.FieldByName('sIdConvenio').AsString;
            zqAvProgramados.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
            zqAvProgramados.ParamByName('Fecha').AsDate      := rxAvancesContrato.FieldByName('dIdFecha').AsDateTime;
            zqAvProgramados.Open;
        end;
    end;
end;

procedure TfrmComparativo2.frxAvancesGetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'ORDEN_TRABAJO') = 0 then
      Value := tsNumeroOrden.Text ;
end;

procedure TfrmComparativo2.Catalogo_001GetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'ORDEN_TRABAJO') = 0 then
      Value := tsNumeroOrden.Text ;
end;

function TfrmComparativo2.GeneraProgramadosActividad;
var
   num     : integer;
   dFactor : double;
begin
    zqAvProgramados.First;
    while not zqAvProgramados.Eof do
    begin
        for num := 1 to 3 do
        begin
            if num = 1 then
               dFactor := 0.208333;

            if num = 2 then
               dFactor := 0.5;

            if num = 3 then
               dFactor := 0.291667;

            if (num = 1) or (num = 2)  then
            begin
                try
                  connection.zCommand.Active := False ;
                  connection.zCommand.SQL.Clear ;
                  connection.zCommand.SQL.Add ( 'INSERT INTO distribuciondeactividades ( sContrato , sIdConvenio, sNumeroOrden, dIdFecha, sWbs, sNumeroActividad, iNumeroGerencial, dCantidad ) ' +
                                              ' VALUES (:contrato, :convenio, :orden, :fecha, :Wbs, :Actividad, :Numero, :cantidad)') ;
                  connection.zCommand.Params.ParamByName('contrato').value     := Global_Contrato ;
                  connection.zCommand.Params.ParamByName('convenio').value     := zqAvProgramados.FieldByName('sIdConvenio').AsString ;
                  connection.zCommand.Params.ParamByName('orden').value        := zqAvProgramados.FieldByName('sNumeroOrden').AsString ;
                  connection.zCommand.Params.ParamByName('Wbs').value          := zqAvProgramados.FieldByName('sWbs').AsString ;
                  connection.zCommand.Params.ParamByName('Actividad').value    := zqAvProgramados.FieldByName('sNumeroActividad').AsString ;
                  connection.zCommand.Params.ParamByName('fecha').value        := zqAvProgramados.FieldByName('dIdFecha').AsDateTime ;
                  connection.zCommand.Params.ParamByName('Numero').value       := num ;
                  connection.zCommand.Params.ParamByName('cantidad').value     := zqAvProgramados.FieldByName('dCantidad').AsFloat * dFactor ;
                  connection.zCommand.ExecSQL ;
                  result := 0;
                Except
                     result := 1;
                     messageDLG('Existe distribución de horarios, Deberá correr la distribución de actividades.', mtInformation, [mbOk], 0);
                     exit;
                end;
            end
            else
            begin
                connection.zCommand.Active := False ;
                connection.zCommand.SQL.Clear ;
                connection.zCommand.SQL.Add ( 'UPDATE distribuciondeactividades SET dCantidad = :Cantidad where sContrato = :contrato And iNumeroGerencial = 3 and ' +
                                              'dIdFecha = :fecha And sIdConvenio = :Convenio And sNumeroOrden = :Orden And sWbs = :Wbs And sNumeroActividad = :Actividad ') ;
                connection.zCommand.Params.ParamByName('contrato').DataType  := ftString ;
                connection.zCommand.Params.ParamByName('contrato').Value     := global_contrato ;
                connection.zCommand.Params.ParamByName('convenio').DataType  := ftString ;
                connection.zCommand.Params.ParamByName('convenio').Value     := zqAvProgramados.FieldByName('sIdConvenio').AsString ;
                connection.zCommand.Params.ParamByName('fecha').DataType     := ftDate ;
                connection.zCommand.Params.ParamByName('fecha').Value        := zqAvProgramados.FieldByName('dIdFecha').AsDateTime ;
                connection.zCommand.Params.ParamByName('Orden').DataType     := ftString ;
                connection.zCommand.Params.ParamByName('Orden').Value        := zqAvProgramados.FieldByName('sNumeroOrden').AsString ;
                connection.zCommand.Params.ParamByName('wbs').DataType       := ftString ;
                connection.zCommand.Params.ParamByName('wbs').Value          := zqAvProgramados.FieldByName('sWbs').AsString  ;
                connection.zCommand.Params.ParamByName('actividad').DataType := ftString ;
                connection.zCommand.Params.ParamByName('actividad').Value    := zqAvProgramados.FieldByName('sNumeroActividad').AsString ;
                connection.zCommand.Params.ParamByName('Cantidad').DataType  := ftFloat ;
                connection.zCommand.Params.ParamByName('Cantidad').value     := zqAvProgramados.FieldByName('dCantidad').AsFloat * dFactor ;
                connection.zCommand.ExecSQL ;
                result := 0;
            end;
        end;
        zqAvProgramados.Next;
    end;
end;

end.
