unit frm_comparativo;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection , global, StdCtrls, DBCtrls, DB, Menus, frxClass,
  frxDBSet, ComCtrls, RxMemDS, RXCtrls, Buttons, StrUtils, fqbClass, fqbSynmemo,
  ExtCtrls, Newpanel, DateUtils, ZAbstractRODataset, ZDataset, utilerias,
  Gauges, ComObj, Excel2000, JvExStdCtrls, UFunctionsGHH, unitexcepciones,
  UnitTBotonesPermisos, DBDateTimePicker, JvCheckBox, Mask, NxCollection,
  RxLookup, RxToolEdit;
type
  TfrmCompara = class(TForm)
    dsInforme: TfrxDBDataset;
    rxPartidasAvance: TRxMemoryData;
    rxPartidasAvancesContrato: TStringField;
    StringField8: TStringField;
    rxPartidasAvancemDescripcion: TMemoField;
    StringField9: TStringField;
    rxPartidasAvancedPonderado: TFloatField;
    rxPartidasAvancedCantidadAnexo: TFloatField;
    rxPartidasAvancedCantidadProgramada: TFloatField;
    rxPartidasAvancedAvanceProgramado: TFloatField;
    rxPartidasAvancedCantidadReal: TFloatField;
    rxPartidasAvancedAvanceReal: TFloatField;
    rxPartidasAvancedFechaInicio: TDateField;
    rxPartidasAvancedFechaFinal: TDateField;
    rxPartidasAvanceiRetraso: TIntegerField;
    rxCantProgramada: TRxMemoryData;
    rxCantProgramadasContrato: TStringField;
    rxCantProgramadasIdClave: TStringField;
    rxCantProgramadasDescripcion: TStringField;
    rxCantProgramadadCantidad: TFloatField;
    rxCantProgramadasRenglon: TStringField;
    rxCantProgramadasMedida: TStringField;
    rxCantProgramadaiMes: TIntegerField;
    rxCantProgramadaiAnno: TIntegerField;
    dsCantProgramada: TfrxDBDataset;
    rxCantProgramadaiItemOrden: TIntegerField;
    rxCostoProgramado: TRxMemoryData;
    StringField1: TStringField;
    IntegerField1: TIntegerField;
    StringField2: TStringField;
    StringField3: TStringField;
    StringField4: TStringField;
    IntegerField2: TIntegerField;
    IntegerField3: TIntegerField;
    StringField5: TStringField;
    rxCostoProgramadodCostoMN: TCurrencyField;
    dbCostoProgramado: TfrxDBDataset;
    rxPartidasAvancedVentaMN: TCurrencyField;
    rxPartidasAvancedVentaDLL: TCurrencyField;
    rxAnexoGenerado: TRxMemoryData;
    StringField6: TStringField;
    StringField7: TStringField;
    MemoField1: TMemoField;
    StringField10: TStringField;
    FloatField1: TFloatField;
    CurrencyField1: TCurrencyField;
    rxAnexoGeneradodGenerado: TFloatField;
    rxAnexoGeneradodPendiente: TFloatField;
    rxAnexoGeneradodAdicional: TFloatField;
    rxAnexoGeneradodPonderado: TFloatField;
    rxSuministroAnexo: TRxMemoryData;
    StringField11: TStringField;
    StringField12: TStringField;
    MemoField2: TMemoField;
    StringField13: TStringField;
    FloatField2: TFloatField;
    CurrencyField2: TCurrencyField;
    FloatField6: TFloatField;
    FloatField8: TFloatField;
    rxSuministroAnexosReferencia: TStringField;
    rxSuministroAnexodCantidad: TFloatField;
    rxSuministroAnexodPReportar: TFloatField;
    rxSuministroAnexodPSuministrar: TFloatField;
    ActividadesxAnexo: TZReadOnlyQuery;
    ActividadesxOrden: TZReadOnlyQuery;
    ActividadesxOrdensContrato: TStringField;
    ActividadesxOrdeniNivel: TIntegerField;
    ActividadesxOrdeniColor: TIntegerField;
    ActividadesxOrdensTipoActividad: TStringField;
    ActividadesxOrdensWbs: TStringField;
    ActividadesxOrdensNumeroActividad: TStringField;
    ActividadesxOrdenmDescripcion: TMemoField;
    ActividadesxOrdensMedida: TStringField;
    ActividadesxOrdendCantidad: TFloatField;
    ActividadesxOrdendPonderado: TFloatField;
    ActividadesxOrdendVentaMN: TFloatField;
    ActividadesxOrdendReportado: TFloatField;
    ActividadesxOrdendSuministrado: TFloatField;
    ActividadesxOrdendGenerado: TFloatField;
    ActividadesxOrdensNumeroOrden: TStringField;
    dbActividadesxAnexo: TfrxDBDataset;
    ActividadesxOrdensWbsAnterior: TStringField;
    ActividadesxOrdendFechaInicio: TDateField;
    ActividadesxOrdendFechaFinal: TDateField;
    ActividadesxOrdendMontoMN: TCurrencyField;
    dbActividadesxOrden: TfrxDBDataset;
    rxPartidasAvanceiNivel: TIntegerField;
    rxPartidasAvancesWbsAnterior: TStringField;
    rxPartidasAvancesNumeroActividad: TStringField;
    rxPartidasAvancesTipoActividad: TStringField;
    Reporte: TZReadOnlyQuery;
    ActividadesxOrdendSubContrato: TFloatField;
    rInforme: TfrxReport;
    ActividadesxAnexosContrato: TStringField;
    ActividadesxAnexoiNivel: TIntegerField;
    ActividadesxAnexoiColor: TIntegerField;
    ActividadesxAnexosTipoActividad: TStringField;
    ActividadesxAnexosWbsAnterior: TStringField;
    ActividadesxAnexosWbs: TStringField;
    ActividadesxAnexosNumeroActividad: TStringField;
    ActividadesxAnexomDescripcion: TMemoField;
    ActividadesxAnexodFechaInicio: TDateField;
    ActividadesxAnexodFechaFinal: TDateField;
    ActividadesxAnexosMedida: TStringField;
    ActividadesxAnexodCantidadAnexo: TFloatField;
    ActividadesxAnexodPonderado: TFloatField;
    ActividadesxAnexodVentaMN: TFloatField;
    ActividadesxAnexodVentaDLL: TFloatField;
    ActividadesxAnexodMontoMN: TCurrencyField;
    ActividadesxAnexodMontoDLL: TCurrencyField;
    ActividadesxAnexodReportado: TFloatField;
    ActividadesxAnexodSuministrado: TFloatField;
    ActividadesxAnexodGenerado: TFloatField;
    ActividadesxAnexodEstimado: TFloatField;
    ActividadesxAnexodSubContrato: TFloatField;
    rxAnexoGeneradolTitulo: TStringField;
    qryBuscaP: TZReadOnlyQuery;
    dsPartidas: TfrxDBDataset;
    ActividadesxOrdendMontoDLL: TFloatField;
    ActividadesxOrdendVentaDLL: TFloatField;
    rxAnexoGeneradodVentaDLL: TFloatField;
    roqOrdenes: TZReadOnlyQuery;
    rx_Isometricos: TRxMemoryData;
    rx_IsometricossContrato: TStringField;
    rx_IsometricossWbs: TStringField;
    rx_IsometricossNumeroOrden: TStringField;
    rx_IsometricosmIsometrico: TStringField;
    rx_IsometricossNumeroActividad: TStringField;
    rx_IsometricosmDescripcion: TStringField;
    rx_IsometricosdIdFecha: TDateField;
    rx_IsometricosdCantidad: TFloatField;
    rx_IsometricosdReportado: TFloatField;
    rx_IsometricosdAvanceReportado: TFloatField;
    rx_IsometricosdGenerado: TFloatField;
    rx_IsometricosdAvanceGenerado: TFloatField;
    rx_IsometricossWbsAnterior: TStringField;
    ds_isometricos: TfrxDBDataset;
    RxMDValida: TRxMemoryData;
    RxMDValidasNumeroActividad: TStringField;
    RxMDValidasWbs: TStringField;
    RxMDValidadCantidad: TStringField;
    RxMDValidasuma: TStringField;
    RxMDValidaaMN: TStringField;
    RxMDValidaaDLL: TStringField;
    RxMDValidabMN: TStringField;
    RxMDValidabDLL: TStringField;
    RxMDValidadCantidadAnexo: TStringField;
    RxMDValidadescripcion: TStringField;
    RxMDValidamensaje: TStringField;
    RxMDValidasNumeroOrden: TStringField;
    RxMDValidasWbs2: TStringField;
    frxDBValida: TfrxDBDataset;
    ActividadesxOrdencancelada: TStringField;
    SaveDialog1: TSaveDialog;
    NxHeaderPanel2: TNxHeaderPanel;
    DescrL: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    tsNumeroOrden: TComboBox;
    Label7: TLabel;
    Label4: TLabel;
    tsFiltro: TComboBox;
    Label5: TLabel;
    chkPeriodo: TJvCheckBox;
    GrupoMoneda: TGroupBox;
    chkMN: TCheckBox;
    chkDLL: TCheckBox;
    GroupBox5: TGroupBox;
    chkadmon: TCheckBox;
    chkPu: TCheckBox;
    NxHeaderPanel1: TNxHeaderPanel;
    tdIdFecha1: TDateEdit;
    tdIdFecha: TDateEdit;
    PanelProgress: TPanel;
    Label15: TLabel;
    Label14: TLabel;
    Label19: TLabel;
    BarraEstado: TProgressBar;
    NxHeaderPanel3: TNxHeaderPanel;
    btnStatus: TButton;
    btnFiltroAnexo: TButton;
    btnAnexoVsEstimaciones: TButton;
    btnPartidasRetraso: TButton;
    btnSuministros: TButton;
    cmdExcedentes: TButton;
    cmdComparativo: TButton;
    Progress: TGauge;
    btnPanel: TNxHeaderPanel;
    btnTerminadas: TButton;
    btnAdicionales: TButton;
    btnTodas: TButton;
    btnPendientes: TButton;
    btnSinGenerar: TButton;
    btnSinReportar: TButton;
    cmdConceptos: TBitBtn;
    CmdProduccion: TBitBtn;
    cmdHistorico: TBitBtn;
    btnSubContratos: TBitBtn;
    cmdProgramado: TBitBtn;
    ActividadesxOrdensHoraInicio: TStringField;
    ActividadesxOrdensHoraFinal: TStringField;
    tsPlataforma: TRxDBLookupCombo;
    ds_plataformas: TDataSource;
    zqPlataformas: TZReadOnlyQuery;
    Label22: TLabel;
    tsReprogramacion: TDBLookupComboBox;
    ds_reprogramacion: TDataSource;
    zqReprogramacion: TZReadOnlyQuery;
    btnRptAvanXFolio: TButton;
    btnRptAvanXFolAcum: TButton;
    btnStaFoliosOT: TButton;
    btnRptActFolios: TButton;
    procedure btnStatusClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure btnFiltroAnexoClick(Sender: TObject);
    procedure tdIdFechaEnter(Sender: TObject);
    procedure tdIdFechaExit(Sender: TObject);
    procedure tdIdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure rptDetalladeGeneradoresGetValue(const VarName: String;
      var Value: Variant);
    procedure rptCmpAnexoGeneradoGetValue(const VarName: String;
      var Value: Variant);
    procedure btnPartidasRetrasoClick(Sender: TObject);
    procedure rptProgramadoGetValue(const VarName: String;
      var Value: Variant);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tsFiltroEnter(Sender: TObject);
    procedure tsFiltroExit(Sender: TObject);
    procedure tsFiltroKeyPress(Sender: TObject; var Key: Char);
    procedure tsOrdenadoKeyPress(Sender: TObject; var Key: Char);
    procedure rptProgramadoOTGetValue(const VarName: String;
      var Value: Variant);
    procedure btnPanelExit(Sender: TObject);
    procedure btnTerminadasClick(Sender: TObject);
    procedure btnAdicionalesClick(Sender: TObject);
    procedure btnPendientesClick(Sender: TObject);
    procedure btnTodasClick(Sender: TObject);
    procedure btnSinGenerarClick(Sender: TObject);
    procedure TabSheet1Show(Sender: TObject);
    procedure TabSheet2Show(Sender: TObject);
    procedure TabSheet4Show(Sender: TObject);
    procedure btnSuministrosClick(Sender: TObject);
    procedure btnAcumuladoTrinomioClick(Sender: TObject);
    procedure ActividadesxOrdenCalcFields(DataSet: TDataSet);
    procedure ActividadesxAnexoCalcFields(DataSet: TDataSet);
    procedure rInformeGetValue(const VarName: String; var Value: Variant);
    procedure btnReportadoVsGeneradoClick(Sender: TObject);
    procedure btnAnexoVsEstimacionesClick(Sender: TObject);
    procedure btnSubContratosClick(Sender: TObject);
    procedure btnSinReportarClick(Sender: TObject);
    procedure cmdConceptosClick(Sender: TObject);
    procedure chkDLLClick(Sender: TObject);
    procedure chkDLLEnter(Sender: TObject);
    procedure cmdExcedentesClick(Sender: TObject);
    procedure btTrinomiodllClick(Sender: TObject);
    procedure tsNumeroOrdenChange(Sender: TObject);
    procedure cmdComparativoClick(Sender: TObject);
    procedure tdIdFecha1Enter(Sender: TObject);
    procedure tdIdFecha1Exit(Sender: TObject);
    procedure tdIdFecha1KeyPress(Sender: TObject; var Key: Char);
    procedure tdIdFechaChange(Sender: TObject);
    procedure tdIdFecha1Change(Sender: TObject);
    procedure chkadmonEnter(Sender: TObject);
    procedure chkPuEnter(Sender: TObject);
    procedure formatoEncabezado();
    procedure CmdProduccionClick(Sender: TObject);
    procedure cmdHistoricoClick(Sender: TObject);
    procedure cmdProgramadoClick(Sender: TObject);
    procedure tsPlataformaEnter(Sender: TObject);
    procedure tsPlataformaExit(Sender: TObject);
    procedure ConsultaReprogramacion;
    procedure btnRptAvanXFolioClick(Sender: TObject);
    procedure btnRptAvanXFolAcumClick(Sender: TObject);
    procedure btnStaFoliosOTClick(Sender: TObject);
    procedure btnRptActFoliosClick(Sender: TObject);

  private
    sMenuP: String;
    { Private declarations }
    procedure acumularDiferencia(suma, sMensaje: string);
    function cantidadesDiferentes(sWBSContrato: string): string;
    procedure ventasDiferentes(sWBSContrato, suma: string);

  public
    { Public declarations }
  end;

var
  frmCompara: TfrmCompara;
  Opcion, cadpua : String ;
  Registro_Actual : String ;
  dCantidadInstalar : Double ;
  sSuperintendente, sSupervisor : String ;
  sPuestoSuperintendente, sPuestoSupervisor : String ;
  sSupervisorTierra, sPuestoSupervisorTierra : String ;
  tsAlcances : Array [1..10] Of String ;
  dTotalGenerado : Real ;
  sOpcion   : String ;
  sConvenio : String ;
  avProgramado,
  avFisico  : Currency ;
  BotonPermiso: TBotonesPermisos;

  //Exporta elementos a Excel..
  Excel, Libro, Hoja: Variant;

  //Matriz de colores
  Colores: array[1..10, 1..2] of integer;
  columnas: array[1..1400] of string;
  meses: array[1..12] of string = ('ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE');

implementation



{$R *.dfm}



procedure TfrmCompara.btnStatusClick(Sender: TObject);
var
  sOrden:string;

begin
   //Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha1.Date>tdIdFecha.Date then
   begin
       showmessage('La fecha final es menor a la fecha inicial' );
       tdIdFecha.SetFocus;
       exit;
   end;

//  try

    sOrden:='';
    if connection.configuracion.FieldByName('lOrdenaItem').AsString = 'Si' then
      sOrden:=' Order by iItemOrden '
    else
      sOrden:=' Order By mysql.udf_NaturalSortFormat(swbs,'+ IntToStr(Global_TamOrden) +  ',' +Quotedstr(Global_SepOrden) +') ';

    cadpua:='';

    If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
    Begin
        if ChkPu.Checked = True Then
           cadpua := 'sTipoAnexo =  "PU" ';

        if chkadmon.Checked then
          if cadpua='' then
            cadpua := 'And sTipoAnexo = "ADM" '
          else
            cadpua := ' And (' + cadpua + 'or sTipoAnexo = "ADM") '
        else
          if cadpua<>'' then
            cadpua:= ' and ' + cadpua;



        ActividadesxAnexo.Active := False ;
        ActividadesxAnexo.SQL.Clear ;
        ActividadesxAnexo.SQL.Add('select sContrato, iNivel, iColor, sTipoActividad, sWbsAnterior, sWbs, sNumeroActividad, mDescripcion, dFechaInicio, dFechaFinal, ' +
                                  'sMedida, dCantidadAnexo, dPonderado, dVentaMN, dVentaDLL, "" as cancelada from actividadesxanexo Where sContrato = :contrato and ' +
                                  'sIdConvenio = :convenio ' + cadpua + sOrden);
        ActividadesxAnexo.Params.ParamByName('Contrato').DataType  := ftString ;
        ActividadesxAnexo.Params.ParamByName('Contrato').Value     := global_contrato ;
        ActividadesxAnexo.Params.ParamByName('Convenio').DataType  := ftString ;
        ActividadesxAnexo.Params.ParamByName('Convenio').Value     := sConvenio ;
        ActividadesxAnexo.Open ;

        //Obtenemos los reportes en M.N. o en Dólares.
        if chkMN.Checked = True then
          begin
            rInforme.LoadFromFile (global_files + 'Estatus Contrato.fr3');
            if not FileExists(global_files + 'Estatus Contrato.fr3') then
              showmessage('El archivo de reporte Estatus Contrato.fr3 no existe, notifique al administrador del sistema');
          end
        else
        begin
           rInforme.LoadFromFile (global_files + 'Estatus ContratoDLL.fr3');
           if not FileExists(global_files + 'Estatus ContratoDLL.fr3') then
             showmessage('El archivo de reporte Estatus ContratoDLL.fr3 no existe, notifique al administrador del sistema');
        end;
        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
    End
    Else
    Begin
       { if ChkPu.Checked = True Then
           cadpua := 'And sTipoAnexo =  "PU" '
        else
           cadpua := 'And sTipoAnexo = "ADM" ' ;  }

        if ChkPu.Checked = True Then
           cadpua := 'sTipoAnexo =  "PU" ';

        if chkadmon.Checked then
          if cadpua='' then
            cadpua := 'And sTipoAnexo = "ADM" '
          else
            cadpua := ' And (' + cadpua + 'or sTipoAnexo = "ADM") '
        else
          if cadpua<>'' then
            cadpua:= ' and ' + cadpua;


        ActividadesxOrden.Active := False ;
        ActividadesxOrden.SQL.Clear ;
        ActividadesxOrden.SQL.Add('select sContrato, sNumeroOrden, iNivel, iColor, sTipoActividad, sWbsAnterior, sWbs, sNumeroActividad, mDescripcion, dFechaInicio, dFechaFinal, sHoraInicio, sHoraFinal, ' +
                                  'sMedida, dCantidad, dPonderado, dVentaMN, dVentaDLL, dCostoMN, dCostoDLL ' +

                                   ',(select lCancelada from bitacoradeactividades b ' +
                                   ' where a.sContrato = b.sContrato and a.sNumeroOrden = b.sNumeroOrden ' +
                                   'and a.swbs = b.swbs and lCancelada = "Si" limit 1) as cancelada ' +

                                   ' from actividadesxorden a Where sContrato = :contrato and ' +
                                   'sIdConvenio = :convenio and sNumeroOrden = :orden '+ cadpua + sOrden );
        ActividadesxOrden.Params.ParamByName('Contrato').DataType := ftString ;
        ActividadesxOrden.Params.ParamByName('Contrato').Value    := global_contrato ;
        ActividadesxOrden.Params.ParamByName('Convenio').DataType := ftString ;
        ActividadesxOrden.Params.ParamByName('Convenio').Value    := zqReprogramacion.FieldByName('sIdConvenio').AsString;
        ActividadesxOrden.Params.ParamByName('Orden').DataType    := ftString ;
        ActividadesxOrden.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text ;
        ActividadesxOrden.Open ;
        //Obtenemos los reportes en M.N. o en Dólares.
        if chkMN.Checked = True then
        begin
           rInforme.LoadFromFile (global_files + 'Estatus Orden.fr3');
           if not FileExists(global_files + 'Estatus Orden.fr3') then
             showmessage('El archivo de reporte Estatus Orden.fr3 no existe, notifique al administrador del sistema');
        end
        else
        begin
           rInforme.LoadFromFile (global_files + 'Estatus OrdenDLL.fr3');
           if not FileExists(global_files + 'Estatus OrdenDLL.fr3') then
             showmessage('El archivo de reporte Estatus OrdenDLL.fr3 no existe, notifique al administrador del sistema');
        end;
        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
    End
//  except
//        on e : exception do begin
//        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al generar Status de conceptos', 0);
//        end;
//  end;
end;

procedure TfrmCompara.FormShow(Sender: TObject);
var
   x,i,y,z : integer;
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'rComparativo');
  BotonPermiso.permisosBotones(nil);

  // ivan - > Llenado del array de las columnas del Excel..
  for x := 1 to 26 do
    columnas[x] := Chr(64 + x);

  i := 27;
  for x := 1 to 26 do
  begin
    for y := 1 to 26 do
    begin
      columnas[i] := Chr(64 + x) + Chr(64 + y);
      i := i + 1;
    end;
  end;

  for x := 1 to 1 do
  begin
    for y := 1 to 26 do
    begin
      for z := 1 to 26 do
      begin
        columnas[i] := Chr(64 + x) + Chr(64 + y) + Chr(64 + z);
        i := i + 1;
      end;
    end;
  end;

  if not BotonPermiso.imprimir then
    Begin
    btnStatus.Enabled:=False;
    btnFiltroAnexo.Enabled:=False;
    btnAnexoVsEstimaciones.Enabled:=False;
    btnSubContratos.Enabled:=False;
    btnPartidasRetraso.Enabled:=False;
    btnSuministros.Enabled:=False;
    cmdExcedentes.Enabled:=False;
    cmdConceptos.Enabled:=False;
    cmdComparativo.Enabled:=False;
    End;
  try
    sMenuP:=stMenu;
    tdIdFecha.Date       := Date ;
    tdIdFecha1.Date      := Date ;
    tsFiltro.ItemIndex   := 0 ;

    tsNumeroOrden.Items.Clear ;
    tsNumeroOrden.Items.Add('CONTRATO No. ' + global_contrato)  ;

    roqOrdenes.Active := False ;
    roqOrdenes.ParamByName('Contrato').AsString := Global_Contrato ;
    roqOrdenes.ParamByName('status').AsString :=  connection.configuracion.FieldValues [ 'sOrdenExtraordinaria' ];
    roqOrdenes.Open;
    While NOT roqOrdenes.Eof Do
    Begin
      tsNumeroOrden.Items.Add(roqOrdenes.FieldValues['sNumeroOrden']) ;
      roqOrdenes.Next
    End ;
    try
       tsNumeroOrden.ItemIndex := 1;
    Except
    end;

    ConsultaReprogramacion;

    if zqReprogramacion.RecordCount > 0 then
       tsReprogramacion.KeyValue := zqReprogramacion.FieldByName('sIdConvenio').AsString;

    zqPlataformas.Active := False;
    zqPlataformas.ParamByName('Contrato').AsString := global_contrato;
    zqPlataformas.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
    zqPlataformas.ParamByName('Convenio').AsString := global_convenio;
    zqPlataformas.Open;

    if zQplataformas.RecordCount > 0  then
        tsPlataforma.KeyValue := zqPlataformas.FieldValues['sIdPlataforma'];

    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select * From reportediario Where sContrato = :Contrato And dIdFecha = :Fecha') ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 Then
        sConvenio := Connection.qryBusca.FieldValues['sIdConvenio']
    Else
        sConvenio := global_convenio;
        
    tdIdFecha1.SetFocus;
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al iniciar el formulario', 0);
    end;
  end;

end;

procedure TfrmCompara.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  BotonPermiso.Free;
  action := cafree ;
end;

procedure TfrmCompara.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsFiltro.SetFocus 
end;

procedure TfrmCompara.btnFiltroAnexoClick(Sender: TObject);
begin


//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
  try
    If rxAnexoGenerado.RecordCount > 0 Then
        rxAnexoGenerado.EmptyTable ;
    sOpcion := '' ;
    btnPanel.Visible := True ;
    btnTerminadas.SetFocus
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al generar cantidad Anexo vs generado', 0);
    end;
  end;
end;

procedure TfrmCompara.tdIdFecha1Change(Sender: TObject);
begin
  tdIdFecha.Date:=tdIdFecha1.Date;
end;

procedure TfrmCompara.tdIdFecha1Enter(Sender: TObject);
begin
tdidfecha1.Color := global_color_entrada
end;

procedure TfrmCompara.tdIdFecha1Exit(Sender: TObject);
begin
tdidfecha1.Color := global_color_salida
end;

procedure TfrmCompara.tdIdFecha1KeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tdidfecha.SetFocus
end;

procedure TfrmCompara.tdIdFechaChange(Sender: TObject);
begin
 // tdIdFecha.Date:=tdIdFecha1.Date;
end;

procedure TfrmCompara.tdIdFechaEnter(Sender: TObject);
begin
    tdIdFecha.Color := global_color_entrada
end;

procedure TfrmCompara.tdIdFechaExit(Sender: TObject);
begin
  try
    tdIdFecha.Color := global_color_salida ;
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select * From reportediario Where sContrato = :Contrato And dIdFecha = :Fecha') ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 Then
        sConvenio := Connection.qryBusca.FieldValues['sIdConvenio']
    Else
        sConvenio := global_convenio ;
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al seleccionar la fecha final', 0);
    end;
  end;
end;

procedure TfrmCompara.tdIdFechaKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsnumeroorden.SetFocus
end;

procedure TfrmCompara.rptDetalladeGeneradoresGetValue(
  const VarName: String; var Value: Variant);
begin
  If CompareText(VarName, 'MI_FECHA') = 0 then
      Value := DateToStr(tdIdFecha.Date) ;
end;

procedure TfrmCompara.rptCmpAnexoGeneradoGetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'MI_FECHA') = 0 then
      Value := DateToStr(tdIdFecha.Date) ;
end;


procedure TfrmCompara.btnPartidasRetrasoClick(Sender: TObject);
Var
    dAvance   : Double ;
    iRetraso  : Integer ;
    lFiltro   : Boolean ;
begin
 if ChkPu.Checked = True Then
        cadpua := 'And a.sTipoAnexo =  "PU" '
    else
        cadpua := 'And a.sTipoAnexo = "ADM" ' ;
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
  try
    // Calculo de los avances de la orden o del contrato ....
    // Avance Programado
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
        Connection.qryBusca.SQL.Add('Select Sum(dAvancePonderadoDia) as dProgramado From avancesglobales ' +
                                    'where sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroOrden = "" and dIdFecha <= :Fecha Group By sNumeroOrden')
    Else
    Begin
        Connection.qryBusca.SQL.Add('Select Sum(dAvancePonderadoDia) as dProgramado From avancesglobales ' +
                                    'where sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroOrden = :orden and dIdFecha <= :Fecha Group By sContrato') ;
        Connection.qryBusca.Params.ParamByName('orden').DataType := ftString ;
        Connection.qryBusca.Params.ParamByName('orden').Value := tsNumeroOrden.Text ;
    End ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString ;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
    Connection.qryBusca.Open ;
    IF Connection.qryBusca.RecordCount > 0 Then
        avProgramado := Connection.qryBusca.FieldValues['dProgramado']
    Else
        avProgramado := 0 ;

    // Avance Fisico
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
        Connection.qryBusca.SQL.Add('Select Sum(dAvance) as dReal From avancesglobalesxorden where sContrato = :Contrato And ' +
                                     'sIdConvenio = :Convenio And sNumeroOrden = "" and dIdFecha <= :Fecha Group By sContrato')
    Else
    Begin
        Connection.qryBusca.SQL.Add('Select Sum(dAvance) as dReal From avancesglobalesxorden where sContrato = :Contrato And ' +
                                    'sIdConvenio = :Convenio And sNumeroOrden = :orden and dIdFecha <= :Fecha Group By sNumeroOrden') ;
        Connection.qryBusca.Params.ParamByName('orden').DataType := ftString ;
        Connection.qryBusca.Params.ParamByName('orden').Value := tsNumeroOrden.Text ;
    End ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString ;
    Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
    Connection.qryBusca.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
    Connection.qryBusca.Open ;
    IF Connection.qryBusca.RecordCount > 0 Then
        avFisico := Connection.qryBusca.FieldValues['dReal']
    Else
        avFisico := 0 ;


    If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
    Begin
        connection.QryBusca.Active := False ;
        connection.QryBusca.SQL.Clear ;
        connection.QryBusca.SQL.Add('Select a.iNivel, a.sTipoActividad, a.sWbsAnterior, a.sWbs, a.sNumeroActividad, a.mDescripcion, a.sMedida, a.dCantidadAnexo, a.dInstalado, a.dExcedente, ' +
                                    'a.dVentaMN, a.dVentaDLL, a.dPonderado, a.dFechaInicio, a.dFechaFinal, sum(d.dCantidad) as dProgramado ' +
                                    'From actividadesxanexo a LEFT JOIN distribuciondeanexo d ON (a.sContrato = d.sContrato And ' +
                                    'a.sIdConvenio = d.sIdConvenio And a.sWbs = d.sWbs And a.sNumeroActividad = d.sNumeroActividad And d.dIdFecha <= :Fecha) Where ' +
                                    'a.sContrato = :Contrato And a.sIdConvenio = :Convenio ' + cadpua  + ' Group By a.sWbs, a.sNumeroActividad Order By a.iItemOrden');
        connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Convenio').Value := sConvenio ;
        connection.QryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
        connection.QryBusca.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
        connection.QryBusca.Open ;

        If rxPartidasAvance.RecordCount > 0 Then
           rxPartidasAvance.EmptyTable ;

        dsInforme.FieldAliases.Clear ;
        dsInforme.DataSet := rxPartidasAvance ;
        Progress.Visible := True ;
        Progress.MinValue := 1 ;
        Progress.MaxValue := connection.QryBusca.RecordCount ;
        While NOT connection.QryBusca.Eof Do
        Begin
            If tdIdFecha.Date > connection.QryBusca.FieldValues['dFechaFinal'] Then
                If (connection.QryBusca.FieldValues['dInstalado'] + connection.QryBusca.FieldValues['dExcedente']) < connection.QryBusca.FieldValues['dCantidadAnexo']Then
                    iRetraso := tdIdFecha.Date - connection.QryBusca.FieldValues['dFechaFinal']
                Else
                    iRetraso := 0
            Else
                iRetraso := 0 ;

            If tsFiltro.Text = 'CON RETRASO' Then
                If iRetraso > 0 then
                    lFiltro := True
                Else
                    lFiltro := False
            Else
                If tsFiltro.Text = 'DESFASADAS' Then
                    If iRetraso = 0 then
                        lFiltro := True
                    Else
                        lFiltro := False
                Else
                    lFiltro := True ;

            If lFiltro Then
            Begin
                rxPartidasAvance.Append ;
                rxPartidasAvance.FieldValues['sContrato'] := global_contrato ;
                rxPartidasAvance.FieldValues['iNivel'] := connection.QryBusca.FieldValues['iNivel'] ;
                rxPartidasAvance.FieldValues['sTipoActividad'] := connection.QryBusca.FieldValues['sTipoActividad'] ;
                rxPartidasAvance.FieldValues['sWbsAnterior'] := connection.QryBusca.FieldValues['sWbsAnterior'] ;
                rxPartidasAvance.FieldValues['sWbs'] := connection.QryBusca.FieldValues['sWbs'] ;
                rxPartidasAvance.FieldValues['sNumeroActividad'] := connection.QryBusca.FieldValues['sNumeroActividad'] ;
                rxPartidasAvance.FieldValues['mDescripcion'] := Copy (connection.QryBusca.FieldValues['mDescripcion'], 1, 100 ) + ' ..' ;
                rxPartidasAvance.FieldValues['dCantidadAnexo'] := connection.QryBusca.FieldValues['dCantidadAnexo'] ;
                rxPartidasAvance.FieldValues['sMedida'] := connection.QryBusca.FieldValues['sMedida'] ;
                rxPartidasAvance.FieldValues['dPonderado'] := connection.QryBusca.FieldValues['dPonderado'] ;
                rxPartidasAvance.FieldValues['dFechaInicio'] := connection.QryBusca.FieldValues['dFechaInicio'] ;
                rxPartidasAvance.FieldValues['dFechaFinal'] := connection.QryBusca.FieldValues['dFechaFinal'] ;
                rxPartidasAvance.FieldValues['dVentaMN'] := connection.QryBusca.FieldValues['dVentaMN'] ;
                rxPartidasAvance.FieldValues['dVentaDLL'] := connection.QryBusca.FieldValues['dVentaDLL'] ;
                If connection.QryBusca.FieldValues['sTipoActividad'] = 'Actividad' Then
                Begin
                    rxPartidasAvance.FieldValues['dCantidadProgramada'] := connection.QryBusca.FieldValues['dProgramado'] ;
                    If ((connection.QryBusca.FieldValues['dProgramado'] / connection.QryBusca.FieldValues['dCantidadAnexo']) * connection.QryBusca.FieldValues['dPonderado']) > 0 Then
                        rxPartidasAvance.FieldValues['dAvanceProgramado'] := ((connection.QryBusca.FieldValues['dProgramado'] / connection.QryBusca.FieldValues['dCantidadAnexo']) * connection.QryBusca.FieldValues['dPonderado'] ) ;
                    rxPartidasAvance.FieldValues['dCantidadReal'] := (connection.QryBusca.FieldValues['dInstalado'] + connection.QryBusca.FieldValues['dExcedente']) ;
                    If rxPartidasAvance.FieldValues['dCantidadReal']  < connection.QryBusca.FieldValues['dCantidadAnexo'] Then
                    Begin
                        connection.QryBusca2.Active := False ;
                        connection.QryBusca2.SQL.Clear ;
                        connection.QryBusca2.SQL.Add('Select sum(b.dAvance) as dAvance, a.dCantidadAnexo, a2.dCantidad From bitacoradeactividades b ' +
                                                     'INNER JOIN actividadesxorden a2 ON (a2.sContrato = b.sContrato And a2.sPaquete = b.sPaquete And a2.sNumeroActividad = b.sNumeroActividad And a2.sIdConvenio = :Convenio And a2.sTipoActividad = "Actividad" )' +
                                                     'INNER JOIN actividadesxanexo a ON (a.sContrato = b.sContrato And a.sNumeroActividad = b.sNumeroActividad And a.sIdConvenio = a2.sIdConvenio) ' +
                                                      'Where b.sContrato = :contrato And b.sNumeroActividad = :Actividad And b.dIdFecha <= :Fecha ' + cadpua  + ' Group By b.sPaquete' ) ;
                        connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
                        connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato ;
                        connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString ;
                        connection.QryBusca2.Params.ParamByName('Convenio').Value := sConvenio ;
                        connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
                        connection.QryBusca2.Params.ParamByName('Actividad').Value := connection.QryBusca.FieldValues['sNumeroActividad'] ;
                        connection.QryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
                        connection.QryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
                        connection.QryBusca2.Open ;
                        dAvance := 0 ;
                        While NOT connection.QryBusca2.Eof Do
                        Begin
                            dAvance := dAvance + (connection.QryBusca2.FieldValues['dAvance'] * connection.QryBusca2.FieldValues['dCantidad']) / connection.QryBusca2.FieldValues['dCantidadAnexo'] ;
                            connection.QryBusca2.Next ;
                        End ;
                        rxPartidasAvance.FieldValues['dAvanceReal'] := dAvance ;
                    End
                    Else
                        rxPartidasAvance.FieldValues['dAvanceReal'] := 100 ;
                    If tdIdFecha.Date > connection.QryBusca.FieldValues['dFechaFinal'] Then
                        If rxPartidasAvance.FieldValues['dCantidadReal'] < rxPartidasAvance.FieldValues['dCantidadAnexo'] Then
                            rxPartidasAvance.FieldValues['iRetraso'] := tdIdFecha.Date - connection.QryBusca.FieldValues['dFechaFinal']
                        Else
                            rxPartidasAvance.FieldValues['iRetraso'] := 0
                    Else
                        rxPartidasAvance.FieldValues['iRetraso'] := 0 ;
                    If (rxPartidasAvance.FieldValues['dAvanceReal'] = 0) And (rxPartidasAvance.FieldValues['dCantidadReal'] > 0) Then
                        If rxPartidasAvance.FieldValues['dCantidadAnexo'] > 0 Then
                            rxPartidasAvance.FieldValues['dAvanceReal'] := (rxPartidasAvance.FieldValues['dCantidadReal'] / rxPartidasAvance.FieldValues['dCantidadAnexo']) * 100 ;
                End ;
                rxPartidasAvance.Post ;
            End ;
            connection.QryBusca.Next ;
            progress.Progress := connection.QryBusca.RecNo ;
        End ;
        progress.Visible := False ;
        progress.Progress := 0 ;
        //Obtenemos Reporte en Dolares y M.N
        if chkMN.Checked = True then
        begin
           rInforme.LoadFromFile (global_files + 'RetrazosContrato.fr3');
           if not FileExists(global_files + 'RetrazosContrato.fr3') then
             showmessage('El archivo de reporte RetrazosContrato.fr3 no existe, notifique al administrador del sistema');
        end
        else
        begin
           rInforme.LoadFromFile (global_files + 'RetrazosContratoDLL.fr3');
           if not FileExists(global_files + 'RetrazosContratoDLL.fr3') then
             showmessage('El archivo de reporte RetrazosContratoDLL.fr3 no existe, notifique al administrador del sistema');
        end;
        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
    End
    Else
    Begin
        // Orden de Trabajo
        connection.QryBusca.Active := False ;
        connection.QryBusca.SQL.Clear ;
        connection.QryBusca.SQL.Add('Select a.iNivel, a.sTipoActividad, a.sWbsAnterior, a.sWbs, a.sNumeroActividad, a.mDescripcion, a.sMedida, a.dCantidad, a.dInstalado, a.dExcedente, ' +
                                        'a.dVentaMN, a.dVentaDLL, a.dPonderado, a.dFechaInicio, a.dFechaFinal, sum(d.dCantidad) as dProgramado ' +
                                        'From actividadesxorden a LEFT JOIN distribuciondeactividades d ON (a.sContrato = d.sContrato And a.sIdConvenio = d.sIdConvenio And ' +
                                        'a.sNumeroOrden = d.sNumeroOrden And a.sWbs = d.sWbs And a.sNumeroActividad = d.sNumeroActividad And d.dIdFecha <= :Fecha) ' +
                                        'Where a.sContrato = :Contrato And a.sIdConvenio = :Convenio ' + cadpua  + ' And a.sNumeroOrden = :Orden  ' +
                                        'Group By a.sWbs, a.sNumeroActividad Order By a.iItemOrden');
        connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString ;
        connection.QryBusca.Params.ParamByName('Orden').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
        connection.QryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
        connection.QryBusca.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
        connection.QryBusca.Open ;
        If rxPartidasAvance.RecordCount > 0 Then
           rxPartidasAvance.EmptyTable ;
        dsInforme.FieldAliases.Clear ;
        dsInforme.DataSet := rxPartidasAvance ;
        Progress.Visible := True ;
        Progress.MinValue := 1 ;
        Progress.MaxValue := connection.QryBusca.RecordCount ;
        While NOT connection.QryBusca.Eof Do
        Begin
            If tdIdFecha.Date > connection.QryBusca.FieldValues['dFechaFinal'] Then
                If (connection.QryBusca.FieldValues['dInstalado'] + connection.QryBusca.FieldValues['dExcedente']) < connection.QryBusca.FieldValues['dCantidad']Then
                    iRetraso := tdIdFecha.Date - connection.QryBusca.FieldValues['dFechaFinal']
                Else
                    iRetraso := 0
            Else
                iRetraso := 0 ;

            If tsFiltro.Text = 'CON RETRASO' Then
                If iRetraso > 0 then
                    lFiltro := True
                Else
                    lFiltro := False
            Else
                If tsFiltro.Text = 'DESFASADAS' Then
                    If iRetraso = 0 then
                        lFiltro := True
                    Else
                        lFiltro := False
                Else
                    lFiltro := True ;

            If lFiltro Then
            Begin
                rxPartidasAvance.Append ;
                rxPartidasAvance.FieldValues['sContrato'] := global_contrato ;
                rxPartidasAvance.FieldValues['iNivel'] := connection.QryBusca.FieldValues['iNivel'] ;
                rxPartidasAvance.FieldValues['sTipoActividad'] := connection.QryBusca.FieldValues['sTipoActividad'] ;
                rxPartidasAvance.FieldValues['sWbsAnterior'] := connection.QryBusca.FieldValues['sWbsAnterior'] ;
                rxPartidasAvance.FieldValues['sWbs'] := connection.QryBusca.FieldValues['sWbs'] ;
                rxPartidasAvance.FieldValues['sNumeroActividad'] := connection.QryBusca.FieldValues['sNumeroActividad'] ;
                rxPartidasAvance.FieldValues['mDescripcion'] := Copy (connection.QryBusca.FieldValues['mDescripcion'], 1, 120 ) + ' .' ;
                rxPartidasAvance.FieldValues['dCantidadAnexo'] := connection.QryBusca.FieldValues['dCantidad'] ;
                rxPartidasAvance.FieldValues['sMedida'] := connection.QryBusca.FieldValues['sMedida'] ;
                rxPartidasAvance.FieldValues['dPonderado'] := connection.QryBusca.FieldValues['dPonderado'] ;
                rxPartidasAvance.FieldValues['dFechaInicio'] := connection.QryBusca.FieldValues['dFechaInicio'] ;
                rxPartidasAvance.FieldValues['dFechaFinal'] := connection.QryBusca.FieldValues['dFechaFinal'] ;
                rxPartidasAvance.FieldValues['dVentaMN'] := connection.QryBusca.FieldValues['dVentaMN'] ;
                rxPartidasAvance.FieldValues['dVentaDLL'] := connection.QryBusca.FieldValues['dVentaDLL'] ;
                rxPartidasAvance.FieldValues['dCantidadProgramada'] := connection.QryBusca.FieldValues['dProgramado'] ;
                If ((connection.QryBusca.FieldValues['dProgramado'] / connection.QryBusca.FieldValues['dCantidad']) * connection.QryBusca.FieldValues['dPonderado']) > 0 Then
                    rxPartidasAvance.FieldValues['dAvanceProgramado'] := ((connection.QryBusca.FieldValues['dProgramado'] / connection.QryBusca.FieldValues['dCantidad']) * connection.QryBusca.FieldValues['dPonderado'] ) ;
                rxPartidasAvance.FieldValues['dCantidadReal'] := (connection.QryBusca.FieldValues['dInstalado'] + connection.QryBusca.FieldValues['dExcedente']) ;
                If rxPartidasAvance.FieldValues['dCantidadReal']  < connection.QryBusca.FieldValues['dCantidad'] Then
                Begin
                    connection.QryBusca2.Active := False ;
                    connection.QryBusca2.SQL.Clear ;
                    connection.QryBusca2.SQL.Add('Select sum(b.dAvance) as dAvance From bitacoradeactividades b ' +
                                      'INNER JOIN actividadesxorden a ON (a.sContrato = b.sContrato And a.sWbs = b.sWbs And a.sNumeroActividad = b.sNumeroActividad And a.sIdConvenio = :Convenio And a.sTipoActividad = "Actividad" )' +
                                      'Where b.sContrato = :contrato And b.sWbs = :Wbs And b.sNumeroActividad = :Actividad ' + cadpua  + ' And b.dIdFecha <= :Fecha Group By b.sWbs, b.sNumeroActividad' ) ;
                    connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
                    connection.QryBusca2.Params.ParamByName('Contrato').Value := global_Contrato ;
                    connection.QryBusca2.Params.ParamByName('Convenio').DataType := ftString ;
                    connection.QryBusca2.Params.ParamByName('Convenio').Value := sConvenio ;
                    connection.QryBusca2.Params.ParamByName('Wbs').DataType := ftString ;
                    connection.QryBusca2.Params.ParamByName('Wbs').Value := connection.QryBusca.FieldValues['sWbs'] ;
                    connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
                    connection.QryBusca2.Params.ParamByName('Actividad').Value := connection.QryBusca.FieldValues['sNumeroActividad'] ;
                    connection.QryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
                    connection.QryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
                    connection.QryBusca2.Open ;
                    dAvance := 0 ;
                    While NOT connection.QryBusca2.Eof Do
                    Begin
                        dAvance := dAvance + connection.QryBusca2.FieldValues['dAvance'] ;
                        connection.QryBusca2.Next ;
                    End ;
                    rxPartidasAvance.FieldValues['dAvanceReal'] := dAvance ;
                End
                Else
                    rxPartidasAvance.FieldValues['dAvanceReal'] := 100 ;

                If tdIdFecha.Date > connection.QryBusca.FieldValues['dFechaFinal'] Then
                    If rxPartidasAvance.FieldValues['dCantidadReal'] < rxPartidasAvance.FieldValues['dCantidadAnexo'] Then
                        rxPartidasAvance.FieldValues['iRetraso'] := tdIdFecha.Date - connection.QryBusca.FieldValues['dFechaFinal']
                    Else
                        rxPartidasAvance.FieldValues['iRetraso'] := 0
                Else
                    rxPartidasAvance.FieldValues['iRetraso'] := 0 ;

                If (rxPartidasAvance.FieldValues['dAvanceReal'] = 0) And (rxPartidasAvance.FieldValues['dCantidadReal'] > 0) Then
                    If rxPartidasAvance.FieldValues['dCantidadAnexo'] > 0 Then
                        rxPartidasAvance.FieldValues['dAvanceReal'] := (rxPartidasAvance.FieldValues['dCantidadReal'] / rxPartidasAvance.FieldValues['dCantidadAnexo']) * 100 ;
                rxPartidasAvance.Post ;
            End ;
            connection.QryBusca.Next ;
            progress.Progress := connection.QryBusca.RecNo ;
        End ;
        progress.Visible := False ;
        progress.Progress := 0 ;
        //Obtenemos Reporte en Dolares y M.N
        if chkMN.Checked = True then
        begin
           rInforme.LoadFromFile (global_files + 'RetrazosOrden.fr3');
           if not FileExists(global_files + 'RetrazosOrden.fr3') then
             showmessage('El archivo de reporte RetrazosOrden.fr3 no existe, notifique al administrador del sistema');
        end
        else
        begin
           rInforme.LoadFromFile (global_files + 'RetrazosOrdenDLL.fr3');
           if not FileExists(global_files + 'RetrazosOrdenDLL.fr3') then
             showmessage('El archivo de reporte RetrazosOrdenDLL.fr3 no existe, notifique al administrador del sistema');
        end;
        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
        
    End ;
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al generar partidas con retraso', 0);
    end;
  end;
end;

procedure TfrmCompara.rptProgramadoGetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'MI_FECHA') = 0 then
      Value := DateToStr(tdIdFecha.Date) ;
  If CompareText(VarName, 'PROGRAMADO') = 0 then
      Value := avProgramado ;

  If CompareText(VarName, 'REAL') = 0 then
     Value := avFisico ;

  If CompareText(VarName, 'TEXTO') = 0 then
     If avProgramado > avFisico Then
          Value := 'ATRASO DEL '
     Else
          Value := 'AVANCE DEL ' ;

  If CompareText(VarName, 'DIFERENCIA') = 0 then
     If avProgramado > avFisico Then
          Value := avProgramado - avFisico
     Else
          Value := avFisico - avProgramado

end;

procedure TfrmCompara.tsNumeroOrdenChange(Sender: TObject);
begin
 try
    roqOrdenes.Locate('snumeroorden', tsNumeroOrden.Text, []);

    zqPlataformas.Active := False;
    zqPlataformas.ParamByName('Contrato').AsString := global_contrato;
    zqPlataformas.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
    zqPlataformas.ParamByName('Convenio').AsString := global_convenio;
    zqPlataformas.Open;

    if zQplataformas.RecordCount > 0  then
        tsPlataforma.KeyValue := zqPlataformas.FieldValues['sIdPlataforma'];
 except
  on e : exception do begin
  UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al seleccionar el frente de trabajo', 0);
  end;
 end;
end;

procedure TfrmCompara.tsNumeroOrdenEnter(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmCompara.tsNumeroOrdenExit(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_salida;
    ConsultaReprogramacion;
end;

procedure TfrmCompara.tsFiltroEnter(Sender: TObject);
begin
    tsFiltro.Color := global_color_entrada
end;

procedure TfrmCompara.tsFiltroExit(Sender: TObject);
begin
    tsFiltro.Color := global_color_salida
end;

procedure TfrmCompara.tsFiltroKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tsPlataforma.SetFocus
end;

procedure TfrmCompara.tsOrdenadoKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        btnPartidasRetraso.SetFocus 
end;

procedure TfrmCompara.tsPlataformaEnter(Sender: TObject);
begin
   tsPlataforma.Color := global_color_entrada
end;

procedure TfrmCompara.tsPlataformaExit(Sender: TObject);
begin
    tsPlataforma.Color := global_color_salida
end;

procedure TfrmCompara.rptProgramadoOTGetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'MI_FECHA') = 0 then
      Value := DateToStr(tdIdFecha.Date) ;

  If CompareText(VarName, 'PROGRAMADO') = 0 then
      Value := avProgramado ;

  If CompareText(VarName, 'REAL') = 0 then
     Value := avFisico ;

  If CompareText(VarName, 'TEXTO') = 0 then
     If avProgramado > avFisico Then
          Value := 'ATRASO DEL '
     Else
          Value := 'AVANCE DEL ' ;

  If CompareText(VarName, 'DIFERENCIA') = 0 then
     If avProgramado > avFisico Then
          Value := avProgramado - avFisico
     Else
          Value := avFisico - avProgramado
end;

procedure TfrmCompara.btnPanelExit(Sender: TObject);
Var
    lFiltro        : Boolean ;
    lCambio        : Boolean ;
    dGenerado      : Double ;
    dCantidadTotal : Double ;
    dReportado     : Double ;
begin

   if ChkPu.Checked = True Then
        cadpua := 'And sTipoAnexo =  "PU" '
    else
        cadpua := 'And sTipoAnexo = "ADM" ' ;

 If rxAnexoGenerado.RecordCount > 0 Then
        rxAnexoGenerado.EmptyTable ;

 If sOpcion <> '' Then
    If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
    Begin
        connection.QryBusca.Active := False ;
        connection.QryBusca.SQL.Clear ;
        connection.QryBusca.SQL.Add('Select sNumeroActividad, mDescripcion, sMedida, dCantidadAnexo, dPonderado, dVentaMN, dVentaDLL, dInstalado ' +
                              'From actividadesxanexo Where sContrato = :contrato and sIdConvenio = :Convenio and sTipoActividad = "Actividad" ' + cadpua + ' Order By iItemOrden');
        connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Convenio').Value := sConvenio ;
        connection.QryBusca.Open ;
        While NOT connection.QryBusca.Eof Do
        Begin
            lFiltro := False ;
            lCambio := False ;

            If Connection.qryBusca.FieldValues['dInstalado'] > 0 Then
                dReportado := Connection.qryBusca.FieldValues['dInstalado']
            Else
                dReportado := 0 ;

            Connection.qryBusca2.Active := False ;
            Connection.qryBusca2.SQL.Clear ;
            Connection.qryBusca2.SQL.Add('Select Sum(a.dCantidad) as dGenerado FROM estimacionxpartida a ' +
                                         'INNER JOIN estimaciones e ON (a.sContrato = e.sContrato And a.sNumeroOrden = e.sNumeroOrden And ' +
                                         'a.sNumeroGenerador = e.sNumeroGenerador And e.dFechaFinal <= :Fecha And e.lStatus = "Autorizado") ' +
                                         'Where a.sContrato = :contrato And a.sNumeroActividad = :Actividad Group By a.sNumeroActividad ') ;
            Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
            Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
            Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
            Connection.qryBusca2.Params.ParamByName('Actividad').Value := Connection.qryBusca.FieldValues['sNumeroActividad'] ;
            Connection.qryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
            Connection.qryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
            Connection.qryBusca2.Open ;
            If Connection.qryBusca2.RecordCount > 0 Then
                dGenerado := Connection.qryBusca2.FieldValues['dGenerado']
            Else
                dGenerado := 0 ;

            If sOpcion = 'Terminadas' Then
            Begin
                If dGenerado = Connection.qryBusca.FieldValues['dCantidadAnexo'] Then
                    lFiltro := True
            End
            Else
                If sOpcion = 'Adicionales' Then
                Begin
                    If dGenerado > Connection.qryBusca.FieldValues['dCantidadAnexo'] Then
                        lFiltro := True
                End
                Else
                    If sOpcion = 'Sin Generar' Then
                    Begin
                        If dGenerado = 0 Then
                            lFiltro := True
                    End
                    Else
                        If sOpcion = 'Pendientes' Then
                        Begin
                            If dGenerado < Connection.qryBusca.FieldValues['dCantidadAnexo'] Then
                                lFiltro := True
                        End
                        Else
                           If sOpcion = 'Generadas' Then
                           Begin
                                If dGenerado > 0 Then
                                   lFiltro := True
                           End
                           Else
                               If sOpcion = 'Reportadas' Then
                               Begin
                                   lFiltro := True;
                                   lCambio := True;
                               End
                               Else
                                   lFiltro := True ;

            If lFiltro and lCambio Then
            begin
               if dReportado > 0 then
               begin
                    rxAnexoGenerado.Append ;
                    rxAnexoGenerado.FieldValues['sNumeroActividad'] := connection.QryBusca.FieldValues['sNumeroActividad'] ;
                    rxAnexoGenerado.FieldValues['sMedida'] := connection.QryBusca.FieldValues['sMedida'] ;
                    rxAnexoGenerado.FieldValues['mDescripcion'] := MidStr(connection.QryBusca.FieldValues['mDescripcion'], 1, 100) + ' ..' ;
                    rxAnexoGenerado.FieldValues['dCantidadAnexo'] := connection.QryBusca.FieldValues['dCantidadAnexo'] ;
                    rxAnexoGenerado.FieldValues['dPonderado'] := connection.QryBusca.FieldValues['dPonderado'] ;
                    rxAnexoGenerado.FieldValues['dVentaMN'] := connection.QryBusca.FieldValues['dVentaMN'] ;
                    rxAnexoGenerado.FieldValues['dVentaDLL'] := connection.QryBusca.FieldValues['dVentaDLL'] ;
                    rxAnexoGenerado.FieldValues['dGenerado'] := dGenerado ;
                    rxAnexoGenerado.FieldValues['dAdicional'] := dReportado;
                    rxAnexoGenerado.FieldValues['lTitulo'] := 'VOL. REPORTADO';

                    If dGenerado < connection.QryBusca.FieldValues['dCantidadAnexo'] Then
                        rxAnexoGenerado.FieldValues['dPendiente'] := dReportado - dGenerado ;
                    rxAnexoGenerado.Post ;
               end;
            End;

            If lFiltro and lCambio = False Then
            begin
                    rxAnexoGenerado.Append ;
                    rxAnexoGenerado.FieldValues['sNumeroActividad'] := connection.QryBusca.FieldValues['sNumeroActividad'] ;
                    rxAnexoGenerado.FieldValues['sMedida'] := connection.QryBusca.FieldValues['sMedida'] ;
                    rxAnexoGenerado.FieldValues['mDescripcion'] := MidStr(connection.QryBusca.FieldValues['mDescripcion'], 1, 100) + ' ..' ;
                    rxAnexoGenerado.FieldValues['dCantidadAnexo'] := connection.QryBusca.FieldValues['dCantidadAnexo'] ;
                    rxAnexoGenerado.FieldValues['dPonderado'] := connection.QryBusca.FieldValues['dPonderado'] ;
                    rxAnexoGenerado.FieldValues['dVentaMN'] := connection.QryBusca.FieldValues['dVentaMN'] ;
                    rxAnexoGenerado.FieldValues['dVentaDLL'] := connection.QryBusca.FieldValues['dVentaDLL'] ;
                    rxAnexoGenerado.FieldValues['dGenerado'] := dGenerado ;
                    rxAnexoGenerado.FieldValues['lTitulo'] := 'VOL. ADICIONAL';

                    If dGenerado >= connection.QryBusca.FieldValues['dCantidadAnexo'] Then
                    begin
                         rxAnexoGenerado.FieldValues['dAdicional'] := dGenerado - connection.QryBusca.FieldValues['dCantidadAnexo'] ;
                    end;

                    If dGenerado < connection.QryBusca.FieldValues['dCantidadAnexo'] Then
                        rxAnexoGenerado.FieldValues['dPendiente'] := connection.QryBusca.FieldValues['dCantidadAnexo'] - dGenerado ;
                    rxAnexoGenerado.Post ;
            End;

            connection.QryBusca.Next
        End ;
        dsInforme.FieldAliases.Clear ;
        dsInforme.DataSet := rxAnexoGenerado ;

        //Reportes en Dolares y en M.N.
        if chkMN.Checked = True then
        begin
           If sOpcion = 'Reportadas' Then
           begin
              rInforme.LoadFromFile (global_files + 'EstatusGeneradoConceptos2.fr3');
        if not FileExists(global_files + 'EstatusGeneradoConceptos2.fr3') then
           showmessage('El archivo de reporte EstatusGeneradoConceptos2.fr3 no existe, notifique al administrador del sistema');
           end
           Else
           begin
              rInforme.LoadFromFile (global_files + 'EstatusGeneradoConceptos.fr3') ;
              if not FileExists(global_files + 'EstatusGeneradoConceptos.fr3') then
           showmessage('El archivo de reporte EstatusGeneradoConceptos.fr3 no existe, notifique al administrador del sistema');
           end;
         end
        else
        begin
            If sOpcion = 'Reportadas' Then
            begin
               rInforme.LoadFromFile (global_files + 'EstatusGeneradoConceptos2DLL.fr3');
            if not FileExists(global_files + 'EstatusGeneradoConceptos2DLL.fr3') then
            showmessage('El archivo de reporte EstatusGeneradoConceptos2DLL.fr3 no existe, notifique al administrador del sistema');

            end
            Else
            begin
               rInforme.LoadFromFile (global_files + 'EstatusGeneradoConceptosDLL.fr3') ;
            if not FileExists(global_files + 'EstatusGeneradoConceptosDLL.fr3') then
            showmessage('El archivo de reporte EstatusGeneradoConceptosDLL.fr3 no existe, notifique al administrador del sistema');


            end;
        end;

        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));

    End
    Else
    Begin
        connection.QryBusca.Active := False ;
        connection.QryBusca.SQL.Clear ;
        connection.QryBusca.SQL.Add('Select sNumeroActividad, mDescripcion, dPonderado, sMedida, Sum(dCantidad) as dCantidadAnexo, dVentaMN, dVentaDLL, dInstalado From actividadesxorden a ' +
                              'Where sContrato = :contrato and sIdConvenio = :Convenio And sTipoActividad = "Actividad" And sNumeroOrden = :Orden ' + cadpua   +
                              'And sTipoActividad = "Actividad"  Group By sNumeroActividad Order By iItemOrden');
        connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString ;
        connection.QryBusca.Params.ParamByName('Orden').DataType := ftString ;
        connection.QryBusca.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
        connection.QryBusca.Open ;
        While NOT connection.QryBusca.Eof Do
        Begin
            lFiltro := False ;
            lCambio := False ;

            If Connection.qryBusca.FieldValues['dInstalado'] > 0 Then
                dReportado := Connection.qryBusca.FieldValues['dInstalado']
            Else
                dReportado := 0 ;

            If NOT connection.QryBusca.FieldByName('dCantidadAnexo').IsNull Then
                dCantidadTotal  := connection.QryBusca.FieldValues['dCantidadAnexo']
            Else
                dCantidadTotal  := 0 ;

            Connection.qryBusca2.Active := False ;
            Connection.qryBusca2.SQL.Clear ;
            Connection.qryBusca2.SQL.Add('Select Sum(a.dCantidad) as dGenerado FROM estimacionxpartida a ' +
                                         'INNER JOIN estimaciones e ON (a.sContrato = e.sContrato And a.sNumeroOrden = e.sNumeroOrden And ' +
                                         'a.sNumeroGenerador = e.sNumeroGenerador And e.dFechaFinal <= :Fecha And e.lStatus = "Autorizado") ' +
                                         'Where a.sContrato = :contrato and a.sNumeroOrden = :Orden And a.sNumeroActividad = :Actividad Group By a.sNumeroActividad ') ;
            Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
            Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
            Connection.qryBusca2.Params.ParamByName('Orden').DataType := ftString ;
            Connection.qryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
            Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
            Connection.qryBusca2.Params.ParamByName('Actividad').Value := connection.QryBusca.FieldValues['sNumeroActividad'] ;
            Connection.qryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
            Connection.qryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
            Connection.qryBusca2.Open ;
            If Connection.qryBusca2.RecordCount > 0 Then
                dGenerado := Connection.qryBusca2.FieldValues['dGenerado']
            Else
                dGenerado := 0 ;

            If sOpcion = 'Terminadas' Then
            Begin
                If dGenerado = dCantidadTotal Then
                    lFiltro := True
            End
            Else
                If sOpcion = 'Adicionales' Then
                Begin
                   If dGenerado > dCantidadTotal Then
                       lFiltro := True
                End
                Else
                   If sOpcion = 'Sin Generar' Then
                   Begin
                       If dGenerado = 0 Then
                           lFiltro := True
                   End
                   Else
                       If tsFiltro.Text = 'Pendientes' Then
                       Begin
                           If dGenerado < dCantidadTotal Then
                               lFiltro := True
                       End
                       Else
                           If sOpcion = 'Generadas' Then
                           Begin
                                If dGenerado > 0 Then
                                   lFiltro := True
                           End
                             Else
                              If sOpcion = 'Reportadas' Then
                               Begin
                                   lFiltro := True;
                                   lCambio := True;
                               End
                                 Else
                                    lFiltro := True ;
            If lFiltro and lCambio Then
            begin
               if dReportado > 0 then
               begin
                    rxAnexoGenerado.Append ;
                    rxAnexoGenerado.FieldValues['sNumeroOrden'] := tsNumeroOrden.Text ;
                    rxAnexoGenerado.FieldValues['sNumeroActividad'] := connection.QryBusca.FieldValues['sNumeroActividad'] ;
                    rxAnexoGenerado.FieldValues['sMedida'] := connection.QryBusca.FieldValues['sMedida'] ;
                    rxAnexoGenerado.FieldValues['mDescripcion'] := MidStr(connection.QryBusca.FieldValues['mDescripcion'],1 ,100 ) + '..' ;
                    rxAnexoGenerado.FieldValues['dCantidadAnexo'] := dCantidadTotal ;
                    rxAnexoGenerado.FieldValues['dPonderado'] := connection.QryBusca.FieldValues['dPonderado'] ;
                    rxAnexoGenerado.FieldValues['dVentaMN'] := connection.QryBusca.FieldValues['dVentaMN'] ;
                    rxAnexoGenerado.FieldValues['dVentaDLL'] := connection.QryBusca.FieldValues['dVentaDLL'] ;
                    rxAnexoGenerado.FieldValues['dGenerado'] := dGenerado ;
                    rxAnexoGenerado.FieldValues['lTitulo'] := 'VOL. REPORTADO';
                    rxAnexoGenerado.FieldValues['dAdicional'] := dReportado ;
                    If dGenerado < dCantidadTotal Then
                        rxAnexoGenerado.FieldValues['dPendiente'] := dReportado - dGenerado ;
                    rxAnexoGenerado.Post ;
               end;
            End;

            If lFiltro and lCambio = False Then
            begin
                rxAnexoGenerado.Append ;
                rxAnexoGenerado.FieldValues['sNumeroOrden'] := tsNumeroOrden.Text ;
                rxAnexoGenerado.FieldValues['sNumeroActividad'] := connection.QryBusca.FieldValues['sNumeroActividad'] ;
                rxAnexoGenerado.FieldValues['sMedida'] := connection.QryBusca.FieldValues['sMedida'] ;
                rxAnexoGenerado.FieldValues['mDescripcion'] := MidStr(connection.QryBusca.FieldValues['mDescripcion'],1 ,100 ) + '..' ;
                rxAnexoGenerado.FieldValues['dCantidadAnexo'] := dCantidadTotal ;
                rxAnexoGenerado.FieldValues['dPonderado'] := connection.QryBusca.FieldValues['dPonderado'] ;
                rxAnexoGenerado.FieldValues['dVentaMN'] := connection.QryBusca.FieldValues['dVentaMN'] ;
                rxAnexoGenerado.FieldValues['dVentaDLL'] := connection.QryBusca.FieldValues['dVentaDLL'] ;
                rxAnexoGenerado.FieldValues['dGenerado'] := dGenerado ;
                rxAnexoGenerado.FieldValues['lTitulo'] := 'VOL. ADICIONAL';
                If dGenerado >= dCantidadTotal Then
                    rxAnexoGenerado.FieldValues['dAdicional'] := dGenerado - dCantidadTotal ;
                If dGenerado < dCantidadTotal Then
                    rxAnexoGenerado.FieldValues['dPendiente'] := dCantidadTotal - dGenerado ;
                rxAnexoGenerado.Post ;
            End ;


            connection.QryBusca.Next
        End ;
        dsInforme.FieldAliases.Clear ;
        dsInforme.DataSet := rxAnexoGenerado ;
        //Reportes en Dolares y Moneda Nacional
        if chkMN.Checked = True then
        begin
            If sOpcion = 'Reportadas' Then
            begin
               rInforme.LoadFromFile (global_files + 'EstatusGeneradoConceptosOrden2.fr3');
        if not FileExists(global_files + 'EstatusGeneradoConceptosOrden2.fr3') then
           showmessage('El archivo de reporte EstatusGeneradoConceptosOrden2.fr3 no existe, notifique al administrador del sistema');

            end
            Else
            begin
               rInforme.LoadFromFile (global_files + 'EstatusGeneradoConceptosOrden.fr3') ;
         if not FileExists(global_files + 'EstatusGeneradoConceptosOrden.fr3') then
           showmessage('El archivo de reporte EstatusGeneradoConceptosOrden.fr3 no existe, notifique al administrador del sistema');

            end;
        end
        else
        begin
           If sOpcion = 'Reportadas' Then
           begin
               rInforme.LoadFromFile (global_files + 'EstatusGeneradoConceptosOrden2DLL.fr3');
        if not FileExists(global_files + 'EstatusGeneradoConceptosOrden2DLL.fr3') then
           showmessage('El archivo de reporte EstatusGeneradoConceptosOrden2DLL.fr3 no existe, notifique al administrador del sistema');

           end
           Else
           begin
               rInforme.LoadFromFile (global_files + 'EstatusGeneradoConceptosOrdenDLL.fr3') ;
        if not FileExists(global_files + 'EstatusGeneradoConceptosOrdenDLL.fr3') then
           showmessage('El archivo de reporte EstatusGeneradoConceptosOrdenDLL.fr3 no existe, notifique al administrador del sistema');

           end;
        end;
        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP))
    End ;

    btnPanel.Visible := False
end;

procedure TfrmCompara.btnTerminadasClick(Sender: TObject);
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
    sOpcion := 'Terminadas' ;
    tdIdFecha.SetFocus
end;

procedure TfrmCompara.btnAdicionalesClick(Sender: TObject);
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
    sOpcion := 'Adicionales' ;
    tdIdFecha.SetFocus
end;

procedure TfrmCompara.btnPendientesClick(Sender: TObject);
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
    sOpcion := 'Pendientes' ;
    tdIdFecha.SetFocus
end;

procedure TfrmCompara.btnTodasClick(Sender: TObject);
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
    sOpcion := 'Generadas' ;
    tdIdFecha.SetFocus
end;

procedure TfrmCompara.btTrinomiodllClick(Sender: TObject);
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
  try
    rDiarioFirmas (global_contrato, '' , 'A',tdIdFecha.Date, frmCompara ) ;
    If MessageDlg('Desea imprimir el consolidado de todas las estimaciones seleccionadas?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    Begin
        Reporte.Active := False ;
        Reporte.SQL.Clear ;
        Reporte.SQL.Add('Select ct.*, e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador, Sum(e.dCantidad * a.dVentaDLL) as dEstimado From estimacionxpartida e ' +
                        'INNER JOIN estimaciones e2 ON (e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And e.sNumeroGenerador = e2.sNumeroGenerador And e2.lStatus = "Autorizado") ' +
                        'INNER JOIN contrato_trinomio ct ON (e.sContrato = ct.sContrato And e.sInstalacion = ct.sInstalacion) ' +
                        'INNER JOIN actividadesxanexo a ON (e.sContrato = a.sContrato And e.sNumeroActividad = a.sNumeroActividad And a.sIdConvenio = :Convenio And a.sTipoActividad = "Actividad") ' +
                        'Where e.sContrato = :Contrato And e.lEstima = "Si" ' +
                        'Group By ct.sInstalacion, e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador') ;
        Reporte.Params.ParamByName('Contrato').DataType := ftString ;
        Reporte.Params.ParamByName('Contrato').Value := global_Contrato ;
        Reporte.Params.ParamByName('Convenio').DataType := ftString ;
        Reporte.Params.ParamByName('Convenio').Value := sConvenio ;
        dsInforme.FieldAliases.Clear ;
        dsInforme.DataSet := Reporte ;
        Reporte.Open ;
    End
    Else
    Begin
     if length(vartostr(Connection.EstimacionPeriodo.FieldValues['dFechaFinal']))>0 then
     begin
        Reporte.Active := False ;
        Reporte.SQL.Clear ;
        Reporte.SQL.Add('Select ct.*, e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador, Sum(e.dCantidad * a.dVentaDLL) as dEstimado From estimacionxpartida e ' +
                        'INNER JOIN estimaciones e2 ON (e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And e.sNumeroGenerador = e2.sNumeroGenerador And e2.lStatus = "Autorizado" ' +
                                                       'And Month(e2.dFechaFinal) = :Mes And Year(e2.dFechaFinal) = :Anno) ' +
                        'INNER JOIN contrato_trinomio ct ON (e.sContrato = ct.sContrato And e.sInstalacion = ct.sInstalacion) ' +
                        'INNER JOIN actividadesxanexo a ON (e.sContrato = a.sContrato And e.sNumeroActividad = a.sNumeroActividad And a.sIdConvenio = :Convenio And a.sTipoActividad = "Actividad") ' +
                        'Where e.sContrato = :Contrato And e.lEstima = "Si" ' +
                        'Group By ct.sInstalacion, e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador') ;
        Reporte.Params.ParamByName('Contrato').DataType := ftString ;
        Reporte.Params.ParamByName('Contrato').Value := global_Contrato ;
        Reporte.Params.ParamByName('Convenio').DataType := ftString ;
        Reporte.Params.ParamByName('Convenio').Value := sConvenio ;
        Reporte.Params.ParamByName('Mes').DataType := ftInteger ;
        Reporte.Params.ParamByName('Mes').Value := MonthOf(Connection.EstimacionPeriodo.FieldValues['dFechaFinal']) ;
        Reporte.Params.ParamByName('Anno').DataType := ftInteger ;
        Reporte.Params.ParamByName('Anno').Value := YearOf(Connection.EstimacionPeriodo.FieldValues['dFechaFinal']) ;
        dsInforme.FieldAliases.Clear ;
        dsInforme.DataSet := Reporte ;
        Reporte.Open ;
     end;
    End ;
    rInforme.LoadFromFile (global_files + 'TrinomioConcentrado.fr3');
    rInforme.PreviewOptions.MDIChild := False ;
    rInforme.PreviewOptions.Modal := True ;
    rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
    rInforme.PreviewOptions.ShowCaptions := False ;
    rInforme.Previewoptions.ZoomMode := zmPageWidth ;
    rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
    if not FileExists(global_files + 'TrinomioConcentrado.fr3') then
     showmessage('El archivo de reporte TrinomioConcentrado.fr3 no existe, notifique al administrador del sistema');
//    frxTrinomio.ShowReport

  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'En el proceso Acumulado de generación trinomio DLL', 0);
    end;
  end;
end;

procedure TfrmCompara.CmdProduccionClick(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// SEGUIMINTO DE AVANCES X PARTIDA DIAVAZ OCTUBRE 2012 //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena: string;
      fs: tStream;
      Alto: Extended;
      MiFechaI, MiFechaF, MiFecha: tDate;
      Ren, nivel, i, total: integer;
      iInicioMes, iFinMes : integer;
      Q_Partidas : TZReadOnlyQuery;
      dVolumen, dAvanceGlobal, dProgramado, dFisico: double;
      Progreso, TotalProgreso: real;
      lEncuentra : boolean;
      sColInicio, sColFinal, sSql1, sSql2 : string;

    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      {$REGION 'ENCABEZADO'}
      Ren := 2;
      //Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 130;

      Excel.Columns['A:A'].ColumnWidth := 10.86;
      Excel.Columns['B:B'].ColumnWidth := 37.29;
      Excel.Columns['C:C'].ColumnWidth := 6.43;
      Excel.Columns['D:D'].ColumnWidth := 8.14;
      Excel.Columns['E:E'].ColumnWidth := 7.57;

      Hoja.Range['A1:A3'].Select;
      Excel.Selection.RowHeight := '15';

      Hoja.Range['A1:A1'].Select;
      Excel.Selection.Value := tsNumeroOrden.Text;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Size := 9;
      Excel.Selection.Font.Name := 'Calibri';

      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Clear;
      connection.QryBusca2.SQL.Add('Select p.sDescripcion, o.dFiProgramado, o.dFfProgramado from ordenesdetrabajo o '+
                                  'inner join plataformas p on (o.sIdPlataforma = o.sIdPlataforma) '+
                                  'where o.sContrato = :contrato and o.sNumeroOrden =:Orden');
      connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato;
      connection.QryBusca2.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
      connection.QryBusca2.Open;

      Hoja.Range['B1:B1'].Select;
      if tsPlataforma.Text <> 'TODAS LAS PLATAFORMAS' then
         Excel.Selection.Value := connection.QryBusca2.FieldByName('sDescripcion').AsString
      else
         Excel.Selection.Value := tsNumeroOrden.Text;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Size := 9;
      Excel.Selection.Font.Name := 'Calibri';

      Hoja.Range['D1:D1'].Select;
      Excel.Selection.Value := 'TC';
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Size := 8;
      Excel.Selection.Font.Name := 'Calibri';

      Ren := 2;
      // Colocar los encabezados de la plantilla...
      Hoja.Range['A'+IntTostr(Ren)+':A'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Partida';
      FormatoEncabezado;
      Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Descripcion';
      FormatoEncabezado;
      Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Unidad';
      FormatoEncabezado;
      Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Precio MXN';
      FormatoEncabezado;
      Hoja.Range['E'+IntTostr(Ren)+':E'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Precio USD';
      FormatoEncabezado;

      Hoja.Range['A'+IntTostr(Ren)+':E'+IntToStr(Ren)].Select;
      Excel.Selection.Interior.ColorIndex := 37;

      {$ENDREGION}

      {$REGION 'DURACION Y TOTAL'}
      connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select a.sAnexo, a.sTipo, aa.* from actividadesxanexo aa '+
                                  'inner join anexos a on (aa.sAnexo = a.sAnexo) ');
      if tsPlataforma.Text <> 'TODAS LAS PLATAFORMAS' then
      begin
          connection.QryBusca.SQL.Add('inner join actividadesxorden o on (o.sContrato = aa.sContrato and o.sIdConvenio =:ConvenioF and o.sNumeroOrden =:Orden  and aa.sNumeroActividad = o.sNumeroActividad and (o.sIdPlataforma =:Plataforma or o.sIdPlataforma = "TODOS")) ');
          connection.QryBusca.ParamByName('Orden').AsString        := tsNumeroOrden.Text;
          connection.QryBusca.ParamByName('Plataforma').AsString   := zqPlataformas.FieldByName('sIdPlataforma').AsString;
      end;
      connection.QryBusca.SQL.Add('where aa.sContrato =:Contrato and aa.sIdConvenio =:Convenio and aa.sTipoActividad = "Actividad"');
      connection.QryBusca.Params.ParamByName('Contrato').DataType  := ftString;
      connection.QryBusca.Params.ParamByName('Contrato').Value     := global_contrato;
      if tsPlataforma.Text <> 'TODAS LAS PLATAFORMAS' then
      begin
         connection.QryBusca.Params.ParamByName('ConvenioF').DataType := ftString;
         connection.QryBusca.Params.ParamByName('ConvenioF').Value    := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      end;
      connection.QryBusca.Params.ParamByName('Convenio').DataType  := ftString;
      connection.QryBusca.Params.ParamByName('Convenio').Value     := global_convenio;
      if connection.configuracion.FieldByName('lOrdenaItem').AsString = 'No' then
      begin
          connection.QryBusca.SQL.Add(' order by mysql.udf_NaturalSortFormat(swbs,:Tam,:Separador)');
          connection.QryBusca.ParamByName('tam').AsInteger      := Global_TamOrden;
          connection.QryBusca.ParamByName('separador').AsString := Global_SepOrden;
      end
      else
         connection.QryBusca.SQL.Add('order by iItemOrden');
      connection.QryBusca.Open;
      if connection.QryBusca.RecordCount > 0 then
      begin
        if chkPeriodo.Checked then
        begin
            MiFecha  := tdIdFecha1.Date;
            MiFechaI := tdIdFecha1.Date;
            MiFechaF := tdIdFecha.Date;
        end
        else
        begin
            MiFecha  := connection.QryBusca2.FieldByName('dFiProgramado').AsDateTime;
            MiFechaI := connection.QryBusca2.FieldByName('dFiProgramado').AsDateTime;
            MiFechaF := connection.QryBusca2.FieldByName('dFfProgramado').AsDateTime;
        end;

        If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
        begin
            MiFecha  := tdIdFecha1.Date;
            MiFechaI := tdIdFecha1.Date;
            MiFechaF := tdIdFecha.Date;
        end;

        iInicioMes := 6;
        for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
        begin
            Hoja.Cells[Ren, 5 + i].Select;
            {Formato de las fechas archivo Excel,, 24/07/2011..}
            Excel.Selection.Value := copy(DateToStr(Mifecha), 0,2);

            iFinMes    := StrToInt(copy(DateToStr(Mifecha), 0,2));
            if iFinMes = 1 then
            begin
                if i = 1 then
                   Hoja.Range[columnas[iInicioMes]+IntTostr(Ren -1)+':'+columnas[i+5]+IntToStr(Ren-1)].Select
                else
                   Hoja.Range[columnas[iInicioMes]+IntTostr(Ren -1)+':'+columnas[i+4]+IntToStr(Ren-1)].Select;
                FormatoEncabezado;
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Value      := meses[StrToInt(copy(DateToStr(Mifecha-1), 4,2))];
                Excel.Selection.MergeCells := True;
                Excel.Selection.Font.Bold  := True;
                Excel.Selection.Font.Size  := 11;
                Excel.Selection.Font.Color := RGB(0,32,96);
                Excel.Selection.Font.Name  := 'Calibri';

                iInicioMes := i + 5;
            end;
            MiFecha := IncDay(MiFecha);
        end;

        if iFinMes < 28 then
        begin
            Hoja.Range[columnas[iInicioMes]+IntTostr(Ren -1)+':'+columnas[i+4]+IntToStr(Ren-1)].Select;
            FormatoEncabezado;
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;
            Excel.Selection.Value      := meses[StrToInt(copy(DateToStr(Mifecha-1), 4,2))];
            Excel.Selection.MergeCells := True;
            Excel.Selection.Font.Bold  := True;
            Excel.Selection.Font.Size  := 11;
            Excel.Selection.Font.Name  := 'Calibri';
        end;
        total := i;

        {Colocamos los anchos de las columnas..}
        Hoja.Range['F'+IntTostr(Ren)+':'+columnas[total+4]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Bold := False;
        Excel.Selection.Font.Size := 8;
        Excel.Selection.Font.Name := 'Calibri';
        Excel.Selection.Interior.ColorIndex := 37;
        FormatoEncabezado;
        Excel.Selection.ColumnWidth := 4;

        {Totales de Concentrato de Produccion}
        Hoja.Range[columnas[total+5]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size := 7;
        Excel.Selection.Font.Name := 'Calibri';

        Hoja.Range[columnas[total+5]+IntTostr(Ren)+':'+columnas[total+5]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Value  := 'Total';
        Excel.Selection.ColumnWidth := 7.57;

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Value := 'Prod. MXN';
        Excel.Selection.ColumnWidth := 9.43;

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Value := 'Prod. USD';
        Excel.Selection.ColumnWidth := 9.43;

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Value := 'Prod. Homol';
        Excel.Selection.ColumnWidth := 9.43;

        Hoja.Range[columnas[total+5]+IntTostr(Ren-1)+':'+columnas[total+8]+IntToStr(Ren-1)].Select;
        FormatoEncabezado;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Bold  := True;
        Excel.Selection.Font.Size  := 11;
        Excel.Selection.Font.Name  := 'Calibri';
        Excel.Selection.Value := 'Producción Real';
        Excel.Selection.MergeCells := True;
        Excel.Selection.Interior.ColorIndex := 37;

        Hoja.Range[columnas[total+9]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        Excel.Selection.ColumnWidth := 0;

        {$ENDREGION}

        inc(Ren);
        dAvanceGlobal := 0;
        connection.QryBusca.First;
        while not connection.QryBusca.Eof do
        begin
            {$REGION 'PARTIDAS DE ANEXO'}
           {Movimiento de la Barra..}
            Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
            TotalProgreso := TotalProgreso + Progreso;
            BarraEstado.Position := Trunc(TotalProgreso);

            {Escritura de Datos en el Archvio de Excel..}
            Hoja.Cells[Ren, 1].Select;
            Excel.Selection.Value := connection.QryBusca.FieldValues['sNumeroActividad'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;

            Hoja.Cells[Ren, 2].Select;
            Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];
            Excel.Selection.HorizontalAlignment := xlJustify;
            Excel.Selection.VerticalAlignment   := xlCenter;
            Excel.Selection.MergeCells := True;
            Excel.Selection.WrapText   := True;

            Hoja.Cells[Ren, 3].Select;
            Excel.Selection.Value := connection.QryBusca.FieldValues['sMedida'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;

            Hoja.Cells[Ren, 4].Select;
            Excel.Selection.NumberFormat := '$ #,##0.00';
            Excel.Selection.Value := connection.QryBusca.FieldValues['dVentaMN'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;

            Hoja.Cells[Ren, 5].Select;
            Excel.Selection.NumberFormat := '#,##0.00';
            Excel.Selection.Value := connection.QryBusca.FieldValues['dVentaDLL'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;

            Hoja.Range['A'+IntTostr(Ren)+':E'+IntToStr(Ren)].Select;
            Excel.Selection.Font.Size := 7;
            Excel.Selection.Font.Bold := False;
            Excel.Selection.Font.Name := 'Calibri';

            {$ENDREGION}

            MiFecha := MiFechaI;

            {$REGION 'CONSULTA DE ANEXOS'}
            {CONSUL DE ANEXOS}
            if connection.QryBusca.FieldByName('sTipo').AsString = 'BARCO' then
            begin
                {Consultamos la partida de barco..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('select dIdFecha, sClasificacion, sum(sFactor) as dCantidad '+
                                   'from movimientosdeembarcacion where sOrden = :Contrato and sClasificacion = :Actividad and dIdFecha >=:Inicial and dIdFecha <= :Final '+
                                   'Group By sClasificacion, dIdFecha Order By sClasificacion,dIdFecha');
                Q_Partidas.Params.ParamByName('Contrato').DataType  := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value     := global_contrato;
                Q_Partidas.Params.ParamByName('Inicial').DataType   := ftDate;
                Q_Partidas.Params.ParamByName('Inicial').Value      := MiFechaI;
                Q_Partidas.Params.ParamByName('Final').DataType     := ftDate;
                Q_Partidas.Params.ParamByName('Final').Value        := MiFechaF;
                Q_Partidas.Params.ParamByName('Actividad').DataType := ftString;
                Q_Partidas.Params.ParamByName('Actividad').Value    := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                Q_Partidas.Open;
            end;


            if connection.QryBusca.FieldByName('sTipo').AsString = 'PERSONAL' then
            begin
                {Consultamos la partida de personal..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('select dIdFecha, sIdPersonal, round(sum(dCantHH),2) as dCantidad '+
                                   'from bitacoradepersonal where sContrato =:Contrato and sIdPersonal = :IdPersonal and dIdFecha >=:Inicial and dIdFecha <=:Final '+
                                   'Group By sIdPersonal, dIdFecha Order By sIdPersonal,dIdFecha');
                Q_Partidas.Params.ParamByName('Contrato').DataType   := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value      := global_contrato;
                Q_Partidas.Params.ParamByName('Inicial').DataType    := ftDate;
                Q_Partidas.Params.ParamByName('Inicial').Value       := MiFechaI;
                Q_Partidas.Params.ParamByName('Final').DataType      := ftDate;
                Q_Partidas.Params.ParamByName('Final').Value         := MiFechaF;
                Q_Partidas.Params.ParamByName('IdPersonal').DataType := ftString;
                Q_Partidas.Params.ParamByName('IdPersonal').Value    := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                Q_Partidas.Open;
            end;

            if connection.QryBusca.FieldByName('sTipo').AsString = 'EQUIPO' then
            begin
                {Consultamos la partida de equipo..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('select dIdFecha, sIdEquipo, round(sum(dCantHH),2) as dCantidad '+
                                   'from bitacoradeequipos where sContrato =:Contrato and sIdEquipo = :IdEquipo and dIdFecha >=:Inicial and dIdFecha <=:Final '+
                                   'Group By sIdEquipo, dIdFecha Order By sIdEquipo,dIdFecha');
                Q_Partidas.Params.ParamByName('Contrato').DataType := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value    := global_contrato;
                Q_Partidas.Params.ParamByName('Inicial').DataType  := ftDate;
                Q_Partidas.Params.ParamByName('Inicial').Value     := MiFechaI;
                Q_Partidas.Params.ParamByName('Final').DataType    := ftDate;
                Q_Partidas.Params.ParamByName('Final').Value       := MiFechaF;
                Q_Partidas.Params.ParamByName('IdEquipo').DataType := ftString;
                Q_Partidas.Params.ParamByName('IdEquipo').Value    := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                Q_Partidas.Open;
            end;

            if connection.QryBusca.FieldByName('sTipo').AsString = 'PERNOCTA' then
            begin
                {Consultamos la partida de barco..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('Select tl.dIdFecha, sum(inacionales) as dCantidad '+
                                  'from tripulaciondiaria_listado tl '+
                                  'inner join cuentas c '+
                                  'on(c.sidcuenta=tl.sidcuenta) '+
                                  'where tl.sContrato =:Contrato and tl.sOrden=:OT and tl.dIdFecha >=:Inicial and tl.didfecha<=:fecha '+
                                  'and c.sidpernocta like :Pernocta group by tl.sContrato, tl.dIdFecha ');
                Q_Partidas.Params.ParamByName('Contrato').DataType   := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value      := global_contrato_barco;
                Q_Partidas.Params.ParamByName('OT').DataType         := ftString;
                Q_Partidas.Params.ParamByName('OT').Value            := global_contrato;
                Q_Partidas.Params.ParamByName('Inicial').DataType    := ftDate;
                Q_Partidas.Params.ParamByName('Inicial').Value       := MiFechaI;
                Q_Partidas.Params.ParamByName('Fecha').DataType      := ftDate;
                Q_Partidas.Params.ParamByName('Fecha').Value         := MiFechaF;
                Q_Partidas.Params.ParamByName('Pernocta').DataType   := ftString;
                Q_Partidas.Params.ParamByName('Pernocta').Value      := '%' + connection.QryBusca.FieldByName('sNumeroActividad').AsString + '%';
                Q_Partidas.Open;
            end;

            if pos('ANEXO',connection.QryBusca.FieldByName('sTipo').AsString) > 0 then
            begin
                {Consultamos partida de PU..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('Select b.sWbs,b.sNumeroActividad, sum(a.dCantidad) as dCantidad, a.dIdFecha, b.dCantidad as dVolumen ' +
                                   'From actividadesxorden b ' +
                                   'inner JOIN bitacoradeactividades a '+
                                   'ON (a.sContrato=b.sContrato and a.sIdConvenio = b.sIdConvenio And a.sWbs=b.sWbs and a.dIdFecha >=:Inicial And a.dIdFecha <=:Final and b.sNumeroOrden=a.sNumeroOrden) '+
                                   'left JOIN tiposdemovimiento t ' +
                                   'ON (b.sContrato=t.sContrato And a.sIdTipoMovimiento=t.sIdTipoMovimiento And t.sClasificacion="Tiempo en Operacion") ' +
                                   'Where b.sContrato=:Contrato And b.sIdConvenio=:Convenio And b.sNumeroOrden like :Orden and b.sNumeroActividad =:Actividad ' +
                                   'Group By b.sNumeroActividad,a.dIdFecha Order By a.dIdFecha');
                Q_Partidas.Params.ParamByName('Contrato').DataType := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value    := global_contrato;
                Q_Partidas.Params.ParamByName('Convenio').DataType := ftString;
                Q_Partidas.Params.ParamByName('Convenio').Value    := zqReprogramacion.FieldByName('sIdConvenio').AsString;
                Q_Partidas.Params.ParamByName('Orden').DataType    := ftString;
                If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
                   Q_Partidas.Params.ParamByName('Orden').Value    := '%'
                else
                   Q_Partidas.Params.ParamByName('Orden').Value    := tsNumeroOrden.Text;
                Q_Partidas.Params.ParamByName('Inicial').DataType  := ftDate;
                Q_Partidas.Params.ParamByName('Inicial').Value     := MiFechaI;
                Q_Partidas.Params.ParamByName('Final').DataType    := ftDate;
                Q_Partidas.Params.ParamByName('Final').Value       := MiFechaF;
                Q_Partidas.Params.ParamByName('Actividad').DataType      := ftString;
                Q_Partidas.Params.ParamByName('Actividad').Value         := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                Q_Partidas.Open;
            end;

            MiFecha := MiFechaI;
            if Q_Partidas.RecordCount > 0 then
            begin
              dVolumen := 0;
              for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
              begin
                  if MiFecha = Q_Partidas.FieldValues['dIdFecha'] then
                  begin
                      Hoja.Cells[Ren, 5 + i].Select;
                      dVolumen := Q_Partidas.FieldByName('dCantidad').AsFloat;

                      Excel.Selection.NumberFormat := '#,##0.00';
                      if (pos('ANEXO',connection.QryBusca.FieldByName('sTipo').AsString) > 0) then
                         Excel.Selection.NumberFormat := '#,##0.000';

                      if (pos('ANEXOEXT',connection.QryBusca.FieldByName('sTipo').AsString) > 0) then
                         Excel.Selection.NumberFormat := '#,##0.0000';

                      if (connection.QryBusca.FieldByName('sTipo').AsString = 'BARCO') then
                          Excel.Selection.NumberFormat := '#,##0.0000';

                      Excel.Selection.Value        := dVolumen;
                      Excel.Selection.Font.Size   := 6;
                      Excel.Selection.HorizontalAlignment := xlCenter;
                      Excel.Selection.VerticalAlignment   := xlCenter;
                      Excel.Selection.Font.Bold := False;
                      Q_Partidas.Next;
                  end;
                  MiFecha := IncDay(MiFecha);
              end;
              dAvanceGlobal := dAvanceGlobal + (connection.QryBusca.FieldValues['dPonderado'] / 100) * dVolumen;
            end;


            {Aplicamos las formulas}
            Hoja.Range[columnas[total+5]+IntTostr(Ren)+':'+columnas[total+5]+IntToStr(Ren)].Select;
            Excel.Selection.NumberFormat := '#,##0.00';
            if (pos('ANEXO',connection.QryBusca.FieldByName('sTipo').AsString) > 0) then
               Excel.Selection.NumberFormat := '#,##0.000';

            if (pos('ANEXOEXT',connection.QryBusca.FieldByName('sTipo').AsString) > 0) then
               Excel.Selection.NumberFormat := '#,##0.0000';

            if (connection.QryBusca.FieldByName('sTipo').AsString = 'BARCO') then
                Excel.Selection.NumberFormat := '#,##0.0000';

            Excel.Selection.Formula        := '=SUM('+'F'+IntTostr(Ren)+':'+columnas[total+4]+IntToStr(Ren)+')';

            Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
            Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
            Excel.Selection.Value        := '=D'+IntTostr(Ren)+'*'+columnas[total+5]+IntToStr(Ren);

            Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
            Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
            Excel.Selection.Value        := '=E'+IntTostr(Ren)+'*'+columnas[total+5]+IntToStr(Ren);

            Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
            Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
            Excel.Selection.Value        := '='+columnas[total+6]+IntToStr(Ren)+'+'+columnas[total+7]+IntToStr(Ren)+'*'+'E1';

            Hoja.Range[columnas[total+9]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
            if (connection.QryBusca.FieldByName('sTipo').AsString = 'PERSONAL') or (connection.QryBusca.FieldByName('sTipo').AsString = 'EQUIPO') or
               (connection.QryBusca.FieldByName('sTipo').AsString = 'PERNOCTA') then
               Excel.Selection.Value       := 'PEQP'
            else
               Excel.Selection.Value       := connection.QryBusca.FieldByName('sTipo').AsString;
            Excel.Selection.Font.Size   := 5;

            {$ENDREGION}
          connection.QryBusca.Next;
          Inc(Ren);
        end;

        {$REGION 'TOTALES'}
        {Suma de Totales..}
        Hoja.Range[columnas[total+3]+IntTostr(Ren)+':'+columnas[total+3]+IntToStr(Ren)].Select;
        Excel.Selection := 'TOTAL';
        Excel.Selection.Font.Size    := 8;
        Hoja.Range[columnas[total+3]+IntTostr(Ren)+':'+columnas[total+5]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.Formula        := '=SUM('+columnas[total+6]+IntToStr(3)+':'+columnas[total+6]+IntToStr(Ren-1)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.Formula        := '=SUM('+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-1)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.Formula        := '=SUM('+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-1)+')';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Hoja.Range[columnas[total+3]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.Interior.Color :=  RGB(255,192,0);;
        {$ENDREGION}

        {$REGION 'BARCO'}
        {Suma de Cantidades de barco.}
        inc(Ren);
        Hoja.Range[columnas[total+3]+IntTostr(Ren)+':'+columnas[total+3]+IntToStr(Ren)].Select;
        Excel.Selection := 'BARCO';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-2)+','+
                                                 '"BARCO",'+columnas[total+6]+IntToStr(3)+':'+columnas[total+6]+IntToStr(Ren-2)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-2)+','+
                                                 '"BARCO",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-2)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-2)+','+
                                                 '"BARCO",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-2)+')';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$REGION 'PERSONAL, EQUIPO, HOTELERIA'}
        {Suma de Cantidades de Personal, Equipo, Pernocta...}
        inc(Ren);
        Hoja.Range[columnas[total+3]+IntTostr(Ren)+':'+columnas[total+3]+IntToStr(Ren)].Select;
        Excel.Selection := 'MO,EQ,HOSP';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-3)+','+
                                                 '"PEQP",'+columnas[total+6]+IntToStr(3)+':'+columnas[total+6]+IntToStr(Ren-3)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-3)+','+
                                                 '"PEQP",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-3)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-3)+','+
                                                 '"PEQP",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-3)+')';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$REGION 'EXTRAORDINARIAS'}
        {Suma de Cantidades de Partidas de PU...}
        inc(Ren);
        Hoja.Range[columnas[total+3]+IntTostr(Ren)+':'+columnas[total+3]+IntToStr(Ren)].Select;
        Excel.Selection := 'EXTRAORDINARIAS';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-4)+','+
                                                 '"ANEXOEXT",'+columnas[total+6]+IntToStr(3)+':'+columnas[total+6]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-4)+','+
                                                 '"ANEXOEXT",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-4)+','+
                                                 '"ANEXOEXT",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$REGION 'PARTIDAS DE ANEXO C'}
        {Suma de Cantidades de Partidas de PU...}
        inc(Ren);
        Hoja.Range[columnas[total+3]+IntTostr(Ren)+':'+columnas[total+3]+IntToStr(Ren)].Select;
        Excel.Selection := 'PARTIDAS DE ANEXO';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-4)+','+
                                                 '"ANEXO",'+columnas[total+6]+IntToStr(3)+':'+columnas[total+6]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-4)+','+
                                                 '"ANEXO",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-4)+','+
                                                 '"ANEXO",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        i := 4;
        while i<=(Ren-5) do
        begin
            Hoja.Range['A'+IntToStr(i)+':'+columnas[total+8]+IntToStr(i)].Select;
            Excel.Selection.Interior.Color := RGB(218,238,243);
            inc(i,2);
        end;

        {$REGION 'FORMATOS FINALES'}
        Hoja.Range[columnas[total+3]+IntTostr(Ren-3)+':'+columnas[total+5]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 8;
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.VerticalAlignment   := xlCenter;

        Hoja.Range[columnas[total+3]+IntTostr(Ren-3)+':'+columnas[total+5]+IntToStr(Ren-3)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+3]+IntTostr(Ren-2)+':'+columnas[total+5]+IntToStr(Ren-2)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+3]+IntTostr(Ren-1)+':'+columnas[total+5]+IntToStr(Ren-1)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+3]+IntTostr(Ren)+':'+columnas[total+5]+IntToStr(Ren)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range['A3:A'+IntToStr(Ren-4)].Select;
        Excel.Selection.Font.Bold    := True;

        Hoja.Range['D3:E'+IntToStr(Ren-4)].Select;
        Excel.Selection.Font.Bold    := True;

        Hoja.Range[columnas[total+5]+IntTostr(2)+':'+columnas[total+8]+IntToStr(2)].Select;
        Excel.Selection.Font.Size    := 8;

        Hoja.Range[columnas[total+5]+IntTostr(3)+':'+columnas[total+8]+IntToStr(Ren-4)].Select;
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Color   := RGB(0,32,96);
        Excel.Selection.Font.Name    := 'Calibri';

        Hoja.Range['A3:'+columnas[total+8]+IntToStr(Ren-5)].Select;
        Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
        Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
        Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
        Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
        Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
        Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
        Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
        Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
        Excel.Selection.Borders[xlInsideVertical].LineStyle   := xlContinuous;
        Excel.Selection.Borders[xlInsideVertical].Weight      := xlThin;
        Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlInsideHorizontal].Weight    := xlThin;
        {$ENDREGION}

      end;
    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'PRODUCCION ' + tsNumeroOrden.Text;
      except
        Hoja.Name := 'PRODUCCION ';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Programa de Trabajo:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
      end;
    end;

    Result := Resultado;
  end;

begin
    //Verificamos si es un frente
    If (MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO') and (chkPeriodo.Checked = False) Then
    Begin
        messageDLG('Seleccione una Fecha de Inicio y Fin', mtInformation, [mbOk], 0);
        chkPeriodo.Checked := True;
        exit;
    End;

    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
    if not SaveDialog1.Execute then
      Exit;
      // Generar el ambiente de excel
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
    end
    else
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;
    end;

    PanelProgress.Visible := True;
    Label15.Refresh;

    BarraEstado.Position := 0;

    Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

      // Verificar si cuenta con las hojas necesarias
    while Libro.Sheets.Count < 2 do
      Libro.Sheets.Add;

      // Verificar si se pasa de hojas necesarias
    Libro.Sheets[1].Select;
    while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

      // Proceder a generar la hoja REPORTE
    CadError := '';

    if GenerarPlantilla then
    begin
          // Grabar el archivo de excel con el nombre dado
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      PanelProgress.Visible := False;
      messageDlg('El Archivo se generó Correctamente!', mtInformation, [mbOk], 0);
    end;

    Excel := '';

    if CadError <> '' then
      showmessage(CadError);
end;

procedure TfrmCompara.cmdProgramadoClick(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// SEGUIMINTO DE AVANCES X PARTIDA DIAVAZ OCTUBRE 2012 //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena: string;
      fs: tStream;
      Alto: Extended;
      MiFechaI, MiFechaF, MiFecha: tDate;
      Ren, nivel, i, total: integer;
      iInicioMes, iFinMes : integer;
      Q_Partidas, Q_Programado : TZReadOnlyQuery;
      dVolumen, dAvanceGlobal, dProgramado, dFisico: double;
      Progreso, TotalProgreso: real;
      lEncuentra : boolean;
      sColInicio, sColFinal : string;

    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      Q_Programado := TZReadOnlyQuery.Create(self);
      Q_Programado.Connection := connection.zConnection;

      {$REGION 'ENCABEZADO'}
      Ren := 2;
      //Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 130;

      Excel.Columns['A:A'].ColumnWidth := 10.86;
      Excel.Columns['B:B'].ColumnWidth := 37.29;
      Excel.Columns['C:C'].ColumnWidth := 6.43;
      Excel.Columns['D:D'].ColumnWidth := 8.14;
      Excel.Columns['E:E'].ColumnWidth := 7.57;
      Excel.Columns['F:F'].ColumnWidth := 4.5;

      Hoja.Range['A1:A3'].Select;
      Excel.Selection.RowHeight := '15';

      Hoja.Range['A1:A1'].Select;
      Excel.Selection.Value := tsNumeroOrden.Text;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Size := 9;
      Excel.Selection.Font.Name := 'Calibri';

      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Clear;
      connection.QryBusca2.SQL.Add('Select p.sDescripcion, o.dFiProgramado, o.dFfProgramado from ordenesdetrabajo o '+
                                  'inner join plataformas p on (o.sIdPlataforma = o.sIdPlataforma) '+
                                  'where o.sContrato = :contrato and o.sNumeroOrden =:Orden');
      connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato;
      connection.QryBusca2.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
      connection.QryBusca2.Open;

      Hoja.Range['B1:B1'].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldByName('sDescripcion').AsString;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Size := 9;
      Excel.Selection.Font.Name := 'Calibri';

      Hoja.Range['D1:D1'].Select;
      Excel.Selection.Value := 'TC';
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Size := 8;
      Excel.Selection.Font.Name := 'Calibri';

      Ren := 2;
      // Colocar los encabezados de la plantilla...
      Hoja.Range['A'+IntTostr(Ren)+':A'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Partida';
      FormatoEncabezado;
      Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Descripcion';
      FormatoEncabezado;
      Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Unidad';
      FormatoEncabezado;
      Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Precio MXN';
      FormatoEncabezado;
      Hoja.Range['E'+IntTostr(Ren)+':E'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Precio USD';
      FormatoEncabezado;
      Hoja.Range['F'+IntTostr(Ren)+':F'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Avance';
      FormatoEncabezado;

      Hoja.Range['A'+IntTostr(Ren)+':F'+IntToStr(Ren)].Select;
      Excel.Selection.Interior.Color := RGB(79,129,189);
      Excel.Selection.Font.Color     := clWhite;

      {$ENDREGION}

      {$REGION 'DURACION Y TOTAL'}
      connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select a.sAnexo, a.sTipo, aa.* from actividadesxanexo aa '+
                                  'inner join anexos a on (aa.sAnexo = a.sAnexo) '+
                                  'where aa.sContrato =:Contrato and aa.sIdConvenio =:Convenio and aa.sTipoActividad = "Actividad"');
      connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Contrato').Value    := global_contrato;
      connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Convenio').Value    := global_convenio;
      if connection.configuracion.FieldByName('lOrdenaItem').AsString = 'No' then
      begin
          connection.QryBusca.SQL.Add(' order by mysql.udf_NaturalSortFormat(swbs,:Tam,:Separador)');
          connection.QryBusca.ParamByName('tam').AsInteger      := Global_TamOrden;
          connection.QryBusca.ParamByName('separador').AsString := Global_SepOrden;
      end
      else
         connection.QryBusca.SQL.Add('order by iItemOrden');
      connection.QryBusca.Open;
      if connection.QryBusca.RecordCount > 0 then
      begin
        if chkPeriodo.Checked then
        begin
            MiFecha  := tdIdFecha1.Date;
            MiFechaI := tdIdFecha1.Date;
            MiFechaF := tdIdFecha.Date;
        end
        else
        begin
            MiFecha  := connection.QryBusca2.FieldByName('dFiProgramado').AsDateTime;
            MiFechaI := connection.QryBusca2.FieldByName('dFiProgramado').AsDateTime;
            MiFechaF := connection.QryBusca2.FieldByName('dFfProgramado').AsDateTime;
        end;

        iInicioMes := 7;
        for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
        begin
            Hoja.Cells[Ren, 6 + i].Select;
            {Formato de las fechas archivo Excel,, 24/07/2011..}
            Excel.Selection.Value := copy(DateToStr(Mifecha), 0,2);

            iFinMes    := StrToInt(copy(DateToStr(Mifecha), 0,2));
            if iFinMes = 1 then
            begin
                Hoja.Range[columnas[iInicioMes]+IntTostr(Ren -1)+':'+columnas[i+5]+IntToStr(Ren-1)].Select;
                FormatoEncabezado;
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Value      := meses[StrToInt(copy(DateToStr(Mifecha-1), 4,2))];
                Excel.Selection.MergeCells := True;
                Excel.Selection.Font.Bold  := True;
                Excel.Selection.Font.Size  := 11;
                Excel.Selection.Font.Color := RGB(0,32,96);
                Excel.Selection.Font.Name  := 'Calibri';
                iInicioMes := i + 6;
            end;
            MiFecha := IncDay(MiFecha);
        end;

        if iFinMes < 28 then
        begin
            Hoja.Range[columnas[iInicioMes]+IntTostr(Ren -1)+':'+columnas[i+5]+IntToStr(Ren-1)].Select;
            FormatoEncabezado;
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;
            Excel.Selection.Value      := meses[StrToInt(copy(DateToStr(Mifecha-1), 4,2))];
            Excel.Selection.MergeCells := True;
            Excel.Selection.Font.Bold  := True;
            Excel.Selection.Font.Size  := 11;
            Excel.Selection.Font.Name  := 'Calibri';
        end;
        total := i;

        {Colocamos los anchos de las columnas..}
        Hoja.Range['G'+IntTostr(Ren)+':'+columnas[total+5]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Bold := False;
        Excel.Selection.Font.Size := 8;
        Excel.Selection.Font.Name := 'Calibri';
        Excel.Selection.Interior.ColorIndex := 37;
        FormatoEncabezado;
        Excel.Selection.ColumnWidth := 4;

        {Totales de Concentrato de Produccion}
        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size := 7;
        Excel.Selection.Font.Name := 'Calibri';

        Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Value  := 'Total';
        Excel.Selection.ColumnWidth := 7.57;

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Value := 'Prod. MXN';
        Excel.Selection.ColumnWidth := 9.43;

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Value := 'Prod. USD';
        Excel.Selection.ColumnWidth := 9.43;

        Hoja.Range[columnas[total+9]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Value := 'Prod. Homol';
        Excel.Selection.ColumnWidth := 9.43;

        Hoja.Range[columnas[total+6]+IntTostr(Ren-1)+':'+columnas[total+9]+IntToStr(Ren-1)].Select;
        FormatoEncabezado;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Bold  := True;
        Excel.Selection.Font.Size  := 11;
        Excel.Selection.Font.Name  := 'Calibri';
        Excel.Selection.Value := 'Producción Real';
        Excel.Selection.MergeCells := True;
        Excel.Selection.Interior.ColorIndex := 37;

        Hoja.Range[columnas[total+10]+IntTostr(Ren)+':'+columnas[total+10]+IntToStr(Ren)].Select;
        Excel.Selection.ColumnWidth := 0;

        {$ENDREGION}

        inc(Ren,2);
        dAvanceGlobal := 0;
        connection.QryBusca.First;
        while not connection.QryBusca.Eof do
        begin
            {$REGION 'PARTIDAS DE ANEXO'}
           {Movimiento de la Barra..}
            Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
            TotalProgreso := TotalProgreso + Progreso;
            BarraEstado.Position := Trunc(TotalProgreso);

            {Escritura de Datos en el Archvio de Excel..}
            Hoja.Cells[Ren, 1].Select;
            Excel.Selection.Value := connection.QryBusca.FieldValues['sNumeroActividad'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;

            Hoja.Cells[Ren, 2].Select;
            Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];
            Excel.Selection.HorizontalAlignment := xlJustify;
            Excel.Selection.VerticalAlignment   := xlCenter;
            Excel.Selection.MergeCells := True;
            Excel.Selection.WrapText   := True;

            Hoja.Cells[Ren, 3].Select;
            Excel.Selection.Value := connection.QryBusca.FieldValues['sMedida'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;

            Hoja.Cells[Ren, 4].Select;
            Excel.Selection.NumberFormat := '$ #,##0.00';
            Excel.Selection.Value := connection.QryBusca.FieldValues['dVentaMN'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;

            Hoja.Cells[Ren, 5].Select;
            Excel.Selection.NumberFormat := '#,##0.00';
            Excel.Selection.Value := connection.QryBusca.FieldValues['dVentaDLL'];
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;

            Hoja.Range['A'+IntTostr(Ren)+':F'+IntToStr(Ren)].Select;
            Excel.Selection.Font.Size := 7;
            Excel.Selection.Font.Bold := False;
            Excel.Selection.Font.Name := 'Calibri';

            {$ENDREGION}

            MiFecha := MiFechaI;

            {$REGION 'CONSULTA DE ANEXOS'}
            {CONSUL DE ANEXOS}
            if connection.QryBusca.FieldByName('sTipo').AsString = 'BARCO' then
            begin
                {Consultamos la partida de barco..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('select dIdFecha, sClasificacion, sum(sFactor) as dCantidad '+
                                   'from movimientosdeembarcacion where sOrden = :Contrato and sClasificacion = :Actividad and dIdFecha <= :Final '+
                                   'Group By sClasificacion, dIdFecha Order By sClasificacion,dIdFecha');
                Q_Partidas.Params.ParamByName('Contrato').DataType  := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value     := global_contrato;
                Q_Partidas.Params.ParamByName('Final').DataType     := ftDate;
                Q_Partidas.Params.ParamByName('Final').Value        := MiFechaF;
                Q_Partidas.Params.ParamByName('Actividad').DataType := ftString;
                Q_Partidas.Params.ParamByName('Actividad').Value    := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                Q_Partidas.Open;
            end;


            if connection.QryBusca.FieldByName('sTipo').AsString = 'PERSONAL' then
            begin
                {Consultamos la partida de personal..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('select dIdFecha, sIdPersonal, round(sum(dCantHH),2) as dCantidad '+
                                   'from bitacoradepersonal where sContrato =:Contrato and sIdPersonal = :IdPersonal and dIdFecha <=:Final '+
                                   'Group By sIdPersonal, dIdFecha Order By sIdPersonal,dIdFecha');
                Q_Partidas.Params.ParamByName('Contrato').DataType   := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value      := global_contrato;
                Q_Partidas.Params.ParamByName('Final').DataType      := ftDate;
                Q_Partidas.Params.ParamByName('Final').Value         := MiFechaF;
                Q_Partidas.Params.ParamByName('IdPersonal').DataType := ftString;
                Q_Partidas.Params.ParamByName('IdPersonal').Value    := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                Q_Partidas.Open;
            end;

            if connection.QryBusca.FieldByName('sTipo').AsString = 'EQUIPO' then
            begin
                {Consultamos la partida de equipo..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('select dIdFecha, sIdEquipo, round(sum(dCantHH),2) as dCantidad '+
                                   'from bitacoradeequipos where sContrato =:Contrato and sIdEquipo = :IdEquipo and dIdFecha <=:Final '+
                                   'Group By sIdEquipo, dIdFecha Order By sIdEquipo,dIdFecha');
                Q_Partidas.Params.ParamByName('Contrato').DataType := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value    := global_contrato;
                Q_Partidas.Params.ParamByName('Final').DataType    := ftDate;
                Q_Partidas.Params.ParamByName('Final').Value       := MiFechaF;
                Q_Partidas.Params.ParamByName('IdEquipo').DataType := ftString;
                Q_Partidas.Params.ParamByName('IdEquipo').Value    := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                Q_Partidas.Open;
            end;

            if connection.QryBusca.FieldByName('sTipo').AsString = 'PERNOCTA' then
            begin
                {Consultamos la partida de barco..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('select dIdFecha, sIdPersonal, round(sum(dCantHH),2) as dCantidad '+
                                   'from bitacoradepersonal where sContrato =:Contrato and sIdPersonal = :IdPersonal and dIdFecha <=:Final '+
                                   'Group By sIdPersonal, dIdFecha Order By sIdPersonal,dIdFecha');
                Q_Partidas.Params.ParamByName('Contrato').DataType   := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value      := global_contrato;
                Q_Partidas.Params.ParamByName('Final').DataType      := ftDate;
                Q_Partidas.Params.ParamByName('Final').Value         := MiFechaF;
                Q_Partidas.Params.ParamByName('IdPersonal').DataType := ftString;
                Q_Partidas.Params.ParamByName('IdPersonal').Value    := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                Q_Partidas.Open;
            end;

            if pos('ANEXO',connection.QryBusca.FieldByName('sTipo').AsString) > 0 then
            begin
                {Consultamos partida de PU..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('Select b.sWbs,b.sNumeroActividad, sum(a.dCantidad) as dCantidad, a.dIdFecha, b.dCantidad as dVolumen ' +
                                  'From actividadesxorden b ' +
                                  'left JOIN bitacoradeactividades a ' +
                                  'ON (a.sContrato=b.sContrato and b.sIdConvenio = a.sIdConvenio And a.sWbs=b.sWbs And a.dIdFecha <=:Final and b.sNumeroOrden=a.sNumeroOrden) ' +
                                  'left JOIN tiposdemovimiento t ' +
                                  'ON (b.sContrato=t.sContrato And a.sIdTipoMovimiento=t.sIdTipoMovimiento And t.sClasificacion="Tiempo en Operacion") ' +
                                  'Where b.sContrato=:Contrato And b.sIdConvenio=:Convenio And b.sNumeroOrden =:Orden and b.sWbsContrato =:Wbs ' +
                                  'Group By b.sWbs,a.dIdFecha Order By b.sNumeroActividad,b.iItemOrden,a.dIdFecha');
                Q_Partidas.Params.ParamByName('Contrato').DataType := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value    := global_contrato;
                Q_Partidas.Params.ParamByName('Convenio').DataType := ftString;
                Q_Partidas.Params.ParamByName('Convenio').Value    := zqReprogramacion.FieldByName('sIdConvenio').AsString;
                Q_Partidas.Params.ParamByName('Orden').DataType    := ftString;
                Q_Partidas.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text;
                Q_Partidas.Params.ParamByName('Final').DataType    := ftDate;
                Q_Partidas.Params.ParamByName('Final').Value       := MiFechaF;
                Q_Partidas.Params.ParamByName('Wbs').DataType      := ftString;
                Q_Partidas.Params.ParamByName('Wbs').Value         := connection.QryBusca.FieldByName('sWbs').AsString;
                Q_Partidas.Open;
            end;

            //Consulta de la distribucion de anexo.
            Q_Programado.Active := False;
            Q_Programado.SQL.Clear;
            Q_Programado.SQL.Add('select dIdFecha, dCantidad from distribuciondeanexo where sContrato =:contrato and sIdConvenio =:convenio and sWbs =:Wbs ');
            Q_Programado.Params.ParamByName('Contrato').DataType := ftString;
            Q_Programado.Params.ParamByName('Contrato').Value    := global_contrato;
            Q_Programado.Params.ParamByName('Convenio').DataType := ftString;
            Q_Programado.Params.ParamByName('Convenio').Value    := global_convenio;
            Q_Programado.Params.ParamByName('Wbs').DataType      := ftString;
            Q_Programado.Params.ParamByName('Wbs').Value         := connection.QryBusca.FieldByName('sWbs').AsString;
            Q_Programado.Open;

            MiFecha := MiFechaI;
            if Q_Partidas.RecordCount > 0 then
            begin
              dVolumen := 0;
              for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
              begin
                  if MiFecha = Q_Partidas.FieldByName('dIdFecha').AsDateTime then
                  begin
                      Hoja.Cells[Ren, 6 + i].Select;
                      dVolumen := Q_Partidas.FieldByName('dCantidad').AsFloat;
                      if (pos('ANEXO',connection.QryBusca.FieldByName('sTipo').AsString) > 0) or
                         (connection.QryBusca.FieldByName('sTipo').AsString = 'BARCO') then
                          Excel.Selection.NumberFormat := '#,##0.0000'
                      else
                          Excel.Selection.NumberFormat := '#,##0.00';
                      Excel.Selection.Value        := dVolumen;
                      Q_Partidas.Next;
                  end;

                  if MiFecha = Q_Programado.FieldByName('dIdFecha').AsDateTime then
                  begin
                      Hoja.Cells[Ren-1, 6 + i].Select;
                      if (pos('ANEXO',connection.QryBusca.FieldByName('sTipo').AsString) > 0) or
                         (connection.QryBusca.FieldByName('sTipo').AsString = 'BARCO') then
                          Excel.Selection.NumberFormat := '#,##0.0000'
                      else
                          Excel.Selection.NumberFormat := '#,##0.00';
                      Excel.Selection.Value       := Q_Programado.FieldByName('dCantidad').AsFloat;
                      Q_Programado.Next;
                  end;

                  MiFecha := IncDay(MiFecha);
              end;
              dAvanceGlobal := dAvanceGlobal + (connection.QryBusca.FieldValues['dPonderado'] / 100) * dVolumen;
            end;

            {Aplicamos las formulas}
            {Real}
            Hoja.Range[columnas[total+6]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
            if (pos('ANEXO',connection.QryBusca.FieldByName('sTipo').AsString) > 0) or
               (connection.QryBusca.FieldByName('sTipo').AsString = 'BARCO') then
                Excel.Selection.NumberFormat := '#,##0.0000'
            else
                Excel.Selection.NumberFormat := '#,##0.00';
            Excel.Selection.Formula        := '=SUM('+'G'+IntTostr(Ren)+':'+columnas[total+5]+IntToStr(Ren)+')';

            Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
            Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
            Excel.Selection.Value        := '=D'+IntTostr(Ren)+'*'+columnas[total+6]+IntToStr(Ren);

            Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
            Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
            Excel.Selection.Value        := '=E'+IntTostr(Ren)+'*'+columnas[total+6]+IntToStr(Ren);

            Hoja.Range[columnas[total+9]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
            Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
            Excel.Selection.Value        := '='+columnas[total+7]+IntToStr(Ren)+'+'+columnas[total+8]+IntToStr(Ren)+'*'+'E1';

            Hoja.Range[columnas[total+10]+IntTostr(Ren)+':'+columnas[total+10]+IntToStr(Ren)].Select;
            if (connection.QryBusca.FieldByName('sTipo').AsString = 'PERSONAL') or (connection.QryBusca.FieldByName('sTipo').AsString = 'EQUIPO') or
               (connection.QryBusca.FieldByName('sTipo').AsString = 'PERNOCTA') then
               Excel.Selection.Value       := 'PEQP'
            else
               Excel.Selection.Value    := connection.QryBusca.FieldByName('sTipo').AsString;
            Excel.Selection.Font.Size   := 5;

          {Programado..}
            Hoja.Range[columnas[total+6]+IntTostr(Ren-1)+':'+columnas[total+6]+IntToStr(Ren-1)].Select;
            if (pos('ANEXO',connection.QryBusca.FieldByName('sTipo').AsString) > 0) or
               (connection.QryBusca.FieldByName('sTipo').AsString = 'BARCO') then
                Excel.Selection.NumberFormat := '#,##0.0000'
            else
                Excel.Selection.NumberFormat := '#,##0.00';
            Excel.Selection.Formula      := '=SUM('+'G'+IntTostr(Ren-1)+':'+columnas[total+5]+IntToStr(Ren-1)+')';

            Hoja.Range[columnas[total+7]+IntTostr(Ren-1)+':'+columnas[total+7]+IntToStr(Ren-1)].Select;
            Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
            Excel.Selection.Value        := '=D'+IntTostr(Ren)+'*'+columnas[total+6]+IntToStr(Ren-1);

            Hoja.Range[columnas[total+8]+IntTostr(Ren-1)+':'+columnas[total+8]+IntToStr(Ren-1)].Select;
            Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
            Excel.Selection.Value        := '=E'+IntTostr(Ren)+'*'+columnas[total+6]+IntToStr(Ren-1);

            Hoja.Range[columnas[total+9]+IntTostr(Ren-1)+':'+columnas[total+9]+IntToStr(Ren-1)].Select;
            Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
            Excel.Selection.Value        := '='+columnas[total+7]+IntToStr(Ren-1)+'+'+columnas[total+8]+IntToStr(Ren-1)+'*'+'E1';

            Hoja.Range[columnas[total+10]+IntTostr(Ren-1)+':'+columnas[total+10]+IntToStr(Ren-1)].Select;
            if (connection.QryBusca.FieldByName('sTipo').AsString = 'PERSONAL') or (connection.QryBusca.FieldByName('sTipo').AsString = 'EQUIPO') or
               (connection.QryBusca.FieldByName('sTipo').AsString = 'PERNOCTA') then
               Excel.Selection.Value       := 'PEQP.'
            else
               Excel.Selection.Value    := connection.QryBusca.FieldByName('sTipo').AsString + '.';
            Excel.Selection.Font.Size   := 5;


            Hoja.Range['F'+IntToStr(Ren-1)+':F'+IntToStr(Ren-1)].Select;
            Excel.Selection.Value := 'PROG.';

            Hoja.Range['F'+IntToStr(Ren)+':F'+IntToStr(Ren)].Select;
            Excel.Selection.Value := 'REAL';

            {$ENDREGION}
          connection.QryBusca.Next;
          Inc(Ren, 2);


        end;

        {$REGION 'TOTALES FISICOS'}

        {$REGION 'TOTALES'}
        {Suma de Totales..}
        Hoja.Range[columnas[total+4]+IntTostr(Ren)+':'+columnas[total+4]+IntToStr(Ren)].Select;
        Excel.Selection := 'TOTAL (REAL)';
        Excel.Selection.Font.Size    := 8;
        Hoja.Range[columnas[total+4]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI(F3:F'+IntToStr(Ren-1)+','+
                                          '"REAL",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-1)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI(F3:F'+IntToStr(Ren-1)+','+
                                          '"REAL",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-1)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI(F3:F'+IntToStr(Ren-1)+','+
                                          '"REAL",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-1)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Hoja.Range[columnas[total+4]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        Excel.Selection.Interior.Color :=  RGB(255,192,0);;
        {$ENDREGION}

        {$REGION 'BARCO'}
        {Suma de Cantidades de barco.}
        inc(Ren);
        Hoja.Range[columnas[total+4]+IntTostr(Ren)+':'+columnas[total+4]+IntToStr(Ren)].Select;
        Excel.Selection := 'BARCO';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-2)+','+
                                                 '"BARCO",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-2)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-2)+','+
                                                 '"BARCO",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-2)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-2)+','+
                                                 '"BARCO",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-2)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$REGION 'PERSONAL, EQUIPO, HOTELERIA'}
        {Suma de Cantidades de Personal, Equipo, Pernocta...}
        inc(Ren);
        Hoja.Range[columnas[total+4]+IntTostr(Ren)+':'+columnas[total+4]+IntToStr(Ren)].Select;
        Excel.Selection := 'MO,EQ,HOSP';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-3)+','+
                                                 '"PEQP",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-3)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-3)+','+
                                                 '"PEQP",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-3)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-3)+','+
                                                 '"PEQP",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-3)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$REGION 'EXTRAORDINARIAS'}
        {Suma de Cantidades de Partidas de PU...}
        inc(Ren);
        Hoja.Range[columnas[total+4]+IntTostr(Ren)+':'+columnas[total+4]+IntToStr(Ren)].Select;
        Excel.Selection := 'EXTRAORDINARIAS';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-4)+','+
                                                 '"ANEXOEXT",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-4)+','+
                                                 '"ANEXOEXT",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-4)+','+
                                                 '"ANEXOEXT",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$REGION 'PARTIDAS DE ANEXO C'}
        {Suma de Cantidades de Partidas de PU...}
        inc(Ren);
        Hoja.Range[columnas[total+4]+IntTostr(Ren)+':'+columnas[total+4]+IntToStr(Ren)].Select;
        Excel.Selection := 'PARTIDAS DE ANEXO';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+7]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-4)+','+
                                                 '"ANEXO",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren)+':'+columnas[total+8]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-4)+','+
                                                 '"ANEXO",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-4)+','+
                                                 '"ANEXO",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-4)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren)+':'+columnas[total+9]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$ENDREGION}


        {$REGION 'TOTALES PROGRAMADOS'}

        {$REGION 'TOTALES'}
        {Suma de Totales..}
        Hoja.Range[columnas[total+4]+IntTostr(Ren+2)+':'+columnas[total+4]+IntToStr(Ren+2)].Select;
        Excel.Selection := 'TOTAL (PROGRAMADO)';
        Excel.Selection.Font.Size    := 8;
        Hoja.Range[columnas[total+4]+IntTostr(Ren+2)+':'+columnas[total+6]+IntToStr(Ren+2)].Select;
        FormatoEncabezado;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+7]+IntTostr(Ren+2)+':'+columnas[total+7]+IntToStr(Ren+2)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI(F3:F'+IntToStr(Ren-5)+','+
                                          '"PROG.",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren+2)+':'+columnas[total+8]+IntToStr(Ren+2)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI(F3:F'+IntToStr(Ren-5)+','+
                                          '"PROG.",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren+2)+':'+columnas[total+9]+IntToStr(Ren+2)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI(F3:F'+IntToStr(Ren-5)+','+
                                          '"PROG.",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren+2)+':'+columnas[total+9]+IntToStr(Ren+2)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Hoja.Range[columnas[total+4]+IntTostr(Ren+2)+':'+columnas[total+9]+IntToStr(Ren+2)].Select;
        Excel.Selection.Interior.Color :=  RGB(204,192,218);;
        {$ENDREGION}

        {$REGION 'BARCO'}
        {Suma de Cantidades de barco.}
        Hoja.Range[columnas[total+4]+IntTostr(Ren+3)+':'+columnas[total+4]+IntToStr(Ren+3)].Select;
        Excel.Selection := 'BARCO';

        Hoja.Range[columnas[total+7]+IntTostr(Ren+3)+':'+columnas[total+7]+IntToStr(Ren+3)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"BARCO.",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren+3)+':'+columnas[total+8]+IntToStr(Ren+3)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"BARCO.",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren+3)+':'+columnas[total+9]+IntToStr(Ren+3)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"BARCO.",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren+3)+':'+columnas[total+9]+IntToStr(Ren+3)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$REGION 'PERSONAL, EQUIPO, HOTELERIA'}
        {Suma de Cantidades de Personal, Equipo, Pernocta...}
        Hoja.Range[columnas[total+4]+IntTostr(Ren+4)+':'+columnas[total+4]+IntToStr(Ren+4)].Select;
        Excel.Selection := 'MO,EQ,HOSP';

        Hoja.Range[columnas[total+7]+IntTostr(Ren+4)+':'+columnas[total+7]+IntToStr(Ren+4)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"PEQP.",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren+4)+':'+columnas[total+8]+IntToStr(Ren+4)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"PEQP.",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren+4)+':'+columnas[total+9]+IntToStr(Ren+4)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"PEQP.",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren+4)+':'+columnas[total+9]+IntToStr(Ren+4)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$REGION 'EXTRAORDINARIAS'}
        {Suma de Cantidades de Partidas de PU...}
        Hoja.Range[columnas[total+4]+IntTostr(Ren+5)+':'+columnas[total+4]+IntToStr(Ren+5)].Select;
        Excel.Selection := 'EXTRAORDINARIAS';

        Hoja.Range[columnas[total+7]+IntTostr(Ren+5)+':'+columnas[total+7]+IntToStr(Ren+5)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"ANEXOEXT.",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren+5)+':'+columnas[total+8]+IntToStr(Ren+5)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"ANEXOEXT.",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren+5)+':'+columnas[total+9]+IntToStr(Ren+5)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"ANEXOEXT.",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren+5)+':'+columnas[total+9]+IntToStr(Ren+5)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$REGION 'PARTIDAS DE ANEXO C'}
        {Suma de Cantidades de Partidas de PU...}
        Hoja.Range[columnas[total+4]+IntTostr(Ren+6)+':'+columnas[total+4]+IntToStr(Ren+6)].Select;
        Excel.Selection := 'PARTIDAS DE ANEXO';

        Hoja.Range[columnas[total+7]+IntTostr(Ren+6)+':'+columnas[total+7]+IntToStr(Ren+6)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"ANEXO.",'+columnas[total+7]+IntToStr(3)+':'+columnas[total+7]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+8]+IntTostr(Ren+6)+':'+columnas[total+8]+IntToStr(Ren+6)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"ANEXO.",'+columnas[total+8]+IntToStr(3)+':'+columnas[total+8]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+9]+IntTostr(Ren+6)+':'+columnas[total+9]+IntToStr(Ren+6)].Select;
        Excel.Selection.NumberFormat := '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)';
        Excel.Selection.FormulaArray        := '=SUMAR.SI('+columnas[total+10]+IntToStr(3)+':'+columnas[total+10]+IntToStr(Ren-5)+','+
                                                 '"ANEXO.",'+columnas[total+9]+IntToStr(3)+':'+columnas[total+9]+IntToStr(Ren-5)+')';

        Hoja.Range[columnas[total+7]+IntTostr(Ren+6)+':'+columnas[total+9]+IntToStr(Ren+6)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Interior.ColorIndex := 37;
        {$ENDREGION}

        {$ENDREGION}


        i := 3;
        while i<=(Ren-4) do
        begin
            Hoja.Range['A'+IntToStr(i)+':'+columnas[total+9]+IntToStr(i)].Select;
            Excel.Selection.Interior.Color := RGB(218,238,243);
            inc(i,2);
        end;

        {$REGION 'FORMATOS FINALES'}
        {Leyendas barco, mo,eq, hosp, extraordinarias.. FISICO}
        Hoja.Range[columnas[total+4]+IntTostr(Ren-3)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 8;
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Interior.Color := RGB(217,217,217);

        Hoja.Range[columnas[total+4]+IntTostr(Ren-3)+':'+columnas[total+6]+IntToStr(Ren-3)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+4]+IntTostr(Ren-2)+':'+columnas[total+6]+IntToStr(Ren-2)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+4]+IntTostr(Ren-1)+':'+columnas[total+6]+IntToStr(Ren-1)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+4]+IntTostr(Ren)+':'+columnas[total+6]+IntToStr(Ren)].Select;
        Excel.Selection.MergeCells := True;


         {Leyendas barco, mo,eq, hosp, extraordinarias.. PROGRAMADO}
        Hoja.Range[columnas[total+4]+IntTostr(Ren+3)+':'+columnas[total+6]+IntToStr(Ren+6)].Select;
        FormatoEncabezado;
        Excel.Selection.Font.Size    := 8;
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Interior.Color := RGB(217,217,217);

        Hoja.Range[columnas[total+4]+IntTostr(Ren+3)+':'+columnas[total+6]+IntToStr(Ren+3)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+4]+IntTostr(Ren+4)+':'+columnas[total+6]+IntToStr(Ren+4)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+4]+IntTostr(Ren+5)+':'+columnas[total+6]+IntToStr(Ren+5)].Select;
        Excel.Selection.MergeCells := True;

        Hoja.Range[columnas[total+4]+IntTostr(Ren+6)+':'+columnas[total+6]+IntToStr(Ren+6)].Select;
        Excel.Selection.MergeCells := True;

        {Columna de avances..}
        Hoja.Range['F3:F'+IntToStr(Ren-5)].Select;
        Excel.Selection.Font.Size    := 5;
        Excel.Selection.Font.Color   := RGB(79,129,189);
        Excel.Selection.Font.Name    := 'Calibri';
        Excel.Selection.HorizontalAlignment := xlRight;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Bold    := True;

        Hoja.Range['G'+IntToStr(3)+':'+columnas[total+5]+IntToStr(Ren-5)].Select;
        Excel.Selection.Font.Size   := 6;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Bold := False;

        Hoja.Range['A3:A'+IntToStr(Ren-5)].Select;
        Excel.Selection.Font.Bold    := True;

        Hoja.Range['D3:E'+IntToStr(Ren-5)].Select;
        Excel.Selection.Font.Bold    := True;

        Hoja.Range[columnas[total+7]+IntTostr(2)+':'+columnas[total+9]+IntToStr(2)].Select;
        Excel.Selection.Font.Size    := 8;

        Hoja.Range[columnas[total+6]+IntTostr(3)+':'+columnas[total+9]+IntToStr(Ren-5)].Select;
        Excel.Selection.Font.Bold    := True;
        Excel.Selection.Font.Size    := 7;
        Excel.Selection.Font.Color   := RGB(0,32,96);
        Excel.Selection.Font.Name    := 'Calibri';

        Hoja.Range['A3:'+columnas[total+9]+IntToStr(Ren-5)].Select;
        Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
        Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
        Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
        Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
        Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
        Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
        Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
        Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
        Excel.Selection.Borders[xlInsideVertical].LineStyle   := xlContinuous;
        Excel.Selection.Borders[xlInsideVertical].Weight      := xlThin;
        Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous;
        Excel.Selection.Borders[xlInsideHorizontal].Weight    := xlThin;
        {$ENDREGION}

      end;
    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'PRODUCCION ' + tsNumeroOrden.Text;
      except
        Hoja.Name := 'PRODUCCION ';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Programa de Trabajo:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
      end;
    end;

    Result := Resultado;
  end;

begin
    //Verificamos si es un frente
    If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
    Begin
        messageDLG('Seleccione un frente de trabajo!', mtInformation, [mbOk], 0);
        exit;
    End;

    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
    if not SaveDialog1.Execute then
      Exit;
      // Generar el ambiente de excel
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
    end
    else
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;
    end;

    PanelProgress.Visible := True;
    Label15.Refresh;

    BarraEstado.Position := 0;

    Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

      // Verificar si cuenta con las hojas necesarias
    while Libro.Sheets.Count < 2 do
      Libro.Sheets.Add;

      // Verificar si se pasa de hojas necesarias
    Libro.Sheets[1].Select;
    while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

      // Proceder a generar la hoja REPORTE
    CadError := '';

    if GenerarPlantilla then
    begin
          // Grabar el archivo de excel con el nombre dado
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      PanelProgress.Visible := False;
      messageDlg('El Archivo se generó Correctamente!', mtInformation, [mbOk], 0);
    end;

    Excel := '';

    if CadError <> '' then
      showmessage(CadError);

end;

procedure TfrmCompara.cmdComparativoClick(Sender: TObject);
var
    wbsContrato: string;
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
  try
    RxMDValida.Active := True;
    If RxMDValida.RecordCount > 0 then
        RxMDValida.EmptyTable ;

    Connection.qryBusca2.Active := False ;
    Connection.qryBusca2.SQL.Clear ;
    Connection.qryBusca2.SQL.Add('select sWbsContrato from actividadesxorden WHERE sContrato = :contrato AND sIdConvenio = :convenio AND sTipoActividad = "Actividad" group by sWbsContrato') ;
    connection.qryBusca2.ParamByName('contrato').Value := global_contrato;
    connection.qryBusca2.ParamByName('convenio').Value := global_convenio;
    Connection.qryBusca2.Open ;

    if Connection.qryBusca2.RecordCount > 0 then
    begin
      Progress.Visible := True ;
      Progress.MinValue := 0 ;
      Progress.MaxValue := connection.QryBusca2.RecordCount ;
    end;

    while not connection.qryBusca2.Eof do begin
        wbsContrato := connection.qryBusca2.FieldByName('sWbsContrato').AsString;
        if wbsContrato <> '' then
            ventasDiferentes(wbsContrato, cantidadesDiferentes(wbsContrato) );
        connection.qryBusca2.Next;
        Progress.Progress := connection.qryBusca2.RecNo;
    end;

    if RxMDValida.RecordCount > 0 then begin
        rInforme.LoadFromFile (global_files + 'validaActOrden.fr3') ;
        rInforme.PreviewOptions.MDIChild := True ;
        rInforme.PreviewOptions.Modal := False ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
        if not FileExists(global_files + 'validaActOrden.fr3') then
           showmessage('El archivo de reporte validaActOrden.fr3 no existe, notifique al administrador del sistema');

    end
    else
       messageDLG('No existen diferencias entre el Anexo C y los Frentes!', mtInformation, [mbOk], 0);
    Progress.Visible := False ;
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al generar comparativo de programa de trabajo del anexo C ', 0);
    end;
  end;
end;

procedure TfrmCompara.chkadmonEnter(Sender: TObject);
begin
  ChkPu.Checked := False ;
end;

procedure TfrmCompara.chkDLLClick(Sender: TObject);
begin
      chkMN.Checked  := False;
end;

procedure TfrmCompara.chkDLLEnter(Sender: TObject);
begin
     chkMN.Checked  := False;
end;

procedure TfrmCompara.chkPuEnter(Sender: TObject);
begin
 ChkAdmon.Checked := False
end;

procedure TfrmCompara.cmdConceptosClick(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// SEGUIMINTO DE AVANCES X PARTIDA DIAVAZ OCTUBRE 2012 //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena: string;
      fs: tStream;
      Alto: Extended;
      MiFechaI, MiFechaF, MiFecha: tDate;
      Ren, nivel, i, total: integer;
      Q_Partidas: TZReadOnlyQuery;
      dVolumen, dAvanceGlobal, dProgramado, dFisico: double;
      Progreso, TotalProgreso: real;
      lEncuentra : boolean;
      sColInicio, sColFinal : string;
      lContinua : boolean;
      sSQLCode  : string;
    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      {$REGION 'ENCABEZADO DEL REPORTE'}
      Ren := 2;
      // Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 65;

      Excel.Columns['A:A'].ColumnWidth := 20;
      Excel.Columns['B:B'].ColumnWidth := 82.71;
      Excel.Columns['C:F'].ColumnWidth := 18;
      Excel.Columns['G:G'].ColumnWidth := 15;
      Excel.Columns['H:H'].ColumnWidth := 8.43;

      Hoja.Range['A1:A2'].Select;
      Excel.Selection.RowHeight := '42';

      Hoja.Range['B'+IntTostr(Ren-1)+':B'+IntToStr(Ren-1)].Select;
      Excel.Selection.Value := global_contrato +' '+ tsNumeroOrden.Text;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Size := 27;
      Excel.Selection.Font.Name := 'Tahoma';

      // Colocar los encabezados de la plantilla...
      Hoja.Range['A'+IntTostr(Ren)+':A'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'NO. ACTIVIDAD';
      FormatoEncabezado;
      Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'DESCRIPCION';
      FormatoEncabezado;
      Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'POND. %';
      FormatoEncabezado;
      Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren)].Select;
      Excel.Selection.Value := '% AVANCE PARCIAL';
      FormatoEncabezado;
      Hoja.Range['E'+IntTostr(Ren)+':E'+IntToStr(Ren)].Select;
      Excel.Selection.Value := '% POR EJECUTAR';
      FormatoEncabezado;
      Hoja.Range['F'+IntTostr(Ren)+':F'+IntToStr(Ren)].Select;
      Excel.Selection.Value := '% AVANCE POND.';
      FormatoEncabezado;
      Hoja.Range['G'+IntTostr(Ren)+':G'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Duración Días (PR)';
      FormatoEncabezado;

      Hoja.Range['A'+IntTostr(Ren)+':G'+IntToStr(Ren)].Select;
      Excel.Selection.Font.Size := 14;

      {$ENDREGION}

      {$REGION 'CONSULTAS PROGRAMADO Y REAL DE LA ORDEN'}
      //Consultamos las fechas del convenio modificatorio para impresion de las cantidades reportadas superiores al programa de trabajo.
      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Clear;
      connection.QryBusca2.SQL.Add('select max(dIdFecha) as dFechaFinal from reportediario where sContrato =:Contrato and sNumeroOrden =:Orden ');
      connection.QryBusca2.ParamByName('contrato').AsString := global_contrato;
      connection.QryBusca2.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
      connection.QryBusca2.Open;

      connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select * from actividadesxorden where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden ');
      connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Contrato').Value    := global_contrato;
      connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString;
      connection.QryBusca.Params.ParamByName('Convenio').Value    := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      connection.QryBusca.Params.ParamByName('Orden').DataType    := ftString;
      connection.QryBusca.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text;
      if chkAdmon.Checked then
         connection.QryBusca.SQL.Add(' and (sTipoAnexo = "ADM" or sTipoActividad = "Paquete") ');
      if chkPU.Checked then
         connection.QryBusca.SQL.Add(' and (sTipoAnexo = "PU" or sTipoActividad = "Paquete") ');
      {Filtros..}
      if tsFiltro.Text = 'SOLO REPORTADAS' then
         connection.QryBusca.SQL.Add(' and dInstalado > 0 ');
      if tsFiltro.Text = 'CON RETRASO' then
         connection.QryBusca.SQL.Add(' and dInstalado = 0 ');
      if tsFiltro.Text = 'DESFASADAS' then
         connection.QryBusca.SQL.Add(' and dInstalado = 0 ');
      if tsFiltro.Text = 'TERMINADAS' then
         connection.QryBusca.SQL.Add(' and dInstalado = dCantidad ');
      if connection.configuracion.FieldByName('lOrdenaItem').AsString = 'No' then
      begin
          connection.QryBusca.SQL.Add(' order by mysql.udf_NaturalSortFormat(swbs,:Tam,:Separador)');
          connection.QryBusca.ParamByName('tam').AsInteger      := Global_TamOrden;
          connection.QryBusca.ParamByName('separador').AsString := Global_SepOrden;
      end
      else
         connection.QryBusca.SQL.Add('order by iItemOrden');
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then
      begin
        if chkPeriodo.Checked then
        begin
            MiFecha  := tdIdFecha1.Date;
            MiFechaI := tdIdFecha1.Date;
            MiFechaF := tdIdFecha.Date;
        end
        else
        begin
             MiFecha  := roqOrdenes.FieldByName('dFiProgramado').AsDateTime;
            MiFechaI := roqOrdenes.FieldByName('dFiProgramado').AsDateTime;
            if connection.QryBusca2.FieldValues['dFechaFinal'] > connection.QryBusca.FieldValues['dFechaFinal'] then
               MiFechaF := connection.QryBusca2.FieldByName('dFechaFinal').AsDateTime
            else
               MiFechaF := roqOrdenes.FieldByName('dFfProgramado').AsDateTime;
        end;

        for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
        begin
            Hoja.Cells[Ren, 8 + i].Select;
                 {Formato de las fechas archivo Excel,, 24/07/2011..}
            Excel.Selection.NumberFormat := '@';
            Excel.Selection.Value := DateToStr(MiFecha);
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment := xlCenter;
            Excel.Selection.Font.Color := clNavy;
            Excel.Selection.Font.Bold := False;
            Excel.Selection.Font.Size := 12;
            Excel.Selection.Font.Name := 'Tahoma';
            Excel.Selection.Interior.ColorIndex := 24;
            MiFecha := IncDay(MiFecha);
        end;
        total := i;

        Hoja.Cells[Ren, 8 + i].Select;
        Excel.Selection.Value := 'Fin';
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment := xlCenter;
        Excel.Selection.Font.Color := clWhite;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Font.Size := 12;
        Excel.Selection.Font.Name := 'Tahoma';
        Excel.Selection.Interior.ColorIndex := 3;

        inc(Ren);
        Hoja.Range['H'+IntTostr(Ren)+':H'+IntToStr(Ren)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Font.Size := 12;
        Excel.Selection.Font.Name := 'Tahoma';
        Excel.Selection.Value := 'PROG.';

        Hoja.Range['H'+IntTostr(Ren + 1)+':H'+IntToStr(Ren + 1)].Select;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment := xlCenter;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Font.Size := 12;
        Excel.Selection.Font.Name := 'Tahoma';
        Excel.Selection.Value := 'REAL';
        {$ENDREGION}

        dAvanceGlobal := 0;
        connection.QryBusca.First;
        while not connection.QryBusca.Eof do
        begin
           {Movimiento de la Barra..}
            Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
            TotalProgreso := TotalProgreso + Progreso;
            BarraEstado.Position := Trunc(TotalProgreso);

            Hoja.Range['A'+IntTostr(Ren)+':A'+IntToStr(Ren)].Select;
            Excel.Selection.RowHeight := '18';

            Hoja.Range['A'+IntTostr(Ren + 1)+':A'+IntToStr(Ren + 1)].Select;
            Excel.Selection.RowHeight := '18';

            Hoja.Range['A'+IntTostr(Ren + 2)+':A'+IntToStr(Ren + 2)].Select;
            Excel.Selection.RowHeight := '24';

            Hoja.Range['A'+IntTostr(Ren + 3)+':A'+IntToStr(Ren + 3)].Select;
            Excel.Selection.RowHeight := '18.75';

            sSQLCode := '';

            //if connection.QryBusca.FieldValues['sTipoActividad'] = 'Paquete' then
               lContinua := True;
            //else
            //   lContinua := False;

            if tsPlataforma.KeyValue <> '-1' then
            begin
               if connection.QryBusca.FieldValues['sTipoActividad'] = 'Actividad' then
                  if connection.QryBusca.FieldValues['sIdPlataforma'] = tsPlataforma.KeyValue then
                  begin
                     lContinua := True;
                     sSQLCode  := ' and sIdPlataforma = "'+ tsPlataforma.KeyValue +'" ';
                  end;
            end;

            if connection.QryBusca.FieldValues['iNivel'] = 0 then
            begin
                {$REGION 'AVANCE REAL Y PROGRAMADO DEL PAQUETE PRINCIPAL'}
                Hoja.Range['A'+IntTostr(Ren)+':G'+IntToStr(Ren)].Select;
                Excel.Selection.Interior.ColorIndex := 17;

                MiFecha := MiFechaI;

                if tsPlataforma.KeyValue <> '-1' then
                   sSQLCode  := ' and b.sIdPlataforma = "'+ tsPlataforma.KeyValue +'" ';

                {Consultamos obtenemos los programados de la orden..}
                Q_Partidas.Active := False;
                Q_Partidas.SQL.Clear;
                Q_Partidas.SQL.Add('select * from avancesglobales where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden');
                Q_Partidas.Params.ParamByName('Contrato').DataType := ftString;
                Q_Partidas.Params.ParamByName('Contrato').Value    := global_contrato;
                Q_Partidas.Params.ParamByName('Convenio').DataType := ftString;
                Q_Partidas.Params.ParamByName('Convenio').Value    := zqReprogramacion.FieldByName('sIdConvenio').AsString;
                Q_Partidas.Params.ParamByName('Orden').DataType    := ftString;
                Q_Partidas.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text;
                Q_Partidas.Open;

                if Q_Partidas.RecordCount > 0 then
                begin
                  sColInicio := '';
                  for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
                  begin
                      if MiFecha = Q_Partidas.FieldByName('dIdFecha').AsDateTime then
                      begin
                          if sColInicio = '' then
                             sColInicio := columnas[8 + i];
                          sColFinal := columnas[8 + i];
                          Hoja.Cells[Ren, 8 + i].Select;
                          Excel.Selection.Value        := Q_Partidas.FieldByName('dAvancePonderadoGlobal').AsFloat ;
                          Q_Partidas.Next;
                      end;
                      MiFecha := IncDay(MiFecha);
                  end;
                  Excel.Selection.NumberFormat := '##0.00%';
                  Hoja.Range[sColInicio + IntToStr(Ren) + ':'+sColFinal + IntToStr(Ren)].Select;
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment := xlCenter;
                  Excel.Selection.Font.Bold := False;
                  Excel.Selection.Interior.ColorIndex := 37;
                  Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                  Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
                  Excel.Selection.Borders[xlEdgeLeft].Color        := clGray;
                  Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                  Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
                  Excel.Selection.Borders[xlEdgeTop].Color         := clGray;
                  Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                  Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
                  Excel.Selection.Borders[xlEdgeBottom].Color      := clGray;
                  Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                  Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
                  Excel.Selection.Borders[xlEdgeRight].Color       := clGray;
                end;

                MiFecha := MiFechaI;

                if Q_Partidas.RecordCount > 0 then
                begin
                  dFisico := 0;
                  sColInicio := '';
                  for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
                  begin

                      Connection.qryBusca2.Active := False ;
                      Connection.qryBusca2.SQL.Clear ;
                      Connection.qryBusca2.SQL.Add('Select a.dIdFecha, sum((b.dPonderado / 100)* a.dCantidad) as dAvanceAcumulado '+
                                 ' From actividadesxorden b '+
                                 ' inner JOIN bitacoradeactividades a '+
                                 ' ON (a.sContrato=b.sContrato and b.sIdConvenio = a.sIdConvenio And a.sWbs=b.sWbs and b.sNumeroOrden=a.sNumeroOrden and a.dIdFecha <=:fecha ) '+
                                 ' left JOIN tiposdemovimiento t '+
                                 ' ON (b.sContrato=t.sContrato And a.sIdTipoMovimiento=t.sIdTipoMovimiento And t.sClasificacion="Tiempo en Operacion") '+
                                 ' Where b.sContrato=:Contrato And b.sIdConvenio=:Convenio And b.sNumeroOrden =:Orden '+sSQLCode +
                                 ' Group By a.sContrato ' ) ;
                      Connection.qryBusca2.ParamByName('Contrato').AsString := global_contrato ;
                      Connection.qryBusca2.ParamByName('Convenio').AsString := zqReprogramacion.FieldByName('sIdConvenio').AsString; ;
                      Connection.qryBusca2.ParamByName('Orden').AsString    := tsNumeroOrden.Text ;
                      Connection.qryBusca2.ParamByName('Fecha').AsDate      := MiFecha ;
                      Connection.qryBusca2.Open ;

                      if connection.QryBusca2.RecordCount > 0 then
                      begin
                          if sColInicio = '' then
                             sColInicio := columnas[8 + i];
                          sColFinal := columnas[8 + i];
                          Hoja.Cells[Ren + 1, 8 + i].Select;
                          Excel.Selection.Value        := connection.QryBusca2.FieldByName('dAvanceAcumulado').AsFloat / 100;
                      end;

                      MiFecha := IncDay(MiFecha);
                  end;
                  Hoja.Range[sColInicio + IntToStr(Ren+1) + ':'+sColFinal + IntToStr(Ren+1)].Select;
                  Excel.Selection.NumberFormat := '##0.00%';
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;
                  Excel.Selection.Font.Bold := False;
                  Excel.Selection.Interior.ColorIndex := 44;
                  Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                  Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
                  Excel.Selection.Borders[xlEdgeLeft].Color        := clGray;
                  Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                  Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
                  Excel.Selection.Borders[xlEdgeTop].Color         := clGray;
                  Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                  Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
                  Excel.Selection.Borders[xlEdgeBottom].Color      := clGray;
                  Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                  Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
                  Excel.Selection.Borders[xlEdgeRight].Color       := clGray;
                  Excel.Selection.Borders[xlInsideVertical].LineStyle   := xlContinuous;
                  Excel.Selection.Borders[xlInsideVertical].Weight      := xlThin;
                  Excel.Selection.Borders[xlInsideVertical].Color       := clGray;
                end;
                {$ENDREGION}
            end;

            if lcontinua then
            begin
                {$REGION 'INFORMACION DE LA ACTIVIDAD'}
                Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
                Excel.Selection.Interior.ColorIndex := 37;

                Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren)].Select;
                Excel.Selection.Interior.ColorIndex := 44;

                Hoja.Range['G'+IntTostr(Ren)+':G'+IntToStr(Ren)].Select;
                Excel.Selection.Interior.ColorIndex := 37;

                {Escritura de Datos en el Archvio de Excel..}
                Hoja.Cells[Ren, 1].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['sNumeroActividad'];
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Font.Size := 14;
                Excel.Selection.Font.Bold := True;
                Excel.Selection.Font.Name := 'Arial';

                Hoja.Range['A'+IntTostr(Ren)+':A'+IntToStr(Ren + 3)].Select;
                Excel.Selection.MergeCells := True;

                Hoja.Cells[Ren, 2].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];
                Excel.Selection.HorizontalAlignment := xlJustify;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.MergeCells := True;
                Excel.Selection.WrapText   := True;
                Excel.Selection.Font.Size := 14;
                Excel.Selection.Font.Bold := False;
                Excel.Selection.Font.Name := 'Arial';

                Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren + 3)].Select;
                Excel.Selection.MergeCells := True;

                Hoja.Cells[Ren, 3].Select;
                Excel.Selection.NumberFormat := '##0.0000%';
                Excel.Selection.Value := connection.QryBusca.FieldValues['dPonderado'] / 100;
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Font.Size := 12;
                Excel.Selection.Font.Bold := True;
                Excel.Selection.Font.Name := 'Tahoma';

                Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren + 3)].Select;
                Excel.Selection.MergeCells := True;

                Hoja.Cells[Ren, 4].Select;
                Excel.Selection.NumberFormat := '##0.00%';
                Excel.Selection.Value := 0;
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Font.Size := 14;
                Excel.Selection.Font.Bold := False;
                Excel.Selection.Font.Name := 'Tahoma';

                Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren + 3)].Select;
                Excel.Selection.MergeCells := True;

                Hoja.Cells[Ren, 5].Select;
                Excel.Selection.NumberFormat := '##0.00%';
                Excel.Selection.Value := 0;
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Font.Size := 14;
                Excel.Selection.Font.Bold := False;
                Excel.Selection.Font.Color:= clRed;
                Excel.Selection.Font.Name := 'Tahoma';

                Hoja.Range['E'+IntTostr(Ren)+':E'+IntToStr(Ren + 3)].Select;
                Excel.Selection.MergeCells := True;

                Hoja.Cells[Ren, 6].Select;
                Excel.Selection.NumberFormat := '##0.0000%';
                Excel.Selection.Value := 0;
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Font.Size := 14;
                Excel.Selection.Font.Bold := False;
                Excel.Selection.Font.Name := 'Tahoma';

                Hoja.Range['F'+IntTostr(Ren)+':F'+IntToStr(Ren + 3)].Select;
                Excel.Selection.MergeCells := True;

                Hoja.Cells[Ren, 7].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['dDuracion'];
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Font.Size := 14;
                Excel.Selection.Font.Bold := True;
                Excel.Selection.Font.Name := 'Tahoma';

                Hoja.Range['G'+IntTostr(Ren)+':G'+IntToStr(Ren + 3)].Select;
                Excel.Selection.MergeCells := True;
                {$ENDREGION}
            end;
           {Colores de los paquetes..}

           if connection.QryBusca.FieldValues['sTipoActividad'] = 'Paquete' then
           begin
              {$REGION 'PAQUETES'}
              Hoja.Range['A' + IntToStr(Ren) + ':B' + IntToStr(Ren)].Select;
              Excel.Selection.Font.Bold := True;


              if connection.QryBusca.FieldValues['iNivel'] > 0 then
              begin
                  //Aqui obtenermos el avance acumulado del paquete..
                  Connection.qryBusca2.Active := False ;
                  Connection.qryBusca2.SQL.Clear ;
                  Connection.qryBusca2.SQL.Add('Select a.dPonderado, '+
                             ' if((select ba.lCancelada from bitacoradeactividades ba where a.sContrato = ba.sContrato and ba.sIdConvenio = a.sIdConvenio and a.sNumeroOrden = ba.sNumeroOrden and a.swbs = ba.swbs and lCancelada = "Si" limit 1) ="Si", 100, sum(b.dAvance)) as dAvance, '+
                             '    if(sum(b.dcantidad) > a.dcantidad, a.dPonderado, '+
                             '    if((select ba.lCancelada from bitacoradeactividades ba where a.sContrato = ba.sContrato and a.sIdConvenio = ba.sIdConvenio and a.sNumeroOrden = ba.sNumeroOrden and a.swbs = ba.swbs and lCancelada = "Si" '+
                             '        and ba.didfecha <= :fecha limit 1) ="Si", a.dPonderado, '+
                             '        sum(b.dcantidad * (a.dPonderado / a.dcantidad))))as dAvancePonderado '+
                             ' From actividadesxorden a inner join bitacoradeactividades b on (b.scontrato = a.scontrato and b.sIdConvenio = a.sIdConvenio and b.snumeroorden = a.snumeroorden and b.swbs = a.swbs) '+
                             ' Where a.sContrato = :contrato and a.sIdConvenio =:Convenio and a.sNumeroOrden = :orden And b.sWbs like concat(:wbs, ".%") group by a.swbs') ;
                  Connection.qryBusca2.Params.ParamByName('Contrato').DataType  := ftString ;
                  Connection.qryBusca2.Params.ParamByName('Contrato').Value     := global_contrato ;
                  Connection.qryBusca2.Params.ParamByName('Convenio').DataType  := ftString ;
                  Connection.qryBusca2.Params.ParamByName('Convenio').Value     := zqReprogramacion.FieldByName('sIdConvenio').AsString; ;
                  Connection.qryBusca2.Params.ParamByName('Orden').DataType     := ftString ;
                  Connection.qryBusca2.Params.ParamByName('Orden').Value        := tsNumeroOrden.Text ;
                  Connection.qryBusca2.Params.ParamByName('Fecha').DataType     := ftDate ;
                  Connection.qryBusca2.Params.ParamByName('Fecha').Value        := MiFechaF ;
                  Connection.qryBusca2.Params.ParamByName('Wbs').DataType       := ftString ;
                  Connection.qryBusca2.Params.ParamByName('Wbs').Value          := connection.QryBusca.FieldByName('sWbs').AsString;
                  Connection.qryBusca2.Open;

                  dVolumen := 0;
                  while Not Connection.QryBusca2.Eof do
                  begin
                      dVolumen := dVolumen + Connection.QryBusca2.FieldByName('dAvancePonderado').AsFloat;
                      Connection.QryBusca2.Next;
                  end;

                  //Avance de la partida..
                  Hoja.Cells[Ren, 4].Select;
                  Excel.Selection.NumberFormat := '##0.00%';
                  if connection.QryBusca.FieldValues['dPonderado'] > 0 then
                     Excel.Selection.Value := ((100 / connection.QryBusca.FieldValues['dPonderado']) * dVolumen)/100
                  else
                     Excel.Selection.Value := 0;
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;

                  //Avance por ejecutar
                  Hoja.Cells[Ren, 5].Select;
                  Excel.Selection.NumberFormat := '##0.00%';
                  If connection.QryBusca.FieldValues['dPonderado'] > 0 then
                     Excel.Selection.Value := (100 - ((100 / connection.QryBusca.FieldValues['dPonderado']) * dVolumen))/100
                  else
                     Excel.Selection.Value := (connection.QryBusca.FieldValues['dPonderado'])/100;
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;
                  Excel.Selection.Font.Bold := False;

                  //Avance Ponderado por partida..
                  Hoja.Cells[Ren, 6].Select;
                  Excel.Selection.NumberFormat := '##0.0000%';
                  Excel.Selection.Value := dVolumen / 100;
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;
                  Excel.Selection.Font.Bold := False;
                  if dVolumen > 0 then
                     Excel.Selection.Font.Color := clBlue;
              end;
              {$ENDREGION}
           end
           else
           begin
               if lContinua then
               begin
                  {$REGION 'PROGRAMADOS ACTIVIDADES FOLIO/ORDEN'}
                  MiFecha := MiFechaI;

                  {Consultamos obtenemos los programados de la orden..}
                  Q_Partidas.Active := False;
                  Q_Partidas.SQL.Clear;
                  Q_Partidas.SQL.Add('select * from distribuciondeactividades where sContrato =:Contrato and sIdConvenio =:Convenio and sNumeroOrden =:Orden and sWbs =:Wbs and sNumeroActividad =:Actividad ');
                  Q_Partidas.Params.ParamByName('Contrato').DataType := ftString;
                  Q_Partidas.Params.ParamByName('Contrato').Value    := global_contrato;
                  Q_Partidas.Params.ParamByName('Convenio').DataType := ftString;
                  Q_Partidas.Params.ParamByName('Convenio').Value    := zqReprogramacion.FieldByName('sIdConvenio').AsString;
                  Q_Partidas.Params.ParamByName('Orden').DataType    := ftString;
                  Q_Partidas.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text;
                  Q_Partidas.Params.ParamByName('Wbs').DataType      := ftString;
                  Q_Partidas.Params.ParamByName('Wbs').Value         := connection.QryBusca.FieldByName('sWbs').AsString;
                  Q_Partidas.Params.ParamByName('Actividad').DataType:= ftString;
                  Q_Partidas.Params.ParamByName('Actividad').Value   := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                  Q_Partidas.Open;

                  if Q_Partidas.RecordCount > 0 then
                  begin
                    dProgramado := 0;
                    sColInicio  := '';
                    for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
                    begin
                        if MiFecha = Q_Partidas.FieldByName('dIdFecha').AsDateTime then
                        begin
                            if sColInicio = '' then
                               sColInicio := columnas[8 + i];
                            sColFinal := columnas[8 + i];
                            Hoja.Cells[Ren, 8 + i].Select;
                            dProgramado := dProgramado + Q_Partidas.FieldByName('dCantidad').AsFloat;
                            if (connection.QryBusca.FieldValues['sMedida'] = 'ACTIVIDAD') and (connection.QryBusca.FieldValues['dCantidad'] >= 1) then
                                Excel.Selection.Value   := dProgramado / 100
                            else
                               Excel.Selection.Value   := dProgramado;
                            Q_Partidas.Next;
                        end;
                        MiFecha := IncDay(MiFecha);
                    end;
                    if sColInicio <> '' then
                    begin
                        Hoja.Range[sColInicio + IntToStr(Ren) + ':'+sColFinal + IntToStr(Ren)].Select;
                        Excel.Selection.NumberFormat := '##0.00%';
                        Excel.Selection.HorizontalAlignment := xlCenter;
                        Excel.Selection.VerticalAlignment := xlCenter;
                        Excel.Selection.Font.Bold := False;
                        Excel.Selection.Interior.ColorIndex := 37;
                        Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                        Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
                        Excel.Selection.Borders[xlEdgeLeft].Color        := clGray;
                        Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                        Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
                        Excel.Selection.Borders[xlEdgeTop].Color         := clGray;
                        Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                        Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
                        Excel.Selection.Borders[xlEdgeBottom].Color      := clGray;
                        Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                        Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
                        Excel.Selection.Borders[xlEdgeRight].Color       := clGray;
                    end;
                  end;
                 {$ENDREGION}

                  {$REGION 'AVANCES FISICOS ACTIVIDADES'}
                  {Consultamos si la partida esta reportada..}
                  Q_Partidas.Active := False;
                  Q_Partidas.SQL.Clear;
                  Q_Partidas.SQL.Add('Select b.sWbs,b.sNumeroActividad, sum(a.dCantidad) as dCantidad, a.dIdFecha, b.dCantidad as dVolumen ' +
                    'From actividadesxorden b ' +
                    'left JOIN bitacoradeactividades a ' +
                    'ON (a.sContrato=b.sContrato and a.sIdconvenio = b.sIdConvenio And a.sWbs=b.sWbs And a.dIdFecha <=:Final and b.sNumeroOrden=a.sNumeroOrden) ' +
                    'left JOIN tiposdemovimiento t ' +
                    'ON (b.sContrato=t.sContrato And a.sIdTipoMovimiento=t.sIdTipoMovimiento And t.sClasificacion="Tiempo en Operacion") ' +
                    'Where b.sContrato=:Contrato And b.sIdConvenio=:Convenio And b.sNumeroOrden =:Orden and a.sWbs =:Wbs ' +
                    'Group By b.sWbs,a.dIdFecha Order By b.sNumeroActividad,b.iItemOrden,a.dIdFecha');
                  Q_Partidas.Params.ParamByName('Contrato').DataType := ftString;
                  Q_Partidas.Params.ParamByName('Contrato').Value := global_contrato;
                  Q_Partidas.Params.ParamByName('Convenio').DataType := ftString;
                  Q_Partidas.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString;
                  Q_Partidas.Params.ParamByName('Orden').DataType := ftString;
                  Q_Partidas.Params.ParamByName('Orden').Value := tsNumeroOrden.Text;
                  Q_Partidas.Params.ParamByName('Final').DataType := ftDate;
                  Q_Partidas.Params.ParamByName('Final').Value := MiFechaF;
                  Q_Partidas.Params.ParamByName('Wbs').DataType := ftString;
                  Q_Partidas.Params.ParamByName('Wbs').Value := connection.QryBusca.FieldByName('sWbs').AsString;
                  Q_Partidas.Open;

                  MiFecha := MiFechaI;
                  if Q_Partidas.RecordCount > 0 then
                  begin
                    dVolumen := 0;
                    sColInicio  := '';
                    for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
                    begin
                        if MiFecha = Q_Partidas.FieldByName('dIdFecha').AsDateTime then
                        begin
                            if sColInicio = '' then
                               sColInicio := columnas[8 + i];
                            sColFinal := columnas[8 + i];

                            Hoja.Cells[Ren + 1, 8 + i].Select;
                            if (connection.QryBusca.FieldValues['sMedida'] = 'ACTIVIDAD') and (connection.QryBusca.FieldValues['dCantidad'] >= 1) then
                               Excel.Selection.Value := Q_Partidas.FieldByName('dCantidad').AsFloat / 100
                            else
                               Excel.Selection.Value := Q_Partidas.FieldByName('dCantidad').AsFloat;

                            Hoja.Cells[Ren + 2, 8 + i].Select;
                            if (connection.QryBusca.FieldValues['sMedida'] = 'ACTIVIDAD') and (connection.QryBusca.FieldValues['dCantidad'] >= 1) then
                               dVolumen := dVolumen + Q_Partidas.FieldByName('dCantidad').AsFloat / 100
                            else
                               dVolumen := dVolumen + Q_Partidas.FieldByName('dCantidad').AsFloat;
                            Excel.Selection.Value        := dVolumen;

                            Q_Partidas.Next;
                        end
                        else
                        begin
                            {$REGION 'FORMATO AVANCES DIARIOS'}
                            if sColInicio <> '' then
                            begin
                                Hoja.Range[sColInicio + IntToStr(Ren+1) + ':'+sColFinal + IntToStr(Ren+1)].Select;
                                Excel.Selection.NumberFormat := '##0.00%';
                                Excel.Selection.HorizontalAlignment := xlCenter;
                                Excel.Selection.VerticalAlignment   := xlCenter;
                                Excel.Selection.Font.Bold := False;
                                Excel.Selection.Interior.ColorIndex := 35;
                                Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                                Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
                                Excel.Selection.Borders[xlEdgeLeft].Color        := clGray;
                                Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                                Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
                                Excel.Selection.Borders[xlEdgeTop].Color         := clGray;
                                Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                                Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
                                Excel.Selection.Borders[xlEdgeBottom].Color      := clGray;
                                Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                                Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
                                Excel.Selection.Borders[xlEdgeRight].Color       := clGray;
                                Excel.Selection.Borders[xlInsideVertical].LineStyle   := xlContinuous;
                                Excel.Selection.Borders[xlInsideVertical].Weight      := xlThin;
                                Excel.Selection.Borders[xlInsideVertical].Color       := clGray;
                                Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous;
                                Excel.Selection.Borders[xlInsideHorizontal].Weight    := xlThin;
                                Excel.Selection.Borders[xlInsideHorizontal].Color       := clGray;

                                Hoja.Range[sColInicio + IntToStr(Ren+2) + ':'+sColFinal + IntToStr(Ren+2)].Select;
                                Excel.Selection.NumberFormat := '##0.00%';
                                Excel.Selection.HorizontalAlignment := xlCenter;
                                Excel.Selection.VerticalAlignment   := xlCenter;
                                Excel.Selection.Font.Bold := False;
                                Excel.Selection.Interior.ColorIndex := 44;
                                Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                                Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
                                Excel.Selection.Borders[xlEdgeLeft].Color        := clGray;
                                Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                                Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
                                Excel.Selection.Borders[xlEdgeTop].Color         := clGray;
                                Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                                Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
                                Excel.Selection.Borders[xlEdgeBottom].Color      := clGray;
                                Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                                Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
                                Excel.Selection.Borders[xlEdgeRight].Color       := clGray;
                                Excel.Selection.Borders[xlInsideVertical].LineStyle   := xlContinuous;
                                Excel.Selection.Borders[xlInsideVertical].Weight      := xlThin;
                                Excel.Selection.Borders[xlInsideVertical].Color       := clGray;
                                Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous;
                                Excel.Selection.Borders[xlInsideHorizontal].Weight    := xlThin;
                                Excel.Selection.Borders[xlInsideHorizontal].Color       := clGray;
                            end;
                            {$ENDREGION}
                            sColInicio  := '';
                        end;

                        MiFecha := IncDay(MiFecha);
                    end;

                    //Avance de la partida..
                    Hoja.Cells[Ren, 4].Select;
                    Excel.Selection.NumberFormat := '##0.00%';
                    Excel.Selection.Value := dVolumen;
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Bold  := False;
                    Excel.Selection.Font.size  := 14;
                    if dVolumen = 100 then
                       Excel.Selection.Font.Color := clBlue;

                    //Avance por ejecutar
                    Hoja.Cells[Ren, 5].Select;
                    Excel.Selection.NumberFormat := '##0.00%';
                    if (dVolumen * 100) < 100 then
                       Excel.Selection.Value := (100 - dVolumen)/100
                    else
                       Excel.Selection.Value := 0;
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.size  := 14;

                    //Avance Ponderado por partida..
                    Hoja.Cells[Ren, 6].Select;
                    Excel.Selection.NumberFormat := '##0.0000%';
                    Excel.Selection.Value := ((connection.QryBusca.FieldValues['dPonderado'] / 100) * dVolumen);
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Bold  := False;
                    Excel.Selection.Font.size  := 14;
                    Excel.Selection.Font.color := clBlue;

                    dAvanceGlobal := dAvanceGlobal + (connection.QryBusca.FieldValues['dPonderado'] / 100) * dVolumen;
                  end;
                  {$ENDREGION}
               end;
          end;

          connection.QryBusca.Next;
          if lContinua then
             Inc(Ren,4);
        end;
      end;

      Hoja.Cells[3, 6].Select;
      Excel.Selection.NumberFormat := '##0.0000%';
      if dAvanceGlobal > 100 then
         Excel.Selection.Value := 100
      else
         Excel.Selection.Value := xRound(dAvanceGlobal,4);
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.size := 12;

      Hoja.Cells[3, 5].Select;
      Excel.Selection.NumberFormat := '##0.00%';
      if dAvanceGlobal > 100 then
         Excel.Selection.Value := 100
      else
         Excel.Selection.Value := (100 - (dAvanceGlobal*100))/100;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.size := 12;

      Hoja.Range['A'+IntTostr(1)+':G'+IntToStr(1)].Select;
      Excel.Selection.Interior.ColorIndex := 15;

      Hoja.Cells[3, 4].Select;
      Excel.Selection.NumberFormat := '##0.00%';
      if dAvanceGlobal > 100 then
         Excel.Selection.Value := 100
      else
         Excel.Selection.Value := dAvanceGlobal;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.size := 12;
    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'AVANCES ' + tsNumeroOrden.Text;
      except
        Hoja.Name := 'AVANCES ';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Programa de Trabajo:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
      end;
    end;

    Result := Resultado;
  end;

begin
    //Verificamos si es un frente
    If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
    Begin
        messageDLG('Seleccione un frente de trabajo!', mtInformation, [mbOk], 0);
        exit;
    End;

    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
    if not SaveDialog1.Execute then
      Exit;
      // Generar el ambiente de excel
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
    end
    else
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;
    end;

    PanelProgress.Visible := True;
    Label15.Refresh;
    BarraEstado.Position := 0;

    Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

      // Verificar si cuenta con las hojas necesarias
    while Libro.Sheets.Count < 2 do
      Libro.Sheets.Add;

      // Verificar si se pasa de hojas necesarias
    Libro.Sheets[1].Select;
    while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

      // Proceder a generar la hoja REPORTE
    CadError := '';

    if GenerarPlantilla then
    begin
          // Grabar el archivo de excel con el nombre dado
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      PanelProgress.Visible := False;
      messageDlg('El Archivo se generó Correctamente!', mtInformation, [mbOk], 0);
    end;

    Excel := '';

    if CadError <> '' then
      showmessage(CadError);

end;

//soad -> Reporte para verificar las paridas Excedidas ...
//******************************************************************************
procedure TfrmCompara.cmdExcedentesClick(Sender: TObject);
var
   lContinua : boolean;
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
  try
    lContinua := False;
    Connection.qryBusca2.Active := False ;
    Connection.qryBusca2.SQL.Clear ;
    Connection.qryBusca2.SQL.Add('select sTipoObra from contratos where sContrato =:Contrato') ;
    Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
    Connection.qryBusca2.Open ;

    if (Connection.qryBusca2.FieldValues['sTipoObra']= 'PROGRAMADA') or (Connection.qryBusca2.FieldValues['sTipoObra']= 'MIXTA') or (Connection.qryBusca2.FieldValues['sTipoObra']= 'OPTATIVA' ) then
    begin
      If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
      Begin
          qryBuscaP.Active := False ;
          qryBuscaP.SQL.Clear ;
          qryBuscaP.SQL.Add('select a.sWbs, b.dIdFecha, b.sNumeroOrden, b.dCantidad, b.dAvance, a.sNumeroActividad, b.dAvanceAnterior, (b.dAvanceActual + b.dAvanceAnterior) as Avance, a.dExcedente, '+
                                'a.mDescripcion, a.dCantidadAnexo, a.dInstalado, a.sMedida, b.dCantidadActual, b.dCantidadAnterior from actividadesxanexo a '+
                                'inner join bitacoradeactividades b '+
                                'on(b.sContrato = a.sContrato and b.dIdFecha >=:FechaI and b.dIdFecha<=:FechaF and b.sNumeroActividad = a.sNumeroActividad '+
                                'and b.dCantidad >0 and b.sNumeroActividad <> "" ) '+
                                'Where a.sContrato =:Contrato and a.sIdConvenio =:Convenio and a.sTipoActividad = "Actividad" and a.dExcedente > 0 '+
                                'order by b.sNumeroActividad, a.iItemOrden');
              qryBuscaP.Params.ParamByName('Contrato').DataType := ftString ;
              qryBuscaP.Params.ParamByName('Contrato').Value := global_contrato ;
              qryBuscaP.Params.ParamByName('Convenio').DataType := ftString ;
              qryBuscaP.Params.ParamByName('Convenio').Value := global_convenio ;
              qryBuscaP.Params.ParamByName('FechaI').DataType := ftDate ;
              qryBuscaP.Params.ParamByName('FechaI').Value := tdIdFecha1.Date ;
              qryBuscaP.Params.ParamByName('FechaF').DataType := ftDate ;
              qryBuscaP.Params.ParamByName('FechaF').Value := tdIdFecha.Date ;
              qryBuscaP.Open ;
              lContinua := True;
      end
      else
      Begin
        qryBuscaP.Active := False ;
        qryBuscaP.SQL.Clear ;
        qryBuscaP.SQL.Add('select a.sWbs, b.dIdFecha, b.sNumeroOrden, b.dCantidad, b.dAvance, a.sNumeroActividad, b.dAvanceAnterior, (b.dAvanceActual + b.dAvanceAnterior) as Avance , a.dExcedente, '+
                              'a.mDescripcion, a.dCantidadAnexo, a.dInstalado, a.sMedida, b.dCantidadActual, b.dCantidadAnterior from actividadesxanexo a '+
                              'inner join bitacoradeactividades b '+
                              'on(b.sContrato = a.sContrato and b.sIdConvenio = a.sIdConvenio and b.dIdFecha>=:FechaI and b.dIdFecha<=:FechaF and b.sNumeroOrden =:Orden and b.sNumeroActividad = a.sNumeroActividad '+
                              'and b.dCantidad >0 and b.sNumeroActividad <> "" ) '+
                              'Where a.sContrato =:Contrato and a.sIdConvenio =:Convenio and a.sTipoActividad = "Actividad" and a.sIdFase = "" and a.dExcedente > 0 '+
                              'order by b.sNumeroActividad, a.iItemOrden');
            qryBuscaP.Params.ParamByName('Contrato').DataType := ftString ;
            qryBuscaP.Params.ParamByName('Contrato').Value := global_contrato ;
            qryBuscaP.Params.ParamByName('Convenio').DataType := ftString ;
            qryBuscaP.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString ;
            qryBuscaP.Params.ParamByName('FechaI').DataType := ftDate ;
            qryBuscaP.Params.ParamByName('FechaI').Value := tdIdFecha1.Date ;
            qryBuscaP.Params.ParamByName('FechaF').DataType := ftDate ;
            qryBuscaP.Params.ParamByName('FechaF').Value := tdIdFecha.Date ;
            qryBuscaP.Params.ParamByName('Orden').DataType := ftString ;
            qryBuscaP.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
            qryBuscaP.Open ;
            lContinua := True;
      end
    end;

    if Connection.qryBusca2.FieldValues['sTipoObra']= 'OPTATIVA' then
    begin
        If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
        Begin
            MessageDlg( 'No se pude Filtrar por esta Opcion, Seleccione Frente de Trabajo', mtWarning, [ mbOk ], 0 );
        end
        else
        Begin
          qryBuscaP.Active := False ;
          qryBuscaP.SQL.Clear ;
          qryBuscaP.SQL.Add('select a.sWbs, b.dIdFecha, b.sNumeroOrden, b.dCantidad, b.dAvance, a.sNumeroActividad, b.dAvanceAnterior, (b.dAvanceActual + b.dAvanceAnterior) as Avance , a.dExcedente, '+
                              'a.mDescripcion, a.dCantidad as dCantidadAnexo, a.dInstalado, a.sMedida, b.dCantidadActual, b.dCantidadAnterior from actividadesxorden a '+
                              'inner join bitacoradeactividades b '+
                              'on(b.sContrato = a.sContrato and a.sIdConvenio = b.sIdConvenio and b.dIdFecha>=:FechaI and b.dIdFecha<=:FechaF and b.sNumeroOrden =:Orden and b.sNumeroActividad = a.sNumeroActividad '+
                              'and b.dCantidad >0 and b.sNumeroActividad <> "" ) '+
                              'Where a.sContrato =:Contrato and a.sIdConvenio =:Convenio and a.sTipoActividad = "Actividad" and '+
                              '(a.sMedida = "ACTIV" or a.sMedida = "ACTIVID" or a.sMedida = "ACTIVIDAD") '+
                              'order by a.sNumeroActividad, a.iItemOrden');
              qryBuscaP.Params.ParamByName('Contrato').DataType := ftString ;
              qryBuscaP.Params.ParamByName('Contrato').Value := global_contrato ;
              qryBuscaP.Params.ParamByName('Convenio').DataType := ftString ;
              qryBuscaP.Params.ParamByName('Convenio').Value := zqReprogramacion.FieldByName('sIdConvenio').AsString ;
              qryBuscaP.Params.ParamByName('FechaI').DataType := ftDate ;
              qryBuscaP.Params.ParamByName('FechaI').Value := tdIdFecha1.Date ;
              qryBuscaP.Params.ParamByName('FechaF').DataType := ftDate ;
              qryBuscaP.Params.ParamByName('FechaF').Value := tdIdFecha.Date ;
              qryBuscaP.Params.ParamByName('Orden').DataType := ftString ;
              qryBuscaP.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
              qryBuscaP.Open ;
              lContinua := True;
        end
    end;

    if lContinua then
    begin
        rInforme.LoadFromFile (global_files + 'Partidas_excedentes.fr3') ;
        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
        if not FileExists(global_files + 'Partidas_excedentes.fr3') then
           showmessage('El archivo de reporte Partidas_excedentes.fr3 no existe, notifique al administrador del sistema');
    end;
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al generando partidas excedidas', 0);
    end;
  end;
end;


procedure TfrmCompara.cmdHistoricoClick(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// SEGUIMINTO DE AVANCES X PARTIDA DIAVAZ OCTUBRE 2012 //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena: string;
      fs: tStream;
      Alto: Extended;
      MiFechaI, MiFechaF, MiFecha: tDate;
      Ren, nivel, i, total: integer;
      iInicioMes, iFinMes : integer;
      Q_Partidas: TZReadOnlyQuery;
      dVolumen, dAvanceGlobal, dProgramado, dFisico: double;
      Progreso, TotalProgreso: real;
      lEncuentra : boolean;
      sColInicio, sColFinal : string;

    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      {$REGION 'ENCABEZADO'}
      Ren := 2;
      //Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 115;

      Excel.Columns['A:A'].ColumnWidth := 7.86;
      Excel.Columns['B:B'].ColumnWidth := 3;
      Excel.Columns['C:C'].ColumnWidth := 10.71;
      Excel.Columns['D:D'].ColumnWidth := 19.71;
      Excel.Columns['E:E'].ColumnWidth := 24.57;
      Excel.Columns['F:F'].ColumnWidth := 39.57;
      Excel.Columns['G:G'].ColumnWidth := 10.71;
      Excel.Columns['H:H'].ColumnWidth := 35.57;
      Excel.Columns['I:I'].ColumnWidth := 3.57;


      Hoja.Range['A1:A1'].Select;
      Excel.Selection.RowHeight := '20.25';

      Hoja.Range['C1:C1'].Select;
      Excel.Selection.Value := 'Histórico de Actividades';
      Excel.Selection.HorizontalAlignment := xlLeft;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := False;
      Excel.Selection.Font.Size := 16;
      Excel.Selection.Font.Color:= RGB(68,114,196);;
      Excel.Selection.Font.Name := 'Calibri Light';
      Hoja.Range['C1:E1'].Select;
      Excel.Selection.MergeCells := True;

      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Clear;
      connection.QryBusca2.SQL.Add('Select p.sDescripcion, o.dFiProgramado, o.dFfProgramado from ordenesdetrabajo o '+
                                  'inner join plataformas p on (o.sIdPlataforma = o.sIdPlataforma) '+
                                  'where o.sContrato = :contrato and o.sNumeroOrden =:Orden');
      connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato;
      connection.QryBusca2.ParamByName('Orden').AsString    := tsNumeroOrden.Text;
      connection.QryBusca2.Open;

      Ren := 3;
      // Colocar los encabezados de la plantilla...
      Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Actividad';
      FormatoEncabezado;
      Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Inicio';
      FormatoEncabezado;
      Hoja.Range['E'+IntTostr(Ren)+':E'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Fin';
      FormatoEncabezado;
      Hoja.Range['F'+IntTostr(Ren)+':F'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Descripcion';
      FormatoEncabezado;
      Hoja.Range['G'+IntTostr(Ren)+':G'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Es Espera';
      FormatoEncabezado;
      Hoja.Range['H'+IntTostr(Ren)+':H'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'Afectacion';
      FormatoEncabezado;

      Hoja.Range['A'+IntTostr(Ren)+':I'+IntToStr(Ren)].Select;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Interior.ColorIndex := 2;
      Hoja.Range['C'+IntTostr(Ren)+':H'+IntToStr(Ren)].Select;
      Excel.Selection.Font.Bold  := True;
      Excel.Selection.Font.Size  := 11;
      Excel.Selection.Font.Name  := 'Calibri';
      Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
      Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;

      {$ENDREGION}

      {$REGION 'DURACION Y TOTAL'}
      connection.QryBusca.Active := False;
      Connection.QryBusca.Filtered := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select a.dIdFecha, ao.sNumeroActividad, ao.dFechaInicio, ao.dFechaFinal, a.sHoraInicio, a.sHoraFinal, a.sIdClasificacion, t.sDescripcion,  a.mDescripcion '+
                                  'from bitacoradeactividades a '+
                                  'inner join actividadesxorden ao on (a.sContrato = ao.sContrato and ao.sIdConvenio = a.sIdConvenio and ao.sNumeroOrden = :Orden '+
                                  'and ao.sWbs = a.sWbs and sTipoActividad = "Actividad" and sTipoAnexo = "ADM") '+
                                  'inner join tiposdemovimiento t on (t.sContrato =:ContratoBarco and t.sIdTipoMovimiento = a.sIdClasificacion) '+
                                  'where a.sContrato = :Contrato and a.sIdConvenio =:Convenio and a.dIdFecha <=:Fecha order by ao.iITemOrden, a.dIdFecha, a.sIdTipoMovimiento, sHoraInicio');
      connection.QryBusca.Params.ParamByName('ContratoBarco').DataType := ftString;
      connection.QryBusca.Params.ParamByName('ContratoBarco').Value    := global_contrato_barco;
      connection.QryBusca.Params.ParamByName('Contrato').DataType      := ftString;
      connection.QryBusca.Params.ParamByName('Contrato').Value         := global_contrato;
      connection.QryBusca.Params.ParamByName('Convenio').DataType      := ftString;
      connection.QryBusca.Params.ParamByName('Convenio').Value         := zqReprogramacion.FieldByName('sIdConvenio').AsString;
      connection.QryBusca.Params.ParamByName('Orden').DataType         := ftString;
      connection.QryBusca.Params.ParamByName('Orden').Value            := tsNumeroOrden.Text;
      connection.QryBusca.Params.ParamByName('Fecha').DataType         := ftDate;
      if chkPeriodo.Checked then
         connection.QryBusca.Params.ParamByName('Fecha').Value         := tdIdFecha.Date
      else
         connection.QryBusca.Params.ParamByName('Fecha').Value         := connection.QryBusca2.FieldByName('dFfProgramado').AsDateTime;
      connection.QryBusca.Open;
      {$ENDREGION}

      inc(Ren);
      dAvanceGlobal := 0;
      connection.QryBusca.First;
      while not connection.QryBusca.Eof do
      begin
          {$REGION 'PARTIDAS DE ANEXO'}
          {Movimiento de la Barra..}
          Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
          TotalProgreso := TotalProgreso + Progreso;
          BarraEstado.Position := Trunc(TotalProgreso);

          {Escritura de Datos en el Archvio de Excel..}
          Hoja.Cells[Ren, 3].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sNumeroActividad'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;

          Hoja.Cells[Ren, 4].Select;
          Excel.Selection.NumberFormat := '@';
          Excel.Selection.Value := DateToStr(connection.QryBusca.FieldByName('dIdFecha').AsDateTime) + ' ' +connection.QryBusca.FieldValues['sHoraInicio'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;

          Hoja.Cells[Ren, 5].Select;
           Excel.Selection.NumberFormat := '@';
          Excel.Selection.Value := DateToStr(connection.QryBusca.FieldByName('dIdFecha').AsDateTime) + ' ' +connection.QryBusca.FieldValues['sHoraFinal'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;

          Hoja.Cells[Ren, 6].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment   := xlCenter;

          Hoja.Cells[Ren, 7].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sIdClasificacion'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;

          Hoja.Cells[Ren, 8].Select;
          Excel.Selection.Value := connection.QryBusca.FieldValues['sDescripcion'];
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;

          {$ENDREGION}
          connection.QryBusca.Next;
          Inc(Ren);
      end;

      Hoja.Range['A'+IntToStr(1)+':J'+IntToStr(Ren+5)].Select;
      Excel.Selection.Interior.Color := clWhite;

      Hoja.Range['A'+IntToStr(2)+':I'+IntToStr(Ren)].Select;
      Excel.Selection.RowHeight := '15';

      i := 4;
      while i<=(Ren-1) do
      begin
          Hoja.Range['C'+IntToStr(i)+':'+columnas[total+8]+IntToStr(i)].Select;
          Excel.Selection.Interior.Color := RGB(238,236,225);
          inc(i,2);
      end;

      Hoja.Range['C'+IntToStr(4)+':I'+IntToStr(Ren)].Select;;
      Excel.Selection.Font.Size  := 10;
      Excel.Selection.Font.Bold  := False;
      Excel.Selection.Font.Color := RGB(38,38,38);
      Excel.Selection.Font.Name  := 'Calibri';

     {$ENDREGION}


    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'PRODUCCION ' + tsNumeroOrden.Text;
      except
        Hoja.Name := 'PRODUCCION ';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Programa de Trabajo:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
      end;
    end;

    Result := Resultado;
  end;

begin
    //Verificamos si es un frente
    If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
    Begin
        messageDLG('Seleccione un frente de trabajo!', mtInformation, [mbOk], 0);
        exit;
    End;

    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
    if not SaveDialog1.Execute then
      Exit;
      // Generar el ambiente de excel
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
    end
    else
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;
    end;

    PanelProgress.Visible := True;
    Label15.Refresh;

    BarraEstado.Position := 0;

    Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

      // Verificar si cuenta con las hojas necesarias
    while Libro.Sheets.Count < 2 do
      Libro.Sheets.Add;

      // Verificar si se pasa de hojas necesarias
    Libro.Sheets[1].Select;
    while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

      // Proceder a generar la hoja REPORTE
    CadError := '';

    if GenerarPlantilla then
    begin
          // Grabar el archivo de excel con el nombre dado
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      PanelProgress.Visible := False;
      messageDlg('El Archivo se generó Correctamente!', mtInformation, [mbOk], 0);
    end;

    Excel := '';

    if CadError <> '' then
      showmessage(CadError);

end;

procedure TfrmCompara.btnSinGenerarClick(Sender: TObject);
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
    sOpcion := 'Sin Generar' ;
    tdIdFecha.SetFocus
end;

procedure TfrmCompara.btnSinReportarClick(Sender: TObject);
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
   sOpcion := 'Reportadas' ;
   tdIdFecha.SetFocus
end;

procedure TfrmCompara.TabSheet1Show(Sender: TObject);
begin
    tdIdFecha.SetFocus
end;

procedure TfrmCompara.TabSheet2Show(Sender: TObject);
begin
    sConvenio := global_convenio ;
end;

procedure TfrmCompara.TabSheet4Show(Sender: TObject);
begin
    sConvenio := global_convenio ;
end;

procedure TfrmCompara.btnSuministrosClick(Sender: TObject);
Var
    dGenerado, dSuministrado : Double ;
begin
   if ChkPu.Checked = True Then
      cadpua := 'And sTipoAnexo =  "PU" '
   else
      cadpua := 'And sTipoAnexo = "ADM" ' ;
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
  try
    If rxSuministroAnexo.RecordCount > 0 Then
        rxSuministroAnexo.EmptyTable ;

    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select sNumeroActividad, mDescripcion, sMedida, dCantidadAnexo, dPonderado, dVentaMN From actividadesxanexo where sContrato = :Contrato ' +
                                'And sIdConvenio = :Convenio and sTipoActividad = "Actividad" ' + cadpua + ' Order By iItemOrden') ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := sConvenio ;
    Connection.qryBusca.Open ;

    dsInforme.FieldAliases.Clear ;
    dsInforme.DataSet := rxSuministroAnexo ;

    if connection.QryBusca.RecordCount>0 then
    begin
      Progress.Visible := True ;
      Progress.MinValue := 1 ;
      Progress.MaxValue := connection.QryBusca.RecordCount ;
    end;
    While NOT Connection.qryBusca.Eof Do
    Begin
        dGenerado := 0 ;
        dSuministrado := 0 ;
        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as dSuministrado From anexo_psuministro e INNER JOIN anexo_suministro e2 ON ' +
                                     '(e.sContrato = e2.sContrato And e.iFolio = e2.iFolio And e2.dFechaAviso <= :Fecha) ' +
                                     'Where e.sContrato = :contrato And e.sNumeroActividad = :Actividad Group By e.sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value := Connection.qryBusca.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
            dSuministrado := Connection.qryBusca2.FieldValues ['dSuministrado'] ;

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as dGenerado From estimacionxpartida e INNER JOIN estimaciones e2 ON ' +
                                    '(e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And e.sNumeroGenerador = e2.sNumeroGenerador And e2.dFechaFinal <= :Fecha) ' +
                                    'Where e.sContrato = :contrato And e.sNumeroActividad = :Actividad And e.lEstima = "Si" Group By e.sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value := Connection.qryBusca.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
             dGenerado := Connection.qryBusca2.FieldValues ['dGenerado'] ;

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select e2.sReferencia, e.dCantidad From anexo_psuministro e INNER JOIN anexo_suministro e2 ON ' +
                                     '(e.sContrato = e2.sContrato And e.iFolio = e2.iFolio And e2.dFechaAviso <= :Fecha) ' +
                                      'Where e.sContrato = :contrato And e.sNumeroActividad = :Actividad Order By e2.dFechaAviso') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value := Connection.qryBusca.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
             While NOT Connection.qryBusca2.Eof Do
             Begin
                  rxSuministroAnexo.Append ;
                  rxSuministroAnexo.FieldValues['sNumeroActividad'] := Connection.qryBusca.FieldValues['sNumeroActividad'] ;
                  rxSuministroAnexo.FieldValues['mDescripcion'] := MidStr(Connection.qryBusca.FieldValues['mDescripcion'],1,100) + ' ..' ;
                  rxSuministroAnexo.FieldValues['sMedida'] := Connection.qryBusca.FieldValues['sMedida'] ;
                  rxSuministroAnexo.FieldValues['dCantidadAnexo'] := Connection.qryBusca.FieldValues['dCantidadAnexo'] ;
                  rxSuministroAnexo.FieldValues['dPonderado'] := Connection.qryBusca.FieldValues['dPonderado'] ;
                  rxSuministroAnexo.FieldValues['dVentaMN'] := Connection.qryBusca.FieldValues['dVentaMN'] ;
                  rxSuministroAnexo.FieldValues['sReferencia'] := Connection.qryBusca2.FieldValues['sReferencia'] ;
                  rxSuministroAnexo.FieldValues['dCantidad'] := Connection.qryBusca2.FieldValues['dCantidad'] ;
                  rxSuministroAnexo.FieldValues['dGenerado'] := dGenerado ;
                  If dGenerado > dSuministrado Then
                  Begin
                      rxSuministroAnexo.FieldValues['dPGenerar'] := 0 ;
                      rxSuministroAnexo.FieldValues['dPReportar'] := dGenerado - dSuministrado ;
                  End
                  Else
                  Begin
                      rxSuministroAnexo.FieldValues['dPGenerar'] := dSuministrado - dGenerado ;
                      rxSuministroAnexo.FieldValues['dPReportar'] := 0 ;
                  End ;
                  If Connection.qryBusca.FieldValues['dCantidadAnexo'] > dSuministrado Then
                      rxSuministroAnexo.FieldValues['dPSuministrar'] := Connection.qryBusca.FieldValues['dCantidadAnexo'] - dSuministrado
                  Else
                      rxSuministroAnexo.FieldValues['dPSuministrar'] := 0 ;
                  rxSuministroAnexo.Post ;
                  Connection.qryBusca2.Next ;
             End
         Else
         Begin
             rxSuministroAnexo.Append ;
             rxSuministroAnexo.FieldValues['sNumeroActividad'] := Connection.qryBusca.FieldValues['sNumeroActividad'] ;
             rxSuministroAnexo.FieldValues['mDescripcion'] := MidStr(Connection.qryBusca.FieldValues['mDescripcion'],1,100) + ' ..' ;
             rxSuministroAnexo.FieldValues['sMedida'] := Connection.qryBusca.FieldValues['sMedida'] ;
             rxSuministroAnexo.FieldValues['dCantidadAnexo'] := Connection.qryBusca.FieldValues['dCantidadAnexo'] ;
             rxSuministroAnexo.FieldValues['dPonderado'] := Connection.qryBusca.FieldValues['dPonderado'] ;
             rxSuministroAnexo.FieldValues['dVentaMN'] := Connection.qryBusca.FieldValues['dVentaMN'] ;
             rxSuministroAnexo.FieldValues['sReferencia'] := 'SIN AVISO DE EMBARQUE' ;
             rxSuministroAnexo.FieldValues['dCantidad'] := 0 ;
             rxSuministroAnexo.FieldValues['dGenerado'] := dGenerado ;
             If dGenerado > dSuministrado Then
             Begin
                 rxSuministroAnexo.FieldValues['dPGenerar'] := 0 ;
                 rxSuministroAnexo.FieldValues['dPReportar'] := dGenerado - dSuministrado ;
             End
             Else
             Begin
                 rxSuministroAnexo.FieldValues['dPGenerar'] := dSuministrado - dGenerado ;
                 rxSuministroAnexo.FieldValues['dPReportar'] := 0 ;
             End ;
             If Connection.qryBusca.FieldValues['dCantidadAnexo'] > dSuministrado Then
                 rxSuministroAnexo.FieldValues['dPSuministrar'] := Connection.qryBusca.FieldValues['dCantidadAnexo'] - dSuministrado
             Else
                 rxSuministroAnexo.FieldValues['dPSuministrar'] := 0 ;
             rxSuministroAnexo.Post ;
         End ;
         Connection.qryBusca.Next ;
         progress.Progress := connection.QryBusca.RecNo ;
    End ;
    progress.Visible := False ;
    progress.Progress := 0 ;
    //Obtenemos Reporte en Dolares y M.N
    if chkMN.Checked = True then
    begin
       rInforme.LoadFromFile (global_files + 'ConcentradodeSuministros.fr3');
       if not FileExists(global_files + 'ConcentradodeSuministros.fr3') then
          showmessage('El archivo de reporte ConcentradodeSuministros.fr3 no existe, notifique al administrador del sistema');
    end
    else
    begin
       rInforme.LoadFromFile (global_files + 'ConcentradodeSuministrosDLL.fr3');
       if not FileExists(global_files + 'ConcentradodeSuministrosDLL.fr3') then
          showmessage('El archivo de reporte ConcentradodeSuministrosDLL.fr3 no existe, notifique al administrador del sistema');
    end;
    rInforme.PreviewOptions.MDIChild := False ;
    rInforme.PreviewOptions.Modal := True ;
    rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
    rInforme.PreviewOptions.ShowCaptions := False ;
    rInforme.Previewoptions.ZoomMode := zmPageWidth ;
    rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al generar concentrado de suministros', 0);
    end;
  end;
end;

procedure TfrmCompara.btnAcumuladoTrinomioClick(Sender: TObject);
begin
//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
  try
    rDiarioFirmas (global_contrato, '' , 'A',tdIdFecha.Date, frmCompara ) ;
    If MessageDlg('Desea imprimir el consolidado de todas las estimaciones seleccionadas?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    Begin
        Reporte.Active := False ;
        Reporte.SQL.Clear ;
        Reporte.SQL.Add('Select ct.*, e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador, Sum(e.dCantidad * a.dVentaMN) as dEstimado From estimacionxpartida e ' +
                        'INNER JOIN estimaciones e2 ON (e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And e.sNumeroGenerador = e2.sNumeroGenerador And e2.lStatus = "Autorizado") ' +
                        'INNER JOIN contrato_trinomio ct ON (e.sContrato = ct.sContrato And e.sInstalacion = ct.sInstalacion) ' +
                        'INNER JOIN actividadesxanexo a ON (e.sContrato = a.sContrato And e.sNumeroActividad = a.sNumeroActividad And a.sIdConvenio = :Convenio And a.sTipoActividad = "Actividad") ' +
                        'Where e.sContrato = :Contrato And e.lEstima = "Si" ' +
                        'Group By ct.sInstalacion, e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador') ;
        Reporte.Params.ParamByName('Contrato').DataType := ftString ;
        Reporte.Params.ParamByName('Contrato').Value := global_Contrato ;
        Reporte.Params.ParamByName('Convenio').DataType := ftString ;
        Reporte.Params.ParamByName('Convenio').Value := sConvenio ;
        dsInforme.FieldAliases.Clear ;
        dsInforme.DataSet := Reporte ;
        Reporte.Open ;
    End
    Else
    Begin
     if length(vartostr(Connection.EstimacionPeriodo.FieldValues['dFechaFinal']))>0 then
     begin
        Reporte.Active := False ;
        Reporte.SQL.Clear ;
        Reporte.SQL.Add('Select ct.*, e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador, Sum(e.dCantidad * a.dVentaMN) as dEstimado From estimacionxpartida e ' +
                        'INNER JOIN estimaciones e2 ON (e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And e.sNumeroGenerador = e2.sNumeroGenerador And e2.lStatus = "Autorizado" ' +
                                                       'And Month(e2.dFechaFinal) = :Mes And Year(e2.dFechaFinal) = :Anno) ' +
                        'INNER JOIN contrato_trinomio ct ON (e.sContrato = ct.sContrato And e.sInstalacion = ct.sInstalacion) ' +
                        'INNER JOIN actividadesxanexo a ON (e.sContrato = a.sContrato And e.sNumeroActividad = a.sNumeroActividad And a.sIdConvenio = :Convenio And a.sTipoActividad = "Actividad") ' +
                        'Where e.sContrato = :Contrato And e.lEstima = "Si" ' +
                        'Group By ct.sInstalacion, e2.iNumeroEstimacion, e2.sNumeroOrden, e2.sNumeroGenerador') ;
        Reporte.Params.ParamByName('Contrato').DataType := ftString ;
        Reporte.Params.ParamByName('Contrato').Value := global_Contrato ;
        Reporte.Params.ParamByName('Convenio').DataType := ftString ;
        Reporte.Params.ParamByName('Convenio').Value := sConvenio ;
        Reporte.Params.ParamByName('Mes').DataType := ftInteger ;
        Reporte.Params.ParamByName('Mes').Value := MonthOf(Connection.EstimacionPeriodo.FieldValues['dFechaFinal']) ;
        Reporte.Params.ParamByName('Anno').DataType := ftInteger ;
        Reporte.Params.ParamByName('Anno').Value := YearOf(Connection.EstimacionPeriodo.FieldValues['dFechaFinal']) ;
        dsInforme.FieldAliases.Clear ;
        dsInforme.DataSet := Reporte ;
        Reporte.Open ;
     end;
    End ;
    rInforme.LoadFromFile (global_files + 'TrinomioConcentrado.fr3');
    rInforme.PreviewOptions.MDIChild := False ;
    rInforme.PreviewOptions.Modal := True ;
    rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
    rInforme.PreviewOptions.ShowCaptions := False ;
    rInforme.Previewoptions.ZoomMode := zmPageWidth ;
    rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
      if not FileExists(global_files + 'TrinomioConcentrado.fr3') then
        showmessage('El archivo de reporte TrinomioConcentrado.fr3 no existe, notifique al administrador del sistema');
//    frxTrinomio.ShowReport

  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'En el proceso Acumulado de generacíon trinomio MN', 0);
    end;
  end;
end;

procedure TfrmCompara.ActividadesxOrdenCalcFields(DataSet: TDataSet);
var
  sPeriodo:String;
begin
  try
    ActividadesxOrdendMontoMN.Value := ActividadesxOrdendCantidad.Value * ActividadesxOrdendVentaMN.Value ;
    sPeriodo:='';

    If ActividadesxOrden.FieldValues['sTipoActividad'] = 'Actividad' Then
    Begin
        if (chkPeriodo.Checked) then
          sPeriodo:=' and didfecha between :Inicio and :Fecha'
        else
          sPeriodo:=' And dIdFecha <= :Fecha';

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(dCantidad) as dReportado From bitacoradeactividades ' +
                                     'Where sContrato = :contrato' + sPeriodo + ' And sNumeroOrden = :Orden And ' +
                                     'sWbs = :wbs And sNumeroActividad = :Actividad Group By sWbs, sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca2.Params.ParamByName('Orden').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
        if (chkPeriodo.Checked) then
        begin
          Connection.qryBusca2.Params.ParamByName('Inicio').DataType := ftDate ;
          Connection.qryBusca2.Params.ParamByName('Inicio').Value := tdIdFecha1.Date ;
        end;
        Connection.qryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
        Connection.qryBusca2.Params.ParamByName('Wbs').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Wbs').Value := ActividadesxOrden.FieldValues['sWbs'] ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad'] ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
             ActividadesxOrdendReportado.Value := Connection.qryBusca2.FieldValues ['dReportado'] ;

        if (chkPeriodo.Checked) then
          sPeriodo:=' and e2.dFechaFinal between :Inicio and :Fecha'
        else
          sPeriodo:=' And e2.dFechaFinal <= :Fecha';

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as dGenerado From estimacionxpartida e ' +
                                     'INNER JOIN estimaciones e2 ON (e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And e.sNumeroGenerador = e2.sNumeroGenerador' + sPeriodo + ' And e2.lStatus = "Autorizado") ' +
                                     'Where e.sContrato = :contrato And e.sNumeroOrden = :Orden And e.sWbs = :Wbs And e.sNumeroActividad = :Actividad And e.lEstima = "Si" Group By e.sWbs, e.sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
        if (chkPeriodo.Checked) then
        begin
          Connection.qryBusca2.Params.ParamByName('Inicio').DataType := ftDate ;
          Connection.qryBusca2.Params.ParamByName('Inicio').Value := tdIdFecha1.Date ;
        end;
        Connection.qryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
        Connection.qryBusca2.Params.ParamByName('Orden').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
        Connection.qryBusca2.Params.ParamByName('Wbs').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Wbs').Value := ActividadesxOrden.FieldValues['sWbs'] ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
             ActividadesxOrdendGenerado.Value := Connection.qryBusca2.FieldValues ['dGenerado'] ;

        if (chkPeriodo.Checked) then
          sPeriodo:=' and e2.dFechaAviso between :Inicio and :Fecha'
        else
          sPeriodo:=' And e2.dFechaAviso <= :Fecha';

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as dSuministrado From anexo_psuministro e INNER JOIN anexo_suministro e2 ON ' +
                                     '(e.sContrato = e2.sContrato And e.iFolio = e2.iFolio' + sPeriodo + ' and e2.sNumeroOrden = :Orden) ' +
                                     'Where e.sContrato = :contrato And e.sNumeroActividad = :Actividad Group By e.sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca2.Params.ParamByName('Orden').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
        if (chkPeriodo.Checked) then
        begin
          Connection.qryBusca2.Params.ParamByName('Inicio').DataType := ftDate ;
          Connection.qryBusca2.Params.ParamByName('Inicio').Value := tdIdFecha1.Date ;
        end;
        Connection.qryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value := ActividadesxAnexo.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
             ActividadesxOrdendSuministrado.Value := Connection.qryBusca2.FieldValues ['dSuministrado'] ;

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as dEstimado From estimacionxpartidaprov e ' +
                                     'INNER JOIN estimacionxproveedor e2 ON (e.sContrato = e2.sContrato And e.sSubContrato = e2.sSubContrato And e.iNumeroEstimacion = e2.iNumeroEstimacion) ' +
                                     'Where e.sContrato = :contrato And e.sNumeroOrden = :Orden and e.sWbs = :Wbs And e.sNumeroActividad = :Actividad Group By e.sNumeroOrden') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca2.Params.ParamByName('Orden').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Orden').Value := ActividadesxOrden.FieldValues['sNumeroOrden']  ;
        Connection.qryBusca2.Params.ParamByName('Wbs').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Wbs').Value := ActividadesxOrden.FieldValues['sWbs']  ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad']  ;

        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value := ActividadesxOrden.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
              ActividadesxOrdendSubContrato.Value := Connection.qryBusca2.FieldValues ['dEstimado'] ;
    End
  except
        on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al ejecutar el proceso actividadesxordencalcfield ', 0);
        end;
  end;

end;

procedure TfrmCompara.ActividadesxAnexoCalcFields(DataSet: TDataSet);
var
  sPeriodo:string;
begin
  try
    ActividadesxAnexodMontoMN.Value := ActividadesxAnexodCantidadAnexo.Value * ActividadesxAnexodVentaMN.Value ;
    ActividadesxAnexodMontoDLL.Value := ActividadesxAnexodCantidadAnexo.Value * ActividadesxAnexodVentaDLL.Value ;
    If ActividadesxAnexo.FieldValues['sTipoActividad'] = 'Actividad' Then
    Begin
        sPeriodo:='';
        if(chkPeriodo.Checked) then
          sPeriodo:=' And dIdFecha between :Inicio and :Fecha'
        else
          sPeriodo:=' And dIdFecha <= :Fecha';


        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(dCantidad) as dReportado From bitacoradeactividades ' +
                                     'Where sContrato = :contrato' + sPeriodo + ' And ' +
                                     'sNumeroActividad = :Actividad and lImprime = "Si" Group By sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType  := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value     := global_contrato ;
        if(chkPeriodo.Checked) then
           Connection.qryBusca2.Params.ParamByName('Inicio').AsDate        := tdIdFecha1.Date ;

        Connection.qryBusca2.Params.ParamByName('Fecha').DataType     := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value        := tdIdFecha.Date ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value    := ActividadesxAnexo.FieldValues['sNumeroActividad'] ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
             ActividadesxAnexodReportado.Value := Connection.qryBusca2.FieldValues ['dReportado'] ;

        sPeriodo:='';
        if(chkPeriodo.Checked) then
          sPeriodo:=' And e2.dFechaAviso between :Inicio and :Fecha'
        else
          sPeriodo:=' And e2.dFechaAviso <= :Fecha';


        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as dSuministrado From anexo_psuministro e INNER JOIN anexo_suministro e2 ON ' +
                                     '(e.sContrato = e2.sContrato And e.iFolio = e2.iFolio' + sPeriodo +') ' +
                                     'Where e.sContrato = :contrato And e.sNumeroActividad = :Actividad Group By e.sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType  := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value     := global_contrato ;

        if(chkPeriodo.Checked) then
          Connection.qryBusca2.Params.ParamByName('Inicio').AsDate        := tdIdFecha1.Date ;

        Connection.qryBusca2.Params.ParamByName('Fecha').DataType     := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value        := tdIdFecha.Date ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value    := ActividadesxAnexo.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
             ActividadesxAnexodSuministrado.Value := Connection.qryBusca2.FieldValues ['dSuministrado'] ;

        sPeriodo:='';
        if(chkPeriodo.Checked) then
          sPeriodo:=' And e2.dFechaFinal between :Inicio and :Fecha'
        else
          sPeriodo:=' And e2.dFechaFinal <= :Fecha';

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as dGenerado From estimacionxpartida e ' +
                                     'INNER JOIN estimaciones e2 ON (e.sContrato = e2.sContrato And e.sNumeroOrden = e2.sNumeroOrden And ' +
                                     'e.sNumeroGenerador = e2.sNumeroGenerador' + sPeriodo +' and e2.lStatus = "Autorizado") ' +
                                     'Where e.sContrato = :contrato And e.sNumeroActividad = :Actividad Group By e.sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType  := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value     := global_contrato ;
        if(chkPeriodo.Checked) then
          Connection.qryBusca2.Params.ParamByName('Inicio').AsDate        := tdIdFecha1.Date ;
        Connection.qryBusca2.Params.ParamByName('Fecha').DataType     := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value        := tdIdFecha.Date ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value    := ActividadesxAnexo.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
              ActividadesxAnexodGenerado.Value := Connection.qryBusca2.FieldValues ['dGenerado'] ;

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as dEstimado From actividadesxestimacion e ' +
                                     'INNER JOIN estimacionperiodo e2 ON (e.sContrato = e2.sContrato And e.iNumeroEstimacion = e2.iNumeroEstimacion and e2.lEstimado = "Si") ' +
                                     'Where e.sContrato = :contrato And e.sNumeroActividad = :Actividad and e.sTipoActividad = "Actividad" Group By e.sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType  := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value     := global_contrato ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value    := ActividadesxAnexo.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
              ActividadesxAnexodEstimado.Value := Connection.qryBusca2.FieldValues ['dEstimado'] ;

        Connection.qryBusca2.Active := False ;
        Connection.qryBusca2.SQL.Clear ;
        Connection.qryBusca2.SQL.Add('Select Sum(e.dCantidad) as dEstimado From estimacionxpartidaprov e ' +
                                     'INNER JOIN estimacionxproveedor e2 ON (e.sContrato = e2.sContrato And e.sSubContrato = e2.sSubContrato And e.iNumeroEstimacion = e2.iNumeroEstimacion) ' +
                                     'Where e.sContrato = :contrato And e.sNumeroActividad = :Actividad Group By e.sNumeroActividad') ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType  := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value     := global_contrato ;
        Connection.qryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Actividad').Value    := ActividadesxAnexo.FieldValues['sNumeroActividad']  ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
              ActividadesxAnexodSubContrato.Value := Connection.qryBusca2.FieldValues ['dEstimado'] ;


    End
  except
        on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al ejecutar el proceso actividadesxanexocalcfields ', 0);
        end;
  end;
end;

procedure TfrmCompara.rInformeGetValue(const VarName: String;
var Value: Variant);
var
    dAvance : Currency ;
    eAvance: Extended;
    sPeriodo:string;
begin
  if roqOrdenes.Found then
  begin
    if CompareText(VarName, 'DESC_CONTRATO') = 0 then
      Value := Connection.contrato.FieldByName('mDescripcion').AsString;
    if CompareText(VarName, 'DESC_ORDEN') = 0 then
      Value := roqOrdenes.FieldByName('mDescripcion').AsString;
  end
  else
  begin
    if CompareText(VarName, 'DESC_CONTRATO') = 0 then
      Value := '';
    if CompareText(VarName, 'DESC_ORDEN') = 0 then
      Value := Connection.contrato.FieldByName('mDescripcion').AsString;
  end;

  If CompareText(VarName, 'ORDEN') = 0 then
      Value := tsNumeroOrden.Text ;
  If CompareText(VarName, 'MI_FECHA') = 0 then
      Value := DateToStr(tdIdFecha.Date) ;
  If CompareText(VarName, 'MI_FECHA1') = 0 then
      Value := DateToStr(tdIdFecha1.Date) ;
  If CompareText(VarName, 'PROGRAMADO') = 0 then
      Value := avProgramado ;

  If CompareText(VarName, 'REAL') = 0 then
     Value := avFisico ;

  If CompareText(VarName, 'TEXTO') = 0 then
     If avProgramado > avFisico Then
          Value := 'ATRASO DEL '
     Else
          Value := 'AVANCE DEL ' ;

  If CompareText(VarName, 'DIFERENCIA') = 0 then
     If avProgramado > avFisico Then
          Value := avProgramado - avFisico
     Else
          Value := avFisico - avProgramado ;

   If CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
       Value := sSupervisorTierra ;

   If CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
       Value := sPuestoSupervisorTierra ;

  If CompareText(VarName, 'ULTIMO_REPORTE') = 0 then
  Begin
      Connection.qryBusca.Active := False ;
      Connection.qryBusca.SQL.Clear ;
      Connection.qryBusca.SQL.Add('Select max(dIdFecha) as dIdFecha From reportediario Where sContrato = :Contrato') ;
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
      Connection.qryBusca.Open ;
      If Connection.qryBusca.RecordCount > 0 Then
          Value := DateToStr(Connection.qryBusca.FieldValues['dIdFecha'] )
      Else
          Value := 'S/REPORTE'
  End ;

  If CompareText(VarName, 'FECHA') = 0 then
      Value := DateToStr(tdIdFecha.Date) ;

  If CompareText(VarName, 'AVANCE_CONTRATO') = 0 then
      If ActividadesxAnexo.FieldValues['sTipoActividad'] = 'Actividad' Then
      Begin
        sPeriodo:='';
        if (chkPeriodo.Checked) then
          sPeriodo:=' And c.didfecha between :Inicio and :Fecha'
        else
          sPeriodo:=' and c.didfecha <= :fecha';

        Connection.QryBusca.Active := False;
        Connection.QryBusca.Sql.Text := 'select a.swbs, b.dCantidad as dCantidadOrden, sum(c.dCantidad) as dCantidad, sum(c.dAvance * (b.dCantidad / a.dCantidadAnexo)) as dAvance ' +
                               'from actividadesxanexo a left join actividadesxorden b on (b.scontrato = a.scontrato and b.sidconvenio = :convenio and b.sWbsContrato = a.sWbs) ' +
                               'left join bitacoradeactividades c on (c.scontrato = b.scontrato and c.snumeroorden = b.snumeroorden and c.swbs = b.sWbs' + sPeriodo + ') ' +
                               'where a.scontrato = :contrato and a.sidconvenio = :convenio and a.sTipoActividad = "Actividad" and a.swbs = :wbs ' +
                               'group by a.swbs order by a.iItemOrden';
        Connection.QryBusca.ParamByName('contrato').AsString := global_contrato;
        Connection.QryBusca.ParamByName('convenio').AsString := global_convenio;
        if (chkPeriodo.Checked) then
        begin
          connection.QryBusca.Params.ParamByName('Inicio').DataType    := ftDate ;
          connection.QryBusca.Params.ParamByName('Inicio').Value       := tdIdFecha1.Date ;
        end;
        Connection.QryBusca.ParamByName('fecha').AsDate      := tdIdFecha.Date;
        Connection.QryBusca.ParamByName('wbs').AsString      := ActividadesxAnexo.FieldByName('sWbs').AsString;
        Connection.QryBusca.Open;
        dAvance := 0;
        while not Connection.QryBusca.Eof do
        begin
          dAvance := dAvance + Connection.QryBusca.FieldByName('dAvance').AsFloat;
          Connection.QryBusca.Next;
        end;
        Value := dAvance;
      End
      Else
      begin
        // Totalizar los avances por paquetes
        sPeriodo:='';
        if (chkPeriodo.Checked) then
          sPeriodo:=' And c.didfecha between :Inicio and :Fecha'
        else
          sPeriodo:=' and c.didfecha <= :fecha';

        Connection.QryBusca.Active := False;
        Connection.QryBusca.Sql.Text := 'select o.swbs, (select sum((c.dAvance * (a.dPonderado / 100)) * (b.dCantidad / a.dCantidadAnexo)) ' +
                                        'from actividadesxanexo a left join actividadesxorden b on (b.scontrato = a.scontrato and b.sidconvenio = :convenio and b.sWbsContrato = a.sWbs) ' +
                                        'left join bitacoradeactividades c on (c.scontrato = b.scontrato and c.snumeroorden = b.snumeroorden and c.swbs = b.sWbs' + sPeriodo +') ' +
                                        'where a.scontrato = :contrato and a.sidconvenio = :convenio and a.sTipoActividad = "Actividad" and a.swbs like concat(o.swbs, ".%")) as dAvance ' +
                                        'from actividadesxanexo o where o.scontrato = :contrato and o.sIdConvenio = :convenio and o.sTipoActividad = "Paquete" and o.swbs = :wbs ' +
                                        'order by o.iItemOrden';
        Connection.QryBusca.ParamByName('contrato').AsString := global_contrato;
        Connection.QryBusca.ParamByName('convenio').AsString := global_convenio;
        if (chkPeriodo.Checked) then
        begin
          connection.QryBusca.Params.ParamByName('Inicio').DataType    := ftDate ;
          connection.QryBusca.Params.ParamByName('Inicio').Value       := tdIdFecha1.Date ;
        end;
        Connection.QryBusca.ParamByName('fecha').AsDate      := tdIdFecha.Date;
        Connection.QryBusca.ParamByName('wbs').AsString      := ActividadesxAnexo.FieldByName('sWbs').AsString;
        Connection.QryBusca.Open;
        dAvance := 0;
        while not Connection.QryBusca.Eof do
        begin
          dAvance := dAvance + Connection.QryBusca.FieldByName('dAvance').AsFloat;
          Connection.QryBusca.Next;
        end;
        Value := dAvance;
      end;

  If CompareText(VarName, 'AVANCE_ORDEN') = 0 then
      If ActividadesxOrden.FieldValues['sTipoActividad'] = 'Actividad' Then
      Begin
        sPeriodo:='';
        if (chkPeriodo.Checked) then
          sPeriodo:=' And b.dIdFecha between :Inicio and :Fecha'
        else
          sPeriodo:=' And b.dIdFecha <= :Fecha';

        if ActividadesxOrden.Fieldbyname('cancelada').AsString<>'Si' then
        begin
          connection.QryBusca.Active := False ;
          connection.QryBusca.SQL.Clear ;
          connection.QryBusca.SQL.Add('Select sum(b.dAvance) as dAvance From bitacoradeactividades b ' +
                                      'Where b.sContrato = :contrato And sNumeroOrden = :Orden And b.sWbs = :Wbs'+
                                      sPeriodo + ' Group By b.sWbs' ) ;
          connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
          connection.QryBusca.Params.ParamByName('Contrato').Value    := global_Contrato ;
          connection.QryBusca.Params.ParamByName('Orden').DataType    := ftString ;
          connection.QryBusca.Params.ParamByName('Orden').Value       := ActividadesxOrden.FieldValues['sNumeroOrden'] ;
          Connection.QryBusca.ParamByName('wbs').AsString             := ActividadesxOrden.FieldByName('swbs').AsString;
          if (chkPeriodo.Checked) then
          begin
            connection.QryBusca.Params.ParamByName('Inicio').DataType    := ftDate ;
            connection.QryBusca.Params.ParamByName('Inicio').Value       := tdIdFecha1.Date ;
          end;
          connection.QryBusca.Params.ParamByName('Fecha').DataType    := ftDate ;
          connection.QryBusca.Params.ParamByName('Fecha').Value       := tdIdFecha.Date ;
          connection.QryBusca.Open ;
          If connection.QryBusca.RecordCount > 0 Then
              Value := connection.QryBusca.FieldValues['dAvance']
          Else
              Value := 0
        end
        else
          Value := 100;
      End
      Else
      begin
        // Totalizar los avances por paquetes

        sPeriodo:='';
        if (chkPeriodo.Checked) then
          sPeriodo:=' And b.dIdFecha between :Inicio and :Fecha'
        else
          sPeriodo:=' and b.didfecha <= :fecha';

        Connection.QryBusca.Active := False;
        Connection.QryBusca.Sql.Clear;
        (*Connection.QryBusca.Sql.Add('Select a.dPonderado, sum(b.dAvance) as dAvance, if(sum(b.dAvance) > 100, a.dPonderado, sum(b.dAvance * (a.dPonderado / 100))) as dAvancePonderado ' +
                                    'From actividadesxorden a inner join bitacoradeactividades b on (b.scontrato = a.scontrato and b.snumeroorden = a.snumeroorden and b.swbs = a.swbs and b.didfecha <= :fecha) ' +
                                    'Where a.sContrato = :contrato and a.sIdConvenio =:Convenio and a.sNumeroOrden = :orden And b.sWbs like concat(:wbs, ".%") group by a.swbs');
        *)

        Connection.QryBusca.Sql.Add('Select a.dPonderado, if((select ba.lCancelada from bitacoradeactividades ba where a.sContrato = ba.sContrato and a.sNumeroOrden = ba.sNumeroOrden ' +
                                    'and a.swbs = ba.swbs and lCancelada = "Si" limit 1) ="Si",100,sum(b.dAvance)) as dAvance, if(sum(b.dcantidad) > a.dcantidad, a.dPonderado, if((select ba.lCancelada from bitacoradeactividades ba ' +
                                    'where a.sContrato = ba.sContrato and a.sNumeroOrden = ba.sNumeroOrden and a.swbs = ba.swbs and lCancelada = "Si" and ba.didfecha <= :fecha limit 1) ="Si",a.dPonderado, ' +
                                    'sum(b.dcantidad * (a.dPonderado / a.dcantidad))))as dAvancePonderado ' +
                                    'From actividadesxorden a inner join bitacoradeactividades b on (b.scontrato = a.scontrato and b.snumeroorden = a.snumeroorden and b.swbs = a.swbs'+speriodo+') ' +
                                    'Where a.sContrato = :contrato and a.sIdConvenio =:Convenio and a.sNumeroOrden = :orden And b.sWbs like concat(:wbs, ".%") group by a.swbs');


        Connection.QryBusca.ParamByName('contrato').AsString := global_contrato;
        Connection.QryBusca.ParamByName('convenio').AsString := global_convenio;
        Connection.QryBusca.ParamByName('orden').AsString    := ActividadesxOrden.FieldValues['sNumeroOrden'];
        Connection.QryBusca.ParamByName('wbs').AsString      := ActividadesxOrden.FieldValues['swbs'];
        Connection.QryBusca.ParamByName('fecha').AsDate      := tdIdFecha.Date;
        if (chkPeriodo.Checked) then
          Connection.QryBusca.ParamByName('Inicio').AsDate      := tdIdFecha1.Date;
        Connection.QryBusca.Open;

        eAvance := 0;
        while Not Connection.QryBusca.Eof do
        begin
          eAvance := eAvance + Connection.QryBusca.FieldByName('dAvancePonderado').AsFloat;
          Connection.QryBusca.Next;
        end;
        Connection.QryBusca.Close;

        Value := eAvance;
      end;
end;

procedure TfrmCompara.btnReportadoVsGeneradoClick(Sender: TObject);
var
   consulta : string;
   sOrden:String;
begin
   { if ChkPu.Checked = True Then
        cadpua := 'And sTipoAnexo =  "PU" '
    else
        cadpua := 'And sTipoAnexo = "ADM" ' ;    }

  cadpua:='';
  if ChkPu.Checked = True Then
    cadpua := 'sTipoAnexo =  "PU" ';

  if chkadmon.Checked then
    if cadpua='' then
      cadpua := 'And sTipoAnexo = "ADM" '
    else
      cadpua := ' And (' + cadpua + 'or sTipoAnexo = "ADM") '
  else
    if cadpua<>'' then
      cadpua:= ' and ' + cadpua;


  sOrden:='';
  if connection.configuracion.FieldByName('lOrdenaItem').AsString = 'Si' then
    sOrden:=' Order by iItemOrden '
  else
    sOrden:=' Order By mysql.udf_NaturalSortFormat(swbs,'+ IntToStr(Global_TamOrden) +  ',' +Quotedstr(Global_SepOrden) +') ';



//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
     showmessage('la fecha final es menor a la fecha inicial' );
     tdIdFecha.SetFocus;
     exit;
   end;
 try

    {StringReplace(before, ' a ', ' THE ',
                          [rfReplaceAll, rfIgnoreCase]);}

    If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
    Begin
        if tsFiltro.Text = 'SOLO REPORTADAS' then
           consulta := 'inner join bitacoradeactividades b '+
                       'on (a.sContrato = b.sContrato  and a.sNumeroActividad = b.sNumeroActividad and sTipoActividad = "Actividad" '+
                       'and b.dIdFecha >= :FechaI and b.dIdFecha <= :FechaF) '+
                       'Where a.sContrato = :Contrato and a.sIdConvenio = :Convenio ' + cadpua + ' group by a.sWbs' + StringReplace(sOrden,'swbs','a.swbs',[rfReplaceAll, rfIgnoreCase])
        else
            consulta := 'Where sContrato = :contrato and sIdConvenio = :convenio ' + cadpua + sOrden;

        ActividadesxAnexo.Active := False ;
        ActividadesxAnexo.SQL.Clear ;
        ActividadesxAnexo.SQL.Add('select a.sContrato, a.iNivel, a.iColor, a.sTipoActividad, a.sWbsAnterior, a.sWbs, a.sNumeroActividad, a.mDescripcion, a.dFechaInicio, a.dFechaFinal, ' +
                                  'a.sMedida, a.dCantidadAnexo, a.dPonderado, a.dVentaMN, a.dVentaDLL, "" as cancelada  from actividadesxanexo a '+ consulta);
        ActividadesxAnexo.Params.ParamByName('Contrato').DataType := ftString ;
        ActividadesxAnexo.Params.ParamByName('Contrato').Value    := global_contrato ;
        ActividadesxAnexo.Params.ParamByName('Convenio').DataType := ftString ;
        ActividadesxAnexo.Params.ParamByName('Convenio').Value    := sConvenio ;
        if tsFiltro.Text = 'SOLO REPORTADAS' then
        begin
            ActividadesxAnexo.Params.ParamByName('FechaI').DataType := ftDate ;
            ActividadesxAnexo.Params.ParamByName('FechaI').Value    := tdIdFecha1.Date;
            ActividadesxAnexo.Params.ParamByName('FechaF').DataType := ftDate;
            ActividadesxAnexo.Params.ParamByName('FechaF').Value    := tdIdFecha.Date  ;
        end;

        ActividadesxAnexo.Open ;
        //Obtenemos reportes en M.N y DLL por contrato..
        if chkMN.Checked = True then
        begin
           rInforme.LoadFromFile (global_files + 'ComparativoReportadoGeneradoContrato.fr3');
           if not FileExists(global_files + 'ComparativoReportadoGeneradoContrato.fr3') then
             showmessage('El archivo de reporte ComparativoReportadoGeneradoContrato.fr3 no existe, notifique al administrador del sistema');
        end
        else
        begin
           rInforme.LoadFromFile (global_files + 'ComparativoReportadoGeneradoContratoDLL.fr3') ;
           if not FileExists(global_files + 'ComparativoReportadoGeneradoContratoDLL.fr3') then
             showmessage('El archivo de reporte ComparativoReportadoGeneradoContratoDLL.fr3 no existe, notifique al administrador del sistema');
        end;
        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
    End
    Else
    Begin
        if tsFiltro.Text = 'SOLO REPORTADAS' then
           consulta := 'inner join bitacoradeactividades b '+
                       'on (a.sContrato = b.sContrato  and a.sNumeroActividad = b.sNumeroActividad and sTipoActividad = "Actividad" '+
                       'and b.dIdFecha >= :FechaI and b.dIdFecha <= :FechaF) '+
                       'Where a.sContrato = :Contrato and a.sIdConvenio = :Convenio and a.sNumeroOrden = :orden group by a.sWbs'  + StringReplace(sOrden,'swbs','a.swbs',[rfReplaceAll, rfIgnoreCase])
        else
        begin
            consulta := 'Where a.sContrato = :contrato and a.sIdConvenio = :convenio and a.sNumeroOrden = :orden' + StringReplace(sOrden,'swbs','a.swbs',[rfReplaceAll, rfIgnoreCase]);
        end;
        ActividadesxOrden.Active := False ;
        ActividadesxOrden.SQL.Clear ;
        ActividadesxOrden.SQL.Add('select a.sContrato, a.sNumeroOrden, a.dCostoMN, a.dCostoDLL, a.iNivel, a.iColor, a.sTipoActividad, a.sWbsAnterior, a.sWbs, a.sNumeroActividad, a.mDescripcion, a.dFechaInicio, a.dFechaFinal, ' +
                                  'a.sMedida, a.dCantidad, a.dPonderado, a.dVentaMN, a.dVentaDLL, "" as cancelada  from actividadesxorden a '+ consulta);
        ActividadesxOrden.Params.ParamByName('Contrato').DataType := ftString ;
        ActividadesxOrden.Params.ParamByName('Contrato').Value    := global_contrato ;
        ActividadesxOrden.Params.ParamByName('Convenio').DataType := ftString ;
        ActividadesxOrden.Params.ParamByName('Convenio').Value    := sConvenio ;
        ActividadesxOrden.Params.ParamByName('Orden').DataType    := ftString ;
        ActividadesxOrden.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text ;
        if tsFiltro.Text = 'SOLO REPORTADAS' then
        begin
            ActividadesxOrden.Params.ParamByName('FechaI').DataType := ftDate ;
            ActividadesxOrden.Params.ParamByName('FechaI').Value    := tdIdFecha1.Date;
            ActividadesxOrden.Params.ParamByName('FechaF').DataType := ftDate;
            ActividadesxOrden.Params.ParamByName('FechaF').Value    := tdIdFecha.Date  ;
        end;


        //Obtenemos reportes en M.N y DLL por frente de trabajo..
        ActividadesxOrden.Open ;
        if chkMN.Checked = True then
        begin
           rInforme.LoadFromFile (global_files + 'ComparativoReportadoGeneradoOrden.fr3');
           if not FileExists(global_files + 'ComparativoReportadoGeneradoOrden.fr3') then
             showmessage('El archivo de reporte ComparativoReportadoGeneradoOrden.fr3 no existe, notifique al administrador del sistema');
        end
        else
        begin
           rInforme.LoadFromFile (global_files + 'ComparativoReportadoGeneradoOrdenDLL.fr3');
           if not FileExists(global_files + 'ComparativoReportadoGeneradoOrdenDLL.fr3') then
             showmessage('El archivo de reporte ComparativoReportadoGeneradoOrdenDLL.fr3 no existe, notifique al administrador del sistema');
        end;
        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
    End

 except
        on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al generar reportado vs generado', 0);
        end;
 end;
end;


Procedure TfrmCompara.btnAnexoVsEstimacionesClick(Sender: TObject);
begin

   if ChkPu.Checked = True Then
       cadpua := 'And sTipoAnexo =  "PU" '
   else
       cadpua := 'And sTipoAnexo = "ADM" ' ;

//Verifica que la fecha final no sea menor que la fecha inicio
   if tdIdFecha.Date<tdIdFecha1.Date then
   begin
   showmessage('la fecha final es menor a la fecha inicial' );
   tdIdFecha.SetFocus;
   exit;
   end;
  try
        ActividadesxAnexo.Active := False ;
        ActividadesxAnexo.SQL.Clear ;
        ActividadesxAnexo.SQL.Add('select sContrato, iNivel, iColor, sTipoActividad, sWbsAnterior, sWbs, sNumeroActividad, mDescripcion, dFechaInicio, dFechaFinal, ' +
                                  'sMedida, dCantidadAnexo, dPonderado, dVentaMN, dVentaDLL, dCostoMN, dCostoDLL  from actividadesxanexo Where sContrato = :contrato and ' +
                                  'sIdConvenio = :convenio ' + cadpua + ' order by iItemOrden') ;
        ActividadesxAnexo.Params.ParamByName('Contrato').DataType := ftString ;
        ActividadesxAnexo.Params.ParamByName('Contrato').Value := global_contrato ;
        ActividadesxAnexo.Params.ParamByName('Convenio').DataType := ftString ;
        ActividadesxAnexo.Params.ParamByName('Convenio').Value := sConvenio ;
        ActividadesxAnexo.Open ;
        //Obtenemos reportes en dolares y M.N.
        if chkMN.Checked = True then
        begin
           rInforme.LoadFromFile (global_files + 'ComparativoAnexovsEstimaciones.fr3');
           if not FileExists(global_files + 'ComparativoAnexovsEstimaciones.fr3') then
             showmessage('El archivo de reporte ComparativoAnexovsEstimaciones.fr3 no existe, notifique al administrador del sistema');
        end
        else
        begin
           rInforme.LoadFromFile (global_files + 'ComparativoAnexovsEstimacionesDLL.fr3');
           if not FileExists(global_files + 'ComparativoAnexovsEstimacionesDLL.fr3') then
             showmessage('El archivo de reporte ComparativoAnexovsEstimacionesDLL.fr3 no existe, notifique al administrador del sistema');
        end;
        rInforme.PreviewOptions.MDIChild := False ;
        rInforme.PreviewOptions.Modal := True ;
        rInforme.PreviewOptions.Maximized := lCheckMaximized () ;
        rInforme.PreviewOptions.ShowCaptions := False ;
        rInforme.Previewoptions.ZoomMode := zmPageWidth ;
        rInforme.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP));
  except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_comparativo', 'Al generar cantidad Anexo vs estimaciones', 0);
    end;
  end;
end;

procedure TfrmCompara.btnSubContratosClick(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// ESTATUS DE ORDNES VIGENTES DIAVAZ OCTUBRE 2012 //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena, sConvenio: string;
      fs: tStream;
      Alto: Extended;
      Ren, nivel, i, total: integer;
      Q_Partidas: TZReadOnlyQuery;
      dAvanceProg, dAvanceFisico: double;
      Progreso, TotalProgreso: real;
      dFecha : tDate;
    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      Ren := 2;
    // Realizar los ajustes visuales y de formato de hoja
      Excel.ActiveWindow.Zoom := 100;

      Excel.Columns['A:A'].ColumnWidth := 5.86;
      Excel.Columns['B:B'].ColumnWidth := 4.29;
      Excel.Columns['C:C'].ColumnWidth := 19.14;
      Excel.Columns['D:D'].ColumnWidth := 12.14;
      Excel.Columns['E:E'].ColumnWidth := 52.00;
      Excel.Columns['F:F'].ColumnWidth := 17.43;
      Excel.Columns['G:G'].ColumnWidth := 11.29;
      Excel.Columns['H:H'].ColumnWidth := 12.86;
      Excel.Columns['I:I'].ColumnWidth := 12.00;
      Excel.Columns['J:J'].ColumnWidth := 12.00;
      Excel.Columns['K:K'].ColumnWidth := 35.00;
      Excel.Columns['L:L'].ColumnWidth := 12.71;
      Excel.Columns['M:M'].ColumnWidth := 12.71;

      Hoja.Range['A1:A2'].Select;
      Excel.Selection.RowHeight := '15';

      //Primero la vigencia de la embarcacion principal
      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Clear;
      connection.QryBusca2.SQL.Add('select sIdEmbarcacion from embarcacion_vigencia ' +
        'where sContrato =:Contrato and dFechaInicio <= :FechaI and dFechaFinal >=:FechaF ');
      connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato_barco;
      connection.QryBusca2.ParamByName('FechaI').AsDate := tdIdFecha.Date;
      connection.QryBusca2.ParamByName('FechaF').AsDate := tdIdFecha.Date;
      connection.QryBusca2.Open;

      if connection.QryBusca2.RecordCount = 1 then
        global_barco := connection.QryBusca2.FieldValues['sIdEmbarcacion'];

      //Consultamos las fechas del convenio modificatorio para impresion de las cantidades reportadas superiores al programa de trabajo.
      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Clear;
      connection.QryBusca2.SQL.Add('select sDescripcion from embarcaciones where sContrato =:Contrato and sIdEmbarcacion =:barco ');
      connection.QryBusca2.ParamByName('contrato').AsString := global_contrato_barco;
      connection.QryBusca2.ParamByName('barco').AsString    := global_barco;
      connection.QryBusca2.Open;

      Hoja.Range['B'+IntTostr(Ren-1)+':B'+IntToStr(Ren-1)].Select;
      Excel.Selection.Value := 'ESTATUS DE ORDENES VIGENTES Y FINALIZADAS EN "'+ connection.QryBusca2.FieldValues['sDescripcion'] + '"';
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Size := 16;
      Excel.Selection.Font.color:= clBlack;
      Excel.Selection.Font.Name := 'Calibri';

      Hoja.Range['B'+IntTostr(Ren-1)+':M'+IntToStr(Ren)].Select;
      Excel.Selection.Interior.ColorIndex := 15;
      Excel.Selection.MergeCells:= True;
      Excel.Selection.WrapText  := True;

      inc(Ren);
      Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'FECHA DE IMPRESIÓN AL DÍA: '+ DateToStr(tdIdFecha.Date);
      Excel.Selection.HorizontalAlignment := xlLeft;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Font.Size := 12;
      Excel.Selection.Font.Name := 'Calibri';
      Excel.Selection.RowHeight := '20.25';

      inc(Ren);
      Hoja.Range['B4:M4'].Select;
      Excel.Selection.RowHeight := '41.25';
      Excel.Selection.Interior.ColorIndex := 36;

      // Colocar los encabezados de la plantilla...
      Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'ID';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'TIPO DE OBRA';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'OT';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['E'+IntTostr(Ren)+':E'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'DESCRIPCION';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['F'+IntTostr(Ren)+':F'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'PLATAFORMA';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['G'+IntTostr(Ren)+':G'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'AVANCE PROG.';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['H'+IntTostr(Ren)+':H'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'AVANCE REAL';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['I'+IntTostr(Ren)+':I'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'FECHA INICIO';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['J'+IntTostr(Ren)+':J'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'FECHA TERMINO';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['K'+IntTostr(Ren)+':K'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'OBSERVACIONES';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['L'+IntTostr(Ren)+':L'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'FECHA INICIO REAL';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;
      Hoja.Range['M'+IntTostr(Ren)+':M'+IntToStr(Ren)].Select;
      Excel.Selection.Value := 'ULTIMO REPORTE DIARIO';
      Excel.Selection.Font.Italic := True;
      FormatoEncabezado;
      Excel.Selection.Font.Size := 11;

      inc(Ren);
      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      connection.QryBusca.SQL.Add('select sContrato, mDescripcion, sTipoObra, mComentarios from contratos ');
      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then
      begin
          total := 1;
          while not connection.QryBusca.Eof do
          begin
              {Movimiento de la Barra..}
              Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
              TotalProgreso := TotalProgreso + Progreso;
              BarraEstado.Position := Trunc(TotalProgreso);

              //Ahora consultamos los frentes de trabajo dados de alta en la consulta..
              connection.QryBusca2.Active := False;
              connection.QryBusca2.SQL.Clear;
              connection.QryBusca2.SQL.Add('select sNumeroOrden, sIdFolio from ordenesdetrabajo where sContrato =:Contrato ');
              connection.QryBusca2.ParamByName('Contrato').AsString := connection.QryBusca.FieldValues['sContrato'];
              connection.QryBusca2.Open;

              while not connection.QryBusca2.Eof do
              begin
                  Hoja.Range['A'+IntTostr(Ren)+':A'+IntToStr(Ren)].Select;
                  Excel.Selection.RowHeight := '60';

                  Hoja.Range['B'+IntTostr(Ren)+':M'+IntToStr(Ren)].Select;
                  Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                  Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
                  Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                  Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
                  Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                  Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
                  Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                  Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
                  Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous;
                  Excel.Selection.Borders[xlInsideVertical].Weight    := xlThin;

                  Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
                  Excel.Selection.Interior.ColorIndex := 24;

                  Hoja.Range['D'+IntTostr(Ren)+':M'+IntToStr(Ren)].Select;
                  Excel.Selection.Interior.ColorIndex := 37;

                  {Escritura de Datos en el Archvio de Excel..}
                  Hoja.Cells[Ren, 2].Select;
                  Excel.Selection.Value := total;
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;
                  Excel.Selection.Font.Size := 11;
                  Excel.Selection.Font.Bold := True;
                  Excel.Selection.Font.Name := 'Tahoma';

                  Hoja.Cells[Ren, 3].Select;
                  if connection.QryBusca.FieldValues['sTipoObra'] ='PROGRAMADA' then
                  begin
                     Excel.Selection.Value     := 'PRECIO UNITARIO';
                     Excel.Selection.Font.Color := clBlue;
                  end
                  else
                  begin
                     Excel.Selection.Value := 'ADMINISTRACION';
                     Excel.Selection.Font.Color := clNavy;
                  end;
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;
                  Excel.Selection.WrapText  := True;
                  Excel.Selection.Font.Size := 12;
                  Excel.Selection.Font.Bold := True;
                  Excel.Selection.Font.Name := 'Calibri';

                  Hoja.Cells[Ren, 4].Select;
                  Excel.Selection.Value := connection.QryBusca.FieldValues['sContrato'];
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;
                  Excel.Selection.Font.Size := 12;
                  Excel.Selection.Font.Bold := True;
                  Excel.Selection.Font.Name := 'Calibri';

                  Hoja.Cells[Ren, 5].Select;
                  Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];
                  Excel.Selection.HorizontalAlignment := xlJustify;
                  Excel.Selection.VerticalAlignment   := xLCenter;
                  Excel.Selection.WrapText  := True;
                  Excel.Selection.Font.Size := 11;
                  Excel.Selection.Font.Bold := False;
                  Excel.Selection.Font.Name := 'Calibri';
                  Excel.Selection.MergeCells := True;

                  Hoja.Cells[Ren, 6].Select;
                  Excel.Selection.Value :=  connection.QryBusca2.FieldValues['sNumeroOrden'];
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;
                  Excel.Selection.WrapText  := True;
                  Excel.Selection.Font.Size := 12;
                  Excel.Selection.Font.Bold := True;
                  Excel.Selection.Font.Name := 'Calibri';
                  Excel.Selection.MergeCells := True;

                  //Buacamos el primer reporte diario registrado para ese frente de trabajo
                  Q_Partidas.Active := False;
                  Q_Partidas.SQL.Clear;
                  Q_Partidas.SQL.Add('select dIdFecha from reportediario where sContrato =:contrato and sNumeroOrden =:orden and dIdFecha <=:fecha order by dIdFecha ASC');
                  Q_Partidas.ParamByName('Contrato').AsString := connection.QryBusca.FieldValues['sContrato'];
                  Q_Partidas.ParamByName('Orden').AsString    := connection.QryBusca2.FieldValues['sNumeroOrden'];
                  Q_Partidas.ParamByName('fecha').AsDate      := tdIdFecha.Date;
                  Q_Partidas.Open;

                  if Q_partidas.RecordCount > 0 then
                  begin
                      Hoja.Cells[Ren, 12].Select;
                      Excel.Selection.Value := Q_Partidas.FieldValues['dIdFecha'];
                      Excel.Selection.HorizontalAlignment := xlCenter;
                      Excel.Selection.VerticalAlignment   := xlCenter;
                      Excel.Selection.Font.Size := 12;
                      Excel.Selection.Font.Bold := True;
                      Excel.Selection.Font.Color:= clBlue;
                      Excel.Selection.Font.Name := 'Calibri';
                  end;

                   //Buacamos el ultimo reporte diario registrado para ese frente de trabajo
                  Q_Partidas.Active := False;
                  Q_Partidas.SQL.Clear;
                  Q_Partidas.SQL.Add('select dIdFecha, sIdConvenio from reportediario where sContrato =:contrato and sNumeroOrden =:orden and dIdFecha <=:fecha order by dIdFecha DESC');
                  Q_Partidas.ParamByName('Contrato').AsString := connection.QryBusca.FieldValues['sContrato'];
                  Q_Partidas.ParamByName('Orden').AsString    := connection.QryBusca2.FieldValues['sNumeroOrden'];
                  Q_Partidas.ParamByName('fecha').AsDate      := tdIdFecha.Date;
                  Q_Partidas.Open;

                  if Q_partidas.RecordCount > 0 then
                  begin
                      dFecha    := Q_Partidas.FieldValues['dIdFecha'];
                      sConvenio := Q_Partidas.FieldValues['sIdConvenio'];
                      Hoja.Cells[Ren, 13].Select;
                      Excel.Selection.Value := Q_Partidas.FieldValues['dIdFecha'];
                      Excel.Selection.HorizontalAlignment := xlCenter;
                      Excel.Selection.VerticalAlignment   := xlCenter;
                      Excel.Selection.Font.Size := 12;
                      Excel.Selection.Font.Bold := True;
                      Excel.Selection.Font.Color:= clBlue;
                      Excel.Selection.Font.Name := 'Calibri';
                  end
                  else
                  begin
                      dFecha    := 0;
                      sConvenio := '';
                  end;

                  //Ahora consultamos el avance programado a la fecha seleccionada.
                  Q_Partidas.Active := False;
                  Q_Partidas.SQL.Clear;
                  Q_Partidas.SQL.Add('select dAvancePonderadoGlobal from avancesglobales where sContrato =:Contrato and sNumeroOrden =:Orden and sIdConvenio =:Convenio and dIdFecha <=:fecha order by dIdFecha DESC ');
                  Q_Partidas.ParamByName('Contrato').AsString := connection.QryBusca.FieldValues['sContrato'];
                  Q_Partidas.ParamByName('Orden').AsString    := connection.QryBusca2.FieldValues['sNumeroOrden'];
                  Q_Partidas.ParamByName('Convenio').AsString := sConvenio;
                  Q_Partidas.ParamByName('Fecha').AsDate      := tdIdFecha.Date;
                  Q_Partidas.Open;

                  dAvanceProg := 0;
                  if Q_partidas.RecordCount > 0 then
                     dAvanceProg := Q_Partidas.FieldValues['dAvancePonderadoGlobal'];

                  //Ahora consultamos el avance programado a la fecha seleccionada.
                  Q_Partidas.Active := False;
                  Q_Partidas.SQL.Clear;
                  Q_Partidas.SQL.Add('select sum(dAvance) as dAvance from avancesglobalesxorden where sContrato =:Contrato and sNumeroOrden =:Orden and sIdConvenio =:Convenio and dIdFecha <=:fecha group by sContrato');
                  Q_Partidas.ParamByName('Contrato').AsString := connection.QryBusca.FieldValues['sContrato'];
                  Q_Partidas.ParamByName('Orden').AsString    := connection.QryBusca2.FieldValues['sNumeroOrden'];
                  Q_Partidas.ParamByName('Convenio').AsString := sConvenio;
                  Q_Partidas.ParamByName('Fecha').AsDate      := dFecha;
                  Q_Partidas.Open;

                  dAvanceFisico := 0;
                  if Q_partidas.RecordCount > 0 then
                     dAvanceFisico := Q_Partidas.FieldValues['dAvance'];

                  Hoja.Cells[Ren, 7].Select;
                  Excel.Selection.NumberFormat := '##0.0000%';
                  Excel.Selection.Value := dAvanceProg / 100;
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;
                  Excel.Selection.Font.Size := 12;
                  Excel.Selection.Font.Bold := True;
                  Excel.Selection.Font.Italic := True;
                  Excel.Selection.Font.Color:= clBlue;
                  Excel.Selection.Font.Name := 'Calibri';

                  Hoja.Cells[Ren, 8].Select;
                  Excel.Selection.NumberFormat := '##0.0000%';
                  if dAvanceFisico > 100 then
                     Excel.Selection.Value := 100 / 100
                  else
                     Excel.Selection.Value := dAvanceFisico / 100;
                  Excel.Selection.HorizontalAlignment := xlCenter;
                  Excel.Selection.VerticalAlignment   := xlCenter;
                  Excel.Selection.Font.Size := 12;
                  Excel.Selection.Font.Bold := True;
                  Excel.Selection.Font.Italic := True;
                  Excel.Selection.Font.Color:= clRed;
                  Excel.Selection.Font.Name := 'Calibri';

                  if dAvanceFisico < 100 then
                  begin
                      Hoja.Cells[Ren, 13].Select;
                      Excel.Selection.Font.Color:= clMaroon;
                  end;

                  //Ahora consultamos las fechas de incio y termino
                  Q_Partidas.Active := False;
                  Q_Partidas.SQL.Clear;
                  Q_Partidas.SQL.Add('select dFechaInicio, dFechaFinal from actividadesxorden where sContrato =:Contrato and sNumeroOrden =:Orden and sIdConvenio =:Convenio and iNivel = 0 ');
                  Q_Partidas.ParamByName('Contrato').AsString := connection.QryBusca.FieldValues['sContrato'];
                  Q_Partidas.ParamByName('Orden').AsString    := connection.QryBusca2.FieldValues['sNumeroOrden'];
                  Q_Partidas.ParamByName('Convenio').AsString := sConvenio;
                  Q_Partidas.Open;

                  if Q_partidas.RecordCount > 0 then
                  begin
                      Hoja.Cells[Ren, 9].Select;
                      Excel.Selection.Value := Q_Partidas.FieldValues['dFechaInicio'];
                      Excel.Selection.HorizontalAlignment := xlCenter;
                      Excel.Selection.VerticalAlignment   := xlCenter;
                      Excel.Selection.Font.Size := 12;
                      Excel.Selection.Font.Bold := False;
                      Excel.Selection.Font.Name := 'Calibri';

                      Hoja.Cells[Ren, 10].Select;
                      Excel.Selection.Value := Q_Partidas.FieldValues['dFechaFinal'];
                      Excel.Selection.HorizontalAlignment := xlCenter;
                      Excel.Selection.VerticalAlignment   := xlCenter;
                      Excel.Selection.Font.Size := 12;
                      Excel.Selection.Font.Bold := False;
                      Excel.Selection.Font.Name := 'Calibri';
                  end;

                  Hoja.Cells[Ren, 11].Select;
                  Excel.Selection.Value := connection.QryBusca.FieldValues['mComentarios'];
                  Excel.Selection.HorizontalAlignment := xlJustify;
                  Excel.Selection.VerticalAlignment   := xLCenter;
                  Excel.Selection.WrapText  := True;
                  Excel.Selection.Font.Size := 11;
                  Excel.Selection.Font.Bold := False;
                  Excel.Selection.Font.Italic:= True;
                  Excel.Selection.Font.Name  := 'Calibri';
                  Excel.Selection.MergeCells := True;

                  connection.QryBusca2.Next;
                  Inc(Ren);
                  inc(Total);
              end;
              connection.QryBusca.Next;
          end;
          Hoja.Range['B5:B5'].Select;
      end;
    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'ESTATUS ORDENES ';
      except
        Hoja.Name := 'ESTATUS ORDEDES ';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Estatus de Ordenes:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
      end;
    end;

    Result := Resultado;
  end;

begin
    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
    if not SaveDialog1.Execute then
      Exit;
      // Generar el ambiente de excel
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
    end
    else
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;
    end;

    PanelProgress.Visible := True;
    Label15.Refresh;
    BarraEstado.Position := 0;

    Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

      // Verificar si cuenta con las hojas necesarias
    while Libro.Sheets.Count < 2 do
      Libro.Sheets.Add;

      // Verificar si se pasa de hojas necesarias
    Libro.Sheets[1].Select;
    while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

      // Proceder a generar la hoja REPORTE
    CadError := '';

    if GenerarPlantilla then
    begin
          // Grabar el archivo de excel con el nombre dado
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      PanelProgress.Visible := False;
      messageDlg('El Archivo se Generó Correctamente!', mtInformation, [mbOk], 0);
    end;

    Excel := '';

    if CadError <> '' then
      showmessage(CadError);
end;


procedure TfrmCompara.btnRptAvanXFolioClick(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// AVANCES X FOLIO MOJICA JUNIO 2016 //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena, sConvenio: string;
      fs: tStream;
      Alto: Extended;
      Ren, nivel, i, total: integer;
      Q_Partidas: TZReadOnlyQuery;
      dAvanceProg, dAvanceFisico: double;
      Progreso, TotalProgreso: real;

    const
      aColumExcel : array[1..36] of string = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ');

    var aDays: array[1..31] of Double;
      nTmp, nAcumAvan: Double;
      myYear, myMonth, myDay : Word;
      tdAvanceAnterior: Double;

    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;

      //Consultamos.
      If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
          Begin
          connection.QryBusca.SQL.Add( 'select b.scontrato, o.sidfolio, b.davanceanterior, b.didfecha, b.davance, b.mDescripcion,' +
          'b.dcantidad, b.snumeroorden, b.sidconvenio,b.snumeroactividad, o.mdescripcion as ordenDesc ' +
          'from ordenesdetrabajo o inner join bitacoradeactividades b ON ' +
          '(o.scontrato = b.scontrato and o.snumeroorden = b.snumeroorden ) ' +
          'where b.scontrato = :Contrato and b.didfecha >= :FechaI and b.didfecha <= :FechaF and b.dAvance > 0 Group by o.sidfolio ' +
          'Order by b.scontrato, o.sidfolio, b.didfecha, b.davance');
          connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
          connection.QryBusca.ParamByName('FechaI').AsDate := tdIdFecha1.Date;
          connection.QryBusca.ParamByName('FechaF').AsDate := tdIdFecha.Date;
          End
      Else
        Begin
          connection.QryBusca.SQL.Add( 'select sContrato, sidFolio, sNumeroOrden, sIdPlataforma, mDescripcion, sCsu from '+
          'ordenesdetrabajo where sContrato = :Contrato and sNumeroOrden = :Orden');
          connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
          connection.QryBusca.ParamByName('Orden').Value       := tsNumeroOrden.Text;
      End;

      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then

          Ren := 2;

          Excel.ActiveWindow.Zoom := 100;
          Excel.Columns['A:A'].ColumnWidth := 5.86;
          Excel.Columns['B:B'].ColumnWidth := 52.29;
          Excel.Columns['C:C'].ColumnWidth := 52.29;
          Excel.Columns['D:D'].ColumnWidth := 12.14;
          Excel.Columns['E:E'].ColumnWidth := 12.14;
          Excel.Columns['F:F'].ColumnWidth := 12.14;
          Excel.Columns['G:G'].ColumnWidth := 12.14;
          Excel.Columns['H:H'].ColumnWidth := 12.14;
          Excel.Columns['I:I'].ColumnWidth := 12.14;
          Excel.Columns['J:J'].ColumnWidth := 12.14;
          Excel.Columns['K:K'].ColumnWidth := 12.14;
          Excel.Columns['L:L'].ColumnWidth := 12.14;
          Excel.Columns['M:M'].ColumnWidth := 12.14;
          Excel.Columns['N:N'].ColumnWidth := 12.14;
          Excel.Columns['O:O'].ColumnWidth := 12.14;
          Excel.Columns['P:P'].ColumnWidth := 12.14;
          Excel.Columns['Q:Q'].ColumnWidth := 12.14;
          Excel.Columns['R:R'].ColumnWidth := 12.14;
          Excel.Columns['S:S'].ColumnWidth := 12.14;
          Excel.Columns['T:T'].ColumnWidth := 12.14;
          Excel.Columns['U:U'].ColumnWidth := 12.14;
          Excel.Columns['V:V'].ColumnWidth := 12.14;
          Excel.Columns['W:W'].ColumnWidth := 12.14;
          Excel.Columns['X:X'].ColumnWidth := 12.14;
          Excel.Columns['Y:Y'].ColumnWidth := 12.14;
          Excel.Columns['Z:Z'].ColumnWidth := 12.14;
          Excel.Columns['AA:AA'].ColumnWidth := 12.14;
          Excel.Columns['AB:AB'].ColumnWidth := 12.14;
          Excel.Columns['AC:AC'].ColumnWidth := 12.14;
          Excel.Columns['AD:AD'].ColumnWidth := 12.14;
          Excel.Columns['AE:AE'].ColumnWidth := 12.14;
          Excel.Columns['AF:AF'].ColumnWidth := 12.14;
          Excel.Columns['AG:AG'].ColumnWidth := 12.14;
          Excel.Columns['AH:AH'].ColumnWidth := 12.14;
          Excel.Columns['AI:AI'].ColumnWidth := 12.14;
          Excel.Columns['AJ:AJ'].ColumnWidth := 12.14;

          //Hoja.Cells[1, 1] := 'Hola mundo';

          Hoja.Range['A1:A2'].Select;
          Excel.Selection.RowHeight := '15';

          Hoja.Range['B'+IntTostr(Ren-1)+':B'+IntToStr(Ren-1)].Select;
          Excel.Selection.Value := 'CONTROL DE AVANCES POR FOLIO DEL '+ DateToStr(tdIdFecha1.Date)+' AL ' + DateToStr(tdIdFecha.Date);
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 16;
          Excel.Selection.Font.color:= clBlack;
          Excel.Selection.Font.Name := 'Calibri';

          Hoja.Range['B'+IntTostr(Ren-1)+':AJ'+IntToStr(Ren)].Select;
          Excel.Selection.Interior.ColorIndex := 15;
          Excel.Selection.MergeCells:= True;
          Excel.Selection.WrapText  := True;

          inc(Ren);
          Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FECHA DE IMPRESIÓN AL DÍA: '+ DateToStr(now);
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Name := 'Calibri';
          Excel.Selection.RowHeight := '20.25';

          inc(Ren);
          Hoja.Range['B4:AJ4'].Select;

          Excel.Selection.RowHeight := '41.25';
          Excel.Selection.Interior.ColorIndex := 36;

          //Colocar los encabezados de la plantilla...
          Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FOLIO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'ACTIVIDAD';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;
          Excel.Selection.RowHeight := '60';

          Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'AVANCE ANTERIOR';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          nivel:=1;
          for i := 5 to 35 do
          begin
              Hoja.Range[aColumExcel[i] + IntTostr(Ren)+':' + aColumExcel[i] + IntToStr(Ren)].Select;
              Excel.Selection.Value := 'DIA ' + IntToStr(nivel);
              Excel.Selection.Font.Italic := True;
              FormatoEncabezado;
              Excel.Selection.Font.Size := 11;
              nivel:=nivel+1;
          end;

          //Redimensionar vector dinamico al no. máximo de dias en un mes
          for i := 1 to 31 do
          begin
              aDays[i] := 0;
          end;

          Hoja.Range[aColumExcel[36] + IntTostr(Ren)+':' + aColumExcel[36] + IntToStr(Ren)].Select;
          Excel.Selection.Value := 'AVANCE GLOBLAL';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          inc(Ren);

          //ShowMessage( IntToStr( DaysBetween(  tdIdFecha1.Date, tdIdFecha.Date ) ) ) ;

      begin
          total := 1;

          while not connection.QryBusca.Eof do
          begin
              {Movimiento de la Barra..}
              Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
              TotalProgreso := TotalProgreso + Progreso;
              BarraEstado.Position := Trunc(TotalProgreso);

              connection.QryBusca2.Active := False;
              connection.QryBusca2.SQL.Clear;
              {
              connection.QryBusca2.SQL.Add( 'select b.scontrato, o.sidfolio, b.davanceanterior, b.didfecha, b.davance,'+
              'b.mDescripcion, b.dcantidad, b.snumeroorden, b.sidconvenio,b.snumeroactividad, b.davanceActual, '+
              'o.mdescripcion as ordenDesc, a.dPonderado ' +
              'from ordenesdetrabajo o inner join bitacoradeactividades b ON ' +
              '(b.scontrato = o.scontrato and b.snumeroorden = o.snumeroorden ) '+
              'inner join actividadesxorden a ON '+
              '(a.scontrato = o.scontrato and a.snumeroorden = o.snumeroorden and '+
              'a.snumeroactividad = b.snumeroactividad) '+
              'where o.sidfolio=:Folio and b.scontrato = :Contrato and b.didfecha >= :FechaI and b.didfecha <= :FechaF ' +
              'Order by b.didfecha, b.snumeroactividad');
              }

              connection.QryBusca2.SQL.Add( 'select b.*, o.sidfolio, o.mDescripcion ' +
              'from ordenesdetrabajo o inner join bitacoradeactividades b ON ' +
              '(o.scontrato = b.scontrato and o.snumeroorden = b.snumeroorden ) '+
              'where o.sidfolio=:Folio and b.scontrato = :Contrato and b.didfecha >= :FechaI and b.didfecha <= :FechaF '+
              'Order by b.snumeroactividad, b.didfecha, time(b.sHoraInicio)');
              connection.QryBusca2.ParamByName('Folio').AsString := connection.QryBusca.FieldValues['sidfolio'];
              connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato;
              connection.QryBusca2.ParamByName('FechaI').AsDate := tdIdFecha1.Date;
              connection.QryBusca2.ParamByName('FechaF').AsDate := tdIdFecha.Date;
              connection.QryBusca2.Open;

              if connection.QryBusca2.recordcount > 0 then
                begin
                  Q_Partidas.Active := False;
                  Q_Partidas.SQL.Clear;
                  Q_Partidas.SQL.Add( 'select b.*, o.sidfolio, o.mDescripcion ' +
                  'from ordenesdetrabajo o inner join bitacoradeactividades b ON ' +
                  '(o.scontrato = b.scontrato and o.snumeroorden = b.snumeroorden ) '+
                  'where o.sidfolio=:Folio and b.scontrato = :Contrato and b.didfecha >= :FechaI and '+
                  'b.didfecha <= :FechaF Group By b.didfecha '+
                  'Order by b.snumeroactividad, b.didfecha, time(b.sHoraInicio)');
                  Q_Partidas.ParamByName('Folio').AsString := connection.QryBusca.FieldValues['sidfolio'];
                  Q_Partidas.ParamByName('Contrato').AsString := global_contrato;
                  Q_Partidas.ParamByName('FechaI').AsDate := tdIdFecha1.Date;
                  Q_Partidas.ParamByName('FechaF').AsDate := tdIdFecha.Date;
                  Q_Partidas.Open;

                  {Avances anteriores}
                  connection.QryBusca2.SQL.Clear;
                  connection.QryBusca2.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd',
                  Q_Partidas.FieldByName('dIdFecha').AsDateTime)+'", :Orden, :Folio), 2) AS dAvanceAnterior;';
                  connection.QryBusca2.ParamByName('Orden').AsString := Q_Partidas.FieldByName('sContrato').AsString;
                  connection.QryBusca2.ParamByName('Folio').AsString := Q_Partidas.FieldByName('sNumeroOrden').AsString;
                  connection.QryBusca2.Open;
                  tdAvanceAnterior := connection.QryBusca2.FieldByName('dAvanceAnterior').AsFloat;

                      {dDiaSiguiente := QryBitacora.FieldByName('dIdFecha').AsDateTime;
                      connection.QryBusca.SQL.Clear;
                      connection.QryBusca.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 4) AS dAvanceAnterior;';
                      connection.QryBusca.ParamByName('Orden').AsString := QryBitacora.FieldByName('sContrato').AsString;
                      connection.QryBusca.ParamByName('Folio').AsString := QryBitacora.FieldByName('sNumeroOrden').AsString;
                      connection.QryBusca.Open;}

                  while not Q_Partidas.Eof do
                      begin
                          connection.QryBusca2.Active := False;
                          connection.QryBusca2.SQL.Clear;
                          connection.QryBusca2.SQL.Add( 'select b.*, o.sidfolio, o.mDescripcion, a.dPonderado, ' +
                          '( SELECT (ifnull(sum(ba.dAvance), 0)) '+
                          'FROM bitacoradeactividades AS ba WHERE '+
                          'ba.sContrato = b.sContrato '+
                          'AND ba.sNumeroOrden = b.sNumeroOrden '+
                          'AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento '+
                          'AND ba.swbs = b.swbs '+
                          'AND ba.sNumeroActividad = b.sNumeroActividad '+
                          'AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
                          '< cast(b.sHoraInicio AS Time))  )	) AS AvanceAnterior '+
                          'from ordenesdetrabajo o inner join bitacoradeactividades b ON ' +
                          '(o.scontrato = b.scontrato and o.snumeroorden = b.snumeroorden ) '+
                          'inner join actividadesxorden a ON '+
                          '(b.scontrato = a.scontrato and b.snumeroorden = a.snumeroorden and '+
                          'b.snumeroactividad = a.snumeroactividad) '+
                          'where o.sidfolio=:Folio and b.scontrato = :Contrato and b.didfecha = :Fecha '+
                          'Order by b.snumeroactividad, b.didfecha, time(b.sHoraInicio)');
                          connection.QryBusca2.ParamByName('Folio').AsString := connection.QryBusca.FieldValues['sidfolio'];
                          connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato;
                          connection.QryBusca2.ParamByName('Fecha').AsDate := Q_Partidas.FieldValues['dIdFecha'];
                          connection.QryBusca2.Open;

                          nAcumAvan:=0;
                          while not connection.QryBusca2.Eof do
                              begin
                                  //Recopilación de fechas, subtraer solo numerico el dia del campo didfecha y compararlo con el array
                                  DecodeDate( connection.QryBusca2.FieldValues['didfecha'], myYear, myMonth, myDay );

                                  nTmp := connection.QryBusca2.FieldValues['dAvanceActual'];

                                  if nTmp > 0 then
                                      begin
                                        aDays[myDay] := nTmp;
                                        Break;
                                      end
                                  else
                                      begin
                                        nTmp := ((connection.QryBusca2.FieldValues['dPonderado'] * connection.QryBusca2.FieldValues['davance'])/100);
                                        aDays[myDay] := aDays[myDay] + nTmp;
                                      end;

                                  connection.QryBusca2.Next;

                              end;

                              Q_Partidas.Next;

                      end;

                end;

              //if Q_partidas.RecordCount > 0 then
                //begin

                //Resultados de la recopilacion en QryBusca2
                      Hoja.Range['A'+IntTostr(Ren)+':A'+IntToStr(Ren)].Select;
                      Excel.Selection.RowHeight := '60';

                      Hoja.Range['B'+IntTostr(Ren)+':AJ'+IntToStr(Ren)].Select;
                      Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                      Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
                      Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                      Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
                      Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                      Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
                      Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                      Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
                      Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous;
                      Excel.Selection.Borders[xlInsideVertical].Weight    := xlThin;

                      {Escritura de Datos en el Archvio de Excel..}
                      Hoja.Cells[Ren, 2].Select;
                      Excel.Selection.Value := connection.QryBusca2.FieldValues['sidfolio'];
                      Excel.Selection.HorizontalAlignment := xlCenter;
                      Excel.Selection.VerticalAlignment   := xlCenter;
                      Excel.Selection.Font.Size := 10;
                      Excel.Selection.Font.Bold := False;
                      Excel.Selection.Font.Name := 'Tahoma';

                      Hoja.Cells[Ren, 3].Select;
                      Excel.Selection.Value := connection.QryBusca2.FieldValues['mdescripcion'];
                      Excel.Selection.HorizontalAlignment := xlCenter;
                      Excel.Selection.VerticalAlignment   := xlCenter;
                      Excel.Selection.Font.Size := 08;
                      Excel.Selection.Font.Bold := False;
                      Excel.Selection.Font.Name := 'Tahoma';

                      Hoja.Cells[Ren, 4].Select;
                      //Excel.Selection.Value := connection.QryBusca2.FieldValues['davanceanterior'];
                      //Excel.Selection.Value := connection.QryBusca2.FieldValues['AvanceAnterior'];
                      Excel.Selection.Value := tdAvanceAnterior;
                      Excel.Selection.HorizontalAlignment := xlCenter;
                      Excel.Selection.VerticalAlignment   := xlCenter;
                      Excel.Selection.Font.Size := 11;
                      Excel.Selection.Font.Bold := False;
                      Excel.Selection.Font.Name := 'Tahoma';

                      nivel := 1;
                      begin
                            for i := 5 to 35 do
                              begin
                                  Hoja.Cells[Ren, i].Select;
                                  Excel.Selection.Value := formatfloat('00.##', aDays[nivel])+'%';
                                  nAcumAvan:=nAcumAvan+aDays[nivel];
                                  Excel.Selection.HorizontalAlignment := xlCenter;
                                  Excel.Selection.VerticalAlignment   := xlCenter;
                                  Excel.Selection.Font.Size := 11;
                                  Excel.Selection.Font.Bold := False;
                                  Excel.Selection.Font.Name := 'Tahoma';
                                  nivel:=nivel+1;
                              end;

                              //Imprimir una celda manual
                              Hoja.Cells[Ren, 36].Select;

                              if nAcumAvan >= 100 then
                                  begin
                                    Excel.Selection.Value := '100.00%';
                                  end
                              else
                                  begin
                                      Excel.Selection.Value := formatfloat('00.##', nAcumAvan)+'%';
                                  end;
                              Excel.Selection.HorizontalAlignment := xlCenter;
                              Excel.Selection.VerticalAlignment   := xlCenter;
                              Excel.Selection.Font.Size := 11;
                              Excel.Selection.Font.Bold := False;
                              Excel.Selection.Font.Name := 'Tahoma';

                              Inc(Ren);
                      end;

                      for i := 1 to 31 do
                      begin
                          aDays[i] := 0;
                      end;

                      nAcumAvan:=0;

              //end;

              connection.QryBusca.Next;

              inc(Total);

          end;
          Hoja.Range['B5:B5'].Select;

      end;

    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'AVANCE POR FOLIO';
      except
        Hoja.Name := 'AVANCE POR FOLIO';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Estatus de Ordenes:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
      end;
    end;

    Result := Resultado;
  end;

begin
    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
    if not SaveDialog1.Execute then
      Exit;
      // Generar el ambiente de excel
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
    end
    else
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;
    end;

    PanelProgress.Visible := True;
    Label15.Refresh;
    BarraEstado.Position := 0;

    Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

      // Verificar si cuenta con las hojas necesarias
    while Libro.Sheets.Count < 2 do
      Libro.Sheets.Add;

      // Verificar si se pasa de hojas necesarias
    Libro.Sheets[1].Select;
    while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

      // Proceder a generar la hoja REPORTE
    CadError := '';

    if GenerarPlantilla then
    begin
          // Grabar el archivo de excel con el nombre dado
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      PanelProgress.Visible := False;
      messageDlg('El Archivo se Generó Correctamente!', mtInformation, [mbOk], 0);
    end;

    Excel := '';

    if CadError <> '' then
      showmessage(CadError);
end;


//=================================================================================================================New buttom

procedure TfrmCompara.btnRptAvanXFolAcumClick(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// AVANCES X FOLIO ACUMULADO MOJICA JUNIO 27 2016 //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena, sConvenio: string;
      fs: tStream;
      Alto: Extended;
      Ren, nivel, i, total, nVald, nz, nDaysMonth : integer;
      Q_Partidas: TZReadOnlyQuery;
      dAvanceProg, dAvanceFisico: double;
      Progreso, TotalProgreso: real;

    const
      aColumExcel : array[1..33] of string = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG');

    var aDays: array[1..31] of Double;
      nTmp: Double;
      nAcumAvan, nAcum2Avan: Double;
      myYear, myMonth, myDay : Word;
      tdAvanceAnterior: Double;

    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      //Consultamos.
      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;
      {
      connection.QryBusca.SQL.Add( 'select b.scontrato, o.sidfolio, b.davanceanterior, b.didfecha, b.davance, b.mDescripcion,' +
      'b.dcantidad, b.snumeroorden, b.sidconvenio,b.snumeroactividad, o.mdescripcion as ordenDesc ' +
      'from ordenesdetrabajo o inner join bitacoradeactividades b ON ' +
      '(o.scontrato = b.scontrato and o.snumeroorden = b.snumeroorden ) ' +
      'where b.scontrato = :Contrato and b.didfecha >= :FechaI and b.didfecha <= :FechaF and b.dAvance > 0 Group by o.sidfolio ' +
      'Order by b.scontrato, o.sidfolio, b.didfecha, b.davance');
      connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
      connection.QryBusca.ParamByName('FechaI').AsDate := tdIdFecha1.Date;
      connection.QryBusca.ParamByName('FechaF').AsDate := tdIdFecha.Date;
      connection.QryBusca.Open;
      }
     If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
        Begin
          connection.QryBusca.SQL.Add( 'select sContrato, sidFolio, sNumeroOrden, sIdPlataforma, mDescripcion, sCsu from '+
          'ordenesdetrabajo where sContrato = :Contrato');
          connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
          End
      Else
        Begin
          connection.QryBusca.SQL.Add( 'select sContrato, sidFolio, sNumeroOrden, sIdPlataforma, mDescripcion, sCsu from '+
          'ordenesdetrabajo where sContrato = :Contrato and sNumeroOrden = :Orden');
          connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
          connection.QryBusca.ParamByName('Orden').Value       := tsNumeroOrden.Text;
      End;

      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then

          nDaysMonth := DayOfTheMonth( tdIdFecha.Date);

          Ren := 2;
          Excel.ActiveWindow.Zoom := 100;
          Excel.Columns['A:A'].ColumnWidth := 5.86;
          Excel.Columns['B:B'].ColumnWidth := 52.29;

          for i := 3 to nDaysMonth+2 do
          begin
              Excel.Columns[  aColumExcel[i]+':'+aColumExcel[i] ].ColumnWidth := 12.14;
          end;

          Hoja.Range['A1:A2'].Select;
          Excel.Selection.RowHeight := '15';

          Hoja.Range['B'+IntTostr(Ren-1)+':B'+IntToStr(Ren-1)].Select;
          Excel.Selection.Value := 'CONTROL DE AVANCES POR FOLIO ACUMULADO DEL  '+ DateToStr(tdIdFecha1.Date)+'  AL  ' + DateToStr(tdIdFecha.Date);
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 16;
          Excel.Selection.Font.color:= clBlack;
          Excel.Selection.Font.Name := 'Calibri';

          //Hoja.Range['B'+IntTostr(Ren-1)+':AG'+IntToStr(Ren)].Select;
          Hoja.Range['B'+IntTostr(Ren-1)+':'+aColumExcel[nDaysMonth+2]+IntToStr(Ren)].Select;
          Excel.Selection.Interior.ColorIndex := 15;
          Excel.Selection.MergeCells:= True;
          Excel.Selection.WrapText  := True;

          inc(Ren);
          Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FECHA DE IMPRESIÓN AL DÍA: '+ DateToStr(now);
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Name := 'Calibri';
          Excel.Selection.RowHeight := '20.25';

          inc(Ren);
          //Hoja.Range['B4:AG4'].Select;
          Hoja.Range['B4:'+aColumExcel[nDaysMonth+2]+'4'].Select;

          Excel.Selection.RowHeight := '41.25';
          Excel.Selection.Interior.ColorIndex := 36;

          //Colocar los encabezados de la plantilla...
          Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FOLIO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 10;

          nivel:=1;
          for i := 3 to nDaysMonth+2 do
          begin
              Hoja.Range[aColumExcel[i] + IntTostr(Ren)+':' + aColumExcel[nDaysMonth+2] + IntToStr(Ren)].Select;
              Excel.Selection.Value := IntToStr(nivel);
              Excel.Selection.Font.Italic := True;
              FormatoEncabezado;
              Excel.Selection.Font.Size := 10;
              nivel:=nivel+1;
          end;

          //Redimensionar vector dinamico al no. máximo de dias en un mes
          for i := 1 to nDaysMonth do
          begin
              aDays[i] := 0;
          end;

          inc(Ren);

      begin
          total := 1;

          while not connection.QryBusca.Eof do
          begin
              {Movimiento de la Barra..}
              Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
              TotalProgreso := TotalProgreso + Progreso;
              BarraEstado.Position := Trunc(TotalProgreso);

              connection.QryBusca2.Active := False;
              connection.QryBusca2.SQL.Clear;
              connection.QryBusca2.SQL.Add('select b.scontrato, o.sidfolio, b.davanceanterior, b.didfecha, b.davance,'+
              'b.mDescripcion, b.dcantidad, b.snumeroorden, b.sidconvenio, b.snumeroactividad, o.sidfolio,'+
              'o.mdescripcion as ordenDesc, a.dPonderado ' +
              'from ordenesdetrabajo o inner join bitacoradeactividades b ON ' +
              '(o.scontrato = b.scontrato and o.snumeroorden = b.snumeroorden ) ' +
              'inner join actividadesxorden a ON '+
              '(b.scontrato = a.scontrato and b.snumeroorden = a.snumeroorden and '+
              'b.snumeroactividad = a.snumeroactividad) ' +
              'where b.scontrato = :Contrato and b.didfecha >= :FechaI and b.didfecha <= :FechaF and ' +
              'o.sidfolio=:Folio ' +
              'Order by b.scontrato, o.sidfolio, b.didfecha, b.davance');
              connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato;
              connection.QryBusca2.ParamByName('FechaI').AsDate := tdIdFecha1.Date;
              connection.QryBusca2.ParamByName('FechaF').AsDate := tdIdFecha.Date;
              connection.QryBusca2.ParamByName('Folio').AsString := connection.QryBusca.FieldValues['sidfolio'];
              connection.QryBusca2.Open;

          if connection.QryBusca2.recordcount > 0 then

              begin
                  Q_Partidas.Active := False;
                  Q_Partidas.SQL.Clear;
                  Q_Partidas.SQL.Add( 'select b.*, o.sidfolio, o.mDescripcion ' +
                  'from ordenesdetrabajo o inner join bitacoradeactividades b ON ' +
                  '(o.scontrato = b.scontrato and o.snumeroorden = b.snumeroorden ) '+
                  'where o.sidfolio=:Folio and b.scontrato = :Contrato and b.didfecha >= :FechaI and '+
                  'b.didfecha <= :FechaF Group By b.didfecha '+
                  'Order by b.snumeroactividad, b.didfecha, time(b.sHoraInicio)');
                  Q_Partidas.ParamByName('Folio').AsString := connection.QryBusca.FieldValues['sidfolio'];
                  Q_Partidas.ParamByName('Contrato').AsString := global_contrato;
                  Q_Partidas.ParamByName('FechaI').AsDate := tdIdFecha1.Date;
                  Q_Partidas.ParamByName('FechaF').AsDate := tdIdFecha.Date;
                  Q_Partidas.Open;

                  {Avances anteriores}
                  connection.QryBusca2.SQL.Clear;
                  connection.QryBusca2.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd',
                  Q_Partidas.FieldByName('dIdFecha').AsDateTime)+'", :Orden, :Folio), 2) AS dAvanceAnterior;';
                  connection.QryBusca2.ParamByName('Orden').AsString := Q_Partidas.FieldByName('sContrato').AsString;
                  connection.QryBusca2.ParamByName('Folio').AsString := Q_Partidas.FieldByName('sNumeroOrden').AsString;
                  connection.QryBusca2.Open;
                  tdAvanceAnterior := connection.QryBusca2.FieldByName('dAvanceAnterior').AsFloat;

                      {dDiaSiguiente := QryBitacora.FieldByName('dIdFecha').AsDateTime;
                      connection.QryBusca.SQL.Clear;
                      connection.QryBusca.SQL.Text := 'SELECT ROUND(AvancesAnteriores("'+FormatDateTime('yyyy-mm-dd', dDiaSiguiente)+'", :Orden, :Folio), 4) AS dAvanceAnterior;';
                      connection.QryBusca.ParamByName('Orden').AsString := QryBitacora.FieldByName('sContrato').AsString;
                      connection.QryBusca.ParamByName('Folio').AsString := QryBitacora.FieldByName('sNumeroOrden').AsString;
                      connection.QryBusca.Open;}

                  while not Q_Partidas.Eof do
                      begin
                          connection.QryBusca2.Active := False;
                          connection.QryBusca2.SQL.Clear;
                          connection.QryBusca2.SQL.Add( 'select b.*, o.sidfolio, o.mDescripcion, a.dPonderado ' +
                          'from ordenesdetrabajo o inner join bitacoradeactividades b ON ' +
                          '(o.scontrato = b.scontrato and o.snumeroorden = b.snumeroorden ) '+
                          'inner join actividadesxorden a ON '+
                          '(b.scontrato = a.scontrato and b.snumeroorden = a.snumeroorden and '+
                          'b.snumeroactividad = a.snumeroactividad) '+
                          'where o.sidfolio=:Folio and b.scontrato = :Contrato and b.didfecha = :Fecha '+
                          'Order by b.snumeroactividad, b.didfecha, time(b.sHoraInicio)');
                          connection.QryBusca2.ParamByName('Folio').AsString := connection.QryBusca.FieldValues['sidfolio'];
                          connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato;
                          connection.QryBusca2.ParamByName('Fecha').AsDate := Q_Partidas.FieldValues['dIdFecha'];
                          connection.QryBusca2.Open;

                          nAcumAvan:=0;
                          while not connection.QryBusca2.Eof do
                              begin
                                  //Recopilación de fechas, subtraer solo numerico el dia del campo didfecha y compararlo con el array
                                  DecodeDate( connection.QryBusca2.FieldValues['didfecha'], myYear, myMonth, myDay );

                                  nTmp := connection.QryBusca2.FieldValues['dAvanceActual'];

                                  if nTmp > 0 then
                                      begin
                                        aDays[myDay] := nTmp;
                                        Break;
                                      end
                                  else
                                      begin
                                        nTmp := ((connection.QryBusca2.FieldValues['dPonderado'] * connection.QryBusca2.FieldValues['davance'])/100);
                                        aDays[myDay] := aDays[myDay] + nTmp;
                                      end;

                                  connection.QryBusca2.Next;

                              end;

                              Q_Partidas.Next;

                      end;
              //end;

              //Resultados de la recopilacion en QryBusca2
                    Hoja.Range['A'+IntTostr(Ren)+':A'+IntToStr(Ren)].Select;
                    Excel.Selection.RowHeight := '60';

                    //Hoja.Range['B'+IntTostr(Ren)+':AG'+IntToStr(Ren)].Select;
                    Hoja.Range['B'+IntTostr(Ren)+':'+aColumExcel[nDaysMonth+2]+IntToStr(Ren)].Select;

                    Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                    Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
                    Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                    Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
                    Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                    Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
                    Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                    Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
                    Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous;
                    Excel.Selection.Borders[xlInsideVertical].Weight    := xlThin;

                    {Escritura de Datos en el Archvio de Excel..}
                    Hoja.Cells[Ren, 2].Select;
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['sidfolio'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    nivel := 1;

                    for i := 3 to nDaysMonth+2 do
                            begin

                                Hoja.Cells[Ren, i].Select;

                                nAcumAvan:=nAcumAvan+aDays[nivel];

                                nVald:=0;
                                nz:=0;

                                if nAcumAvan > 0 then
                                  begin

                                      nVald := nivel;
                                      nAcum2Avan:=0;

                                      for nz := nVald to nDaysMonth do
                                          begin
                                            //nAcum2Avan:=nAcum2Avan + aDays[nz];
                                          end;

                                      if nAcum2Avan = 0 then
                                          begin
                                          if nAcumAvan >= 100  then
                                             begin
                                                //Excel.Selection.Value := '>';  // Salto

                                                //Hoja.Cells[Ren, i-1].Select;
                                                Hoja.Cells[Ren, i].Select;
                                                Excel.Selection.Value := '100.00';  // Salto
                                                Excel.Selection.HorizontalAlignment := xlCenter;
                                                Excel.Selection.VerticalAlignment   := xlCenter;
                                                Excel.Selection.Font.Size := 10;
                                                Excel.Selection.Font.Bold := False;
                                                Excel.Selection.Font.Name := 'Tahoma';
                                                Excel.Selection.Interior.ColorIndex := 36;

                                                Break;

                                             end
                                           else
                                              begin
                                                //Excel.Selection.Value := FloatToStr(nAcumAvan) + '>';
                                                Excel.Selection.Value := formatfloat('00.##', nAcumAvan);
                                              end
                                           end

                                      else
                                          begin

                                            if i = nDaysMonth+2 then
                                                begin
                                                if nAcumAvan >= 100  then
                                                    begin
                                                      Excel.Selection.Value := '100.00';
                                                      Excel.Selection.Interior.ColorIndex := 36;
                                                    end
                                                else
                                                    begin
                                                      Excel.Selection.Value := FloatToStr(nAcumAvan) + '*';
                                                    end;
                                                end
                                            else
                                              begin
                                                  Excel.Selection.Value := FloatToStr(nAcumAvan) + '**';
                                              end;

                                            end;

                                          Excel.Selection.HorizontalAlignment := xlCenter;
                                          Excel.Selection.VerticalAlignment   := xlCenter;
                                          Excel.Selection.Font.Size := 10;
                                          Excel.Selection.Font.Bold := False;
                                          Excel.Selection.Font.Name := 'Tahoma';

                                          end
                                    else
                                        begin
                                          Excel.Selection.Value := formatfloat('00.##', nAcumAvan);
                                          Excel.Selection.HorizontalAlignment := xlCenter;
                                          Excel.Selection.VerticalAlignment   := xlCenter;
                                          Excel.Selection.Font.Size := 10;
                                          Excel.Selection.Font.Bold := False;
                                          Excel.Selection.Font.Name := 'Tahoma';
                                        end;

                                    nivel:=nivel+1;

                                end;

                                {
                                //Mientras
                                Excel.Selection.Value := aDays[nivel];
                                Excel.Selection.HorizontalAlignment := xlCenter;
                                Excel.Selection.VerticalAlignment   := xlCenter;
                                Excel.Selection.Font.Size := 10;
                                Excel.Selection.Font.Bold := False;
                                Excel.Selection.Font.Name := 'Tahoma';
                                }

                            Inc(Ren);

                    for i := 1 to nDaysMonth do
                    begin
                        aDays[i] := 0;
                    end;
                    nAcumAvan:=0;
           end;

           connection.QryBusca.Next;

           inc(Total);

          end;

          Hoja.Range['B5:B5'].Select;

      end;
  end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'AVANCE ACUNULADO POR FOLIO';
      except
        Hoja.Name := 'AVANCE ACUNULADO POR FOLIO';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Estatus de Ordenes:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
      end;
    end;

    Result := Resultado;
  end;

begin
    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
    if not SaveDialog1.Execute then
      Exit;
      // Generar el ambiente de excel
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
    end
    else
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;
    end;

    PanelProgress.Visible := True;
    Label15.Refresh;
    BarraEstado.Position := 0;

    Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

      // Verificar si cuenta con las hojas necesarias
    while Libro.Sheets.Count < 2 do
      Libro.Sheets.Add;

      // Verificar si se pasa de hojas necesarias
    Libro.Sheets[1].Select;
    while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

      // Proceder a generar la hoja REPORTE
    CadError := '';

    if GenerarPlantilla then
    begin
          // Grabar el archivo de excel con el nombre dado
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      PanelProgress.Visible := False;
      messageDlg('El Archivo se Generó Correctamente!', mtInformation, [mbOk], 0);
    end;

    Excel := '';

    if CadError <> '' then
      showmessage(CadError);
end;


procedure TfrmCompara.btnStaFoliosOTClick(Sender: TObject);
var
  CadError, OrdenVigencia: string;
//////////////////////////////////// AVANCES ESTATUS DE FOLIO GLOBAL MOJICA JUNIO 2016 //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena, sConvenio: string;
      fs: tStream;
      Alto: Extended;
      Ren, i, total: integer;
      Q_Partidas: TZReadOnlyQuery;
      dAvanceProg, dAvanceFisico: double;
      Progreso, TotalProgreso: real;

      nAcumAvan: Double;
      myYear, myMonth, myDay : Word;

    var aFields: array[1..7] of String;
      cContrato, cOrden, cTurno: String;
      dDateRptDiario: TDate;

    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      //Consultamos.
      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;

      If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
          Begin
          connection.QryBusca.SQL.Add( 'select sContrato, sidFolio, sNumeroOrden, sIdPlataforma, mDescripcion, sCsu from '+
          'ordenesdetrabajo where sContrato = :Contrato');
          connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
          End
      Else
        Begin
          connection.QryBusca.SQL.Add( 'select sContrato, sidFolio, sNumeroOrden, sIdPlataforma, mDescripcion, sCsu from '+
          'ordenesdetrabajo where sContrato = :Contrato and sNumeroOrden = :Orden');
          connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
          connection.QryBusca.ParamByName('Orden').Value       := tsNumeroOrden.Text;
      End;

      connection.QryBusca.Open;

      if connection.QryBusca.RecordCount > 0 then

          Ren := 2;

          Excel.ActiveWindow.Zoom := 100;
          Excel.Columns['A:A'].ColumnWidth := 5.86;
          Excel.Columns['B:B'].ColumnWidth := 52.29;
          Excel.Columns['C:C'].ColumnWidth := 14.00;
          Excel.Columns['D:D'].ColumnWidth := 60.29;
          Excel.Columns['E:E'].ColumnWidth := 20.00;
          Excel.Columns['F:F'].ColumnWidth := 18.14;
          Excel.Columns['G:G'].ColumnWidth := 20.00;
          Excel.Columns['H:H'].ColumnWidth := 12.14;
          Excel.Columns['I:I'].ColumnWidth := 18.14;
          Excel.Columns['J:J'].ColumnWidth := 10.14;
          Excel.Columns['K:K'].ColumnWidth := 30.14;
          Excel.Columns['L:L'].ColumnWidth := 30.14;
          Excel.Columns['M:M'].ColumnWidth := 30.14;
          Excel.Columns['N:N'].ColumnWidth := 30.14;
          Excel.Columns['O:O'].ColumnWidth := 30.14;

          Hoja.Range['A1:A2'].Select;
          Excel.Selection.RowHeight := '15';

          Hoja.Range['B'+IntTostr(Ren-1)+':B'+IntToStr(Ren-1)].Select;
          //Excel.Selection.Value := 'ESTATUS DE FOLIOS GLOBAL '+ DateToStr(tdIdFecha1.Date)+' AL ' + DateToStr(tdIdFecha.Date);
          Excel.Selection.Value := 'ESTATUS DE FOLIOS GLOBAL';
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 16;
          Excel.Selection.Font.color:= clBlack;
          Excel.Selection.Font.Name := 'Calibri';

          Hoja.Range['B'+IntTostr(Ren-1)+':O'+IntToStr(Ren)].Select;
          Excel.Selection.Interior.ColorIndex := 15;
          Excel.Selection.MergeCells:= True;
          Excel.Selection.WrapText  := True;

          inc(Ren);
          Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FECHA DE IMPRESIÓN AL DÍA: '+ DateToStr(now);
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Name := 'Calibri';
          Excel.Selection.RowHeight := '20.25';

          inc(Ren);
          Hoja.Range['B4:O4'].Select;

          Excel.Selection.RowHeight := '41.25';
          Excel.Selection.Interior.ColorIndex := 36;

          //Colocar los encabezados de la plantilla...
          Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FOLIO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'INSTALACION';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'DESCRIPCION';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;
          //cel.Selection.RowHeight := '100';

          Hoja.Range['E'+IntTostr(Ren)+':E'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FECHA INICIO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['F'+IntTostr(Ren)+':F'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'HORA. INICIO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['G'+IntTostr(Ren)+':G'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FECHA TERMINO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['H'+IntTostr(Ren)+':H'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'HORA TERMINO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['I'+IntTostr(Ren)+':I'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'AVANCE ACTUAL';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['J'+IntTostr(Ren)+':J'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'CSU';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['K'+IntTostr(Ren)+':K'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'REPTTE. PEMEX INICIO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['L'+IntTostr(Ren)+':L'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'REPTTE. PEMEX TERMINO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['M'+IntTostr(Ren)+':M'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'REPTTE. CIA. INICIO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['N'+IntTostr(Ren)+':N'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'REPTTE. CIA. TERMINO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['O'+IntTostr(Ren)+':O'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'OBSERVACIONES';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          inc(Ren);

      begin

          aFields[1] := '';
          aFields[2] := '';
          aFields[3] := '';
          aFields[4] := '';
          aFields[5] := '';
          aFields[6] := '';
          aFields[7] := '';

          total := 1;

          while not connection.QryBusca.Eof do
          begin
              {Movimiento de la Barra..}
              Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
              TotalProgreso := TotalProgreso + Progreso;
              BarraEstado.Position := Trunc(TotalProgreso);

              connection.QryBusca2.Active := False;
              connection.QryBusca2.SQL.Clear;
              connection.QryBusca2.SQL.Add('select b.sContrato, b.sNumeroOrden, b.dIdFecha, b.sHoraInicio, max(b.sHoraFinal) as sHoraFinal,' +
              'max(b.dAvance) as dAvance from bitacoradeactividades b '+
              'where b.sContrato = :Contrato and ' +
              'b.sNumeroOrden = :NoOrden And b.sIdTipoMovimiento="ED" Group by dIdFecha');
              connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato;
              connection.QryBusca2.ParamByName('NoOrden').AsString := connection.QryBusca.FieldValues['sNumeroOrden'];
              connection.QryBusca2.Open;

              //while not connection.QryBusca2.Eof do
              //begin

                  if connection.QryBusca2.RecordCount = 1 then
                    begin

                      //Buscamos en reporte diario los datos
                      Q_Partidas.Active := False;
                      Q_Partidas.SQL.Clear;
                      Q_Partidas.SQL.Add('select dIdFecha, sOrden, sNumeroOrden, sIdTurno from reportediario where '+
                      'sOrden =:sOrden and sNumeroOrden =:NoOrden and dIdFecha = :fecha');
                      Q_Partidas.ParamByName('sOrden').AsString    := connection.QryBusca2.FieldValues['sContrato'];
                      Q_Partidas.ParamByName('NoOrden').AsString    := connection.QryBusca2.FieldValues['sContrato'];
                      Q_Partidas.ParamByName('fecha').AsDate      := connection.QryBusca2.FieldValues['didFecha'];
                      Q_Partidas.Open;

                      if Q_partidas.RecordCount > 0 then
                        begin
                              cOrden := Q_Partidas.FieldValues['sNumeroOrden'];
                              cTurno := Q_Partidas.FieldValues['sIdTurno'];
                              dDateRptDiario := Q_Partidas.FieldValues['didfecha'];

                              //Buscamos en firmas
                              Q_Partidas.Active := False;
                              Q_Partidas.SQL.Clear;
                              Q_Partidas.SQL.Add('Select * from firmas where sContrato = :Contrato and sNumeroOrden = :Orden and '+
                              'sIdTurno = :Turno And dIdFecha <= :Fecha Order By dIdFecha DESC');
                              Q_Partidas.ParamByName('Contrato').AsString    := cOrden;
                              Q_Partidas.ParamByName('Orden').AsString    := cOrden;
                              Q_Partidas.ParamByName('Turno').AsString    := cTurno;
                              Q_Partidas.ParamByName('Fecha').AsDate      := dDateRptDiario;
                              Q_Partidas.Open;

                              aFields[1] := formatdatetime('dddddd', connection.QryBusca2.FieldValues['didFecha'] );
                              aFields[2] := connection.QryBusca2.FieldValues['sHoraInicio'];

                              if Q_Partidas.RecordCount > 0 then
                                  begin
                                    aFields[3] := Q_Partidas.FieldValues['sFirmante1'];
                                    aFields[4] := Q_Partidas.FieldValues['sFirmante5'];
                                  end
                              else
                                  begin
                                    aFields[3] := '';
                                    aFields[4] := '';
                                  end;

                              aFields[5] := formatdatetime('dddddd', connection.QryBusca2.FieldValues['dIdFecha'] );
                              aFields[6] := connection.QryBusca2.FieldValues['sHoraFinal'];
                              aFields[7] := connection.QryBusca2.FieldValues['dAvance'];


                        end;

                  end;



              if connection.QryBusca2.RecordCount >= 2 then

                  for i := 1 to 2 do
                    begin
                        if i=1 then
                            begin

                              //Buscamos en reporte diario los datos
                              Q_Partidas.Active := False;
                              Q_Partidas.SQL.Clear;
                              Q_Partidas.SQL.Add('select dIdFecha, sOrden, sNumeroOrden, sIdTurno from reportediario where '+
                              'sOrden =:sOrden and sNumeroOrden =:NoOrden and dIdFecha = :fecha');
                              Q_Partidas.ParamByName('sOrden').AsString    := connection.QryBusca2.FieldValues['sContrato'];
                              Q_Partidas.ParamByName('NoOrden').AsString    := connection.QryBusca2.FieldValues['sContrato'];
                              Q_Partidas.ParamByName('fecha').AsDate      := connection.QryBusca2.FieldValues['didFecha'];
                              Q_Partidas.Open;

                              if Q_partidas.RecordCount > 0 then
                                  begin

                                      cOrden := Q_Partidas.FieldValues['sNumeroOrden'];
                                      cTurno := Q_Partidas.FieldValues['sIdTurno'];
                                      dDateRptDiario := Q_Partidas.FieldValues['didfecha'];

                                      //Buscamos en firmas
                                      Q_Partidas.Active := False;
                                      Q_Partidas.SQL.Clear;
                                      Q_Partidas.SQL.Add('Select * from firmas where sContrato = :Contrato and '+
                                      'sNumeroOrden = :Orden and '+
                                      'sIdTurno = :Turno And dIdFecha <= :Fecha Order By dIdFecha DESC');
                                      Q_Partidas.ParamByName('Contrato').AsString    := cOrden;
                                      Q_Partidas.ParamByName('Orden').AsString    := cOrden;
                                      Q_Partidas.ParamByName('Turno').AsString    := cTurno;
                                      Q_Partidas.ParamByName('Fecha').AsDate      := dDateRptDiario;
                                      Q_Partidas.Open;

                                      aFields[1] := formatdatetime('dddddd', connection.QryBusca2.FieldValues['didFecha'] );
                                      aFields[2] := connection.QryBusca2.FieldValues['sHoraInicio'];

                                      if Q_Partidas.RecordCount > 0 then
                                          begin
                                            aFields[3] := Q_Partidas.FieldValues['sFirmante1'];
                                          end
                                      else
                                          begin
                                            aFields[3] := '';
                                          end;

                                      aFields[5] := formatdatetime('dddddd', connection.QryBusca2.FieldValues['dIdFecha'] );
                                      aFields[6] := connection.QryBusca2.FieldValues['sHoraFinal'];
                                      aFields[7] := connection.QryBusca2.FieldValues['dAvance'];

                                  end;
                            end;
                        if i=2 then
                            begin

                              //Buscamos en reporte diario los datos
                              Q_Partidas.Active := False;
                              Q_Partidas.SQL.Clear;
                              Q_Partidas.SQL.Add('select dIdFecha, sOrden, sNumeroOrden, sIdTurno from reportediario where '+
                              'sOrden =:sOrden and sNumeroOrden =:NoOrden and dIdFecha = :fecha');
                              Q_Partidas.ParamByName('sOrden').AsString    := connection.QryBusca2.FieldValues['sContrato'];
                              Q_Partidas.ParamByName('NoOrden').AsString    := connection.QryBusca2.FieldValues['sContrato'];
                              Q_Partidas.ParamByName('fecha').AsDate      := connection.QryBusca2.FieldValues['didFecha'];
                              Q_Partidas.Open;

                              if Q_partidas.RecordCount > 0 then
                                  begin

                                      cOrden := Q_Partidas.FieldValues['sNumeroOrden'];
                                      cTurno := Q_Partidas.FieldValues['sIdTurno'];
                                      dDateRptDiario := Q_Partidas.FieldValues['didfecha'];

                                      //Buscamos en firmas
                                      Q_Partidas.Active := False;
                                      Q_Partidas.SQL.Clear;
                                      Q_Partidas.SQL.Add('Select * from firmas where sContrato = :Contrato and '+
                                      'sNumeroOrden = :Orden and '+
                                      'sIdTurno = :Turno And dIdFecha <= :Fecha Order By dIdFecha DESC');
                                      Q_Partidas.ParamByName('Contrato').AsString    := cOrden;
                                      Q_Partidas.ParamByName('Orden').AsString    := cOrden;
                                      Q_Partidas.ParamByName('Turno').AsString    := cTurno;
                                      Q_Partidas.ParamByName('Fecha').AsDate      := dDateRptDiario;
                                      Q_Partidas.Open;

                                      //aFields[1] := formatdatetime('dddddd', connection.QryBusca2.FieldValues['didFecha'] );
                                      //aFields[2] := connection.QryBusca2.FieldValues['sHoraInicio'];

                                      if Q_Partidas.RecordCount > 0 then
                                          begin
                                            aFields[4] := Q_Partidas.FieldValues['sFirmante5'];
                                          end
                                      else
                                          begin
                                            aFields[4] := '';
                                          end;

                                      aFields[5] := formatdatetime('dddddd', connection.QryBusca2.FieldValues['dIdFecha'] );
                                      aFields[6] := connection.QryBusca2.FieldValues['sHoraFinal'];
                                      aFields[7] := connection.QryBusca2.FieldValues['dAvance'];
                                  end;
                            end;

                        connection.QryBusca2.Next;

                    end;

              //end;

              //Resultados de la recopilacion en QryBusca2
                    Hoja.Range['A'+IntTostr(Ren)+':A'+IntToStr(Ren)].Select;
                    Excel.Selection.RowHeight := '60';

                    Hoja.Range['B'+IntTostr(Ren)+':O'+IntToStr(Ren)].Select;
                    Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
                    Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
                    Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
                    Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
                    Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
                    Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
                    Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
                    Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
                    Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous;
                    Excel.Selection.Borders[xlInsideVertical].Weight    := xlThin;

                    {Escritura de Datos en el Archvio de Excel..}
                    Hoja.Cells[Ren, 2].Select;
                    Excel.Selection.Value := connection.QryBusca.FieldValues['sidfolio'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 3].Select;
                    Excel.Selection.Value := connection.QryBusca.FieldValues['sIdPlataforma'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 4].Select;
                    Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 5].Select;
                    Excel.Selection.Value := aFields[1];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 6].Select;
                    Excel.Selection.Value := aFields[2];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 7].Select;
                    Excel.Selection.Value := aFields[5];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 8].Select;
                    Excel.Selection.Value := aFields[6];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 9].Select;
                    Excel.Selection.Value := aFields[7];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 10].Select;
                    Excel.Selection.Value := connection.QryBusca.FieldValues['sCsu'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 11].Select;
                    Excel.Selection.Value := aFields[3];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 12].Select;
                    Excel.Selection.Value := aFields[3];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 13].Select;
                    Excel.Selection.Value := aFields[4];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 14].Select;
                    Excel.Selection.Value := aFields[4];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 08;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Inc(Ren);


              connection.QryBusca.Next;

              inc(Total);

          end;
          Hoja.Range['B5:B5'].Select;

      end;

    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'ESTATUS DE FOLIO GLOBAL';
      except
        Hoja.Name := 'ESTATUS DE FOLIO GLOBAL';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Estatus de Ordenes:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
      end;
    end;

    Result := Resultado;
  end;

begin
    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
    if not SaveDialog1.Execute then
      Exit;
      // Generar el ambiente de excel
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
    end
    else
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;
    end;

    PanelProgress.Visible := True;
    Label15.Refresh;
    BarraEstado.Position := 0;

    Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

      // Verificar si cuenta con las hojas necesarias
    while Libro.Sheets.Count < 2 do
      Libro.Sheets.Add;

      // Verificar si se pasa de hojas necesarias
    Libro.Sheets[1].Select;
    while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

      // Proceder a generar la hoja REPORTE
    CadError := '';

    if GenerarPlantilla then
    begin
          // Grabar el archivo de excel con el nombre dado
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      PanelProgress.Visible := False;
      messageDlg('El Archivo se Generó Correctamente!', mtInformation, [mbOk], 0);
    end;

    Excel := '';

    if CadError <> '' then
      showmessage(CadError);
end;


procedure TfrmCompara.btnRptActFoliosClick(Sender: TObject);

var
  CadError, OrdenVigencia: string;
//////////////////////////////////// RPT ACTIVIDADES GRAL DE FOLIOS JUNIO 2016 //////////////////
  function GenerarPlantilla: Boolean;
  var
    Resultado: Boolean;

    procedure DatosPlantilla;
    var
      CadFecha, tmpNombre, cadena, sConvenio: string;
      fs: tStream;
      Alto: Extended;
      Ren, i, total: integer;
      Q_Partidas: TZReadOnlyQuery;
      dAvanceProg, dAvanceFisico: double;
      Progreso, TotalProgreso: real;

      myYear, myMonth, myDay : Word;

    begin
      Q_Partidas := TZReadOnlyQuery.Create(self);
      Q_Partidas.Connection := connection.zConnection;

      //Consultamos.
      connection.QryBusca.Active := False;
      connection.QryBusca.SQL.Clear;

      If MidStr(tsNumeroOrden.Text, 1 , 8) = 'CONTRATO' Then
          Begin
          connection.QryBusca.SQL.Add( 'select sContrato, sidFolio, sNumeroOrden, sIdPlataforma, mDescripcion, sCsu from '+
          'ordenesdetrabajo where sContrato = :Contrato');
          connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
          End
      Else
        Begin
          connection.QryBusca.SQL.Add( 'select sContrato, sidFolio, sNumeroOrden, sIdPlataforma, mDescripcion, sCsu from '+
          'ordenesdetrabajo where sContrato = :Contrato and sNumeroOrden = :Orden');
          connection.QryBusca.ParamByName('Contrato').AsString := global_contrato;
          connection.QryBusca.ParamByName('Orden').Value       := tsNumeroOrden.Text;
      End;

      connection.QryBusca.open;

      if connection.QryBusca.RecordCount > 0 then

          Ren := 2;

          Excel.ActiveWindow.Zoom := 100;
          Excel.Columns['A:A'].ColumnWidth := 5.86;
          Excel.Columns['B:B'].ColumnWidth := 12.14;
          Excel.Columns['C:C'].ColumnWidth := 60.29;
          Excel.Columns['D:D'].ColumnWidth := 10.00;
          Excel.Columns['E:E'].ColumnWidth := 60.29;
          Excel.Columns['F:F'].ColumnWidth := 5.86;
          Excel.Columns['G:G'].ColumnWidth := 5.86;
          Excel.Columns['H:H'].ColumnWidth := 12.14;
          Excel.Columns['I:I'].ColumnWidth := 12.14;
          Excel.Columns['J:J'].ColumnWidth := 12.14;
          Excel.Columns['K:K'].ColumnWidth := 12.14;
          Excel.Columns['L:L'].ColumnWidth := 13.00;

          Hoja.Range['A1:A2'].Select;
          Excel.Selection.RowHeight := '15';

          Hoja.Range['B'+IntTostr(Ren-1)+':B'+IntToStr(Ren-1)].Select;
          //Excel.Selection.Value := 'REPORTE DE ACTIVIDADE'+ DateToStr(tdIdFecha1.Date)+' AL ' + DateToStr(tdIdFecha.Date);
          Excel.Selection.Value := 'REPORTE DE ACTIVIDADES';
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 16;
          Excel.Selection.Font.color:= clBlack;
          Excel.Selection.Font.Name := 'Calibri';

          Hoja.Range['B'+IntTostr(Ren-1)+':L'+IntToStr(Ren)].Select;
          Excel.Selection.Interior.ColorIndex := 15;
          Excel.Selection.MergeCells:= True;
          Excel.Selection.WrapText  := True;

          inc(Ren);
          Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FECHA DE IMPRESIÓN AL DÍA: '+ DateToStr(now);
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.Font.Size := 12;
          Excel.Selection.Font.Name := 'Calibri';
          Excel.Selection.RowHeight := '20.25';

          inc(Ren);
          Hoja.Range['B4:L4'].Select;

          Excel.Selection.RowHeight := '41.25';
          Excel.Selection.Interior.ColorIndex := 36;

          //Colocar los encabezados de la plantilla...
          Hoja.Range['B'+IntTostr(Ren)+':B'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FECHA TERMINO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['C'+IntTostr(Ren)+':C'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'FOLIO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['D'+IntTostr(Ren)+':D'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'ID ACTIVIDAD';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;
          //cel.Selection.RowHeight := '100';

          Hoja.Range['E'+IntTostr(Ren)+':E'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'ACTIVIDAD';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['F'+IntTostr(Ren)+':F'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'TIPO NEC';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['G'+IntTostr(Ren)+':G'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'F.T.';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['H'+IntTostr(Ren)+':H'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'HORA INICIO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['I'+IntTostr(Ren)+':I'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'HORA TERMINO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['J'+IntTostr(Ren)+':J'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'AVANCE ANTERIOR';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['K'+IntTostr(Ren)+':K'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'AVANCE ACTUAL';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          Hoja.Range['L'+IntTostr(Ren)+':L'+IntToStr(Ren)].Select;
          Excel.Selection.Value := 'AVANCE ACUMULADO';
          Excel.Selection.Font.Italic := True;
          FormatoEncabezado;
          Excel.Selection.Font.Size := 11;

          inc(Ren);

      begin

          total := 1;

          while not connection.QryBusca.Eof do
          begin
              {Movimiento de la Barra..}
              Progreso := (1 / (connection.QryBusca.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
              TotalProgreso := TotalProgreso + Progreso;
              BarraEstado.Position := Trunc(TotalProgreso);

              connection.QryBusca2.Active := False;
              connection.QryBusca2.SQL.Clear;
              connection.QryBusca2.SQL.Add('select b.didfecha, o.sidfolio, b.sNumeroActividad, b.mDescripcion, '+
              'b.sidClasificacion, b.sHoraInicio, b.sHoraFinal, b.dAvanceAnterior, b.dAvanceActual, b.dAvance, '+
              '( SELECT (ifnull(sum(ba.dAvance), 0)) FROM bitacoradeactividades AS ba WHERE '+
              'ba.sContrato = b.sContrato '+
              'AND ba.sNumeroOrden = b.sNumeroOrden '+
              'AND ba.sIdTipoMovimiento = b.sIdTipoMovimiento '+
              'AND ba.swbs = b.swbs '+
              'AND ba.sNumeroActividad = b.sNumeroActividad '+
              'AND ( ba.didfecha < b.didfecha OR (ba.didfecha = b.didfecha AND cast(ba.sHoraInicio AS Time) '+
              '< cast(b.sHoraInicio AS Time))  )	) AS AvanceAnterior '+
              'from ordenesdetrabajo o inner join bitacoradeactividades b ON '+
              '(o.scontrato = b.scontrato and o.snumeroorden = b.snumeroorden) '+
              'where b.scontrato = :Contrato and o.sidfolio=:Folio '+
              'and b.sIdTipoMovimiento="ED" Order by b.sNumeroActividad, b.didfecha, time(b.sHoraInicio)');
              connection.QryBusca2.ParamByName('Contrato').AsString := global_contrato;
              connection.QryBusca2.ParamByName('Folio').AsString := connection.QryBusca.FieldValues['sidfolio'];
              //connection.QryBusca2.ParamByName('Orden').Value       := tsNumeroOrden.Text ;
              connection.QryBusca2.Open;

              while not connection.QryBusca2.Eof do
                begin
                    {Escritura de Datos en el Archvio de Excel..}
                    Hoja.Cells[Ren, 2].Select;
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['didfecha'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 3].Select;
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['sIdFolio'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 4].Select;
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['sNumeroActividad'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 5].Select;
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['mDescripcion'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 6].Select;
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['sidClasificacion'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 7].Select;
                    Excel.Selection.Value := ' ';
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 8].Select;
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['sHoraInicio'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 9].Select;
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['sHoraFinal'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 10].Select;
                    //Excel.Selection.Value := connection.QryBusca2.FieldValues['dAvanceAnterior'];
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['AvanceAnterior'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Hoja.Cells[Ren, 11].Select;
                    //Excel.Selection.Value := connection.QryBusca2.FieldValues['dAvanceActual'];
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['dAvance'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                     Hoja.Cells[Ren, 12].Select;
                    //Excel.Selection.Value := connection.QryBusca2.FieldValues['dAvanceActual'];
                    Excel.Selection.Value := connection.QryBusca2.FieldValues['AvanceAnterior']+connection.QryBusca2.FieldValues['dAvance'];
                    Excel.Selection.HorizontalAlignment := xlCenter;
                    Excel.Selection.VerticalAlignment   := xlCenter;
                    Excel.Selection.Font.Size := 10;
                    Excel.Selection.Font.Bold := False;
                    Excel.Selection.Font.Name := 'Tahoma';

                    Inc(Ren);

                    connection.QryBusca2.Next;

              end;

          connection.QryBusca.Next;

          inc(Total);

          end;

          Hoja.Range['B5:B5'].Select;

      end;

    end;

  begin
    Resultado := True;
    try
      Hoja := Libro.Sheets[1];
      Hoja.Select;
      try
        Hoja.Name := 'REPORTE DE ACTIVIDADES';
      except
        Hoja.Name := 'REPORTE DE ACTIVIDADES';
      end;
      Excel.ActiveWorkbook.SaveAs(SaveDialog1.FileName);
      DatosPlantilla;
    except
      on e: exception do
      begin
        Resultado := False;
        CadError := 'Se ha producido el siguiente error al Generar el Estatus de Ordenes:' + #10 + #10 + e.Message;
        PanelProgress.Visible := False;
      end;
    end;

    Result := Resultado;
  end;

begin
    // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
    if not SaveDialog1.Execute then
      Exit;
      // Generar el ambiente de excel
    try
      Excel := CreateOleObject('Excel.Application');
    except
      FreeAndNil(Excel);
      showmessage('No es posible generar el archivo de EXCEL, informe de esto al administrador del sistema.');
      Exit;
    end;

    if MessageDlg('Deseas visualizar el diseño del Archivo de Excel?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := True;
    end
    else
    begin
      Excel.Visible := True;
      Excel.DisplayAlerts := False;
      Excel.ScreenUpdating := False;
    end;

    PanelProgress.Visible := True;
    Label15.Refresh;
    BarraEstado.Position := 0;

    Libro := Excel.Workbooks.Add; // Crear el libro sobre el que se ha de trabajar

      // Verificar si cuenta con las hojas necesarias
    while Libro.Sheets.Count < 2 do
      Libro.Sheets.Add;

      // Verificar si se pasa de hojas necesarias
    Libro.Sheets[1].Select;
    while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

      // Proceder a generar la hoja REPORTE
    CadError := '';

    if GenerarPlantilla then
    begin
          // Grabar el archivo de excel con el nombre dado
      Excel.Visible := True;
      Excel.DisplayAlerts := True;
      Excel.ScreenUpdating := True;
      PanelProgress.Visible := False;
      messageDlg('El Archivo se Generó Correctamente!', mtInformation, [mbOk], 0);
    end;

    Excel := '';

    if CadError <> '' then
      showmessage(CadError);
end;



//*********************************BRITO 28-03-11*******************************
procedure TfrmCompara.ventasDiferentes(sWBSContrato, suma: string);
var
    sSQL: string;
    lError1, lError2: boolean;
begin
          sSQL := 'SELECT ' +
          'b.sNumeroActividad, b.sWbs, a.dCantidad, substr(b.mDescripcion,1,255) as descripcion, ' +
          'a.dVentaMN as aMN, a.dVentaDLL as aDLL, a.sTipoActividad, a.sNumeroOrden, a.sWbs as wbs2, ' +
          'b.dCantidadAnexo,  b.dVentaMN as bMN, b.dVentaDLL as bDLL  ' +
          'FROM actividadesxorden a ' +
          'INNER JOIN  actividadesxanexo b ' +
          'ON a.sContrato = b.sContrato ' +
          'AND a.sIdConvenio = b.sIdConvenio ' +
          'AND a.sWbsContrato = b.sWbs ' +
          'AND a.sTipoActividad = "Actividad" ' +
          'WHERE a.sContrato = :contrato ' +
          'AND a.sIdConvenio = :convenio ' +
          'AND a.sWbsContrato = :wbscontrato ' +
          'AND a.sTipoActividad = "Actividad" '+
          'ORDER BY b.sWbs';

          connection.QryBusca.Active := false;
          connection.QryBusca.SQL.Clear;
          connection.QryBusca.SQL.Add(sSQL);
          connection.QryBusca.ParamByName('wbscontrato').Value := sWBSContrato;
          connection.QryBusca.ParamByName('contrato').Value := global_contrato;
          connection.QryBusca.ParamByName('convenio').Value := global_convenio;
          connection.QryBusca.Open;

          lError1 := false;
          lError2 := false;
          while not connection.QryBusca.Eof do begin
              if (connection.QryBusca.FieldByName('aMN').Value <>
                connection.QryBusca.FieldByName('bMN').Value)
              or (connection.QryBusca.FieldByName('aDLL').Value <>
                connection.QryBusca.FieldByName('bDLL').Value) then begin
                  acumularDiferencia(suma, 'Existe diferencia entre los valores de ventas');
                  lError1 := true;
              end
              else begin
                  if (not lError1) and (not lError2) then begin
                      if (connection.QryBusca.FieldByName('dCantidadAnexo').Value <> suma) then
                          lError2 := true;
                  end;
              end;
              connection.QryBusca.Next;
          end;
          if (not lError1) and (lError2) then begin
              connection.QryBusca.First;
              while not connection.QryBusca.Eof do begin
                  acumularDiferencia(suma, 'Existe diferencia entre la suma total de las cantidades y la cantidad del anexo');
                  connection.QryBusca.Next;
              end;
          end;
end;



function TfrmCompara.cantidadesDiferentes(sWBSContrato: string): string;
var
    sSQL: string;
begin
          result := '';

          sSQL := 'SELECT ' +
          'sum(a.dCantidad) as suma ' +
          'FROM actividadesxorden a ' +
          'INNER JOIN  actividadesxanexo b ' +
          'ON a.sContrato = b.sContrato ' +
          'AND a.sIdConvenio = b.sIdConvenio ' +
          'AND a.sWbsContrato = b.sWbs ' +
          'AND a.sTipoActividad = "Actividad" ' +
          'WHERE a.sContrato = :contrato ' +
          'AND a.sIdConvenio = :convenio ' +
          'AND a.sWbsContrato = :wbscontrato ' +
          'AND a.sTipoActividad = "Actividad"';

          connection.QryBusca.Active := false;
          connection.QryBusca.SQL.Clear;
          connection.QryBusca.SQL.Add(sSQL);
          connection.QryBusca.ParamByName('wbscontrato').Value := sWBSContrato;
          connection.QryBusca.ParamByName('contrato').Value := global_contrato;
          connection.QryBusca.ParamByName('convenio').Value := global_convenio;
          connection.QryBusca.Open;

          if connection.QryBusca.RecordCount > 0 then
              result :=  connection.QryBusca.FieldByName('suma').AsString
end;

procedure TfrmCompara.acumularDiferencia(suma, sMensaje: string);
begin
    RxMDValida.Append;
    RxMDValida.FieldByName('sNumeroActividad').Value := connection.QryBusca.FieldByName('sNumeroActividad').AsString;
    RxMDValida.FieldByName('sWbs').Value             := connection.QryBusca.FieldByName('sWbs').AsString;
    RxMDValida.FieldByName('dCantidad').Value        := connection.QryBusca.FieldByName('dCantidad').AsString;
    RxMDValida.FieldByName('suma').Value             := suma;
    RxMDValida.FieldByName('aMN').Value              := connection.QryBusca.FieldByName('aMN').AsString;
    RxMDValida.FieldByName('aDLL').Value             := connection.QryBusca.FieldByName('aDLL').AsString;
    RxMDValida.FieldByName('dCantidadAnexo').Value   := connection.QryBusca.FieldByName('dCantidadAnexo').AsString;
    RxMDValida.FieldByName('bMN').Value              := connection.QryBusca.FieldByName('bMN').AsString;
    RxMDValida.FieldByName('bDLL').Value             := connection.QryBusca.FieldByName('bDLL').AsString;
    RxMDValida.FieldByName('descripcion').Value      := connection.QryBusca.FieldByName('descripcion').AsString;
    RxMDValida.FieldByName('mensaje').Value          := sMensaje;
    RxMDValida.FieldByName('sNumeroOrden').Value     := connection.QryBusca.FieldByName('sNumeroOrden').AsString;
    RxMDValida.FieldByName('sWbs2').Value            := connection.QryBusca.FieldByName('wbs2').AsString;
    RxMDValida.Post;
end;


procedure TfrmCompara.formatoEncabezado;
begin
  Excel.Selection.MergeCells := False;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;
  Excel.Selection.WrapText  := True;
  Excel.Selection.Font.Size := 8;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Font.Name := 'Calibri';
  Excel.Selection.Borders[xlEdgeLeft].LineStyle    := xlContinuous;
  Excel.Selection.Borders[xlEdgeLeft].Weight       := xlThin;
  Excel.Selection.Borders[xlEdgeTop].LineStyle     := xlContinuous;
  Excel.Selection.Borders[xlEdgeTop].Weight        := xlThin;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle  := xlContinuous;
  Excel.Selection.Borders[xlEdgeBottom].Weight     := xlThin;
  Excel.Selection.Borders[xlEdgeRight].LineStyle   := xlContinuous;
  Excel.Selection.Borders[xlEdgeRight].Weight      := xlThin;
end;

procedure TfrmCompara.ConsultaReprogramacion;
begin
    zqReprogramacion.Active := False;
    zqReprogramacion.Params.ParamByName('Contrato').DataType := ftString;
    zqReprogramacion.Params.ParamByName('Contrato').Value    := Global_Contrato;
    zqReprogramacion.Params.ParamByName('Folio').DataType    := ftString;
    zqReprogramacion.Params.ParamByName('folio').Value       := tsNumeroOrden.Text;
    zqReprogramacion.Open;
end;

end.
