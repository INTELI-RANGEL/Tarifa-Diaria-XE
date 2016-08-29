unit frm_OpcionesActividades;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, frm_connection, global, DateUtils,
  ExtCtrls, DBCtrls, Mask, db, Menus, OleCtrls, frxClass, frxDBSet,
  Buttons, RxMemDS, utilerias, RXCtrls, ExportaExcel,
  ZAbstractRODataset, ZDataset, unitTarifa,
  ZAbstractDataset,
  Grids, DBGrids, AdvPageControl, NxCollection, cxGraphics, cxControls,
  cxLookAndFeels, cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore,
  dxSkinDevExpressStyle, dxSkinFoggy, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxStyles, dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxNavigator, cxDBData, cxGridLevel, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxClasses, cxGridCustomView, cxGrid, cxButtons,
  cxLookupEdit, cxDBLookupEdit, cxDBExtLookupComboBox, dxSkinBlack, dxSkinBlue,
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
  TfrmOpcionesActividades = class(TForm)
    pgOpciones: TPageControl;
    QryConfiguracion: TZReadOnlyQuery;
    QryConfiguracioniFirmasGeneradores: TStringField;
    QryConfiguracioniFirmas: TStringField;
    QryConfiguracionsOrdenPerEq: TStringField;
    QryConfiguracionsTipoPartida: TStringField;
    QryConfiguracionsImprimePEP: TStringField;
    QryConfiguracionsClaveSeguridad: TStringField;
    QryConfiguracioncStatusProceso: TStringField;
    QryConfiguracionsOrdenExtraordinaria: TStringField;
    QryConfiguracionlLicencia: TStringField;
    QryConfiguracionsDireccion1: TStringField;
    QryConfiguracionsDireccion2: TStringField;
    QryConfiguracionsDireccion3: TStringField;
    QryConfiguracionsCiudad: TStringField;
    QryConfiguracionsTelefono: TStringField;
    QryConfiguracionbImagen: TBlobField;
    QryConfiguracionsContrato: TStringField;
    QryConfiguracionsNombre: TStringField;
    QryConfiguracionsPiePagina: TStringField;
    QryConfiguracionsEmail: TStringField;
    QryConfiguracionsWeb: TStringField;
    QryConfiguracionsSlogan: TStringField;
    QryConfiguracionsFirmasElectronicas: TStringField;
    QryConfiguracionlImprimeExtraordinario: TStringField;
    QryConfiguracionsCodigo: TStringField;
    QryConfiguracionmDescripcion: TMemoField;
    QryConfiguracionsTitulo: TMemoField;
    QryConfiguracionmCliente: TMemoField;
    QryConfiguracionbImagenPEP: TBlobField;
    dsConfiguracion: TfrxDBDataset;
    qryOrdenes: TZReadOnlyQuery;
    rReporte: TfrxReport;
    zOrdenes: TZQuery;
    ds_ordenes: TDataSource;
    QryConfiguracionsContratoBarco: TStringField;
    QryConfiguracionmDescripcionBarco: TMemoField;
    pgDatos: TTabSheet;
    ds_actividades: TDataSource;
    zqActividades: TZQuery;
    zqActividadessContrato: TStringField;
    zqActividadesdIdFecha: TDateField;
    zqActividadessNumeroOrden: TStringField;
    zqActividadessIdFolio: TStringField;
    zqActividadesmDescripcion: TMemoField;
    zqActividadessNumeroActividad: TStringField;
    zqActividadessHorainicio: TStringField;
    zqActividadessHorafinal: TStringField;
    zqActividadessIdClasificacion: TStringField;
    zqActividadessTiempo: TStringField;
    zqActividadesdAnterior: TFloatField;
    zqActividadesdActual: TFloatField;
    zqActividadesdAcumulado: TFloatField;
    zqActividadesanterior: TFloatField;
    zqActividadesactual: TFloatField;
    zqActividadesAnteriorDia: TFloatField;
    cxGridActividades: TcxGrid;
    cxViewActividades: TcxGridDBTableView;
    dIdFecha: TcxGridDBColumn;
    sFolio: TcxGridDBColumn;
    Descripcion: TcxGridDBColumn;
    Actividad: TcxGridDBColumn;
    Inicio: TcxGridDBColumn;
    Fin: TcxGridDBColumn;
    Tiempo: TcxGridDBColumn;
    Afectacion: TcxGridDBColumn;
    Anterior: TcxGridDBColumn;
    Actual: TcxGridDBColumn;
    Acumulado: TcxGridDBColumn;
    cxGridLevelActividades: TcxGridLevel;
    cxButton1: TcxButton;
    dlgSave1: TSaveDialog;
    Orden: TcxGridDBColumn;
    procedure FormShow(Sender: TObject);
    function  CalculaAvances(sParamContrato, sParamConvenio, sParamOrden, sParamWbs, sParamTipo, sParamHoraI, sParamHoraF : string;  dParamFecha :tDate; dParamPonderado : double; dParamNivel : integer) : double;
    procedure NotasGerencial(sParamContrato, sParamOrden, sParamInicio, sParamFinal :string; dParamFechaI, dParamFechaF :tdate);
    procedure EditPaquetesEnter(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure zqActividadesCalcFields(DataSet: TDataSet);
    procedure cxViewActividadesDblClick(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmOpcionesActividades: TfrmOpcionesActividades;
  //Variables locales de Gerencial
  num, total_orden  : integer;
  indiceH : integer;
  //Notas de barco
  sUbicacion, sInicio, sFinal, sMovimiento, sArribo : string;
  lInserta, lAplica, lAlerta : boolean;
  sConceptoTiempos  : string;
  sHoraFinal, sMensaje : string;

implementation


{$R *.dfm}

procedure TfrmOpcionesActividades.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := cafree;
end;

procedure TfrmOpcionesActividades.FormShow(Sender: TObject);
begin
   //Obtenemos las fechas minimas y maximas de las actividades.
   connection.zCommand.Active := False;
   connection.zCommand.SQL.Clear;
   connection.zCommand.SQL.Add('select max(dIdFecha) as maxima, min(dIdFecha) as minima '+
                               'from bitacoradeactividades where sContrato =:contrato group by sContrato');
   connection.zCommand.ParamByName('Contrato').AsString := global_contrato;
   connection.zCommand.Open;

   zOrdenes.Active := False;
   zOrdenes.ParamByName('contrato').AsString := global_Contrato;
   zOrdenes.Open;

   if connection.zCommand.RecordCount > 0 then
   begin
       zqActividades.Active := False;
       zqActividades.ParamByName('contrato').AsString := global_Contrato;
       zqActividades.ParamByName('Folio').AsString    := '%';
       zqActividades.ParamByName('fechaI').AsDate     := connection.zCommand.FieldByName('minima').AsDateTime;
       zqActividades.ParamByName('fechaF').AsDate     := connection.zCommand.FieldByName('maxima').AsDateTime;
       zqActividades.Open;
   end;

end;

procedure TfrmOpcionesActividades.cxButton1Click(Sender: TObject);
var
   QueryImagen: TZQuery;
begin
  try
    QueryImagen := TZQuery.Create(Self);

    QueryImagen.Connection := connection.ZConnection;
    QueryImagen.Active:=False;
    QueryImagen.SQL.Clear;
    QueryImagen.SQL.Add('SELECT bImagen FROM configuracion WHERE sContrato=:sContrato');
    QueryImagen.ParamByName('sContrato').AsString:=global_contrato;
    QueryImagen.Open;

    Orden.Visible := True;

    //dxRibbonRadialMenu1
    ExportExcelPersonalizado(QueryImagen,cxViewActividades,'Actividades','Actividades Diarias x Folio');

  finally
       Orden.Visible := False;
  end;
end;

procedure TfrmOpcionesActividades.cxViewActividadesDblClick(Sender: TObject);
begin
    if cxViewActividades.OptionsView.CellAutoHeight then
       cxViewActividades.OptionsView.CellAutoHeight := False
    else
       cxViewActividades.OptionsView.CellAutoHeight := True;
end;

procedure TfrmOpcionesActividades.EditPaquetesEnter(Sender: TObject);
begin

end;

//Procedimiento para guardar datos del Rx,
function TfrmOpcionesActividades.CalculaAvances;
var
     zQAvances, zQCalcula : TzReadOnlyQuery;
     dAvance   : double;
     sSelect, sCondicion, sCondicionHora, sParamCalcula : string;
begin
    zQAvances := TZReadOnlyQuery.Create(self);
    zQAvances.Connection := connection.zConnection;

    zQCalcula := TZReadOnlyQuery.Create(self);
    zQCalcula.Connection := connection.zConnection;

    sParamCalcula := '';
    //Buscamos que ordenes se calculan en automatico y cuales obtienen datos de paquetes,
    zQAvances.Active := False ;
    zQAvances.SQL.Clear ;
    zQAvances.SQL.Add('select lCalculaPaquete from gerencial_diario '+
                      'where sContrato =:Contrato and dIdFecha =:Fecha and sHoraInicio =:Inicio and sHoraFinal =:Final ') ;
    zQAvances.ParamByName('Contrato').AsString := sParamContrato;
    if sParamHoraI > sParamHoraF then
       zQAvances.ParamByName('fecha').AsDate   := dParamFecha + 1
    else
       zQAvances.ParamByName('fecha').AsDate   := dParamFecha;
    zQAvances.ParamByName('Inicio').AsString   := sParamHoraI;
    zQAvances.ParamByName('Final').AsString    := sParamHoraF;
    zQAvances.Open;

    if zQAvances.RecordCount > 0 then
       sParamCalcula := zQAvances.FieldValues['lCalculaPaquete'];

    sSelect := '';
    if sParamTipo = 'Paquete' then
    begin
        dAvance := 0;
        if sParamCalcula = 'Si' then
        begin
             sSelect    := 'Select sum((o.dPonderado/o.dCantidad)* b.dCantidad) as AvanceFisico ';
             sCondicion := 'And b.lAlcance = "No" and b.sWbs like :Wbs group by o.sContrato order by o.sWbs ';
             sParamWbs  := sParamWbs + '.%';
        end
        else
        begin
            //Obtenermos los avances de la bitacora de paquetes..
            zQAvances.Active := False ;
            zQAvances.SQL.Clear ;
            zQAvances.SQL.Add('select dAvance from bitacoradepaquetes '+
                              'where sContrato =:Contrato and sIdConvenio =:Convenio and dIdFecha =:Fecha '+
                              'and sNumeroOrden =:Orden and sWbs =:Wbs ') ;
            zQAvances.ParamByName('Contrato').AsString := sParamContrato;
            zQAvances.ParamByName('Convenio').AsString := sParamConvenio;
            zQAvances.ParamByName('Orden').AsString    := sParamOrden;
            zQAvances.ParamByName('Fecha').AsDate      := dParamFecha ;
            zQAvances.ParamByName('Wbs').AsString      := sParamWbs;
            zQAvances.Open;

            if zqAvances.RecordCount > 0 then
               dAvance := zqAvances.FieldValues['dAvance'];

            CalculaAvances := dAvance;
        end;
    end;

    if sParamTipo = 'Partida' then
    begin
        sSelect    := 'Select sum((100/o.dCantidad)* b.dCantidad) as AvanceFisico ';
        sCondicion := 'And b.lAlcance = "No" and b.sWbs = :Wbs group by o.sContrato order by o.sWbs ';
        dAvance := 0;
    end;

    if sSelect <> '' then
    begin
        try
            //Ahora calculamos los avances anteriores por paquetes o partidas..
            zQAvances.Active := False ;
            zQAvances.SQL.Clear ;
            zQAvances.SQL.Add(sSelect +' From bitacoradeactividades b '+
                      'inner join actividadesxorden o on (b.sContrato = o.sContrato And o.sIdConvenio =:Convenio And o.sNumeroOrden = b.sNumeroOrden And b.sNumeroActividad = o.sNumeroActividad and b.sWbs = o.sWbs and o.sTipoActividad = "Actividad") '+
                      'Where b.sContrato =:Contrato and b.sNumeroOrden =:Orden and b.dIdFecha < :Fecha '+ sCondicion) ;
            zQAvances.ParamByName('Contrato').AsString := sParamContrato;
            zQAvances.ParamByName('Convenio').AsString := sParamConvenio;
            zQAvances.ParamByName('Orden').AsString    := sParamOrden;
            zQAvances.ParamByName('Fecha').AsDate      := dParamFecha;
            zQAvances.ParamByName('Wbs').AsString      := sParamWbs;
            zQAvances.Open;

            if zQAvances.RecordCount > 0 then
               if zQAvances.FieldValues['AvanceFisico'] > 100 then
                  dAvance := 100
               else
                  dAvance := zQAvances.FieldValues['AvanceFisico'];

            //Calculo de avance porcentual paquete
            if sParamTipo = 'Paquete' then
            begin
                dAvance := ((dAvance / dParamPonderado)* 100);
                if dAvance > 100 then
                   dAvance := 100;
            end;

            if sParamHoraI > sParamHoraF then
               sCondicionHora := ' '
            else
               sCondicionHora := ' and b.sHoraInicio =:hInicio and b.sHoraFinal =:hFinal ';

            //Ahora calculamos los avances actuales por paquetes o partidas..
            zQAvances.Active := False ;
            zQAvances.SQL.Clear ;
            zQAvances.SQL.Add(sSelect + ' From bitacoradeactividades b '+
                      'inner join actividadesxorden o on (b.sContrato = o.sContrato And o.sIdConvenio =:Convenio And o.sNumeroOrden = b.sNumeroOrden And b.sNumeroActividad = o.sNumeroActividad and b.sWbs = o.sWbs and o.sTipoActividad = "Actividad") '+
                      'Where b.sContrato =:Contrato and b.sNumeroOrden =:Orden and b.dIdFecha =:Fecha '+ sCondicionHora +sCondicion) ;
            zQAvances.ParamByName('Contrato').AsString := sParamContrato;
            zQAvances.ParamByName('Convenio').AsString := sParamConvenio;
            zQAvances.ParamByName('Orden').AsString    := sParamOrden;
            zQAvances.ParamByName('Fecha').AsDate      := dParamFecha;
            zQAvances.ParamByName('Wbs').AsString      := sParamWbs ;
            if sParamHoraI < sParamHoraF then
            begin
               zQAvances.ParamByName('hInicio').AsString  := sParamHoraI;
               zQAvances.ParamByName('hFinal').AsString   := sParamHoraF;
            end;
            zQAvances.Open;

            if zQAvances.RecordCount > 0 then
            begin
                if sParamTipo = 'Partida' then
                begin
                    if zQAvances.FieldValues['AvanceFisico'] > 100 then
                      dAvance := 100
                    else
                      dAvance := dAvance + zQAvances.FieldValues['AvanceFisico'];
                end;

                //Calculo de avance porcentual paquete
                if sParamTipo = 'Paquete' then
                begin
                    dAvance := dAvance + ((zQAvances.FieldValues['AvanceFisico'] / dParamPonderado)* 100);
                    if dAvance > 100 then
                       dAvance := 100;
                end;
            end;

            CalculaAvances := dAvance;
        Except
        end;
    end;
    zQAvances.Destroy;
    zQCalcula.Destroy;
end;

procedure TfrmOpcionesActividades.zqActividadesCalcFields(DataSet: TDataSet);
begin
    if zqActividades.RecordCount > 0 then
    begin
        zqActividades.FieldByName('sTiempo').AsString   := sfnRestaHoras(zqActividades.FieldByName('sHorafinal').AsString, zqActividades.FieldByName('sHorainicio').AsString) ;
        zqActividades.FieldByName('dAnterior').AsFloat  := (zqActividades.FieldByName('AnteriorDia').AsFloat + zqActividades.FieldByName('Anterior').AsFloat);
        zqActividades.FieldByName('dActual').AsFloat    := zqActividades.FieldByName('Actual').AsFloat;
        zqActividades.FieldByName('dAcumulado').AsFloat := (zqActividades.FieldByName('Anterior').AsFloat + zqActividades.FieldByName('AnteriorDia').AsFloat + zqActividades.FieldByName('Actual').AsFloat);
    end;
end;

procedure TfrmOpcionesActividades.NotasGerencial(sParamContrato: string; sParamOrden: string; sParamInicio: string; sParamFinal: string; dParamFechaI: TDate; dParamFechaF: TDate);
var
   zNotas : tzReadOnlyQuery;
begin
      zNotas := tzReadOnlyQuery.Create(self);
      zNotas.Connection := connection.zConnection;

      //Ahora Consultamos las notas de Reporte Diario..
      zNotas.Active := False;
      zNotas.SQL.Clear;
      zNotas.SQL.Add('select sContrato, dIdFecha, iIdDiario, sNumeroOrden, sHoraInicio, sHoraFinal '+
                     'from bitacoradeactividades b '+
                     'where sContrato =:Contrato and b.sNumeroOrden =:Orden and dIdFecha >=:FechaI and dIdFecha <=:FechaF '+
                     'and sIdTipoMovimiento = "N" and sHoraInicio =:HoraInicio and sHoraFinal =:HoraFinal group by sContrato, dIdfecha, iIdDiario');
      zNotas.ParamByName('Contrato').AsString   := sParamContrato;
      zNotas.ParamByName('Orden').AsString      := sParamOrden;
      zNotas.ParamByName('FechaI').AsDate       := dParamFechaI;
      zNotas.ParamByName('FechaF').AsDate       := dParamFechaI;
      zNotas.ParamByName('HoraInicio').AsString := sParamInicio;
      zNotas.ParamByName('HoraFinal').AsString  := sParamFinal;
      zNotas.Open;

      while not zNotas.Eof do
      begin
          //Ahora consultamos todas las notas del gerencial contenidas en Notas de Reportes Diarios.
          connection.QryBusca.Active := False;
          connection.QryBusca.SQL.Clear;
          connection.QryBusca.SQL.Add('select sHoraInicio, sHoraFinal, mDescripcion, sConceptoGerencial, lImprime '+
                                      'from bitacoradeactividades where sContrato =:Contrato and dIdFecha =:Fecha '+
                                      'and iIdDiarioNota =:Diario and sNumeroOrden =:Orden and sIdTipoMovimiento = "G"');
          connection.QryBusca.ParamByName('Contrato').AsString := sParamContrato;
          connection.QryBusca.ParamByName('Fecha').AsDate      := zNotas.FieldValues['dIdFecha'];
          connection.QryBusca.ParamByName('Diario').AsInteger  := zNotas.FieldValues['iIdDiario'];
          connection.QryBusca.ParamByName('Orden').AsString    := zNotas.FieldValues['sNumeroOrden'];
          connection.QryBusca.Open;


          zNotas.Next;
      end;
      zNotas.Destroy;
end;




end.
