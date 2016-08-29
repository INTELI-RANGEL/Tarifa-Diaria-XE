unit frm_abrereporte;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, Grids, DBGrids, frm_connection, StdCtrls, Buttons, global,
   DBCtrls, RXDBCtrl,  frxClass, utilerias, masUtilerias, UnitExcepciones,
  Menus, Gauges, ZAbstractRODataset, ZDataset, RXCtrls, ExtCtrls,
  ComCtrls, AdvGlowButton, udbgrid;

type                       
  TfrmAbreReporte = class(TForm)
    ds_reportediario: TDataSource;
    ds_ordenesdetrabajo: TDataSource;
    pgValidacion: TPageControl;
    TabReportes: TTabSheet;
    TabGeneradores: TTabSheet;
    ds_estimaciones: TDataSource;
    Grid_Generadores: TRxDBGrid;
    frGenerador: TfrxReport;
    rDiario: TfrxReport;
    Progress: TGauge;
    ordenesdetrabajo: TZReadOnlyQuery;
    ReporteDiario: TZReadOnlyQuery;
    ReporteDiariosContrato: TStringField;
    ReporteDiariosNumeroOrden: TStringField;
    ReporteDiariodIdFecha: TDateField;
    ReporteDiariosNumeroReporte: TStringField;
    ReporteDiariosIdTurno: TStringField;
    ReporteDiariosIdConvenio: TStringField;
    ReporteDiariolStatus: TStringField;
    ReporteDiariosIdUsuario: TStringField;
    ReporteDiariosIdUsuarioValida: TStringField;
    ReporteDiariosIdUsuarioAutoriza: TStringField;
    ReporteDiariosTiempoMuerto: TStringField;
    ReporteDiariosOrigen: TStringField;
    ReporteDiariosOrigenTierra: TStringField;
    ReporteDiariosDescripcion: TStringField;
    Estimaciones: TZReadOnlyQuery;
    SecretPanel1: TSecretPanel;
    tsNumeroOrden: TDBLookupComboBox;
    Label2: TLabel;
    TabEstimaciones: TTabSheet;
    Grid_Estimaciones: TRxDBGrid;
    EstimacionPeriodo: TZReadOnlyQuery;
    EstimacionPeriodoiNumeroEstimacion: TStringField;
    EstimacionPeriodosIdTipoEstimacion: TStringField;
    EstimacionPeriodolEstimado: TStringField;
    EstimacionPeriododFechaInicio: TDateField;
    EstimacionPeriododFechaFinal: TDateField;
    EstimacionPeriododMontoMN: TFloatField;
    EstimacionPeriododRetencionMN: TFloatField;
    EstimacionPeriodosDescripcion: TStringField;
    dsEstimacionPeriodo: TDataSource;
    EstimacionPeriodosIdUsuarioAutoriza: TStringField;
    TabRequisicion: TTabSheet;
    TabOrdenCompra: TTabSheet;
    Grid_requisicion: TRxDBGrid;
    Grid_OrdenCompra: TRxDBGrid;
    Requisicion: TZReadOnlyQuery;
    ds_requisicion: TDataSource;
    OrdenCompra: TZReadOnlyQuery;
    ds_OrdenCompra: TDataSource;
    BtnValida: TAdvGlowButton;
    AdvGlowButton2: TAdvGlowButton;
    btnAutoriza: TAdvGlowButton;
    btnExit: TAdvGlowButton;
    Grid_reportes: TRxDBGrid;
    PopSistemas: TPopupMenu;
    mnTiemposMuertos: TMenuItem;
    mnRegeneraAvances: TMenuItem;
    mnValidacionReportes: TMenuItem;
    mnAsignaAvfisico: TMenuItem;
    DesvalidacionTodos: TMenuItem;
    DesautorizacionTodos: TMenuItem;
    PanelProgress: TPanel;
    Label15: TLabel;
    Label16: TLabel;
    Label14: TLabel;
    Label19: TLabel;
    BarraEstado: TProgressBar;
    PanelProgress2: TPanel;
    Label1: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    BarraEstado2: TProgressBar;
    OrdenComprasContrato: TStringField;
    OrdenCompraiFolioPedido: TIntegerField;
    OrdenComprasOrdenCompra: TStringField;
    OrdenComprasFolioRequisicion: TStringField;
    OrdenComprasIdProveedor: TStringField;
    OrdenComprasNumeroOrden: TStringField;
    OrdenCompradIdFecha: TDateField;
    OrdenCompradFechaEntrega: TDateField;
    OrdenComprasReferencia: TStringField;
    OrdenComprasElaboro: TStringField;
    OrdenComprasReviso1: TStringField;
    OrdenComprasReviso2: TStringField;
    OrdenComprasAutorizo: TStringField;
    OrdenComprasMedioTransporte: TStringField;
    OrdenComprasFormaPago: TStringField;
    OrdenCompraiPeriodoPago: TIntegerField;
    OrdenCompralUnicoProveedor: TStringField;
    OrdenComprasMoneda: TStringField;
    OrdenCompradCambio: TFloatField;
    OrdenCompradIVA: TFloatField;
    OrdenCompradDescuento: TFloatField;
    OrdenCompramComentarios: TMemoField;
    OrdenComprasStatus: TStringField;
    OrdenComprasLugarEntrega: TStringField;
    OrdenComprasCondiciones: TStringField;
    OrdenComprasEntrega: TStringField;
    OrdenComprasPrecios: TStringField;
    OrdenComprasVigencia: TStringField;
    OrdenComprasVendedor: TStringField;
    OrdenComprasMail: TStringField;
    RequisicionsContrato: TStringField;
    RequisicioniFolioRequisicion: TIntegerField;
    RequisicionsNumeroOrden: TStringField;
    RequisiciondIdFecha: TDateField;
    RequisiciondFechaSolicitado: TDateField;
    RequisiciondFechaRequerido: TDateField;
    RequisicionsRequisita: TStringField;
    RequisicionsReferencia: TStringField;
    RequisicionsRevision: TStringField;
    RequisicionsSolicito: TStringField;
    RequisicionsStatus: TStringField;
    RequisicionsAutorizo: TStringField;
    RequisicionsVerificacion: TStringField;
    RequisicionsRecibido: TStringField;
    RequisicionsidDepartamento: TStringField;
    RequisicionmComentarios: TMemoField;
    RequisicionsMotivo: TStringField;
    RequisicionsEstado: TStringField;
    RequisicionsLugarEntrega: TStringField;
    RequisicionsNumeroSolicitud: TStringField;
    RequisicionsCodigoMaterial: TStringField;
    ReporteDiariosOrden: TStringField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BtnExitClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure btnValidaClick(Sender: TObject);
    procedure btnAutorizaClick(Sender: TObject);
    procedure Grid_reportesGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure FormActivate(Sender: TObject);
    procedure rDiarioGetValue(const VarName: String; var Value: Variant);
    procedure frGeneradorGetValue(const VarName: String;
      var Value: Variant);
    procedure Grid_reportesDblClick(Sender: TObject);
    procedure Grid_GeneradoresDblClick(Sender: TObject);
    procedure ReporteDiarioCalcFields(DataSet: TDataSet);
    procedure mnTiemposMuertosClick(Sender: TObject);
    procedure mnRegeneraAvancesClick(Sender: TObject);
    procedure mnValidacionReportesClick(Sender: TObject);
    procedure procAjustaBitacoraAlcances (sParamContrato, sParamOrden, sParamTurno : String ; dParamFecha : tDate) ;
    procedure procAjustaBitacoraActividades (sParamContrato, sParamOrden, sParamTurno, sParamConvenio : String ; dParamFecha : tDate) ;
    procedure Grid_GeneradoresGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure EstimacionPeriodoCalcFields(DataSet: TDataSet);
    procedure mnAsignaAvfisicoClick(Sender: TObject);
    procedure Grid_requisicionGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure Grid_OrdenCompraGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure pgValidacionChange(Sender: TObject);
    procedure Grid_EstimacionesGetCellParams(Sender: TObject; Field: TField;
      AFont: TFont; var Background: TColor; Highlight: Boolean);
    procedure Grid_reportesMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_reportesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_reportesTitleClick(Column: TColumn);
    procedure Grid_GeneradoresMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_GeneradoresMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_GeneradoresTitleClick(Column: TColumn);
    procedure Grid_EstimacionesTitleClick(Column: TColumn);
    procedure Grid_EstimacionesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_EstimacionesMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_requisicionMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Grid_requisicionMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_requisicionTitleClick(Column: TColumn);
    procedure Grid_OrdenCompraTitleClick(Column: TColumn);
    procedure Grid_OrdenCompraMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure Grid_OrdenCompraMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure DesvalidacionTodosClick(Sender: TObject);

    procedure DesautorizaTodos(sParamContrato : string);
    procedure DesvalidaTodos(sParamContrato : string);
    procedure DesautorizacionTodosClick(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;
Const
  iPausa = 1000 ;

var
  frmAbreReporte: TfrmAbreReporte;
  sJornada : String ;
  iRecord  : Integer ;
  utgrid:ticdbgrid;
  utgrid2:ticdbgrid;
  utgrid3:ticdbgrid;
  utgrid4:ticdbgrid;
  utgrid5:ticdbgrid;
implementation

uses frm_seguridad;

{$R *.dfm}
procedure TfrmAbreReporte.pgValidacionChange(Sender: TObject);
begin
    //Valida si tiene permisos para desautorizar y desvalidar..
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('select lDesValidaRD, lDesAutorizaRD, lDesValidaGeneradores, lDesAutorizaGeneradores, lDesvalidaEstimacion, lDesautorizaEstimacion from usuarios where sIdUsuario =:usuario ');
    connection.QryBusca.ParamByName('Usuario').AsString := global_usuario;
    connection.QryBusca.Open;

     if pgValidacion.ActivePageIndex = 0 then
     begin
         if connection.QryBusca.FieldValues['lDesValidaRD'] = 'No' then
            btnValida.Enabled   := False;

         if connection.QryBusca.FieldValues['lDesAutorizaRD'] = 'No' then
            btnAutoriza.Enabled := False;

         btnValida.Caption := '&Valida Apertura Reportes Diarios' ;
         btnAutoriza.Caption := '&Autoriza Apertura Reportes Diarios' ;
     end;
     if pgValidacion.ActivePageIndex = 1 then
     begin
         if connection.QryBusca.FieldValues['lDesValidaGeneradores'] = 'No' then
            btnValida.Enabled   := False;

         if connection.QryBusca.FieldValues['lDesAutorizaGeneradores'] = 'No' then
            btnAutoriza.Enabled := False;

         btnValida.Caption := '&Valida Apertura Generadores' ;
         btnAutoriza.Caption := '&Autoriza Apertura Generadores' ;
     end;
     if pgValidacion.ActivePageIndex = 2 then
     begin
         if connection.QryBusca.FieldValues['lDesValidaEstimacion'] = 'No' then
            btnValida.Enabled   := False;

         if connection.QryBusca.FieldValues['lDesAutorizaEstimacion'] = 'No' then
            btnAutoriza.Enabled := False;

         btnValida.Enabled := True ;
         btnValida.Caption := '&Valida Apertura Estimaciones' ;
         btnAutoriza.Caption := '&Autoriza Apertura Estimaciones' ;
     end;
     if pgValidacion.ActivePageIndex = 3 then
     begin
          btnValida.Enabled := True ;
          btnValida.Caption := '&Valida Apertura Requisiciones' ;
          btnAutoriza.Caption := '&Autoriza Apertura Requisiciones' ;
     end;
     if pgValidacion.ActivePageIndex = 4 then
     begin
          btnValida.Enabled := True ;
          btnValida.Caption := '&Valida Apertura Orden de Compra' ;
          btnAutoriza.Caption := '&Autoriza Apertura Orden de Compra' ;
     end;

end;

procedure TfrmAbreReporte.procAjustaBitacoraActividades (sParamContrato, sParamOrden, sParamTurno, sParamConvenio : String ; dParamFecha : tDate) ;
Var
    qryBitacora      : tzReadOnlyQuery ;
    sPaqueteBusqueda,
    sPartidaOriginal : String ;
    dCantidadAnterior,
    dAvanceAnterior  : Currency ;
begin
    qryBitacora := tzReadOnlyQuery.Create(self) ;
    qryBitacora.Connection := connection.ConnTrx;

    // Inicializo la Bitacora Principal

    connection.zCommand.Active := False ;
    connection.zCommand.SQL.Clear ;
    connection.zCommand.SQL.Add ( 'Update bitacoradeactividades SET dCantidadAnterior = 0, dAvanceAnterior = 0, dCantidadActual = 0, dAvanceActual = 0 ' +
                                  'Where sContrato = :Contrato And sNumeroOrden = :Orden And dIdFecha = :fecha And sIdTurno = :Turno') ;
    connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
    connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
    connection.zCommand.Params.ParamByName('Orden').DataType := ftString ;
    connection.zCommand.Params.ParamByName('Orden').Value := sParamOrden ;
    connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate ;
    connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha ;
    connection.zCommand.Params.ParamByName('Turno').DataType := ftString ;
    connection.zCommand.Params.ParamByName('Turno').Value := sParamTurno ;
    connection.zCommand.ExecSQL ;
    qryBitacora.Active := False ;
    qryBitacora.SQL.Clear ;
    qryBitacora.SQL.Add('select b.sWbs, b.sNumeroActividad, b.iIdDiario , Sum(b.dCantidad) as dCantidadActual, Sum(b.dAvance) as dAvanceActual, ' +
                        'sum((b.dAvance * a.dPonderado)/100) as dAvanceReal from bitacoradeactividades b ' +
                        'INNER JOIN actividadesxorden a ON (b.sContrato = a.sContrato And a.sIdConvenio = :Convenio And b.sNumeroOrden = a.sNumeroOrden And ' +
                        'b.sWbs = a.sWbs And b.sNumeroActividad = a.sNumeroActividad) ' +
                        'INNER JOIN tiposdemovimiento t ON (b.sContrato = t.sContrato And b.sIdTipoMovimiento = t.sIdTipoMovimiento And t.sClasificacion = :clasificacion) ' +
                        'where b.sContrato = :contrato and b.dIdFecha = :fecha And b.sNumeroOrden = :Orden And b.sIdTurno = :Turno ' +
                        'Group by b.sWbs, b.sNumeroActividad order by a.iItemOrden, a.sNumeroActividad asc') ;
    qryBitacora.Params.ParamByName('contrato').DataType := ftString ;
    qryBitacora.Params.ParamByName('contrato').Value := global_contrato ;
    qryBitacora.Params.ParamByName('convenio').DataType := ftString ;
    qryBitacora.Params.ParamByName('convenio').Value := sParamConvenio ;
    qryBitacora.Params.ParamByName('Orden').DataType := ftString ;
    qryBitacora.Params.ParamByName('Orden').Value := sParamOrden ;
    qryBitacora.Params.ParamByName('fecha').DataType := ftDate ;
    qryBitacora.Params.ParamByName('fecha').Value := dParamFecha ;
    qryBitacora.Params.ParamByName('Turno').DataType := ftString ;
    qryBitacora.Params.ParamByName('Turno').Value := sParamTurno ;
    qryBitacora.Params.ParamByName('clasificacion').DataType := ftString ;
    qryBitacora.Params.ParamByName('clasificacion').Value := 'Tiempo en Operacion' ;
    qryBitacora.Open ;
    If QryBitacora.RecordCount > 0 then
    Begin
        Progress.Visible := True ;
        Progress.Progress := 1 ;
        Progress.MinValue := 1 ;
        Progress.MaxValue := QryBitacora.RecordCount ;
        QryBitacora.First ;
        For iRecord := 1 to Progress.MaxValue Do
        Begin
            try
                // Aqui almaceno el avance anterior acumulado .........
                connection.qryBusca.Active := False ;
                connection.qryBusca.SQL.Clear ;
                connection.qryBusca.SQL.Add('Select sum(dCantidad) as dInstalado, sum(dAvance) as dAvance from bitacoradeactividades where sContrato = :Contrato and ' +
                                            'dIdFecha < :fecha And sNumeroOrden = :Orden And sWbs = :Wbs And sNumeroActividad = :Actividad Group By sWbs, sNumeroActividad') ;
                connection.qryBusca.Params.ParamByName('contrato').DataType := ftString ;
                connection.qryBusca.Params.ParamByName('contrato').Value := global_contrato ;
                connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
                connection.qryBusca.Params.ParamByName('Fecha').Value := dParamFecha ;
                connection.qryBusca.Params.ParamByName('Orden').DataType := ftString ;
                connection.qryBusca.Params.ParamByName('Orden').Value := sParamOrden ;
                connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString ;
                connection.qryBusca.Params.ParamByName('Wbs').Value := qryBitacora.FieldValues['sWbs'] ;
                connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString ;
                connection.qryBusca.Params.ParamByName('Actividad').Value := qryBitacora.FieldValues['sNumeroActividad'] ;
                connection.qryBusca.Open ;
                dCantidadAnterior := 0 ;
                dAvanceAnterior := 0 ;
                If connection.qryBusca.RecordCount > 0 Then
                Begin
                    dCantidadAnterior := connection.qryBusca.FieldValues['dInstalado'] ;
                    dAvanceAnterior := connection.qryBusca.FieldValues['dAvance'] ;
                End ;

                connection.zCommand.Active := False ;
                connection.zCommand.SQL.Clear ;
                connection.zCommand.SQL.Add ( 'UPDATE bitacoradepaquetes SET dAvance = dAvance + :Avance  ' +
                                              'Where sContrato = :Contrato And sNumeroOrden = :Orden And dIdFecha = :Fecha And sIdConvenio = :convenio And InStr(:wbs, concat(sWbs,".")) > 0') ;
                connection.zCommand.Params.ParamByName('contrato').DataType := ftString ;
                connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato ;
                connection.zCommand.Params.ParamByName('Orden').DataType := ftString ;
                connection.zCommand.Params.ParamByName('Orden').value := sParamOrden ;
                connection.zCommand.Params.ParamByName('convenio').DataType := ftString ;
                connection.zCommand.Params.ParamByName('convenio').value := sParamConvenio ;
                connection.zCommand.Params.ParamByName('fecha').DataType := ftDate ;
                connection.zCommand.Params.ParamByName('fecha').value := dParamFecha ;
                connection.zCommand.Params.ParamByName('wbs').DataType := ftString ;
                connection.zCommand.Params.ParamByName('wbs').value := QryBitacora.FieldValues['sWbs'] ;
                connection.zCommand.Params.ParamByName('Avance').DataType := ftFloat ;
                connection.zCommand.Params.ParamByName('Avance').value := QryBitacora.FieldValues['dAvanceReal'] ;
                connection.zCommand.ExecSQL ;
            Except
                 MessageDlg('Ocurrio un error al actualizar el registro en la bitacora de actividades', mtWarning, [mbOk], 0);
            End ;
            Progress.Progress := iRecord ;
            QryBitacora.Next ;
        End ;
        Progress.Visible := False ;
    End ;
    QryBitacora.Destroy ;
end ;

procedure TfrmAbreReporte.procAjustaBitacoraAlcances (sParamContrato, sParamOrden, sParamTurno : String ; dParamFecha : tDate) ;
Var
    qryBitacora : tzReadOnlyQuery ;
    i : Integer ;
    dCantidadAnterior,
    dAvanceAnterior  : Currency ;

begin
    qryBitacora := tzReadOnlyQuery.Create(self) ;
    qryBitacora.Connection := connection.ConnTrx;

    // Inicializo los acumulados historicos de la bitacora de Alcances ...
    connection.zCommand.Active := False ;
    connection.zCommand.SQL.Clear ;
    connection.zCommand.SQL.Add ( 'Update bitacoradealcances SET dCantidadAnterior = 0, dAvanceAnterior = 0, dCantidadActual = 0, dAvanceActual = 0 ' +
                                  'Where sContrato = :Contrato And sNumeroOrden = :Orden And dIdFecha = :fecha And sIdTurno = :Turno') ;
    connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
    connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
    connection.zCommand.Params.ParamByName('Orden').DataType := ftString ;
    connection.zCommand.Params.ParamByName('Orden').Value := sParamOrden ;
    connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate ;
    connection.zCommand.Params.ParamByName('Fecha').Value := dParamFecha ;
    connection.zCommand.Params.ParamByName('Turno').DataType := ftString ;
    connection.zCommand.Params.ParamByName('Turno').Value := sParamTurno ;
    connection.zCommand.ExecSQL ;

    // 1. Acumulados de la Bitacora de Alcances .... los almaceno en sus historicos ...
    qryBitacora.Active := False ;
    qryBitacora.SQL.Clear ;
    qryBitacora.SQL.Add('select iIdDiario, sWbs, sNumeroActividad, iFase, dCantidad, dAvance From bitacoradealcances where sContrato = :contrato and ' +
                        'dIdFecha = :fecha And sNumeroOrden = :Orden and sIdTurno = :Turno order by sWbs, sNumeroActividad asc') ;
    qryBitacora.Params.ParamByName('contrato').DataType := ftString ;
    qryBitacora.Params.ParamByName('contrato').Value := global_contrato ;
    qryBitacora.Params.ParamByName('Orden').DataType := ftString ;
    qryBitacora.Params.ParamByName('Orden').Value := sParamOrden ;
    qryBitacora.Params.ParamByName('fecha').DataType := ftDate ;
    qryBitacora.Params.ParamByName('fecha').Value := dParamFecha ;
    qryBitacora.Params.ParamByName('Turno').DataType := ftString ;
    qryBitacora.Params.ParamByName('Turno').Value := sParamTurno ;
    qryBitacora.Open ;
    If qryBitacora.RecordCount > 0 then
    Begin
        Progress.Visible := True ;
        Progress.Progress := 1 ;
        Progress.MinValue :=1 ;
        Progress.MaxValue := qryBitacora.RecordCount ;
        qryBitacora.First ;

        For iRecord := 1 to Progress.MaxValue Do
        Begin
            try
                 connection.qryBusca.Active := False ;
                 connection.qryBusca.SQL.Clear ;
                 connection.qryBusca.SQL.Add('Select sum(dCantidad) as dInstalado, sum(dAvance) as dAvance from bitacoradealcances where sContrato = :Contrato and ' +
                                             'dIdFecha < :fecha And sNumeroOrden = :Orden And sWbs = :Wbs And sNumeroActividad = :Actividad And iFase = :Fase Group By sWbs, sNumeroActividad') ;
                 connection.qryBusca.Params.ParamByName('contrato').DataType := ftString ;
                 connection.qryBusca.Params.ParamByName('contrato').Value := global_contrato ;
                 connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
                 connection.qryBusca.Params.ParamByName('Fecha').Value := dParamFecha ;
                 connection.qryBusca.Params.ParamByName('Orden').DataType := ftString ;
                 connection.qryBusca.Params.ParamByName('Orden').Value := sParamOrden ;
                 connection.qryBusca.Params.ParamByName('Wbs').DataType := ftString ;
                 connection.qryBusca.Params.ParamByName('Wbs').Value := qryBitacora.FieldValues['sWbs'] ;
                 connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString ;
                 connection.qryBusca.Params.ParamByName('Actividad').Value := qryBitacora.FieldValues['sNumeroActividad'] ;
                 connection.qryBusca.Params.ParamByName('Fase').DataType := ftInteger;
                 connection.qryBusca.Params.ParamByName('Fase').Value := qryBitacora.FieldValues['iFase'] ;
                 connection.qryBusca.Open ;
                 dCantidadAnterior := 0 ;
                 dAvanceAnterior := 0 ;
                 If connection.qryBusca.RecordCount > 0 Then
                 Begin
                      dCantidadAnterior := connection.qryBusca.FieldValues['dInstalado'] ;
                      dAvanceAnterior := connection.qryBusca.FieldValues['dAvance'] ;
                 End ;
                 connection.zCommand.Active := False ;
                 connection.zCommand.SQL.Clear ;
                 connection.zCommand.SQL.Add ( 'UPDATE bitacoradealcances SET dCantidadAnterior = :CantidadAnterior, dAvanceAnterior = :AvanceAnterior, ' +
                                               'dCantidadActual = :CantidadActual, dAvanceActual = :AvanceActual ' +
                                               'Where sContrato = :Contrato And dIdFecha = :Fecha And iIdDiario = :Diario') ;
                 connection.zCommand.Params.ParamByName('contrato').DataType := ftString ;
                 connection.zCommand.Params.ParamByName('contrato').value := Global_Contrato ;
                 connection.zCommand.Params.ParamByName('fecha').DataType := ftDate ;
                 connection.zCommand.Params.ParamByName('fecha').value := dParamFecha ;
                 connection.zCommand.Params.ParamByName('diario').DataType := ftInteger ;
                 connection.zCommand.Params.ParamByName('diario').value := qryBitacora.FieldValues['iIdDiario'] ;
                 connection.zCommand.Params.ParamByName('CantidadAnterior').DataType := ftFloat ;
                 connection.zCommand.Params.ParamByName('CantidadAnterior').value := dCantidadAnterior ;
                 connection.zCommand.Params.ParamByName('AvanceAnterior').DataType := ftFloat ;
                 connection.zCommand.Params.ParamByName('AvanceAnterior').value := dAvanceAnterior ;
                 connection.zCommand.Params.ParamByName('CantidadActual').DataType := ftFloat ;
                 connection.zCommand.Params.ParamByName('CantidadActual').value := qryBitacora.FieldValues['dCantidad'] ;
                 connection.zCommand.Params.ParamByName('AvanceActual').DataType := ftFloat ;
                 connection.zCommand.Params.ParamByName('AvanceActual').value := qryBitacora.FieldValues['dAvance'] ;
                 connection.zCommand.ExecSQL ;
            Except
                 MessageDlg('Ocurrio un error al actualizar el registro en la bitacora de actividades', mtWarning, [mbOk], 0);
            End ;
            Progress.Progress := iRecord;
            QryBitacora.Next ;
       End ;
       Progress.Visible := False ;
    End ;
    QryBitacora.Destroy ;
end ;


procedure TfrmAbreReporte.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  frmSeguridad.tsIdUsuarioValida.Text := '' ;
  frmSeguridad.tsPasswordValida.Text := '' ;
  action := cafree ;
  utgrid.destroy;
  utgrid2.destroy;
  utgrid3.destroy;
  utgrid4.destroy;
  utgrid5.destroy;
end;

procedure TfrmAbreReporte.BtnExitClick(Sender: TObject);
begin
    close
end;

procedure TfrmAbreReporte.FormShow(Sender: TObject);
begin
  UtGrid:=TicdbGrid.create(grid_reportes);
  UtGrid2:=TicdbGrid.create(grid_generadores);
  UtGrid3:=TicdbGrid.create(grid_estimaciones);
  UtGrid4:=TicdbGrid.create(grid_requisicion);
  UtGrid5:=TicdbGrid.create(grid_ordencompra);
  pgValidacion.ActivePageIndex := 0 ;

  Requisicion.Active := False;
  Requisicion.ParamByName('Contrato').DataType := ftString;
  Requisicion.ParamByName('Contrato').Value    := global_contrato;
  Requisicion.Open;

  OrdenCompra.Active := False;
  OrdenCompra.ParamByName('Contrato').DataType := ftString;
  OrdenCompra.ParamByName('Contrato').Value    := global_contrato;
  OrdenCompra.Open;

  EstimacionPeriodo.Active := False ;
  EstimacionPeriodo.Params.ParamByName('contrato').DataType := ftString ;
  EstimacionPeriodo.Params.ParamByName('contrato').Value := global_contrato ;
  EstimacionPeriodo.Open ;

  Estimaciones.Active := False ;
  Estimaciones.Params.ParamByName('contrato').DataType := ftString ;
  Estimaciones.Params.ParamByName('contrato').Value := global_contrato ;
  Estimaciones.Open ;

  If global_orden_general <> '' then
  Begin
      OrdenesdeTrabajo.Active := False ;
      OrdenesdeTrabajo.SQL.Clear ;
      OrdenesdeTrabajo.SQL.Add('Select sNumeroOrden, iJornada from ordenesdetrabajo where sContrato = :Contrato and ' +
                               'sNumeroOrden = :orden') ;
      OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString ;
      OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := Global_Contrato ;
      ordenesdetrabajo.Params.ParamByName('orden').DataType := ftString ;
      ordenesdetrabajo.Params.ParamByName('orden').Value := global_orden_general ;
      OrdenesdeTrabajo.Open ;
  End
  Else
   begin
      OrdenesdeTrabajo.Active := False ;
      OrdenesdeTrabajo.SQL.Clear ;
      If (global_grupo = 'INTEL-CODE') Then
          OrdenesdeTrabajo.SQL.Add('Select sNumeroOrden, iJornada from ordenesdetrabajo where sContrato = :Contrato and ' +
                               'cIdStatus = :status order by sNumeroOrden')
      Else
          OrdenesdeTrabajo.SQL.Add('Select  ot.sNumeroOrden, ot.sIdPlataforma, ot.sDescripcionCorta, ot.sIdPernocta ' +
                            'from ordenesdetrabajo ot ' +
                            'INNER JOIN ordenesxusuario ou On (ou.sContrato=ot.sContrato '  +
                            'And ou.sNumeroOrden=ot.sNumeroOrden) ' +
                            'where ot.sContrato =:Contrato And ou.sDerechos<>"BLOQUEADO" ' +
                            'And ou.sIdUsuario =:Usuario And ot.cIdStatus =:Status order by ot.sNumeroOrden') ;
      OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString ;
      OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := Global_Contrato ;
      OrdenesdeTrabajo.Params.ParamByName('status').DataType := ftString ;
      OrdenesdeTrabajo.Params.ParamByName('status').Value :=  connection.configuracion.FieldValues [ 'cStatusProceso' ];
      If (global_grupo <> 'INTEL-CODE') Then
       begin
        OrdenesdeTrabajo.Params.ParamByName('Usuario').DataType  := ftString ;
        OrdenesdeTrabajo.Params.ParamByName('Usuario').Value     := Global_Usuario ;
       end;
      OrdenesdeTrabajo.Open ;
   end ;

  tsNumeroOrden.KeyValue := OrdenesdeTrabajo.FieldValues['sNumeroOrden'] ;
  ReporteDiario.Active := False ;
  ReporteDiario.Params.ParamByName('Contrato').DataType := ftString ;
  ReporteDiario.Params.ParamByName('Contrato').Value := global_contrato ;
  //ReporteDiario.Params.ParamByName('Orden').DataType := ftString ;
  //ReporteDiario.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
  ReporteDiario.Open ;


  //Valida si tiene permisos para desautorizar y desvalidar..
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('select lDesValidaRD, lDesAutorizaRD, lDesValidaGeneradores, lDesAutorizaGeneradores, lDesvalidaEstimacion, lDesautorizaEstimacion from usuarios where sIdUsuario =:usuario ');
  connection.QryBusca.ParamByName('Usuario').AsString := global_usuario;
  connection.QryBusca.Open;

  if connection.QryBusca.FieldValues['lDesValidaRD'] = 'No' then
     btnValida.Enabled   := False;

  if connection.QryBusca.FieldValues['lDesAutorizaRD'] = 'No' then
     btnAutoriza.Enabled := False;
end;

procedure TfrmAbreReporte.tsNumeroOrdenExit(Sender: TObject);
begin
  tsNumeroOrden.Color := global_color_salida ;
  ReporteDiario.Active := False ;
  ReporteDiario.Params.ParamByName('Contrato').DataType := ftString ;
  ReporteDiario.Params.ParamByName('Contrato').Value := global_contrato_barco ;
  //ReporteDiario.Params.ParamByName('Orden').DataType := ftString ;
  //ReporteDiario.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
  ReporteDiario.Open ;

  Estimaciones.Active := False ;
  Estimaciones.Params.ParamByName('contrato').DataType := ftString ;
  Estimaciones.Params.ParamByName('contrato').Value := global_contrato ;
  Estimaciones.Open ;
end;

procedure TfrmAbreReporte.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
   {If Key = #13 Then
        If pgValidacion.ActivePageIndex = 0 Then
            Grid_Reportes.SetFocus
        Else
            Grid_Generadores.SetFocus   }
end;

procedure TfrmAbreReporte.tsNumeroOrdenEnter(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmAbreReporte.btnValidaClick(Sender: TObject);
Var
    lPoder      : Boolean ;
    iGrid       : Integer ;
    SavePlace   : TBookmark;
    sMuertoReal : String ;
    lRecordChange : Boolean ;
    Q_User      : TZReadOnlyQuery;
begin
  try
    Connection.CommandTrx.Active := False;
    Connection.CommandTrx.SQL.Text := 'START TRANSACTION';
    Connection.CommandTrx.ExecSQL;

    {En esta parte se asigna el usuario que abre el Reporte de Barco.. 23 de Febrero de 2011}
    If pgValidacion.ActivePageIndex = 0 Then
    begin
        if global_contrato = global_contrato_barco then
        begin
           Q_User := TZReadOnlyQuery.Create(self);
           Q_User.Connection := connection.ConnTrx;

           Q_User.Active := False;
           Q_User.SQL.Clear;
           Q_User.SQL.Add('Update reportediario set sIdUsuarioBarco =:Usuario where sContrato =:Contrato and dIdFecha =:Fecha and sNumeroOrden =:Orden and sIdTurno =:Turno');
           Q_User.ParamByName('Contrato').AsString := global_contrato;
           Q_User.ParamByName('Fecha').AsDate      := reporteDiario.FieldValues['dIdFecha'] ;
           Q_User.ParamByName('Orden').AsString    := reporteDiario.FieldValues['sNumeroOrden'];
           Q_User.ParamByName('Turno').AsString    := reporteDiario.FieldValues['sIdTurno'];
           Q_User.ParamByName('Usuario').AsString  := global_usuario;
           Q_User.ExecSQL;
        end;
        {Termina asignacion de usuario de barco.. 23 Febrero de 2011}
    end;

     //soad -> Proceso de Apertura Validacion de Requisiciones..
    If pgValidacion.ActivePageIndex = 3 Then
    Begin
      If Requisicion.RecordCount > 0 Then
        If Grid_requisicion.SelectedRows.Count > 0 then
          If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
          Begin
            frmSeguridad.ShowModal  ;
            If (global_valida <> '') Then
              lPoder := True
            Else
              lPoder := False
          End
          Else
          Begin
            lPoder := True ;
            global_valida := global_usuario ;
          End
        Else
          raise Exception.Create('-Seleccione por lo menos una Requisición.');

        If lPoder Then
        Begin
            lRecordChange := False ;
            SavePlace := Grid_requisicion.DataSource.DataSet.GetBookmark ;
            with Grid_requisicion.DataSource.DataSet do
                for iGrid := 0 To Grid_requisicion.SelectedRows.Count-1 do
                Begin
                    GotoBookmark(pointer(Grid_requisicion.SelectedRows.Items[iGrid]));
                    If FieldValues['sStatus'] = 'VALIDADO' then
                    Begin
                        lRecordChange := True ;
                        connection.CommandTrx.Active := False;
                        connection.CommandTrx.SQL.Clear;
                        connection.CommandTrx.SQL.Add('Update anexo_requisicion set sStatus ="PENDIENTE" where sContrato =:Contrato and iFolioRequisicion =:Requisicion ');
                        connection.CommandTrx.ParamByName('Contrato').DataType    := ftString;
                        connection.CommandTrx.ParamByName('Contrato').Value       := global_contrato;
                        connection.CommandTrx.ParamByName('Requisicion').DataType := ftInteger;
                        connection.CommandTrx.ParamByName('Requisicion').Value    := Requisicion.FieldValues['iFolioRequisicion'];
                        connection.CommandTrx.ExecSQL;
                    End
                    Else
                      Raise Exception.CreateFmt('-La Requisicon [%i] se encuentra en estado de Autorizado', [Requisicion ['iFolioRequisicion']]);
                        //MessageDlg('La Requisicon [' + IntToStr(Requisicion ['iFolioRequisicion']) + '] se encuentra en estado de Autorizado' , mtInformation, [mbOk], 0) ;
            End;
            If lRecordChange Then
            Begin
                Requisicion.Refresh;
                Try
                    Grid_requisicion.DataSource.DataSet.GotoBookmark(SavePlace);
                Except
                Else
                    Grid_requisicion.DataSource.DataSet.FreeBookmark(SavePlace);
                End ;
                MessageDlg('Proceso terminado con Exito.', mtInformation, [mbOk], 0);
            End
        End;
    End;

    //soad -> Proceso de Apertura Validacion de Ordenes de Compra..
    If pgValidacion.ActivePageIndex = 4 Then
    Begin
         If OrdenCompra.RecordCount > 0 Then
           If Grid_OrdenCompra.SelectedRows.Count > 0 then
              If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
              Begin
                   frmSeguridad.ShowModal  ;
                   If (global_valida <> '') Then
                        lPoder := True
                   Else
                        lPoder := False
              End
              Else
              Begin
                   lPoder := True ;
                    global_valida := global_usuario ;
              End
            Else
              Raise Exception.Create('-Seleccione por lo menos una Orden de Compra.');

        If lPoder Then
        Begin
            lRecordChange := False ;
            SavePlace := Grid_OrdenCompra.DataSource.DataSet.GetBookmark ;
            with Grid_OrdenCompra.DataSource.DataSet do
                for iGrid := 0 To Grid_OrdenCompra.SelectedRows.Count-1 do
                Begin
                    GotoBookmark(pointer(Grid_OrdenCompra.SelectedRows.Items[iGrid]));
                    If FieldValues['sStatus'] = 'VALIDADO' then
                    Begin
                        lRecordChange := True ;
                         connection.CommandTrx.Active := False;
                         connection.CommandTrx.SQL.Clear;
                         connection.CommandTrx.SQL.Add('Update anexo_pedidos set sStatus ="PENDIENTE" where sContrato =:Contrato and iFolioPedido =:Pedido ');
                         connection.CommandTrx.ParamByName('Contrato').DataType := ftString;
                         connection.CommandTrx.ParamByName('Contrato').Value    := global_contrato;
                         connection.CommandTrx.ParamByName('Pedido').DataType   := ftInteger;
                         connection.CommandTrx.ParamByName('Pedido').Value      := OrdenCompra.FieldValues['iFolioPedido'];
                         connection.CommandTrx.ExecSQL;
                    End
                    Else
                      Raise Exception.CreateFmt('-La Orden de Compra No. [%i] se encuentra en estado de Autorizado.',[OrdenCompra ['iFolioPedido']]);
                        //MessageDlg('La Orden de Compra No. [' + IntToStr(OrdenCompra ['iFolioPedido']) + '] se encuentra en estado de Autorizado' , mtInformation, [mbOk], 0) ;
            End ;
            If lRecordChange Then
            Begin
                OrdenCompra.Refresh;
                Try
                    Grid_OrdenCompra.DataSource.DataSet.GotoBookmark(SavePlace);
                Except
                Else
                    Grid_OrdenCompra.DataSource.DataSet.FreeBookmark(SavePlace);
                End ;
                MessageDlg('Proceso terminado con Exito.', mtInformation, [mbOk], 0);
            End
        End  ;
    End;

    frmSeguridad.tsPasswordValida.Text := '' ;
    global_tipo_autorizacion := 'Validación' ;
    lPoder := False ;
    If pgValidacion.ActivePageIndex = 0 Then
    Begin
        If ReporteDiario.RecordCount > 0 Then
            If Grid_reportes.SelectedRows.Count > 0 then
               If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
               Begin
                    frmSeguridad.ShowModal  ;
                    If (global_valida <> '') Then
                        lPoder := True
                    Else
                        lPoder := False
               End
               Else
               Begin
                    lPoder := True ;
                    global_valida := global_usuario ;
               End
            Else
              Raise Exception.Create('-Seleccione por lo menos un Reporte Diario.');

        If lPoder Then
        Begin
            lRecordChange := False ;
            SavePlace := Grid_reportes.DataSource.DataSet.GetBookmark ;
            with Grid_reportes.DataSource.DataSet do
                for iGrid := 0 To Grid_reportes.SelectedRows.Count-1 do
                Begin
                    GotoBookmark(pointer(Grid_reportes.SelectedRows.Items[iGrid]));
                    If (FieldValues['lStatus'] = 'Validado') And (FieldValues['sIdConvenio'] = Global_Convenio) then
                    Begin
                        lRecordChange := True ;
                        connection.zcommand.Active:= False;
                        connection.zcommand.SQL.Clear;
                        connection.zcommand.SQL.Add('Update reportediario SET lStatus = :Status , sIdUsuarioValida = null, iPersonal = 0, ' +
                                                      'sTiempoAdicional = "00:00", sTiempoEfectivo = "00:00", sTiempoMuertoReal = "00:00" ' +
                                                      'Where sContrato = :Contrato And sOrden = :Orden And dIdFecha = :fecha And sIdTurno = :Turno') ;
                        connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                        connection.zcommand.Params.ParamByName('Contrato').Value := Global_Contrato_barco ;
                        connection.zcommand.Params.ParamByName('Orden').DataType := ftString ;
                        connection.zcommand.Params.ParamByName('Orden').Value := FieldValues['sOrden'] ;
                        connection.zcommand.Params.ParamByName('Fecha').DataType := ftDate ;
                        connection.zcommand.Params.ParamByName('Fecha').Value := FieldValues['dIdFecha'] ;
                        connection.zcommand.Params.ParamByName('Turno').DataType := ftString ;
                        connection.zcommand.Params.ParamByName('Turno').Value := FieldValues['sIdTurno'] ;
                        connection.zcommand.Params.ParamByName('Status').DataType := ftString ;
                        connection.zcommand.Params.ParamByName('Status').Value := 'Pendiente' ;
                        connection.zcommand.ExecSQL ;

                        Kardex('Otros Movimientos', 'Autoriza Validacion del Reporte Diario No. [' + FieldValues['sNumeroReporte'] + ']. AUTORIZA ' + global_valida, '', '', '', '', '','Tarifa Diaria','Desvalida Reporte' );
                    End
                    Else
                    begin
                        Connection.ConnTrx.Commit;
                        Raise Exception.CreateFmt('-El Reporte Diario [%s] se encuentra en estado de AUTORIZADO o el Reporte Diario pertenece a otro Convenio.', [FieldValues['sNumeroReporte']]);
                        ReporteDiario.Active := False ;
                        ReporteDiario.Open ;
                        Grid_reportes.UnselectAll;
                        Grid_reportes.DataSource.DataSet.FreeBookmark(SavePlace);
                    end;
                End ;
            If lRecordChange Then
            Begin
                ReporteDiario.Active := False ;
                ReporteDiario.Open ;
                Try
                    Grid_reportes.DataSource.DataSet.GotoBookmark(SavePlace);
                Except
                Else
                    Grid_reportes.DataSource.DataSet.FreeBookmark(SavePlace);
                End ;
                MessageDlg('Proceso terminado con Exito.', mtInformation, [mbOk], 0);
            End ;
        End
    End;

    if (pgValidacion.ActivePageIndex = 1) or (pgValidacion.ActivePageIndex = 2) then
    Begin
        If Estimaciones.RecordCount > 0 Then
           If Grid_Generadores.SelectedRows.Count > 0 then
           Begin
                If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
                Begin
                    frmSeguridad.ShowModal  ;
                    If (global_valida <> '') Then
                        lPoder := True
                    Else
                        lPoder := False
                End
                Else
                Begin
                    lPoder := True ;
                    global_valida := global_usuario ;
                End
           End
           Else
             Raise Exception.Create('-Seleccione por lo menos un Generador.');

        If lPoder Then
        Begin
            lRecordChange := False ;
            SavePlace := Grid_Generadores.DataSource.DataSet.GetBookmark ;
            with Grid_Generadores.DataSource.DataSet do
                for iGrid := 0 To Grid_Generadores.SelectedRows.Count-1 do
                Begin
                    GotoBookmark(pointer(Grid_Generadores.SelectedRows.Items[iGrid]));
                    If FieldValues['lStatus'] = 'Validado' Then
                    begin
                        lRecordChange := True ;
                        connection.CommandTrx.Active := False ;
                        connection.CommandTrx.SQL.Clear ;
                        connection.CommandTrx.SQL.Add ( 'Update estimaciones SET lStatus = :Status, sIdUsuarioValida =null, dMontoMN = 0, dMontoDLL = 0 ' +
                                                      'Where sContrato = :Contrato And sNumeroOrden = :Orden And iNumeroEstimacion = :Estimacion And sNumeroGenerador = :Generador') ;
                        connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Contrato').Value := Global_Contrato ;
                        connection.CommandTrx.Params.ParamByName('Orden').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Orden').Value := FieldValues['sNumeroOrden'] ;
                        connection.CommandTrx.Params.ParamByName('Estimacion').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Estimacion').Value := FieldValues['iNumeroEstimacion'] ;
                        connection.CommandTrx.Params.ParamByName('Generador').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Generador').Value := FieldValues['sNumeroGenerador'] ;
                        connection.CommandTrx.Params.ParamByName('Status').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Status').Value := 'Pendiente' ;
                      //  connection.CommandTrx.Params.ParamByName('Valida').DataType := ftString ;
                      //  connection.CommandTrx.Params.ParamByName('Valida').Value := '' ;
                        connection.CommandTrx.ExecSQL ;

                        // Actualizo Kardex del Sistema ....
                        //Sleep(iPausa) ;
                        connection.CommandTrx.Active := False ;
                        connection.CommandTrx.SQL.Clear ;
                        connection.CommandTrx.SQL.Add ( 'Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                                      'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
                        connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Contrato').Value := Global_Contrato ;
                        connection.CommandTrx.Params.ParamByName('Usuario').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Usuario').Value := Global_Usuario ;
                        connection.CommandTrx.Params.ParamByName('Fecha').DataType := ftDate ;
                        connection.CommandTrx.Params.ParamByName('Fecha').Value := Date ;
                        connection.CommandTrx.Params.ParamByName('Hora').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss.zzz', Now) ;
                        connection.CommandTrx.Params.ParamByName('Descripcion').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Descripcion').Value := 'Apertura del Generador No. [' +  FieldValues ['sNumeroGenerador']  + '] de la Orden [' + tsNumeroOrden.Text + ']. VALIDA ' + global_valida ;
                        connection.CommandTrx.Params.ParamByName('Origen').DataType := ftString ;
                        connection.CommandTrx.Params.ParamByName('Origen').Value := 'Generadores' ;
                        connection.CommandTrx.ExecSQL ;
                    End
                    Else
                      Raise Exception.CreateFmt('-El Numero de Generador: [%s] se encuentra en estado de AUTORIZADO.', [FieldValues['sNumeroGenerador']]);
                End;

            If lRecordChange Then
            Begin
                Estimaciones.Active := False ;
                Estimaciones.Open ;
                Try
                    Grid_Generadores.DataSource.DataSet.GotoBookmark(SavePlace);
                Except
                Else
                    Grid_Generadores.DataSource.DataSet.FreeBookmark(SavePlace);
                End ;
                MessageDlg('Proceso terminado con Exito.', mtInformation, [mbOk], 0);
            End
        End
    End;

    Connection.ConnTrx.Commit;
  Except
    on e:exception do
    begin
      Connection.ConnTrx.Rollback;
      if e.Message[1] = '-' then
        MessageDlg(e.Message, mtWarning, [mbOk], 0)
      else
        MessageDlg('Ha ocurrido un error al tratar de registrar los cambios solicitados.' + #10 + #10 +
                   'Informe del siguiente error al administrador del sistema:' + #10 +
                   e.Message, mtWarning, [mbOk], 0);
    end;
  end;
end;

procedure TfrmAbreReporte.DesautorizacionTodosClick(Sender: TObject);
var
     Progreso, TotalProgreso : real;
     zOrdenes : TzReadOnlyQuery;
begin
    PanelProgress2.Visible := True;
    BarraEstado.Position  := 0;

    zOrdenes := TZReadOnlyQuery.Create(self);
    zOrdenes.Connection := connection.zConnection ;

    zOrdenes.Active := False;
    zOrdenes.SQL.Clear;
    zOrdenes.SQL.Add('select sContrato, sNumeroOrden, dIdFecha from reportediario '+
                     'where dIdFecha =:fecha group by sContrato, sNumeroOrden');
    zOrdenes.ParamByName('fecha').AsDate := ReporteDiario.FieldValues['dIdFecha'];
    zOrdenes.Open;

    while not zOrdenes.Eof do
    begin
        ReporteDiario.Active := False ;
        ReporteDiario.SQL.Clear;
        ReporteDiario.SQL.Add('select r.sContrato, r.sNumeroOrden, r.dIdFecha, r.sNumeroReporte, r.sIdTurno, r.sIdConvenio, r.lStatus, r.sIdUsuario, r.sIdUsuarioValida, '+
                              'r.sIdUsuarioAutoriza, r.sTiempoMuerto, t.sDescripcion as sOrigen, t.sOrigenTierra from reportediario r '+
                              'inner join turnos t on (r.sContrato = t.sContrato And r.sIdTurno = t.sIdTurno) '+
                              'where r.sContrato = :contrato And r.sNumeroOrden = :Orden and dIdFecha =:Fecha And r.lStatus <> "Pendiente" order by r.dIdFecha desc');
        ReporteDiario.Params.ParamByName('Contrato').DataType := ftString ;
        ReporteDiario.Params.ParamByName('Contrato').Value    := zOrdenes.FieldValues['sContrato'];
        ReporteDiario.Params.ParamByName('Orden').DataType    := ftString ;
        ReporteDiario.Params.ParamByName('Orden').Value       := zOrdenes.FieldValues['sNumeroOrden'];
        ReporteDiario.Params.ParamByName('fecha').DataType    := ftDate ;
        ReporteDiario.Params.ParamByName('fecha').Value       := zOrdenes.FieldValues['dIdFecha'];
        ReporteDiario.Open ;

        if ReporteDiario.RecordCount > 0then
           DesautorizaTodos(zOrdenes.FieldValues['sContrato']);

        Label1.Caption := 'Desautorizando.. ' + zOrdenes.FieldValues['sContrato'];
        Label1.Refresh;
        PanelProgress2.Refresh;
        Progreso := (1 / (zOrdenes.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
        TotalProgreso        := TotalProgreso + Progreso;
        BarraEstado2.Position := Trunc(TotalProgreso);

        zOrdenes.Next;
    end;
    ReporteDiario.Active := False ;
    ReporteDiario.SQL.Clear;
    ReporteDiario.SQL.Add('select r.sContrato, r.sNumeroOrden, r.dIdFecha, r.sNumeroReporte, r.sIdTurno, r.sIdConvenio, r.lStatus, r.sIdUsuario, r.sIdUsuarioValida, '+
                              'r.sIdUsuarioAutoriza, r.sTiempoMuerto, t.sDescripcion as sOrigen, t.sOrigenTierra from reportediario r '+
                              'inner join turnos t on (r.sContrato = t.sContrato And r.sIdTurno = t.sIdTurno) '+
                              'where r.sContrato = :contrato And r.sNumeroOrden = :Orden And r.lStatus <> "Pendiente" order by r.dIdFecha desc');
    ReporteDiario.Params.ParamByName('Contrato').DataType := ftString ;
    ReporteDiario.Params.ParamByName('Contrato').Value := global_contrato ;
    ReporteDiario.Params.ParamByName('Orden').DataType := ftString ;
    ReporteDiario.Params.ParamByName('Orden').Value    := tsNumeroOrden.Text;
    ReporteDiario.Open ;

    PanelProgress2.Visible := False;
    messageDLG('Proceso Terminado con Exito!', mtInformation, [mbOk], 0);

end;

procedure TfrmAbreReporte.DesvalidacionTodosClick(Sender: TObject);
var
     Progreso, TotalProgreso : real;
     zOrdenes : TzReadOnlyQuery;
begin
    PanelProgress2.Visible := True;
    BarraEstado.Position  := 0;

    zOrdenes := TZReadOnlyQuery.Create(self);
    zOrdenes.Connection := connection.zConnection ;

    zOrdenes.Active := False;
    zOrdenes.SQL.Clear;
    zOrdenes.SQL.Add('select sContrato, sNumeroOrden, dIdFecha from reportediario '+
                     'where dIdFecha =:fecha group by sContrato, sNumeroOrden');
    zOrdenes.ParamByName('fecha').AsDate := ReporteDiario.FieldValues['dIdFecha'];
    zOrdenes.Open;

    while not zOrdenes.Eof do
    begin
        ReporteDiario.Active := False ;
        ReporteDiario.SQL.Clear;
        ReporteDiario.SQL.Add('select r.sContrato, r.sNumeroOrden, r.dIdFecha, r.sNumeroReporte, r.sIdTurno, r.sIdConvenio, r.lStatus, r.sIdUsuario, r.sIdUsuarioValida, '+
                              'r.sIdUsuarioAutoriza, r.sTiempoMuerto, t.sDescripcion as sOrigen, t.sOrigenTierra from reportediario r '+
                              'inner join turnos t on (r.sContrato = t.sContrato And r.sIdTurno = t.sIdTurno) '+
                              'where r.sContrato = :contrato And r.sNumeroOrden = :Orden and dIdFecha =:Fecha And r.lStatus <> "Pendiente" order by r.dIdFecha desc');
        ReporteDiario.Params.ParamByName('Contrato').DataType := ftString ;
        ReporteDiario.Params.ParamByName('Contrato').Value    := zOrdenes.FieldValues['sContrato'];
        ReporteDiario.Params.ParamByName('Orden').DataType    := ftString ;
        ReporteDiario.Params.ParamByName('Orden').Value       := zOrdenes.FieldValues['sNumeroOrden'];
        ReporteDiario.Params.ParamByName('fecha').DataType    := ftDate ;
        ReporteDiario.Params.ParamByName('fecha').Value       := zOrdenes.FieldValues['dIdFecha'];
        ReporteDiario.Open ;

        if ReporteDiario.RecordCount > 0then
           DesvalidaTodos(zOrdenes.FieldValues['sContrato']);

        Label1.Caption := 'Desvalidando.. ' + zOrdenes.FieldValues['sContrato'];
        Label1.Refresh;
        PanelProgress2.Refresh;
        Progreso := (1 / (zOrdenes.RecordCount + 1)) * (BarraEstado.Max - BarraEstado.Min);
        TotalProgreso        := TotalProgreso + Progreso;
        BarraEstado2.Position := Trunc(TotalProgreso);

        zOrdenes.Next;
    end;
    ReporteDiario.Active := False ;
    ReporteDiario.SQL.Clear;
    ReporteDiario.SQL.Add('select r.sContrato, r.sNumeroOrden, r.dIdFecha, r.sNumeroReporte, r.sIdTurno, r.sIdConvenio, r.lStatus, r.sIdUsuario, r.sIdUsuarioValida, '+
                          'r.sIdUsuarioAutoriza, r.sTiempoMuerto, t.sDescripcion as sOrigen, t.sOrigenTierra from reportediario r '+
                          'inner join turnos t on (r.sContrato = t.sContrato And r.sIdTurno = t.sIdTurno) '+
                          'where r.sContrato = :contrato And r.sNumeroOrden = :Orden And r.lStatus <> "Pendiente" order by r.dIdFecha desc');
    ReporteDiario.Params.ParamByName('Contrato').DataType := ftString ;
    ReporteDiario.Params.ParamByName('Contrato').Value := global_contrato ;
    ReporteDiario.Params.ParamByName('Orden').DataType := ftString ;
    ReporteDiario.Params.ParamByName('Orden').Value    := tsNumeroOrden.Text;
    ReporteDiario.Open ;

    PanelProgress2.Visible := False;
    messageDLG('Proceso Terminado con Exito!', mtInformation, [mbOk], 0);

end;

procedure TfrmAbreReporte.btnAutorizaClick(Sender: TObject);
Var
    lPoder : Boolean ;
    iGrid  : Integer ;
    SavePlace : TBookmark;
    dFechaUltReporte : tDate;
    lAutorizaResidente,
    lRecordChange : Boolean;
    QryBusca: TZQuery;
begin
  Try
    QryBusca := TZQuery.Create(Nil);
    QryBusca.Connection := Connection.ConnTrx;

    Try
      Connection.CommandTrx.Active := False;
      Connection.CommandTrx.SQL.Text := 'START TRANSACTION';
      Connection.CommandTrx.ExecSQL;

      //soad -> Proceso de Apertura Autorizacion de Requisiciones..
      If pgValidacion.ActivePageIndex = 3 Then
      Begin
         If Requisicion.RecordCount > 0 Then
           If Grid_requisicion.SelectedRows.Count > 0 then
              If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
              Begin
                   frmSeguridad.ShowModal  ;
                   If (global_valida <> '') Then
                        lPoder := True
                   Else
                        lPoder := False
              End
              Else
              Begin
                   lPoder := True ;
                    global_valida := global_usuario ;
              End
            Else
              Raise Exception.Create('-Seleccione por lo menos una Requisicion.');

        If lPoder Then
        Begin
            lRecordChange := False ;
            SavePlace := Grid_requisicion.DataSource.DataSet.GetBookmark ;
            with Grid_requisicion.DataSource.DataSet do
                for iGrid := 0 To Grid_requisicion.SelectedRows.Count-1 do
                Begin
                    GotoBookmark(pointer(Grid_requisicion.SelectedRows.Items[iGrid]));
                    If FieldValues['sStatus'] = 'AUTORIZADO' then
                    Begin
                        lRecordChange := True ;
                         connection.CommandTrx.Active := False;
                         connection.CommandTrx.SQL.Clear;
                         connection.CommandTrx.SQL.Add('Update anexo_requisicion set sStatus ="VALIDADO" where sContrato =:Contrato and iFolioRequisicion =:Requisicion ');
                         connection.CommandTrx.ParamByName('Contrato').DataType    := ftString;
                         connection.CommandTrx.ParamByName('Contrato').Value       := global_contrato;
                         connection.CommandTrx.ParamByName('Requisicion').DataType := ftInteger;
                         connection.CommandTrx.ParamByName('Requisicion').Value    := Requisicion.FieldValues['iFolioRequisicion'];
                         connection.CommandTrx.ExecSQL;
                    End
                    Else
                      Raise Exception.CreateFmt('-La Requisicon [%i] se encuentra en estado de Validado.' , [Requisicion ['iFolioRequisicion']]);
            End ;
            If lRecordChange Then
            Begin
                Requisicion.Refresh;
                Try
                    Grid_requisicion.DataSource.DataSet.GotoBookmark(SavePlace);
                Except
                Else
                    Grid_requisicion.DataSource.DataSet.FreeBookmark(SavePlace);
                End ;
                MessageDlg('Proceso terminado con Exito.', mtInformation, [mbOk], 0);
            End
        End;
      End;

      //soad -> Proceso de Apertura Autoriza de Ordenes de Compra..
      If pgValidacion.ActivePageIndex = 4 Then
      Begin
         If OrdenCompra.RecordCount > 0 Then
           If Grid_OrdenCompra.SelectedRows.Count > 0 then
              If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
              Begin
                   frmSeguridad.ShowModal  ;
                   If (global_valida <> '') Then
                        lPoder := True
                   Else
                        lPoder := False
              End
              Else
              Begin
                   lPoder := True ;
                    global_valida := global_usuario ;
              End
            Else
              Raise Exception.Create('Seleccione por lo menos una Orden de Compra.');

        If lPoder Then
        Begin
            lRecordChange := False ;
            SavePlace := Grid_OrdenCompra.DataSource.DataSet.GetBookmark ;
            with Grid_OrdenCompra.DataSource.DataSet do
                for iGrid := 0 To Grid_OrdenCompra.SelectedRows.Count-1 do
                Begin
                    GotoBookmark(pointer(Grid_OrdenCompra.SelectedRows.Items[iGrid]));
                    If FieldValues['sStatus'] = 'AUTORIZADO' then
                    Begin
                        lRecordChange := True ;
                         connection.CommandTrx.Active := False;
                         connection.CommandTrx.SQL.Clear;
                         connection.CommandTrx.SQL.Add('Update anexo_pedidos set sStatus ="VALIDADO" where sContrato =:Contrato and iFolioPedido =:Pedido ');
                         connection.CommandTrx.ParamByName('Contrato').DataType := ftString;
                         connection.CommandTrx.ParamByName('Contrato').Value    := global_contrato;
                         connection.CommandTrx.ParamByName('Pedido').DataType   := ftInteger;
                         connection.CommandTrx.ParamByName('Pedido').Value      := OrdenCompra.FieldValues['iFolioPedido'];
                         connection.CommandTrx.ExecSQL;
                    End
                    else
                      Raise Exception.CreateFmt('-La Orden de Compra No. [%i] se encuentra en estado de Validado.' , [OrdenCompra ['iFolioPedido']]);
            End ;
            If lRecordChange Then
            Begin
                OrdenCompra.Refresh;
                Try
                    Grid_OrdenCompra.DataSource.DataSet.GotoBookmark(SavePlace);
                Except
                Else
                    Grid_OrdenCompra.DataSource.DataSet.FreeBookmark(SavePlace);
                End ;
                MessageDlg('Proceso terminado con Exito.', mtInformation, [mbOk], 0);
            End
        End;
      End;

      frmSeguridad.tsPasswordValida.Text := '' ;
      global_tipo_autorizacion := 'Autorización' ;
      lPoder := False ;
      If pgValidacion.ActivePageIndex = 0 Then
      Begin
        If ReporteDiario.RecordCount > 0 Then
            If Grid_reportes.SelectedRows.Count > 0 then
            Begin
                If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
                Begin
                    frmSeguridad.ShowModal  ;
                    If (global_autoriza <> '') Then
                        lPoder := True
                    Else
                        lPoder := False
                End
                Else
                Begin
                    lPoder := True ;
                    global_autoriza := global_usuario ;
                End
            End
            Else
              Raise Exception.Create('-Seleccione por lo menos un Reporte Diario.');

        If lPoder Then
        Begin
            lRecordChange := False ;
            SavePlace := Grid_reportes.DataSource.DataSet.GetBookmark ;
            with Grid_reportes.DataSource.DataSet do
                for iGrid := 0 To Grid_reportes.SelectedRows.Count-1 do
                Begin
                    GotoBookmark(pointer(Grid_reportes.SelectedRows.Items[iGrid]));
                    If (FieldValues['lStatus'] = 'Autorizado')  And (FieldValues['sIdConvenio'] = Global_Convenio) then
                    Begin
                        If lPoder Then
                        Begin
                            lRecordChange := True ;
                            connection.zcommand.Active := False ;
                            connection.zcommand.SQL.Clear ;
                            connection.zcommand.SQL.Add ( 'Update reportediario SET lStatus = :Status , sIdUsuarioAutoriza = null, sIdUsuarioResidente = null ' +
                                                          'Where sContrato = :Contrato And sNumeroOrden = :Orden And dIdFecha = :fecha And sIdTurno = :Turno') ;
                            connection.zcommand.Params.ParamByName('Contrato').DataType := ftString ;
                            connection.zcommand.Params.ParamByName('Contrato').Value := Global_Contrato_barco ;
                            connection.zcommand.Params.ParamByName('Orden').DataType := ftString ;
                            connection.zcommand.Params.ParamByName('Orden').Value := FieldValues['sNumeroOrden'] ;
                            connection.zcommand.Params.ParamByName('Fecha').DataType := ftDate ;
                            connection.zcommand.Params.ParamByName('Fecha').Value := FieldValues['dIdFecha'] ;
                            connection.zcommand.Params.ParamByName('Turno').DataType := ftString ;
                            connection.zcommand.Params.ParamByName('Turno').Value := FieldValues['sIdTurno'] ;
                            connection.zcommand.Params.ParamByName('Status').DataType := ftString ;
                            connection.zcommand.Params.ParamByName('Status').Value := 'Validado' ;
                            connection.zcommand.ExecSQL ;

                            Kardex('Otros Movimientos', 'Autoriza Apertura del Reporte Diario No. [' + FieldValues['sNumeroReporte'] + ']. AUTORIZA ' + global_valida, '', '', '', '', '','Tarifa Diaria','Desautoriza Reporte' );

                        End
                    End
                    Else
                    begin
                        Raise Exception.CreateFmt('-El Reporte Diario: [%s] se encuentra en estado de VALIDADO o bien corresponde a otro convenio.' , [FieldValues ['sNumeroReporte']]);
                        ReporteDiario.Active := False ;
                        ReporteDiario.Open ;
                        Grid_reportes.UnselectAll;
                        Grid_reportes.DataSource.DataSet.FreeBookmark(SavePlace);
                    end;
                End ;
            If lRecordChange Then
            Begin
                ReporteDiario.Active := False ;
                ReporteDiario.Open ;
                Try
                    Grid_reportes.DataSource.DataSet.GotoBookmark(SavePlace);
                Except
                    Grid_reportes.DataSource.DataSet.FreeBookmark(SavePlace);
                End ;
                MessageDlg('Proceso terminado con Exito.', mtInformation, [mbOk], 0);
            End ;
        End
      End
      Else
        If pgValidacion.ActivePageIndex = 1 Then
        Begin
            If Estimaciones.RecordCount > 0 Then
                If Grid_Generadores.SelectedRows.Count > 0 then
                Begin
                    If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
                    Begin
                        frmSeguridad.ShowModal  ;
                        If (global_autoriza <> '') Then
                            lPoder := True
                        Else
                            lPoder := False
                    End
                    Else
                    Begin
                        lPoder := True ;
                        global_autoriza := global_usuario ;
                    End
                End
                Else
                  Raise Exception.Create('-Seleccione por lo menos un Generador.');

            If lPoder Then
            Begin
                lRecordChange := False ;
                SavePlace := Grid_Generadores.DataSource.DataSet.GetBookmark ;
                with Grid_Generadores.DataSource.DataSet do
                    for iGrid := 0 To Grid_Generadores.SelectedRows.Count-1 do
                    Begin
                        GotoBookmark(pointer(Grid_Generadores.SelectedRows.Items[iGrid]));
                        If (FieldValues['lStatus'] = 'Autorizado') Then
                        Begin
                            qryBusca.Active := False ;
                            qryBusca.SQL.Clear ;
                            qryBusca.SQL.Add('Select iNumeroEstimacion, dFechaInicio, dFechaFinal From estimacionperiodo Where sContrato = :Contrato And iNumeroEstimacion = :Estimacion And lEstimado = "Si"') ;
                            qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
                            qryBusca.Params.ParamByName('Contrato').Value := global_Contrato ;
                            qryBusca.Params.ParamByName('Estimacion').DataType := ftString ;
                            qryBusca.Params.ParamByName('Estimacion').Value := FieldValues['iNumeroEstimacion'] ;
                            qryBusca.Open ;
                            lAutorizaResidente := False ;
                            lPoder := True ;
                            If qryBusca.RecordCount > 0 Then
                              Raise Exception.CreateFmt('-No es posible realizar la apertura de un Generador de Obra que pertenezca al periodo de estimacion del ' +
                                                        '%s al %s de la Estimación No. %s.' + #10 +
                                                        'Realice la DesAutorización de la Estimación para poder realizar esta acción.', [QryBusca.fieldByName('dFechaInicio').AsString, QryBusca.fieldByName('dFechaFinal').AsString, QryBusca.fieldByName('iNumeroEstimacion').AsString]);
                            {Begin
                                 lPoder := False ;
                                 MessageDlg('No es posible realizar la apertura de un Generador de Obra que pertenezca al periodo de estimacion del ' +
                                             connection.QryBusca.fieldByName('dFechaInicio').AsString  + ' al ' + connection.QryBusca.fieldByName('dFechaFinal').AsString +
                                            ' de la Estimacion No. ' + connection.QryBusca.fieldByName('iNumeroEstimacion').AsString + '. ' + chr (13) +
                                            'Realiza la DesAutorizacion de la Estimacion para poder realizar esta accion.', mtWarning, [mbOk], 0);
                            End;}
                            If lPoder Then
                            Begin
                                lRecordChange := True ;
                                connection.CommandTrx.Active := False ;
                                connection.CommandTrx.SQL.Clear ;
                                connection.CommandTrx.SQL.Add ( 'Update estimaciones SET lStatus = :Status, sIdUsuarioAutoriza = null ' +
                                                                'Where sContrato = :Contrato And sNumeroOrden = :Orden And iNumeroEstimacion = :Estimacion And sNumeroGenerador = :Generador') ;
                                connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Contrato').Value := Global_Contrato ;
                                connection.CommandTrx.Params.ParamByName('Orden').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Orden').Value := FieldValues['sNumeroOrden'] ;
                                connection.CommandTrx.Params.ParamByName('Estimacion').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Estimacion').Value := FieldValues['iNumeroEstimacion'] ;
                                connection.CommandTrx.Params.ParamByName('Generador').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Generador').Value := FieldValues['sNumeroGenerador'] ;
                                connection.CommandTrx.Params.ParamByName('Status').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Status').Value := 'Validado' ;
                               // connection.CommandTrx.Params.ParamByName('Valida').DataType := ftString ;
                               // connection.CommandTrx.Params.ParamByName('Valida').Value :=null ;
                                connection.CommandTrx.ExecSQL ;

                                // Actualizo Kardex del Sistema ....
                                //Sleep(iPausa) ;
                                connection.CommandTrx.Active := False ;
                                connection.CommandTrx.SQL.Clear ;
                                connection.CommandTrx.SQL.Add ( 'Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                                                'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
                                connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Contrato').Value := Global_Contrato ;
                                connection.CommandTrx.Params.ParamByName('Usuario').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Usuario').Value := Global_Usuario ;
                                connection.CommandTrx.Params.ParamByName('Fecha').DataType := ftDate ;
                                connection.CommandTrx.Params.ParamByName('Fecha').Value := Date ;
                                connection.CommandTrx.Params.ParamByName('Hora').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss.zzz', Now) ;
                                connection.CommandTrx.Params.ParamByName('Descripcion').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Descripcion').Value := 'Apertura del Generador No. [' +  Estimaciones.FieldValues ['sNumeroGenerador']  + '] de la Orden [' + tsNumeroOrden.Text + ']. AUTORIZA ' + global_autoriza ;
                                connection.CommandTrx.Params.ParamByName('Origen').DataType := ftString ;
                                connection.CommandTrx.Params.ParamByName('Origen').Value := 'Generadores' ;
                                connection.CommandTrx.ExecSQL ;
                            End
                        End
                        Else
                          Raise Exception.CreateFmt('-El Numero de Generador : %s se encuentra en estado de VALIDADO.', [FieldValues ['sNumeroGenerador']]);
                    End ;
                If lRecordChange Then
                Begin
                    Estimaciones.Active := False ;
                    Estimaciones.Open ;
                    Try
                        Grid_Generadores.DataSource.DataSet.GotoBookmark(SavePlace);
                    Except
                    Else
                        Grid_Generadores.DataSource.DataSet.FreeBookmark(SavePlace);
                    End ;
                    MessageDlg('Proceso terminado con Exito.', mtInformation, [mbOk], 0);
                End
            End
        End
        Else
        Begin
            If EstimacionPeriodo.RecordCount > 0 Then
               If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
               Begin
                   frmSeguridad.ShowModal  ;
                   If (global_autoriza <> '') Then
                       lPoder := True
                   Else
                       lPoder := False
               End
               Else
               Begin
                   lPoder := True ;
                   global_autoriza := global_usuario ;
               End
            Else
              Raise Exception.Create('-Seleccione por lo menos un generador.');

            If lPoder Then
            Begin
                connection.CommandTrx.Active := False ;
                connection.CommandTrx.SQL.Clear ;
                connection.CommandTrx.SQL.Add ( 'UPDATE estimacionperiodo SET lEstimado = "No", dMontoMNGeneradores = 0, dMontoDLLGeneradores = 0, ' +
                                                'dMontoMN = 0 , dMontoDLL = 0, dMontoAcumuladoMN = 0, dMontoAcumuladoDLL = 0, sIdUsuarioAutoriza = "" ' +
                                                'Where sContrato = :Contrato And iNumeroEstimacion = :Estimacion') ;
                Connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                Connection.CommandTrx.Params.ParamByName('Contrato').Value := global_contrato ;
                Connection.CommandTrx.Params.ParamByName('Estimacion').DataType := ftString ;
                Connection.CommandTrx.Params.ParamByName('Estimacion').Value := EstimacionPeriodo.FieldValues['iNumeroEstimacion'] ;
                Connection.CommandTrx.ExecSQL ;
                SavePlace := Grid_Estimaciones.DataSource.DataSet.GetBookmark ;
                EstimacionPeriodo.Active := False ;
                EstimacionPeriodo.Open ;
                Try
                   Grid_Estimaciones.DataSource.DataSet.GotoBookmark(SavePlace);
                Except
                Else
                   Grid_Estimaciones.DataSource.DataSet.FreeBookmark(SavePlace);
                End ;
            End
        End;

        Connection.ConnTrx.Commit;
    Except
      on e:Exception do
      begin
        Connection.ConnTrx.Rollback;

        if e.Message[1] = '-' then
          MessageDlg(e.Message, mtWarning, [mbOk], 0)
        else
          MessageDlg('Ha ocurrido un error al tratar de registrar los cambios solicitados.' + #10 + #10 +
                     'Informe del siguiente error al administrador del sistema:' + #10 +
                     e.Message, mtWarning, [mbOk], 0);
      end;
    End;
  Finally
    QryBusca.Close;
    QryBusca.Destroy;
  End;
end;

procedure TfrmAbreReporte.Grid_reportesGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
    If (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('lStatus').AsString = 'Autorizado' then
        Background := $00FFB66C
    Else
        If (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('lStatus').AsString = 'Validado' then
            Background := $00D0AD9F ;
End ;

procedure TfrmAbreReporte.Grid_reportesMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmAbreReporte.Grid_reportesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmAbreReporte.Grid_reportesTitleClick(Column: TColumn);
begin
UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmAbreReporte.Grid_requisicionGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
If (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sStatus').AsString = 'AUTORIZADO' then
        Background := $00FFB66C
    Else
        If (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sStatus').AsString = 'VALIDADO' then
            Background := $00D0AD9F ;
end;

procedure TfrmAbreReporte.Grid_requisicionMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid4.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmAbreReporte.Grid_requisicionMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid4.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmAbreReporte.Grid_requisicionTitleClick(Column: TColumn);
begin
 UtGrid4.DbGridTitleClick(Column);
end;

procedure TfrmAbreReporte.FormActivate(Sender: TObject);
begin
    ReporteDiario.Active := False ;
    Reportediario.Open ;
    Estimaciones.Active := False ;
    Estimaciones.Open ;

    If global_grupo = 'INTEL-CODE' Then
    begin
        mnTiemposMuertos.Enabled := True ;
        mnRegeneraAvances.Enabled := True ;
        mnValidacionReportes.Enabled := True ;
    end
    Else
    begin
        mnTiemposMuertos.Enabled := False ;
        mnRegeneraAvances.Enabled := False ;
        mnValidacionReportes.Enabled := False ;
    End
end;

procedure TfrmAbreReporte.rDiarioGetValue(const VarName: String;
  var Value: Variant);
begin
  If CompareText(VarName, 'DESCRIPCION_CORTA') = 0 then
      Value := sDiarioDescripcionCorta ;

  If CompareText(VarName, 'IMPRIME_AVANCES') = 0 then
      Value := sDiarioComentario ;

  If CompareText(VarName, 'sNewTexto') = 0 then
      Value := sDiarioTitulo ;

  If CompareText(VarName, 'PERIODO') = 0 then
      Value := sDiarioPeriodo ;


  If CompareText(VarName, 'SUPERINTENDENTE') = 0 then
      Value := sSuperIntendente ;
  If CompareText(VarName, 'SUPERVISOR') = 0 then
      Value := sSupervisor ;
  If CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
      Value := sSupervisorTierra ;

  If CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
      Value := sPuestoSuperIntendente ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
      Value := sPuestoSupervisor ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
      Value := sPuestoSupervisorTierra ;

  If CompareText(VarName, 'REAL_ANTERIOR') = 0 then
      Value := dRealGlobalAnterior ;
  If CompareText(VarName, 'REAL_ACTUAL') = 0 then
      Value := dRealGlobalActual ;
  If CompareText(VarName, 'REAL_ACUMULADO') = 0 then
      Value := dRealGlobalAcumulado ;
  If CompareText(VarName, 'PROGRAMADO_ANTERIOR') = 0 then
      Value := dProgramadoGlobalAnterior ;
  If CompareText(VarName, 'PROGRAMADO_ACTUAL') = 0 then
      Value := dProgramadoGlobalActual ;
  If CompareText(VarName, 'PROGRAMADO_ACUMULADO') = 0 then
      Value := dProgramadoGlobalAcumulado;


  If CompareText(VarName, 'REAL_ANTERIOR_MULTIPLE') = 0 then
      Value := dRealOrdenAnterior ;
  If CompareText(VarName, 'REAL_ACTUAL_MULTIPLE') = 0 then
      Value := dRealOrdenActual ;
  If CompareText(VarName, 'REAL_ACUMULADO_MULTIPLE') = 0 then
      Value := dRealOrdenAcumulado ;
  If CompareText(VarName, 'PROGRAMADO_ANTERIOR_MULTIPLE') = 0 then
      Value := dProgramadoOrdenAnterior ;
  If CompareText(VarName, 'PROGRAMADO_ACTUAL_MULTIPLE') = 0 then
      Value := dProgramadoOrdenActual ;
  If CompareText(VarName, 'PROGRAMADO_ACUMULADO_MULTIPLE') = 0 then
      Value := dProgramadoOrdenAcumulado ;

end;

procedure TfrmAbreReporte.frGeneradorGetValue(const VarName: String;
  var Value: Variant);
Var
  sIsometricos : String ;
begin
  If CompareText(VarName, 'ISOMETRICOS') = 0 then
  Begin
      sIsometricos := '' ;
      Connection.qryBusca.Active := False ;
      Connection.qryBusca.SQL.Clear ;
      Connection.qryBusca.SQL.Add('Select distinct sIsometrico, sPrefijo From estimacionxpartida Where sContrato = :Contrato And sNumeroOrden = :Orden And ' +
                                      'sNumeroGenerador = :Generador And sNumeroActividad = :Actividad And sIsometricoReferencia = :Referencia') ;
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
      Connection.qryBusca.Params.ParamByName('Orden').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('Orden').Value := Estimaciones.FieldValues['sNumeroOrden'] ;
      Connection.qryBusca.Params.ParamByName('Generador').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('Generador').Value := Estimaciones.FieldValues['sNumeroGenerador'] ;
      Connection.qryBusca.Params.ParamByName('Actividad').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('Actividad').Value := Estimaciones.FieldValues['sNumeroActividad'] ;
      Connection.qryBusca.Params.ParamByName('Referencia').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('Referencia').Value := Estimaciones.FieldValues['sIsometricoReferencia'] ;
      Connection.qryBusca.Open ;
      While NOT Connection.qryBusca.Eof Do
      Begin
          If sIsometricos <> '' Then
              sIsometricos := sIsometricos + ', ' ;
          sIsometricos := sIsometricos + Connection.qryBusca.FieldValues['sIsometrico'] + ' ' + Connection.qryBusca.FieldValues['sPrefijo'] ;
          Connection.qryBusca.Next
      End ;
      Value := sIsometricos ;
  End ;

  If CompareText(VarName, 'ANEXO') = 0 then
  Begin
      Connection.qryBusca.Active := False ;
      Connection.qryBusca.SQL.Clear ;
      Connection.qryBusca.SQL.Add('Select sAnexo From convenios Where sContrato = :Contrato And sIdConvenio = :convenio') ;
      Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
      Connection.qryBusca.Params.ParamByName('convenio').DataType := ftString ;
      Connection.qryBusca.Params.ParamByName('convenio').Value := global_convenio ;
      Connection.qryBusca.Open ;
      If Connection.qryBusca.RecordCount > 0 Then
          Value := Connection.qryBusca.FieldValues ['sAnexo']
      Else
          Value := '' ;
  End ;
  If CompareText(VarName, 'SUPERINTENDENTE') = 0 then
      Value := sSuperIntendente ;
  If CompareText(VarName, 'SUPERVISOR') = 0 then
      Value := sSupervisorGenerador ;
  If CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
      Value := sSupervisorTierra ;
  If CompareText(VarName, 'SUPERVISOR_RESIDENTE') = 0 then
      Value := sResidente ;

  If CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
      Value := sPuestoSuperIntendente ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
      Value := sPuestoSupervisorGenerador ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
      Value := sPuestoSupervisorTierra ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR_RESIDENTE') = 0 then
      Value := sPuestoResidente ;
end;

procedure TfrmAbreReporte.Grid_reportesDblClick(Sender: TObject);
begin
 If ReporteDiario.RecordCount > 0 Then
   // procReporteDiario (ReporteDiario.FieldValues['sContrato'] , ReporteDiario.FieldValues['sNumeroOrden'], ReporteDiario.FieldValues['sNumeroReporte'], ReporteDiario.FieldValues['sIdTurno'], ReporteDiario.FieldValues['sIdConvenio'] , ReporteDiario.FieldValues['dIdFecha'], '' , frmAbreReporte, rDiario.OnGetValue )
end;

procedure TfrmAbreReporte.Grid_EstimacionesGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
      Background := $006FF8FF;
end;

procedure TfrmAbreReporte.Grid_EstimacionesMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid3.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmAbreReporte.Grid_EstimacionesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid3.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmAbreReporte.Grid_EstimacionesTitleClick(Column: TColumn);
begin
UtGrid3.DbGridTitleClick(Column);
end;

procedure TfrmAbreReporte.Grid_GeneradoresDblClick(Sender: TObject);
begin
  try
      If Estimaciones.RecordCount > 0 Then
          If lfnValidaGenerador (global_contrato, global_convenio, Estimaciones.FieldValues['sNumeroOrden'], Estimaciones.FieldValues['sNumeroGenerador'] , frmAbreReporte ) Then
              If OrdenesdeTrabajo.RecordCount > 1 Then
                    procCaratulaGenerador (0, global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'] , Estimaciones.FieldValues['sNumeroOrden'] , Estimaciones.FieldValues['sNumeroGenerador'] , global_convenio, frmAbreReporte, frGenerador.OnGetValue, True)
              Else
                    procCaratulaGenerador (0, global_contrato, Estimaciones.FieldValues['iNumeroEstimacion'] , Estimaciones.FieldValues['sNumeroOrden'] , Estimaciones.FieldValues['sNumeroGenerador'] , global_convenio, frmAbreReporte, frGenerador.OnGetValue, False)
          Else
              MessageDlg('Existen partidas adicionales de anexo en el generador seleccionado, es necesario adicionar la nota de cambio acerca del motivo por el cual se excedieron las partidas.' , mtWarning, [mbOk], 0) ;
  except
      on e:Exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Abrir Reportes Diarios/Generadores/Estimaciones', 'Al doble click en la cuadricula generadores', 0);
      end;
  end;
end;

procedure TfrmAbreReporte.ReporteDiarioCalcFields(DataSet: TDataSet);
begin
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select sDescripcion From convenios Where sContrato = :Contrato And sIdConvenio = :Convenio') ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := ReporteDiario.FieldValues['sContrato'] ;
    Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Convenio').Value := ReporteDiario.FieldValues['sIdConvenio'] ;
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 Then
        ReporteDiariosDescripcion.Value := Connection.qryBusca.FieldValues['sDescripcion']
    Else
        ReporteDiariosDescripcion.Value := ''
end;

procedure TfrmAbreReporte.mnTiemposMuertosClick(Sender: TObject);
Var
    iJornada,
    iGrid  : Integer ;
begin
  try
    If ReporteDiario.RecordCount > 0 Then
        with Grid_reportes.DataSource.DataSet do
            for iGrid := 0 To Grid_reportes.SelectedRows.Count-1 do
            Begin
                GotoBookmark(pointer(Grid_reportes.SelectedRows.Items[iGrid]));
                If OrdenesdeTrabajo.FieldValues['iJornada'] = 0 Then
                    iJornada := ifnJornadaDia (global_contrato, FieldValues['dIdFecha'], frmAbreReporte)
                Else
                    iJornada := OrdenesdeTrabajo.FieldValues['iJornada'] ;

                If iJornada < 10 Then
                   sJornada := '0' + Trim(IntToStr(iJornada)) + ':00'
                Else
                   sJornada := Trim(IntToStr(iJornada)) + ':00' ;

                If FieldValues['sOrigenTierra'] = 'No' Then
                Begin
                     procInicializaJornadas ( global_contrato, FieldValues['sNumeroOrden'], FieldValues['sIdTurno'], sJornada, FieldValues['dIdFecha'], frmAbreReporte) ;
                     procActualizaJornadas ( global_contrato, FieldValues['sNumeroOrden'], FieldValues['dIdFecha'], frmAbreReporte) ;
                End
            End
  except
      on e:Exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Abrir Reportes Diarios/Generadores/Estimaciones', 'Al click en Regenera Tiempos Muertos en las Fechas Seleccionadas', 0);
      end;
  end;
end;

procedure TfrmAbreReporte.mnRegeneraAvancesClick(Sender: TObject);
var
    iGrid : Integer ;
begin
  try
    If ReporteDiario.RecordCount > 0 Then
        with Grid_reportes.DataSource.DataSet do
            for iGrid := 0 To Grid_reportes.SelectedRows.Count-1 do
            Begin
                GotoBookmark(pointer(Grid_reportes.SelectedRows.Items[iGrid]));
                connection.zCommand.Active := False ;
                connection.zCommand.SQL.Clear ;
                connection.zCommand.SQL.Add ( 'UPDATE avancesglobalesxorden SET dAvance = 0.00 ' +
                                              'WHERE sContrato = :contrato And sIdConvenio = :Convenio And sNumeroOrden = "" ' +
                                              'And dIdFecha = :Fecha');
                connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
                connection.zCommand.Params.ParamByName('Contrato').Value := global_contrato ;
                connection.zCommand.Params.ParamByName('Convenio').DataType := ftString ;
                connection.zCommand.Params.ParamByName('Convenio').Value := FieldValues['sIdConvenio'] ;
                connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate ;
                connection.zCommand.Params.ParamByName('Fecha').Value := FieldValues['dIdFecha'] ;
                connection.zCommand.ExecSQL () ;
                cfnCalculaAvances (global_contrato, '', FieldValues['sIdConvenio'] , 'XXX', False, FieldValues['dIdFecha'], 'Avanzada' , frmAbreReporte)
            End
  except 
      on e:Exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Abrir Reportes Diarios/Generadores/Estimaciones', 'Al click en Regenera Avances Fisicos del Contrato', 0);
      end;
  end;
end;

procedure TfrmAbreReporte.mnValidacionReportesClick(Sender: TObject);
var
    iGrid         : Integer ;
begin
    If ReporteDiario.RecordCount > 0 Then
        with Grid_reportes.DataSource.DataSet do
            for iGrid := 0 To Grid_reportes.SelectedRows.Count-1 do
            Begin
                GotoBookmark(pointer(Grid_reportes.SelectedRows.Items[iGrid]));
                // Primero Elimino todo de la Bitacora de Paquetes de ese dia ...
                connection.zCommand.Active := False ;
                connection.zCommand.SQL.Clear ;
                connection.zCommand.SQL.Add ( 'Delete from bitacoradepaquetes where sContrato = :contrato And sIdConvenio = :convenio And sNumeroOrden = :Orden ' +
                                              'And dIdFecha = :fecha') ;
                connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
                connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
                connection.zCommand.Params.ParamByName('convenio').DataType := ftString ;
                connection.zCommand.Params.ParamByName('convenio').Value := FieldValues['sIdConvenio'] ;
                connection.zCommand.Params.ParamByName('Orden').DataType := ftString ;
                connection.zCommand.Params.ParamByName('Orden').Value := FieldValues['sNumeroOrden'] ;
                connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate ;
                connection.zCommand.Params.ParamByName('Fecha').Value := FieldValues['dIdFecha'] ;
                connection.zCommand.ExecSQL ;

                // Inserccion de todos los paquetes en 0 a la fecha seleccionada ....
                connection.zCommand.Active := False ;
                connection.zCommand.SQL.Clear ;
                connection.zCommand.SQL.Add ( 'insert into bitacoradepaquetes (sContrato, sIdConvenio, dIdFecha, sNumeroOrden, sWbs, sNumeroActividad, dAvance) ' +
                                              'select sContrato, sIdConvenio, :fecha, sNumeroOrden, sWbs, sNumeroActividad, 0 from actividadesxorden ' +
                                              'Where sContrato = :contrato And sIdConvenio = :convenio And sNumeroOrden = :orden And sTipoActividad = "Paquete"') ;
                connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
                connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
                connection.zCommand.Params.ParamByName('convenio').DataType := ftString ;
                connection.zCommand.Params.ParamByName('convenio').Value := FieldValues['sIdConvenio'] ;
                connection.zCommand.Params.ParamByName('Orden').DataType := ftString ;
                connection.zCommand.Params.ParamByName('Orden').Value := FieldValues['sNumeroOrden'] ;
                connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate ;
                connection.zCommand.Params.ParamByName('Fecha').Value := FieldValues['dIdFecha'] ;
                connection.zCommand.ExecSQL () ;

                // Inicia Proceso de Reajuste de Paquetes ....
                // Primero la Bitacora de Alcances
                // ajusto los historicos a 0 y calculo los nuevos historicos ...
                procAjustaBitacoraAlcances (global_contrato, FieldValues['sNumeroOrden'], FieldValues['sIdTurno'], FieldValues['dIdFecha']) ;

                // Ahora la Bitacora de Actividades
                // ajusto los historicos a 0 y calculo los nuevos historicos ...
                procAjustaBitacoraActividades (global_contrato, FieldValues['sNumeroOrden'], FieldValues['sIdTurno'], FieldValues['sIdConvenio'], FieldValues['dIdFecha']) ;
            End
end;

procedure TfrmAbreReporte.Grid_GeneradoresGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
    If (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('lStatus').AsString = 'Autorizado' then
        Background := $00FFB66C
    Else
        If (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('lStatus').AsString = 'Validado' then
            Background := $00D0AD9F ;
end;

procedure TfrmAbreReporte.Grid_GeneradoresMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid2.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmAbreReporte.Grid_GeneradoresMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid2.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmAbreReporte.Grid_GeneradoresTitleClick(Column: TColumn);
begin
UtGrid2.DbGridTitleClick(Column);
end;

procedure TfrmAbreReporte.Grid_OrdenCompraGetCellParams(Sender: TObject;
  Field: TField; AFont: TFont; var Background: TColor; Highlight: Boolean);
begin
If (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sStatus').AsString = 'AUTORIZADO' then
        Background := $00FFB66C
    Else
        If (Sender as TrxDBGrid).DataSource.DataSet.FieldByName('sStatus').AsString = 'VALUDADO' then
            Background := $00D0AD9F ;
end;

procedure TfrmAbreReporte.Grid_OrdenCompraMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid5.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmAbreReporte.Grid_OrdenCompraMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid5.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmAbreReporte.Grid_OrdenCompraTitleClick(Column: TColumn);
begin
 UtGrid5.DbGridTitleClick(Column);
end;

procedure TfrmAbreReporte.EstimacionPeriodoCalcFields(DataSet: TDataSet);
begin
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select sDescripcion From tiposdeestimacion ' +
                                'Where sIdTipoEstimacion = :Tipo') ;
    Connection.qryBusca.Params.ParamByName('Tipo').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Tipo').Value := EstimacionPeriodo.FieldValues['sIdTipoEstimacion'] ;
    Connection.qryBusca.Open ;
    If Connection.qryBusca.RecordCount > 0 then
        EstimacionPeriodosDescripcion.Value := Connection.qryBusca.FieldValues['sDescripcion']
    Else
        EstimacionPeriodosDescripcion.Value := ''
end;

procedure TfrmAbreReporte.mnAsignaAvfisicoClick(Sender: TObject);
var
    iGrid : Integer ;
begin
  try
    If ReporteDiario.RecordCount > 0 Then
        with Grid_reportes.DataSource.DataSet do
            for iGrid := 0 To Grid_reportes.SelectedRows.Count-1 do
            Begin
                GotoBookmark(pointer(Grid_reportes.SelectedRows.Items[iGrid]));
                procAvancesHistorico (global_contrato, FieldValues ['sNumeroOrden'], FieldValues['sIdConvenio'], FieldValues['sIdTurno'], FieldValues ['dIdFecha'], False,  frmAbreReporte ) ;

                connection.zCommand.Active := False ;
                connection.zCommand.SQL.Clear ;
                connection.zCommand.SQL.Add ( 'Update reportediario SET dAvProgAnteriorContrato = :ProgAntContrato, ' +
                                              'dAvProgActualContrato = :ProgActContrato, dAvRealAnteriorContrato = :RealAntContrato, ' +
                                              'dAvRealActualContrato = :RealActcontrato ' +
                                              'Where sContrato = :Contrato And dIdFecha = :fecha') ;
                connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
                connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
                connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate ;
                connection.zCommand.Params.ParamByName('Fecha').Value := FieldValues['dIdFecha'] ;
                connection.zCommand.Params.ParamByName('ProgAntContrato').DataType := ftFloat;
                connection.zCommand.Params.ParamByName('ProgAntContrato').Value := dProgramadoGlobalAnterior ;
                connection.zCommand.Params.ParamByName('ProgActContrato').DataType := ftFloat;
                connection.zCommand.Params.ParamByName('ProgActContrato').Value := dProgramadoGlobalActual ;
                connection.zCommand.Params.ParamByName('RealAntContrato').DataType := ftFloat;
                connection.zCommand.Params.ParamByName('RealAntContrato').Value := dRealGlobalAnterior ;
                connection.zCommand.Params.ParamByName('RealActContrato').DataType := ftFloat;
                connection.zCommand.Params.ParamByName('RealActContrato').Value := dRealGlobalActual ;
                connection.zCommand.ExecSQL () ;
            End
  except
      on e:Exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Abrir Reportes Diarios/Generadores/Estimaciones', 'Al reportar isometricos', 0);
      end;
  end;
end;

procedure TfrmAbreReporte.DesautorizaTodos(sParamContrato: string);
Var
    lPoder : Boolean ;
    iGrid  : Integer ;
    dFechaUltReporte : tDate;
    lAutorizaResidente,
    lRecordChange : Boolean;
    Convenio : string;
    QryBusca: TZQuery;
begin
  Try
    QryBusca := TZQuery.Create(Nil);
    QryBusca.Connection := Connection.ConnTrx;

    Try
      Connection.CommandTrx.Active := False;
      Connection.CommandTrx.SQL.Text := 'START TRANSACTION';
      Connection.CommandTrx.ExecSQL;

      frmSeguridad.tsPasswordValida.Text := '' ;
      global_tipo_autorizacion := 'Autorización' ;
      lPoder := False ;

      dFechaUltReporte := Date ;
      Connection.CommandTrx.Active := False ;
      Connection.CommandTrx.SQL.Clear ;
      Connection.CommandTrx.SQL.Add('Select Max(dIdFecha) as dIdFecha From reportediario Where sContrato = :Contrato') ;
      Connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
      Connection.CommandTrx.Params.ParamByName('Contrato').Value := sParamContrato ;
      Connection.CommandTrx.Open ;
      If Connection.CommandTrx.RecordCount > 0 Then
          dFechaUltReporte := Connection.CommandTrx.FieldValues['dIdFecha'] ;
      dFechaUltReporte := dFechaUltReporte - Connection.configuracion.FieldValues['iReportesSinValid'] ;

      If ReporteDiario.RecordCount > 0 Then
         If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
         Begin
              frmSeguridad.ShowModal  ;
              If (global_autoriza <> '') Then
                  lPoder := True
              Else
                  lPoder := False
          End
          Else
          Begin
              lPoder := True ;
              global_autoriza := global_usuario ;
          End;


        If lPoder Then
        Begin
            lRecordChange := False ;

            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select sIdConvenio from configuracion where sContrato =:Contrato');
            connection.zCommand.ParamByName('Contrato').AsString := sParamContrato;
            connection.zCommand.Open;

            if connection.zCommand.RecordCount > 0 then
               convenio := connection.zCommand.FieldValues['sIdConvenio'];

            If (ReporteDiario.FieldValues['lStatus'] = 'Autorizado')  And (ReporteDiario.FieldValues['sIdConvenio'] = convenio) then
            Begin
                // Checar que no exista un generador autorizado que abarque la fecha del reporte diarios
                QryBusca.Active := False ;
                QryBusca.SQL.Clear ;
                QryBusca.SQL.Add('select sNumeroGenerador, dFechaInicio, dFechaFinal ' +
                                 'from estimaciones Where sContrato = :Contrato and sNumeroOrden = :Orden and dFechaInicio <= :FechaI And dFechaFinal >= :FechaF And lStatus = "Autorizado"') ;
                QryBusca.Params.ParamByName('contrato').DataType := ftString ;
                QryBusca.Params.ParamByName('contrato').Value := sParamContrato ;
                QryBusca.Params.ParamByName('orden').DataType := ftString ;
                QryBusca.Params.ParamByName('orden').Value := ReporteDiario.FieldValues['sNumeroOrden'] ;
                QryBusca.Params.ParamByName('fechai').DataType := ftDate ;
                QryBusca.Params.ParamByName('fechai').Value := ReporteDiario.FieldValues['dIdFecha'] ;
                QryBusca.Params.ParamByName('fechaf').DataType := ftDate ;
                QryBusca.Params.ParamByName('fechaf').Value := ReporteDiario.FieldValues['dIdFecha'] ;
                QryBusca.Open ;

                //Aqui Borro la PERNOCTA
                 connection.CommandTrx.Active := False ;
                 connection.CommandTrx.SQL.Clear ;
                 connection.CommandTrx.SQL.Add ( 'DELETE FROM bitacoradepernocta Where sContrato = :contrato ' +
                                                 'And dIdFecha = :Fecha And sNumeroOrden = :Orden ') ;
                 connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                 connection.CommandTrx.Params.ParamByName('Contrato').Value    := sParamContrato ;
                 connection.CommandTrx.Params.ParamByName('Fecha').DataType    := ftDate ;
                 connection.CommandTrx.Params.ParamByName('Fecha').Value       := ReporteDiario.FieldValues['dIdFecha'] ;
                 connection.CommandTrx.Params.ParamByName('Orden').DataType    := ftString ;
                 connection.CommandTrx.Params.ParamByName('Orden').Value       := ReporteDiario.FieldValues['sNumeroOrden'] ;
                 connection.CommandTrx.ExecSQL () ;
                //TERMINO DE BORRAR LA PERNOCTA

                lPoder := True ;
                If QryBusca.RecordCount > 0 Then
                  Raise Exception.CreateFmt('-No es posible realizar la apertura de un Reporte Diario que pertenezca al periodo de generacion del ' +
                                            '%s al %s del generador de obra No. %s.' + #10 + 'Realice la DesAutorización del generador de obra para poder realizar esta acción.',
                                            [QryBusca.fieldByName('dFechaInicio').AsString, QryBusca.fieldByName('dFechaFinal').AsString, QryBusca.fieldByName('sNumeroGenerador').AsString]);
                {Begin
                  lPoder := False ;
                  MessageDlg('No es posible realizar la apertura de un Reporte Diario que pertenezca al periodo de generacion del ' +
                             connection.QryBusca.fieldByName('dFechaInicio').AsString  + ' al ' + connection.QryBusca.fieldByName('dFechaFinal').AsString +
                             ' del generador de obra No. ' + connection.QryBusca.fieldByName('sNumeroGenerador').AsString + '. ' + chr (13) +
                             'Realiza la DesAutorizacion del generador de obra para poder realizar esta accion.', mtWarning, [mbOk], 0);
                end;}
                If lPoder Then
                Begin
                    lRecordChange := True ;
                    // Actualizo Kardex del Sistema ....
                    connection.CommandTrx.Active := False ;
                    connection.CommandTrx.SQL.Clear ;
                    connection.CommandTrx.SQL.Add ( 'Update reportediario SET lStatus = :Status , sIdUsuarioAutoriza = null, sIdUsuarioResidente = null ' +
                                                  'Where sContrato = :Contrato And sNumeroOrden = :Orden And dIdFecha = :fecha And sIdTurno = :Turno') ;
                    connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                    connection.CommandTrx.Params.ParamByName('Contrato').Value := sParamContrato ;
                    connection.CommandTrx.Params.ParamByName('Orden').DataType := ftString ;
                    connection.CommandTrx.Params.ParamByName('Orden').Value := ReporteDiario.FieldValues['sNumeroOrden'] ;
                    connection.CommandTrx.Params.ParamByName('Fecha').DataType := ftDate ;
                    connection.CommandTrx.Params.ParamByName('Fecha').Value := ReporteDiario.FieldValues['dIdFecha'] ;
                    connection.CommandTrx.Params.ParamByName('Turno').DataType := ftString ;
                    connection.CommandTrx.Params.ParamByName('Turno').Value := ReporteDiario.FieldValues['sIdTurno'] ;
                    connection.CommandTrx.Params.ParamByName('Status').DataType := ftString ;
                    connection.CommandTrx.Params.ParamByName('Status').Value := 'Validado' ;
                    connection.CommandTrx.ExecSQL ;

                    //sleep(iPausa) ;
                    connection.CommandTrx.Active := False ;
                    connection.CommandTrx.SQL.Clear ;
                    connection.CommandTrx.SQL.Add ( 'Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                                  'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
                    connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                    connection.CommandTrx.Params.ParamByName('Contrato').Value := sParamContrato ;
                    connection.CommandTrx.Params.ParamByName('Usuario').DataType := ftString ;
                    connection.CommandTrx.Params.ParamByName('Usuario').Value := Global_Usuario ;
                    connection.CommandTrx.Params.ParamByName('Fecha').DataType := ftDate ;
                    connection.CommandTrx.Params.ParamByName('Fecha').Value := Date ;
                    connection.CommandTrx.Params.ParamByName('Hora').DataType := ftString ;
                    connection.CommandTrx.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss.zzz', Now) ;
                    connection.CommandTrx.Params.ParamByName('Descripcion').DataType := ftString ;
                    connection.CommandTrx.Params.ParamByName('Descripcion').Value := 'Apertura del Reporte Diario No. [' + ReporteDiario.FieldValues ['sNumeroReporte'] + ']. AUTORIZA ' + global_autoriza ;
                    connection.CommandTrx.Params.ParamByName('Origen').DataType := ftString ;
                    connection.CommandTrx.Params.ParamByName('Origen').Value := 'Reporte Diario' ;
                    connection.CommandTrx.ExecSQL
                End
            End
            Else
            begin
                Raise Exception.CreateFmt('-El Reporte Diario: [%s] se encuentra en estado de VALIDADO o bien corresponde a otro convenio.' , [ReporteDiario.FieldValues ['sNumeroReporte']]);
                ReporteDiario.Active := False ;
                ReporteDiario.Open ;
            end;
        End ;                  
        Connection.ConnTrx.Commit;
    Except
      on e:Exception do
      begin
        Connection.ConnTrx.Rollback;

        if e.Message[1] = '-' then
          MessageDlg(e.Message, mtWarning, [mbOk], 0)
        else
          MessageDlg('Ha ocurrido un error al tratar de registrar los cambios solicitados.' + #10 + #10 +
                     'Informe del siguiente error al administrador del sistema:' + #10 +
                     e.Message, mtWarning, [mbOk], 0);
      end;
    End;
  Finally
    QryBusca.Close;
    QryBusca.Destroy;
  End;

end;

procedure TfrmAbreReporte.DesvalidaTodos(sParamContrato: string);
var
    lPoder : Boolean ;
    iGrid  : Integer ;
    dFechaUltReporte : tDate;
    lAutorizaResidente,
    lRecordChange : Boolean;
    Convenio, sMuertoReal : string;
    QryBusca: TZQuery;
begin

    QryBusca := TZQuery.Create(Nil);
    QryBusca.Connection := Connection.ConnTrx;

    Try
       Connection.CommandTrx.Active := False;
       Connection.CommandTrx.SQL.Text := 'START TRANSACTION';
       Connection.CommandTrx.ExecSQL;
        frmSeguridad.tsPasswordValida.Text := '' ;
        global_tipo_autorizacion := 'Validación' ;
        lPoder := False ;
        If ReporteDiario.RecordCount > 0 Then
           If Connection.configuracion.FieldValues['sTipoSeguridad'] = 'Avanzada' Then
           Begin
               frmSeguridad.ShowModal  ;
               If (global_valida <> '') Then
                   lPoder := True
               Else
                   lPoder := False
           End
           Else
           Begin
                lPoder := True ;
                global_valida := global_usuario ;
           End;

        If lPoder Then
        Begin
            lRecordChange := False ;

            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('select sIdConvenio from configuracion where sContrato =:Contrato');
            connection.zCommand.ParamByName('Contrato').AsString := sParamContrato;
            connection.zCommand.Open;

            if connection.zCommand.RecordCount > 0 then
               convenio := connection.zCommand.FieldValues['sIdConvenio'];

            If (ReporteDiario.FieldValues['lStatus'] = 'Validado') And (ReporteDiario.FieldValues['sIdConvenio'] = convenio) then
            Begin
                lRecordChange := True ;
                // Elimino los Tiempo Muertos Reales del Contrato
                sMuertoReal := '00:00' ;
                connection.CommandTrx.Active := False ;
                connection.CommandTrx.SQL.Clear ;

                connection.CommandTrx.SQL.Add ( 'UPDATE reportediario SET sTiempoMuertoReal = :Real, iPersonal= 0 ' +
                'Where sContrato = :Contrato And dIdFecha = :Fecha') ;
                connection.CommandTrx.Params.ParamByName('Real').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Real').value := sMuertoReal ;
                connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Contrato').value := sParamContrato ;
                connection.CommandTrx.Params.ParamByName('Fecha').DataType := ftDate ;
                connection.CommandTrx.Params.ParamByName('Fecha').value := ReporteDiario.FieldValues['dIdFecha'] ;
                connection.CommandTrx.ExecSQL ;

                // Primero Elimino todo de la Bitacora de Paquetes de ese dia ...
                connection.CommandTrx.Active := False ;
                connection.CommandTrx.SQL.Clear ;
                connection.CommandTrx.SQL.Add ( 'Delete from bitacoradepaquetes where sContrato = :contrato And sIdConvenio = :convenio And ' +
                                              'sNumeroOrden = :Orden and dIdFecha = :fecha ') ;
                connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Contrato').Value := sParamContrato ;
                connection.CommandTrx.Params.ParamByName('Convenio').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Convenio').Value := ReporteDiario.FieldValues['sIdConvenio'] ;
                connection.CommandTrx.Params.ParamByName('Orden').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Orden').Value := ReporteDiario.FieldValues['sNumeroOrden'] ;
                connection.CommandTrx.Params.ParamByName('Fecha').DataType := ftDate ;
                connection.CommandTrx.Params.ParamByName('Fecha').Value := ReporteDiario.FieldValues['dIdFecha'] ;
                connection.CommandTrx.ExecSQL ;

                connection.CommandTrx.Active := False ;
                connection.CommandTrx.SQL.Clear ;

                connection.CommandTrx.SQL.Add('Update reportediario SET lStatus = :Status , sIdUsuarioValida = null, iPersonal = 0, ' +
                                              'sTiempoAdicional = "00:00", sTiempoEfectivo = "00:00", sTiempoMuertoReal = "00:00" ' +
                                              'Where sContrato = :Contrato And sNumeroOrden = :Orden And dIdFecha = :fecha And sIdTurno = :Turno') ;
                connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Contrato').Value := sParamContrato ;
                connection.CommandTrx.Params.ParamByName('Orden').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Orden').Value := ReporteDiario.FieldValues['sNumeroOrden'] ;
                connection.CommandTrx.Params.ParamByName('Fecha').DataType := ftDate ;
                connection.CommandTrx.Params.ParamByName('Fecha').Value := ReporteDiario.FieldValues['dIdFecha'] ;
                connection.CommandTrx.Params.ParamByName('Turno').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Turno').Value := ReporteDiario.FieldValues['sIdTurno'] ;
                connection.CommandTrx.Params.ParamByName('Status').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Status').Value := 'Pendiente' ;
                connection.CommandTrx.ExecSQL ;

                // Actualizo Kardex del Sistema ....
                //sleep (iPausa) ;
                connection.CommandTrx.Active := False ;
                connection.CommandTrx.SQL.Clear ;
                connection.CommandTrx.SQL.Add ( 'Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen) ' +
                                              'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
                connection.CommandTrx.Params.ParamByName('Contrato').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Contrato').Value := sParamContrato ;
                connection.CommandTrx.Params.ParamByName('Usuario').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Usuario').Value := Global_Usuario ;
                connection.CommandTrx.Params.ParamByName('Fecha').DataType := ftDate ;
                connection.CommandTrx.Params.ParamByName('Fecha').Value := Date ;
                connection.CommandTrx.Params.ParamByName('Hora').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss.zzz', Now) ;
                connection.CommandTrx.Params.ParamByName('Descripcion').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Descripcion').Value := 'Apertura del Reporte Diario No. [ ' + ReporteDiario.FieldValues ['sNumeroReporte'] + ']. VALIDA ' + global_valida ;
                connection.CommandTrx.Params.ParamByName('Origen').DataType := ftString ;
                connection.CommandTrx.Params.ParamByName('Origen').Value := 'Reporte Diario' ;
                connection.CommandTrx.ExecSQL ;

                // Eliminar los avances globales reportados
                connection.CommandTrx.Active := False;
                connection.CommandTrx.SQL.Text := 'delete from avancesglobalesxorden where scontrato = :contrato and sidconvenio = :convenio and (snumeroorden = "" or snumeroorden = :Orden) and sIdTurno = :Turno and dIdFecha = :fecha';
                connection.CommandTrx.ParamByName('contrato').AsString := sParamContrato;
                connection.CommandTrx.ParamByName('convenio').AsString := ReporteDiario.FieldByName('sIdConvenio').AsString;
                connection.CommandTrx.ParamByName('orden').AsString := ReporteDiario.FieldByName('snumeroorden').AsString;
                connection.CommandTrx.ParamByName('fecha').AsDate := ReporteDiario.FieldByName('dIdFecha').AsDateTime;
                connection.CommandTrx.ParamByName('turno').AsString := ReporteDiario.FieldByName('sIdTurno').AsString;
                connection.CommandTrx.ExecSQL;
            End
            Else
            begin
                Connection.ConnTrx.Commit;
                Raise Exception.CreateFmt('-El Reporte Diario [%s] se encuentra en estado de AUTORIZADO o el Reporte Diario pertenece a otro Convenio.', [ReporteDiario.FieldValues['sNumeroReporte']]);
                ReporteDiario.Active := False ;
                ReporteDiario.Open ;
            end;
        End;
        Connection.ConnTrx.Commit;
  Except
    on e:exception do
    begin
      Connection.ConnTrx.Rollback;
      if e.Message[1] = '-' then
        MessageDlg(e.Message, mtWarning, [mbOk], 0)
      else
        MessageDlg('Ha ocurrido un error al tratar de registrar los cambios solicitados.' + #10 + #10 +
                   'Informe del siguiente error al administrador del sistema:' + #10 +
                   e.Message, mtWarning, [mbOk], 0);
    end;
  end;

end;

end.
