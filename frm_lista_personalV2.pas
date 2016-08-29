unit frm_lista_personalV2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, ComCtrls, AdvDateTimePicker, JvExControls, JvDBLookup,
  cxGraphics, cxControls, cxLookAndFeels, cxLookAndFeelPainters, cxContainer,
  cxEdit, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel,
  dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide, dxSkinDevExpressDarkStyle,
  dxSkinDevExpressStyle, dxSkinFoggy, dxSkinGlassOceans, dxSkinHighContrast,
  dxSkiniMaginary, dxSkinLilian, dxSkinLiquidSky, dxSkinLondonLiquidSky,
  dxSkinMcSkin, dxSkinMetropolis, dxSkinMetropolisDark, dxSkinMoneyTwins,
  dxSkinOffice2007Black, dxSkinOffice2007Blue, dxSkinOffice2007Green,
  dxSkinOffice2007Pink, dxSkinOffice2007Silver, dxSkinOffice2010Black,
  dxSkinOffice2010Blue, dxSkinOffice2010Silver, dxSkinOffice2013DarkGray,
  dxSkinOffice2013LightGray, dxSkinOffice2013White, dxSkinPumpkin, dxSkinSeven,
  dxSkinSevenClassic, dxSkinSharp, dxSkinSharpPlus, dxSkinSilver,
  dxSkinSpringTime, dxSkinStardust, dxSkinSummer2008, dxSkinTheAsphaltWorld,
  dxSkinsDefaultPainters, dxSkinValentine, dxSkinVS2010, dxSkinWhiteprint,
  dxSkinXmas2008Blue, cxLabel, DB, ZAbstractRODataset, Excel2000, ZDataset, cxStyles,
  dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxNavigator, cxDBData, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  ZAbstractDataset,dateutils, cxTextEdit, cxDBLookupComboBox, ZSqlUpdate,
  StdCtrls, Buttons,Frm_BuscaPersonal, Menus, ComObj, StrUtils, OleServer,
  ExcelXP, AdvOfficeButtons, cxCheckBox, cxDropDownEdit,cxCheckComboBox,
  cxMaskEdit;

type
  TfrmListaPersonalV2 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    AvDtpFecha: TAdvDateTimePicker;
    jDblCmbTurnos: TJvDBLookupCombo;
    cxLabel1: TcxLabel;
    QrTurnos: TZReadOnlyQuery;
    dsTurnos: TDataSource;
    jDblCmbCategoria: TJvDBLookupCombo;
    cxLabel2: TcxLabel;
    QrCategorias: TZReadOnlyQuery;
    dsCategorias: TDataSource;
    cxDbGridListadoDBTable: TcxGridDBTableView;
    cxgrdListadoLevel1: TcxGridLevel;
    cxgrdListado: TcxGrid;
    QTripulacion: TZQuery;
    dsTripulacion: TDataSource;
    Ficha: TcxGridDBColumn;
    Nombre: TcxGridDBColumn;
    Id_cat: TcxGridDBColumn;
    Categoria: TcxGridDBColumn;
    Rfc: TcxGridDBColumn;
    Ot: TcxGridDBColumn;
    Compania: TcxGridDBColumn;
    Cabina: TcxGridDBColumn;
    Tipo_pernocta: TcxGridDBColumn;
    QrCompanias: TZReadOnlyQuery;
    QrPernoctas: TZReadOnlyQuery;
    dsCompanias: TDataSource;
    dsPernoctas: TDataSource;
    USqlTripulacion: TZUpdateSQL;
    QrOrdenes: TZReadOnlyQuery;
    dsOrdenes: TDataSource;
    cmdImportar: TBitBtn;
    btnExportar: TBitBtn;
    btnPrinter: TBitBtn;
    btnDeleteAll: TBitBtn;
    btnDelete: TBitBtn;
    btnNuevo: TBitBtn;
    pmPrincipal: TPopupMenu;
    raerPersonaldelDiaAnterios1: TMenuItem;
    raerPersonalxFecha1: TMenuItem;
    tsArchivo: TEdit;
    cxLabel3: TcxLabel;
    OpenXLS: TOpenDialog;
    zq_Esp: TZQuery;
    ExcelApplication1: TExcelApplication;
    ExcelWorkbook1: TExcelWorkbook;
    ExcelWorksheet1: TExcelWorksheet;
    btnExportarDtos: TBitBtn;
    cmdActualizar: TBitBtn;
    AvOChkBActualiza: TcxCheckBox;
    Imprime: TcxGridDBColumn;
    cxStyleRepository1: TcxStyleRepository;
    cxStyle1: TcxStyle;
    cxStyleRepository2: TcxStyleRepository;
    cxStyle2: TcxStyle;
    Grupo: TcxGridDBColumn;
    Pernocta: TcxGridDBColumn;
    ds_pernoctan: TDataSource;
    QPernoctan: TZReadOnlyQuery;
    cmdImportarLista: TBitBtn;
    procedure FormShow(Sender: TObject);
    procedure AvDtpFechaExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnNuevoClick(Sender: TObject);
    procedure jDblCmbTurnosExit(Sender: TObject);
    procedure jDblCmbCategoriaExit(Sender: TObject);
    procedure btnDeleteAllClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure raerPersonaldelDiaAnterios1Click(Sender: TObject);
    procedure raerPersonalxFecha1Click(Sender: TObject);
    procedure AvDtpFechaKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure QTripulacionBeforePost(DataSet: TDataSet);
    procedure btnExportarClick(Sender: TObject);
    procedure cmdImportarClick(Sender: TObject);
    procedure btnExportarDtosClick(Sender: TObject);
    procedure cmdActualizarClick(Sender: TObject);
    procedure QTripulacionAfterPost(DataSet: TDataSet);
    procedure QTripulacionBeforeDelete(DataSet: TDataSet);
    procedure AvOChkBActualizaEditing(Sender: TObject; var CanEdit: Boolean);
    procedure ProcCategoria(sParamGrupo :string);
    procedure ProcCategoriasTripulacion;
    procedure ProcPernoctas;
    function formatoFecha(fecha: TDate) : string;
    procedure cmdImportarListaClick(Sender: TObject);
  private
    { Private declarations }
    TmpFecha:Tdate;
    LParamContrato,LParamIdPersonal,
    LParamCuenta:String;
    lParamFecha:TDate;
    Procedure TraerPersonal(Fecha:Tdate);
    Procedure ActualizarResumen(ParamContrato:String;ParamIdPersonal:String;ParamCuenta:string;ParamFecha:tDate;ParamImprime:String='Si');
  public
    { Public declarations }
  end;

var
  frmListaPersonalV2: TfrmListaPersonalV2;
  flcid, Fila: Integer;

  //Variables globales del formulario
  global_sIdCategoria,
  global_sIdPernocta : string;
  dFechaVigencia : tDate;

implementation

uses frm_connection, global, Frm_EligeFecha;

{$R *.dfm}

Procedure TfrmListaPersonalV2.ActualizarResumen(ParamContrato: string; ParamIdPersonal: string;ParamCuenta:string; ParamFecha: TDate;ParamImprime:String='Si');
Var
  QrDatos:TzReadOnlyQuery;
  QGuardar:TzQuery;
  QConsulta:TzQuery;
begin
  QrDatos   := TzReadOnlyQuery.Create(nil);
  QGuardar  := TzQuery.Create(nil);
  QConsulta := TzQuery.Create(nil);
  try
    if ParamContrato<>'' then
    begin
      QrDatos.Connection:=Connection.zConnection;
      QGuardar.Connection:=Connection.zConnection;
      QConsulta.Connection:=Connection.zConnection;

//      QrDatos.sql.Text:='select sIdCategoria, sDescripcion,sum(iNacionales) as total from tripulaciondiaria_listado where sContrato=:Contrato and '+
//                        'sIdTurno=:Turno and dIdFecha=:fecha and sOrden=:Orden and sIdPersonal=:Personal and sIdCuenta=:Cuenta and sIdCategoria =:Categoria '+
//                        'and sIdPernocta =:Pernocta group by sIdPersonal';
      QrDatos.sql.Text:='select sIdCategoria, sDescripcion,sum(iNacionales) as total, dCantidad, dSolicitado, sIdPernocta from tripulaciondiaria_listado where sContrato=:Contrato and '+
                        'sIdTurno=:Turno and dIdFecha=:fecha and sOrden=:Orden and sIdPersonal=:Personal and sIdCuenta=:Cuenta and sIdCategoria =:Categoria '+
                        'group by sIdPersonal';
      QrDatos.ParamByName('Contrato').AsString  := global_Contrato_Barco;
      QrDatos.ParamByName('Turno').AsString     :=          global_Turno;
      QrDatos.ParamByName('Fecha').AsDate       :=            ParamFecha;
      QrDatos.ParamByName('Orden').AsString     :=         ParamContrato;
      QrDatos.ParamByName('Personal').AsString  :=       ParamIdPersonal;
      QrDatos.ParamByName('Cuenta').AsString    :=           ParamCuenta;
      QrDatos.ParamByName('Categoria').AsString :=   global_sIdCategoria;
      //QrDatos.ParamByName('Pernocta').AsString:=    global_sIdPernocta;
      try
        QrDatos.Open;
      except
        raise;
      end;

      QGuardar.SQL.Text:= 'select * from bitacoradepersonal_cuadre where sContrato=:Contrato and ' +
                          'dIdFecha=:fecha and sIdPersonal=:Personal and sIdPlataforma="@" and sNumeroOrden="@" and '+
                          'sTipoPernocta=:Tipo and sAgrupaPersonal =:categoria and sIdPernocta =:Pernocta ';
      QGuardar.ParamByName('Contrato').AsString  :=       ParamContrato;
      QGuardar.ParamByName('Fecha').AsDate       :=          Paramfecha;
      QGuardar.ParamByName('Personal').AsString  :=     ParamIdPersonal;
      QGuardar.ParamByName('Tipo').AsString      :=         ParamCuenta;
      QGuardar.ParamByName('Categoria').AsString := global_sIdCategoria;
      QGuardar.ParamByName('Pernocta').AsString  :=  QrDatos.FieldByName('sIdPernocta').AsString;
      QGuardar.Open;

      if (QrDatos.Recordcount=0) and (QGuardar.recordcount=1) then
        QGuardar.Delete
      else
        if (QrDatos.Recordcount=1) then
          if QGuardar.recordcount=0 then
          begin
            QGuardar.Append;
            QGuardar.FieldByName('sContrato').AsString    :=                               ParamContrato;
            QGuardar.FieldByName('dIdFecha').AsDateTime   :=                                  ParamFecha;
            QGuardar.FieldByName('sIdPersonal').AsString  :=                             ParamIdPersonal;
            QGuardar.FieldByName('iItemOrden').AsInteger  :=                                           0;
            QGuardar.FieldByName('sDescripcion').AsString :=QrDatos.fieldByName('sDescripcion').AsString;

//            //se carga el sidpernocta y el sidplataroama de acuerdo al folio
//            QConsulta.Active := false;
//            QConsulta.SQL.Clear;
//            QConsulta.sql.Text := ''+
//                       ' select * from bitacoradepersonal where dIdFecha = :dIdFecha and sContrato = :scontrato '+
//                       ' limit 1 ';
//            QConsulta.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', Paramfecha);
//            QConsulta.ParamByName('scontrato').AsString := ParamContrato;
//            QConsulta.Open;
//
//
//            if connection.QryBusca.recordcount = 1 then
//            begin
//              QGuardar.FieldByName('sIdPernocta').AsString:= connection.QryBusca.FieldByName('sIdPernocta').asstring;
//              QGuardar.FieldByName('sIdPlataforma').AsString:= '*';
//            end
//            else
           // begin
              QGuardar.FieldByName('sIdPernocta').AsString:= '@';
              QGuardar.FieldByName('sIdPlataforma').AsString:= '@';
            //end;

            QGuardar.FieldByName('sNumeroOrden').AsString:='@';
            QGuardar.FieldByName('sHoraInicio').AsString:='00:00';
            QGuardar.FieldByName('sHoraFinal').AsString:='00:00';
            QGuardar.FieldByName('dCantidad').AsInteger      :=QrDatos.FieldByName('dCantidad').AsInteger;
            QGuardar.FieldByName('sAgrupaPersonal').AsString :=QrDatos.FieldByName('sIdCategoria').AsString;
            QGuardar.FieldByName('lAplicaPernocta').AsString:='Si';
            QGuardar.FieldByName('sTipoPernocta').AsString:=ParamCuenta;
            QGuardar.FieldByName('dSolicitado').AsFloat:=QrDatos.FieldByName('dSolicitado').AsFloat;
            QGuardar.FieldByName('dCantHH').AsFloat:=0;

            QConsulta.Active:=false;
            QConsulta.SQL.Text:='select * from personal where sContrato=:Contrato and sIdPersonal=:Personal';
            QConsulta.ParamByName('Contrato').AsString:=Global_Contrato_Barco;
            QConsulta.ParamByName('Personal').AsString:=ParamIdPersonal;
            QConsulta.open;
            if QConsulta.recordcount=1 then
              QGuardar.FieldByName('sTipoObra').AsString:=QConsulta.FieldByName('sIdTipoPersonal').AsString
            else
              QGuardar.FieldByName('sTipoObra').AsString:='';

            QGuardar.FieldByName('lImprimeResumen').AsString:=ParamImprime;
            QGuardar.FieldByName('sIdPernocta').AsString    :=global_sIdPernocta;

            QGuardar.Post;
          end
          else
          begin
            QGuardar.Edit;
            QGuardar.FieldByName('dCantidad').AsInteger:=QrDatos.FieldByName('dCantidad').AsInteger;
            QGuardar.FieldByName('dSolicitado').AsFloat:=QrDatos.FieldByName('dSolicitado').AsFloat;
            QGuardar.FieldByName('sAgrupaPersonal').AsString :=QrDatos.FieldByName('sIdCategoria').AsString;
            QGuardar.FieldByName('lImprimeResumen').AsString:=ParamImprime;               
            QGuardar.Post;
          end;
    end;
  finally
    QrDatos.Destroy;
    QGuardar.Destroy;
    QConsulta.Destroy;
  end;


end;



Procedure TfrmListaPersonalV2.TraerPersonal(Fecha:Tdate);
var
  Resp:Word;
  QTraerDatos:TzReadOnlyQuery;
  Grabar:Boolean;
begin
  Resp:=  (MessageDlg('¿Desea eliminar todos los datos Actuales?',
          mtConfirmation, [mbYes, mbNo], 0));
  QTraerDatos:=TzReadOnlyQuery.Create(nil);
  try
    QTraerDatos.Connection:=connection.zConnection;
    QTraerDatos.SQL.Text:='Select a.*, cp.sDescripcion as Compania,c.sidPernocta,c.sdescripcion as pernocta' + #10 +
                          'from' + #10 +
                          'tripulaciondiaria_listado as a' + #10 +
                          'left join' + #10 +
                          'compersonal as cp' + #10 +
                          'on(cp.sIdCompania = a.sIdCompania)' + #10 +
                          'left join cuentas c' + #10 +
                          'on(c.sidcuenta=a.sidcuenta)' + #10 +
                          'where a.sContrato =:Contrato' + #10 +
                          'And a.dIdFecha =:Fecha' + #10 +
                          'and a.sIdTurno =:Turno';
    if connection.contrato.FieldByName('sTipoObra').AsString='BARCO' then
      QTraerDatos.ParamByName('Contrato').AsString:=global_contrato
    else
      QTraerDatos.ParamByName('Contrato').AsString:=global_contrato_Barco;
    QTraerDatos.ParamByName('Fecha').AsDate:=Fecha;
    QTraerDatos.ParamByName('Turno').AsString:=Global_turno;
    if Resp=MrYes then
    begin
      Connection.zCommand.Active:=false;
      Connection.zCommand.SQL.Text:='delete from tripulaciondiaria_listado' + #10 +
                                    'where sContrato=:Contrato and  sIdTurno=:Turno' + #10 +
                                    'and dIdFecha=:Fecha';
      if connection.contrato.FieldByName('sTipoObra').AsString='BARCO' then
        Connection.zCommand.ParamByName('Contrato').AsString:=global_contrato
      else
        Connection.zCommand.ParamByName('Contrato').AsString:=global_contrato_Barco;

      Connection.zCommand.ParamByName('Turno').AsString:=Global_turno;
      Connection.zCommand.ParamByName('Fecha').AsDate:=AvDtpFecha.Date;
      Connection.zCommand.ExecSQL;
    end;

    QTraerDatos.Open;
    while not QTraerDatos.Eof do
    begin
      Grabar:=true;
      if Resp=MrNO then
      begin
        Connection.QryBusca.Active:=false;
        Connection.QryBusca.SQL.Text:='select * from tripulaciondiaria_listado' + #10 +
                                      'where sContrato=:Contrato and  sIdTurno=:Turno' + #10 +
                                      'and dIdFecha=:Fecha and sIdCategoria=:Categoria' + #10 +
                                      'and sIdTripulacion=:Tripulacion';
        if connection.contrato.FieldByName('sTipoObra').AsString='BARCO' then
          Connection.QryBusca.ParamByName('Contrato').AsString:=global_contrato
        else
          Connection.QryBusca.ParamByName('Contrato').AsString:=global_contrato_Barco;
        Connection.QryBusca.ParamByName('Turno').AsString:=Global_turno;
        Connection.QryBusca.ParamByName('Fecha').AsDate:=AvDtpFecha.Date;
        Connection.QryBusca.ParamByName('Categoria').AsString:=QTraerDatos.FieldByName('sIdCategoria').AsString;
        Connection.QryBusca.ParamByName('Tripulacion').AsString:=QTraerDatos.FieldByName('sIdTripulacion').AsString;
        Connection.QryBusca.Open;
        if Connection.QryBusca.RecordCount=1 then
          Grabar:=false;
      end;

      if Grabar then
      begin
        Connection.zCommand.Active:=false;
        Connection.zCommand.SQL.Text:='insert into tripulaciondiaria_listado ' + #10 + 
                                      '(sContrato,sIdTurno,dIdFecha,sIdCategoria,' + #10 +
                                      'sIdTripulacion,sOrden,sNombre,sIdpersonal,' + #10 +
                                      'sDescripcion,dSolicitado,sNacionalidad,' + #10 +
                                      'iNacionales,iExtranjeros,sIdCompania,dCantidad,sIdCuenta, sIdPernocta)' + #10 +
                                      'values(:Contrato,:Turno,:Fecha,:Categoria,' + #10 +
                                      ':Tripulacion,:Orden,:Nombre,:personal,' + #10 +
                                      ':Descripcion,:Solicitado,:Nacionalidad,1,0,' + #10 +
                                      ':Compania,:Cantidad,:Cuenta, :pernocta)';
        Connection.zCommand.ParamByName('Contrato').AsString:=QTraerDatos.fieldByName('sContrato').AsString;
        Connection.zCommand.ParamByName('Turno').AsString:=QTraerDatos.fieldByName('sIdTurno').AsString;
        Connection.zCommand.ParamByName('Fecha').AsDate:=AvDtpFecha.Date;
        Connection.zCommand.ParamByName('Categoria').AsString:=QTraerDatos.fieldByName('sIdCategoria').AsString;
        Connection.zCommand.ParamByName('Tripulacion').AsString:=QTraerDatos.fieldByName('sIdTripulacion').AsString;
        Connection.zCommand.ParamByName('Orden').AsString:=global_contrato;
        Connection.zCommand.ParamByName('Nombre').AsString:=QTraerDatos.fieldByName('sNombre').AsString;
        Connection.zCommand.ParamByName('personal').AsString:=QTraerDatos.fieldByName('sIdpersonal').AsString;
        Connection.zCommand.ParamByName('Descripcion').AsString:=QTraerDatos.fieldByName('sDescripcion').AsString;
        Connection.zCommand.ParamByName('Solicitado').AsFloat:=QTraerDatos.fieldByName('dSolicitado').AsFloat;
        Connection.zCommand.ParamByName('Nacionalidad').AsString:=QTraerDatos.fieldByName('sNacionalidad').AsString;
        Connection.zCommand.ParamByName('Compania').AsString:=QTraerDatos.fieldByName('sIdCompania').AsString;
        Connection.zCommand.ParamByName('Cantidad').AsFloat:=QTraerDatos.fieldByName('dCantidad').AsFloat;
        Connection.zCommand.ParamByName('Cuenta').AsString:=QTraerDatos.fieldByName('sIdCuenta').AsString;
        Connection.zCommand.ParamByName('pernocta').AsString:=QTraerDatos.fieldByName('sIdPernocta').AsString;
        Connection.zCommand.ExecSQL;
      end;

      QTraerDatos.Next;
    end;
  finally
    QTraerDatos.Destroy;
    QTripulacion.Refresh;
  end;

end;

procedure TfrmListaPersonalV2.AvDtpFechaExit(Sender: TObject);
begin
  if TmpFecha<>AvDtpFecha.Date then
  begin
    QrOrdenes.Active:=false;
    QrOrdenes.ParamByName('Fecha').AsDate:=AvDtpFecha.Date;


    QTripulacion.Active:=false;
    if connection.contrato.FieldByName('sTipoObra').AsString='BARCO' then
    begin
      QTripulacion.ParamByName('contrato').AsString:= global_contrato;
      QrOrdenes.ParamByName('Orden').AsString:= global_contrato;

    end
    else
    begin
      QTripulacion.ParamByName('contrato').AsString:= global_contrato_Barco;
      QrOrdenes.ParamByName('Orden').AsString:= global_contrato_Barco;
    end;
    QTripulacion.ParamByName('Fecha').AsDate         :=AvDtpFecha.Date;
    QTripulacion.ParamByName('FechaVigencia').AsDate :=dFechaVigencia;
    QTripulacion.ParamByName('Turno').AsString:=jDblCmbTurnos.KeyValue;
    if jDblCmbCategoria.KeyValue=null then
      QTripulacion.ParamByName('Categoria').AsInteger:=2123
    else
      QTripulacion.ParamByName('Categoria').AsString:=jDblCmbCategoria.KeyValue;
    QTripulacion.Open;
    QrOrdenes.Open;

    TmpFecha:=AvDtpFecha.Date;
  end;
end;

procedure TfrmListaPersonalV2.AvDtpFechaKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if key = 13 then
    AvDtpFechaExit(Sender);
end;

procedure TfrmListaPersonalV2.AvOChkBActualizaEditing(Sender: TObject;
  var CanEdit: Boolean);
begin
  if AvOChkBActualiza.Checked then
  begin
    if MessageDlg('Si Deshabilita esta opcion, el resumen de personal no se Actualizara.' + #13 + #10 + #13 +
                  '¿Desea Continuar?',mtConfirmation,[MbYes,MbNo],0)=mrNo then
      CanEdit:=false;
  end;
end;

procedure TfrmListaPersonalV2.cmdActualizarClick(Sender: TObject);
begin
   ProcPernoctas;
end;

procedure TfrmListaPersonalV2.btnDeleteAllClick(Sender: TObject);
begin
  try
    if (not QTripulacion.active) or (QTripulacion.RecordCount < 1) then
      exit;

    if (MessageDlg('¿Está seguro que desea eliminar todos los datos?', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
    begin
      QTripulacion.DisableControls;
      try
        QTripulacion.First;
        while not QTripulacion.Eof do
          QTripulacion.Delete;

        if AvOChkBActualiza.Checked then
        begin                           
            //Se elimina todo el personal de bitacoradepersonal_cuadre..
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('delete from bitacoradepersonal_cuadre where sContrato=:contrato and dIdFecha =:fecha ');
            connection.zCommand.ParamByName('contrato').AsString := global_contrato;
            connection.zCommand.ParamByName('fecha').AsDate      := AvDtpFecha.Date;
            connection.zCommand.ExecSQL;
        end;
      finally
        QTripulacion.EnableControls;
      end;
    end;
  except
    on e: Exception do
      MessageDlg('Ha ocurrido el siguiente error, informar al administrador del sistema: ' + e.Message, mtError, [mbOK], 0);
  end;

end;

procedure TfrmListaPersonalV2.btnDeleteClick(Sender: TObject);
var
   sOrden, sIdPersonal, sIdCuenta, sIdCategoria, sIdPernocta :string;
begin
  try
    if (not QTripulacion.active) or (QTripulacion.RecordCount < 1) then
      exit;

    if (MessageDlg('¿Está seguro que desea eliminar [' +  QTripulacion.FieldByName('sNombre').AsString + ']?', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
    begin
       sOrden       :=QTripulacion.FieldByName('sOrden').AsString;
       sIdPersonal  :=QTripulacion.FieldByName('sIdPersonal').AsString;
       sIdCuenta    :=QTripulacion.FieldByName('sIdCuenta').AsString;
       sIdCategoria :=QTripulacion.FieldByName('sIdCategoria').AsString;
       sIdPernocta  :=QTripulacion.FieldByName('sIdPernocta').AsString;

       QTripulacion.Delete;
       QTripulacion.Refresh;
       jDblCmbCategoria.KeyValue :=  sIdCategoria;

       if AvOChkBActualiza.Checked then
       begin
          connection.QryBusca.Active := False;
          connection.QryBusca.SQL.Clear;
          connection.QryBusca.SQL.Add('select sIdCategoria, sDescripcion,sum(iNacionales) as total from tripulaciondiaria_listado where sContrato=:Contrato and '+
                        'sIdTurno=:Turno and dIdFecha=:fecha and sOrden=:Orden and sIdPersonal=:Personal and sIdCuenta=:Cuenta and sIdCategoria =:Categoria and sIdPernocta =:Pernocta '+
                        'group by sIdPersonal');
          connection.QryBusca.ParamByName('Contrato').AsString  :=global_Contrato_Barco;
          connection.QryBusca.ParamByName('Turno').AsString     :=global_Turno;
          connection.QryBusca.ParamByName('Fecha').AsDate       :=AvDtpFecha.Date;
          connection.QryBusca.ParamByName('Orden').AsString     :=sOrden;
          connection.QryBusca.ParamByName('Personal').AsString  :=sIdPersonal;
          connection.QryBusca.ParamByName('Cuenta').AsString    :=sIdCuenta;
          connection.QryBusca.ParamByName('Categoria').AsString :=sIdCategoria;
          connection.QryBusca.ParamByName('Pernocta').AsString  :=sIdPernocta;
          connection.QryBusca.Open;

          if connection.QryBusca.RecordCount > 0 then
          begin
               //Se elimina todo el personal de bitacoradepersonal_cuadre..
             connection.zCommand.Active := False;
             connection.zCommand.SQL.Clear;
             connection.zCommand.SQL.Add('update bitacoradepersonal_cuadre set dCantidad =:cantidad where sContrato=:contrato and dIdFecha =:fecha and sIdPersonal =:Id and sAgrupaPersonal =:Categoria and sTipoPernocta =:Tipo and sIdPernocta =:Pernocta ');
             connection.zCommand.ParamByName('contrato').AsString  := global_contrato;
             connection.zCommand.ParamByName('fecha').AsDate       := AvDtpFecha.Date;
             connection.zCommand.ParamByName('Id').AsString        := sIdPersonal;
             connection.zCommand.ParamByName('Categoria').AsString := sIdCategoria;
             connection.zCommand.ParamByName('Tipo').AsString      := sIdCuenta;
             connection.zCommand.ParamByName('Pernocta').AsString  := sIdPernocta;
             connection.zCommand.ParamByName('Cantidad').AsFloat   := connection.QryBusca.FieldByName('total').AsFloat;
             connection.zCommand.ExecSQL;

          end
          else
          begin
             //Se elimina todo el personal de bitacoradepersonal_cuadre..
             connection.zCommand.Active := False;
             connection.zCommand.SQL.Clear;
             connection.zCommand.SQL.Add('delete from bitacoradepersonal_cuadre where sContrato=:contrato and dIdFecha =:fecha and sIdPersonal =:Id and sAgrupaPersonal =:Categoria and sTipoPernocta =:pernocta ');
             connection.zCommand.ParamByName('contrato').AsString  := global_contrato;
             connection.zCommand.ParamByName('fecha').AsDate       := AvDtpFecha.Date;
             connection.zCommand.ParamByName('Id').AsString        := sIdPersonal;
             connection.zCommand.ParamByName('Categoria').AsString := sIdCategoria;
             connection.zCommand.ParamByName('Pernocta').AsString  := sIdCuenta;
             connection.zCommand.ExecSQL;
          end;
       end;

    end;
  except
    on e: Exception do
      MessageDlg('Ha ocurrido el siguiente error, informar al administrador del sistema: ' + e.Message, mtError, [mbOK], 0);
  end;

end;

procedure TfrmListaPersonalV2.btnExportarClick(Sender: TObject);
const
  NombreCols: Array[1..14] of String = ('Ficha', 'Nombre Completo', 'Grupo', 'Id Grupo', 'Categoria', 'Id Categoria', 'Cantidad', 'OT', 'Compañia', 'Id Compañia', 'Solicitado', 'Tipo de Pernocta', 'Id Pernocta', 'Pernocta');
  NombreFields: Array[1..14] of string = ('sidTripulacion', 'sNombre', 'grupo', 'sIdCategoria', 'sDescripcion', 'sIdPersonal', 'dCantidad', 'sOrden', '', 'sIdCompania', 'dSolicitado', '', 'sidcuenta', 'pernocta2');
var
  Excel, workSheet, WorkBook: Variant;
  Cursor: TCursor;
  i, j: Integer;
  ExcepcionesFields: TStringList;

Procedure createComboExcel(Var Hoja: Variant; NombreHoja: String; PosCol: String; ListaDatos:string);
begin
  Hoja.Sheets[NombreHoja].Select;
  Hoja.Range[PosCol].Select;
  hoja.Selection.Validation.Delete;
  hoja.Selection.Validation.add(xlValidateList,xlValidAlertStop,xlBetween,ListaDatos);
  hoja.Selection.Validation.IgnoreBlank := True;
  hoja.Selection.Validation.InCellDropdown := True;
  hoja.Selection.Validation.InputTitle := '';
  hoja.Selection.Validation.ErrorTitle := '';
  hoja.Selection.Validation.ErrorMessage := '';
  hoja.Selection.Validation.ShowInput := True;
  hoja.Selection.Validation.ShowError := True;
end;
Procedure AsignaFormulas(var Hoja: Variant; NombreHoja: string; Celda: String; Formula: String; Rango_AutoFill: string;
                         sLocked: Boolean; sFormulaOculta: Boolean; OcultarColumna: String);
begin
  Hoja.Sheets[NombreHoja].Select;
  Hoja.Range[Celda].Select;
  Hoja.Selection.FormulaR1C1 := Formula;
  if Length(Trim(Rango_AutoFill)) > 0 then
    Hoja.Selection.Autofill(Hoja.range[Rango_AutoFill], xlFillDefault);
  if Length(Trim(OcultarColumna)) > 0 then
  begin
    Hoja.Columns[OcultarColumna].select;
    Hoja.Selection.Locked := sLocked;
    Hoja.Selection.FormulaHidden := sFormulaOculta;
    Hoja.Selection.EntireColumn.hidden := sFormulaOculta;
  end;
end;

begin
  try
    //Trae todos las pernoctas que existan
    with connection.zCommand do
    begin
      Active := False;
      SQL.Clear;
      SQL.Add('select sIdPernocta from pernoctan');
      Open;
    end;

    //Traen todas los GRUPOS(categorias) que existan
    with connection.QryBusca2 do
    begin
      Active := False;
      SQL.Clear;
      SQL.Add('select sIdCategoria, sDescripcion from categorias');
      Open;
    end;

    i := 0;
    Cursor := Screen.Cursor;

    try
      //Columnas descartadas
      ExcepcionesFields := TStringList.Create;
      ExcepcionesFields.Add('9');
      ExcepcionesFields.Add('12');

      Excel := CreateOleObject('Excel.Application');
      Excel.Visible := True;
      Excel.WorkBooks.Add(xlWBATWorksheet);

      //Columnas Excel LookUpComboExcel Pernoctas
      Excel.ActiveSheet.Name := leftStr('Pernoctas', 31);
      Excel.Sheets['Pernoctas'].Select;

      connection.zCommand.First;

      while Not connection.zCommand.Eof do
      begin
        Excel.Cells[connection.zCommand.RecNo,1] := connection.zCommand.FieldByName('sIdPernocta').AsString;
        connection.zCommand.Next;
      end;

      //Columnas Excel LookUpComboExcel Grupos
      Excel.sheets.add;
      Excel.ActiveSheet.Name := leftStr('Grupos', 31);
      Excel.Sheets['Grupos'].Select;
      connection.QryBusca2.First;
      while Not connection.QryBusca2.Eof do
      begin
        Excel.Cells[connection.QryBusca2.RecNo,2] := connection.QryBusca2.FieldByName('sIdCategoria').AsString;
        Excel.Cells[connection.QryBusca2.RecNo,1] := connection.QryBusca2.FieldByName('sDescripcion').AsString;
        connection.QryBusca2.Next;
      end;

      //Columnas Excel LookUpComboExcel Tipos Pernoctas
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('TiposPernoctas',31);
      Excel.Sheets['TiposPernoctas'].Select;
      QrPernoctas.First;
      while Not QrPernoctas.Eof do
      begin
        Excel.Cells[QrPernoctas.RecNo,2] := QrPernoctas.FieldByName('sIdCuenta').AsString;
        Excel.Cells[QrPernoctas.RecNo,1] := QrPernoctas.FieldByName('sDescripcion').AsString;
        QrPernoctas.Next;
      end;

      with connection.zCommand do
      begin
        Active := False;
        SQL.Clear;
        SQL.Add('select snumeroorden from reportediario ' +
                'group by sNumeroOrden');
        Open;
      end;

      //Columnas Excel LookUpComboExcel Ordenes de Trabajo
      Excel.sheets.add;
      Excel.ActiveSheet.Name := leftStr('Ordenes', 31);
      Excel.Sheets['Ordenes'].Select;
      connection.zCommand.First;
      while Not connection.zCommand.Eof do
      begin
        Excel.Cells[connection.zCommand.RecNo,1] := connection.zCommand.FieldByName('snumeroorden').AsString;
        connection.zCommand.Next;
      end;

       //Columnas Excel LookUpComboExcel Companias
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('Compañias',31);
      Excel.Sheets['Compañias'].Select;
      QrCompanias.First;
      while Not QrCompanias.Eof do
      begin
        Excel.Cells[QrCompanias.RecNo,2] := QrCompanias.FieldByName('sIdCompania').AsString;
        Excel.Cells[QrCompanias.RecNo,1] := QrCompanias.FieldByName('sDescripcion').AsString;
        QrCompanias.Next;
      end;

      //Columnas Excel LookUpComboExcel Categorias
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('Categorias',31);
      Excel.Sheets['Categorias'].Select;
      zq_Esp.Active:=False;
      zq_Esp.ParamByName('Contrato').AsString:=global_contrato_barco;
      zq_Esp.ParamByName('Per').AsString:='-1';
      zq_Esp.Open;
      zq_Esp.First;
      while Not zq_Esp.Eof do
      begin
        Excel.Cells[zq_Esp.RecNo,2] := zq_Esp.FieldByName('sIdPersonal').AsString;
        Excel.Cells[zq_Esp.RecNo,1] := zq_Esp.FieldByName('sDescripcion').AsString;
        zq_Esp.Next;
      end;

      //Llenar el La plantilla de Excel
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('Plantilla',31);
      Excel.Sheets['Plantilla'].Select;

      for i  := 1 to 14 do
        Excel.Cells[1, i] := nombreCols[i];

      createComboExcel(Excel, 'Plantilla', 'C2', '=Grupos!$A:$A');
      createComboExcel(Excel, 'Plantilla', 'E2', '=Categorias!$A:$A');
      createComboExcel(Excel, 'Plantilla', 'H2', '=Ordenes!$A:$A');
      createComboExcel(Excel, 'Plantilla', 'I2', '=Compañias!$A:$A');
      createComboExcel(Excel, 'Plantilla', 'L2', '=TiposPernoctas!$A:$A');
      createComboExcel(Excel, 'Plantilla', 'N2', '=Pernoctas!$A:$A');

      //Comentarios
      Excel.Cells[QTripulacion.RecNo + 1, 4].AddComment;
      Excel.Cells[QTripulacion.RecNo + 1, 4].Comment.Visible := False;
      Excel.Cells[QTripulacion.RecNo + 1, 4].Comment.Text('No tipiar en esta columna, ya que se llena en base al Grupo que elija');

      Excel.Cells[QTripulacion.RecNo + 1, 6].AddComment;
      Excel.Cells[QTripulacion.RecNo + 1, 6].Comment.Visible := False;
      Excel.Cells[QTripulacion.RecNo + 1, 6].Comment.Text('No tipiar en esta columna, ya que se llena en base a la Categoría que elija');

      Excel.Cells[QTripulacion.RecNo + 1, 10].AddComment;
      Excel.Cells[QTripulacion.RecNo + 1, 10].Comment.Visible := False;
      Excel.Cells[QTripulacion.RecNo + 1, 10].Comment.Text('No tipiar en esta columna, ya que se llena en base a la Compañía que elija');

      Excel.Cells[QTripulacion.RecNo + 1, 13].AddComment;
      Excel.Cells[QTripulacion.RecNo + 1, 13].Comment.Visible := False;
      Excel.Cells[QTripulacion.RecNo + 1, 13].Comment.Text('No tipiar en esta columna, ya que se llena en base al Tipo de Pernocta que elija');

      AsignaFormulas(Excel, 'Plantilla', 'D2', '=VLOOKUP(RC[-1], Grupos!C[-4]:C[-3], 2, FALSE)', 'D2:' + 'D' + IntToStr(3), false, False, 'D:D');
      AsignaFormulas(Excel, 'Plantilla', 'F2', '=VLOOKUP(RC[-1], Categorias!C[-5]:C[-3], 2, FALSE)', 'F2:' + 'F' + IntToStr(3), false, False, 'F:F');
      AsignaFormulas(Excel, 'Plantilla', 'J2', '=VLOOKUP(RC[-1],Compañias!C[-9]:C[-7],2,FALSE)', 'J2:' + 'J' + IntToStr(3), false, False, 'J:J');
      AsignaFormulas(Excel, 'Plantilla', 'M2', '=VLOOKUP(RC[-1],TiposPernoctas!C[-12]:C[-10],2,FALSE)', 'M2:' + 'M' + IntToStr(3), false, False, 'M:M');

      //Formato
      Excel.Range['A:N'].NumberFormat := '@';
      Excel.Range['D:D'].Style := 'Énfasis2';
      Excel.Range['F:F'].Style := 'Énfasis2';
      Excel.Range['J:J'].Style := 'Énfasis2';
      Excel.Range['M:M'].Style := 'Énfasis2';
      Excel.Range['A1:N1'].select;
      Excel.Range['A1:N1'].Style := 'Énfasis1';
      Excel.Range['A1:N1'].Font.Bold := true;
      Excel.Range['A1:N1'].HorizontalAlignment := xlCenter;
      Excel.Range['A:N'].columns.autofit;
      MessageDlg('La Informacion se Exporto Correctamente.',mtInformation,[MbOk],0);

    finally
      Screen.Cursor :=     Cursor;
      QrCompanias.Filtered:=False;
      QrPernoctas.Filtered:=False;
    end;
  except
    on e: Exception do
    begin
      Excel.Quit;
      Excel := Null;
      ShowMessage(e.Message);
    end;
  end;
end;

procedure TfrmListaPersonalV2.btnExportarDtosClick(Sender: TObject);
const
  NombreCols: Array[1..13] of String = ('Ficha', 'Nombre Completo', 'Grupo','Categoria', 'Id Categoria', 'Cantidad', 'OT', 'Compañia', 'Id Compañia', 'Solicitado', 'Tipo de Pernocta', 'Id Pernocta', 'Pernocta');
  NombreFields: Array[1..13] of string = ('sidTripulacion', 'sNombre', 'grupo','sDescripcion', 'sIdPersonal', 'dCantidad', 'sOrden', '', 'sIdCompania', 'dSolicitado', '', 'sidcuenta', 'pernocta2');
var
  Excel, workSheet, WorkBook: Variant;
  Cursor: TCursor;
  i, j: Integer;
  ExcepcionesFields: TStringList;

Procedure createComboExcel(Var Hoja: Variant; NombreHoja: String; PosCol: String; ListaDatos:string);
begin
  Hoja.Sheets[NombreHoja].Select;
  Hoja.Range[PosCol].Select;
  hoja.Selection.Validation.Delete;
  hoja.Selection.Validation.add(xlValidateList,xlValidAlertStop,xlBetween,ListaDatos);
  hoja.Selection.Validation.IgnoreBlank := True;
  hoja.Selection.Validation.InCellDropdown := True;
  hoja.Selection.Validation.InputTitle := '';
  hoja.Selection.Validation.ErrorTitle := '';
  hoja.Selection.Validation.ErrorMessage := '';
  hoja.Selection.Validation.ShowInput := True;
  hoja.Selection.Validation.ShowError := True;
end;
Procedure AsignaFormulas(var Hoja: Variant; NombreHoja: string; Celda: String; Formula: String; Rango_AutoFill: string;
                         sLocked: Boolean; sFormulaOculta: Boolean; OcultarColumna: String);
begin
  Hoja.Sheets[NombreHoja].Select;
  Hoja.Range[Celda].Select;
  Hoja.Selection.FormulaR1C1 := Formula;
  if Length(Trim(Rango_AutoFill)) > 0 then
    Hoja.Selection.Autofill(Hoja.range[Rango_AutoFill], xlFillDefault);
  if Length(Trim(OcultarColumna)) > 0 then
  begin
    Hoja.Columns[OcultarColumna].select;
    Hoja.Selection.Locked := sLocked;
    Hoja.Selection.FormulaHidden := sFormulaOculta;
    Hoja.Selection.EntireColumn.hidden := sFormulaOculta;
  end;
end;

begin
  try
    i := 0;
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;
    if (Not QTripulacion.Active) or (QTripulacion.RecordCount = 0) then
    begin
      MessageDlg('No existen datos a exportar, Filtrar por fecha.',mtWarning,[mbOK], 0);
      Exit;
    end;
    try
      //Columnas descartadas
      ExcepcionesFields := TStringList.Create;
      ExcepcionesFields.Add('8');
      ExcepcionesFields.Add('11');

      Excel := CreateOleObject('Excel.Application');
      Excel.Visible := True;
      Excel.WorkBooks.Add(xlWBATWorksheet);
      Excel.ActiveSheet.Name := leftStr('Ordenes', 31);

      //Columnas Excel LookUpComboExcel Ordenes de Trabajo
      Excel.Sheets['Ordenes'].Select;
      QrOrdenes.First;
      while Not QrOrdenes.Eof do
      begin
        Excel.Cells[QrOrdenes.RecNo,1] := QrOrdenes.FieldByName('snumeroorden').AsString;
        QrOrdenes.Next;
      end;

       //Columnas Excel LookUpComboExcel Companias
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('Compañias',31);
      Excel.Sheets['Compañias'].Select;
      QrCompanias.First;
      while Not QrCompanias.Eof do
      begin
        Excel.Cells[QrCompanias.RecNo,2] := QrCompanias.FieldByName('sIdCompania').AsString;
        Excel.Cells[QrCompanias.RecNo,1] := QrCompanias.FieldByName('sDescripcion').AsString;
        QrCompanias.Next;
      end;

      zq_Esp.Active:=False;
      zq_Esp.ParamByName('Contrato').AsString:=global_contrato_barco;
      zq_Esp.ParamByName('Per').AsString:='-1';
      zq_Esp.Open;

      //Columnas Excel LookUpComboExcel Categorias
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('Categorias',31);
      Excel.Sheets['Categorias'].Select;
      zq_Esp.First;
      while Not zq_Esp.Eof do
      begin
        Excel.Cells[zq_Esp.RecNo,2] := zq_Esp.FieldByName('sIdPersonal').AsString;
        Excel.Cells[zq_Esp.RecNo,1] := zq_Esp.FieldByName('sDescripcion').AsString;
        zq_Esp.Next;
      end;

      //Columnas Excel LookUpComboExcel Categorias
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('Pernoctas',31);
      Excel.Sheets['Pernoctas'].Select;
      QrPernoctas.First;
      while Not QrPernoctas.Eof do
      begin
        Excel.Cells[QrPernoctas.RecNo,2] := QrPernoctas.FieldByName('sIdCuenta').AsString;
        Excel.Cells[QrPernoctas.RecNo,1] := QrPernoctas.FieldByName('sDescripcion').AsString;
        QrPernoctas.Next;
      end;

      //Llenar el La plantilla de Excel
      Excel.sheets.add;
      Excel.ActiveSheet.Name := LeftStr('Plantilla',31);
      Excel.Sheets['Plantilla'].Select;

      for i  := 1 to 13 do
        Excel.Cells[1, i] := nombreCols[i];

      if (QTripulacion.Active) and (QTripulacion.RecordCount > 0) then
      begin
        QTripulacion.First;
        while not QTripulacion.Eof  do
        begin
          for i := 1 to 13 do
            for j := 0 to ExcepcionesFields.Count - 1 do
              if ExcepcionesFields[j] <> IntToStr(i) then
                if Length(trim(NombreFields[i])) > 0 then
                  Excel.Cells[QTripulacion.RecNo + 1, i] := QTripulacion.FieldByName(NombreFields[i]).AsString
                else
                  Excel.Cells[QTripulacion.RecNo + 1, i] := '';
          createComboExcel(Excel, 'Plantilla', 'D' + IntToStr(QTripulacion.RecNo + 1), '=Categorias!$A:$A');
          createComboExcel(Excel, 'Plantilla', 'G' + IntToStr(QTripulacion.RecNo + 1), '=Ordenes!$A:$A');
          createComboExcel(Excel, 'Plantilla', 'H' + IntToStr(QTripulacion.RecNo + 1), '=Compañias!$A:$A');
          createComboExcel(Excel, 'Plantilla', 'I' + IntToStr(QTripulacion.RecNo + 1), '=Pernoctas!$A:$A');

          QrCompanias.Filtered:=False;
          QrCompanias.Filter:='sIdCompania='+QuotedStr(QTripulacion.FieldByName('sIdCompania').AsString);
          QrCompanias.Filtered:=True;

          QrPernoctas.Filtered:=False;
          QrPernoctas.Filter:='sidcuenta='+QuotedStr(QTripulacion.FieldByName('sidcuenta').AsString);
          QrPernoctas.Filtered:=True;

          Excel.Cells[QTripulacion.RecNo + 1, 8] := QrCompanias.FieldByName('sDescripcion').AsString;
          Excel.Cells[QTripulacion.RecNo + 1, 11] := QrPernoctas.FieldByName('sDescripcion').AsString;

          //Comentarios
          Excel.Cells[QTripulacion.RecNo + 1, 5].AddComment;
          Excel.Cells[QTripulacion.RecNo + 1, 5].Comment.Visible := False;
          Excel.Cells[QTripulacion.RecNo + 1, 5].Comment.Text('No tipiar en esta columna, ya que se llena en base a la Categoría que elija');

          Excel.Cells[QTripulacion.RecNo + 1, 9].AddComment;
          Excel.Cells[QTripulacion.RecNo + 1, 9].Comment.Visible := False;
          Excel.Cells[QTripulacion.RecNo + 1, 9].Comment.Text('No tipiar en esta columna, ya que se llena en base a la Compañía que elija');

          Excel.Cells[QTripulacion.RecNo + 1, 12].AddComment;
          Excel.Cells[QTripulacion.RecNo + 1, 12].Comment.Visible := False;
          Excel.Cells[QTripulacion.RecNo + 1, 12].Comment.Text('No tipiar en esta columna, ya que se llena en base a la Tipo de Pernocta que elija');

          QTripulacion.Next;
        end;
      end;
      if QTripulacion.RecNo=1 then
      begin
        AsignaFormulas(Excel, 'Plantilla', 'E2', '=VLOOKUP(RC[-1], Categorias!C[-4]:C[-3], 2, FALSE)', 'E2:' + 'E' + IntToStr(3), false, False, 'E:E');
        AsignaFormulas(Excel, 'Plantilla', 'I2', '=VLOOKUP(RC[-1],Compañias!C[-8]:C[-7],2,FALSE)', 'I2:' + 'I' + IntToStr(3), false, False, 'I:I');
        AsignaFormulas(Excel, 'Plantilla', 'L2', '=VLOOKUP(RC[-1],Pernoctas!C[-11]:C[-10],2,FALSE)', 'L2:' + 'L' + IntToStr(3), false, False, 'L:L');

      end
      else
      begin
        AsignaFormulas(Excel, 'Plantilla', 'E2', '=VLOOKUP(RC[-1], Categorias!C[-4]:C[-3], 2, FALSE)', 'E2:' + 'E' + IntToStr(QTripulacion.RecNo + 1), false, False, 'E:E');
        AsignaFormulas(Excel, 'Plantilla', 'I2', '=VLOOKUP(RC[-1],Compañias!C[-8]:C[-7],2,FALSE)', 'I2:' + 'I' + IntToStr(QTripulacion.RecNo + 1), false, False, 'I:I');
        AsignaFormulas(Excel, 'Plantilla', 'L2', '=VLOOKUP(RC[-1],Pernoctas!C[-11]:C[-10],2,FALSE)', 'L2:' + 'L' + IntToStr(QTripulacion.RecNo + 1), false, False, 'L:L');



      end;
      //Formato
      Excel.Range['A:M'].NumberFormat := '@';
      Excel.Range['E:E'].Style := 'Énfasis2';
      Excel.Range['I:I'].Style := 'Énfasis2';
      Excel.Range['L:L'].Style := 'Énfasis2';
      Excel.Range['A1:M1'].select;
      Excel.Range['A1:M1'].Style := 'Énfasis1';
      Excel.Range['A1:M1'].Font.Bold := true;
      Excel.Range['A1:M1'].HorizontalAlignment := xlCenter;
      Excel.Range['A:M'].columns.autofit;
      MessageDlg('La Informacion se Exporto Correctamente.',mtInformation,[MbOk],0);

    finally
      Screen.Cursor := Cursor;
      QrCompanias.Filtered:=False;
      QrPernoctas.Filtered:=False;
    end;
  except
    on e: Exception do
    begin
      Excel.Quit;
      Excel := Null;
      ShowMessage(e.Message);
    end;
  end;
end;

procedure TfrmListaPersonalV2.btnNuevoClick(Sender: TObject);
begin
  if jDblCmbCategoria.KeyValue=null then
  begin
    MessageDlg('Debe Elegir Un Grupo de Categoria.',mtInformation,[MbOk],0);
    jDblCmbCategoria.SetFocus;
  end
  else
  begin
    application.CreateForm(TFrmBuscaPersonal,FrmBuscaPersonal);
    try
      FrmBuscaPersonal.ParamFecha:=AvDtpFecha.Date;
      FrmBuscaPersonal.sCategoria:=QrCategorias.FieldByName('sIdCategoria').AsString;
      FrmBuscaPersonal.QExterno:=QTripulacion;
      FrmBuscaPersonal.showmodal;
    finally
      FrmBuscaPersonal.Destroy;
    end;
  end;
end;

procedure TfrmListaPersonalV2.cmdImportarClick(Sender: TObject);
var
  x, y, i: Integer;
  turno, sIdPersonal, ficha, nombre, grupo, idGrupo, categoria, idCategoria, rfc, OT, compania, idCompania, noCabina, tipoPernocta, idPernocta, pernocta: string;
  QryTriList:TzReadOnlyQuery;
  datosExistenEnBD : Boolean;
begin
  OpenXLS.Title := 'Inserta Archivo de Consulta';
  if OpenXLS.Execute then
  begin
    tsArchivo.Text := OpenXLS.FileName;

    try
      flcid := GetUserDefaultLCID;
      ExcelApplication1.Connect;
      ExcelApplication1.Visible[flcid] := true;
      ExcelApplication1.UserControl := true;

      ExcelWorkbook1.ConnectTo(ExcelApplication1.Workbooks.Open(tsArchivo.Text,
                                       emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam,
      emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, flcid));

      ExcelWorksheet1.ConnectTo(ExcelWorkbook1.Sheets.Item[1] as ExcelWorkSheet);
      try
        QryTriList:=TzREadOnlyQuery.Create(nil);
        QryTriList.Connection:=connection.zConnection;

        Fila := 2;
        datosExistenEnBD := True;
        //Checar si la pernocta existe en la tabla pernoctan, sino existe, entonces no subir el archivo hasta que el usuario ponga
        //una pernocta que si existe en la BD
        ficha       := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;      //Ficha Empleado
        while ficha <> '' do
        begin
          OT          := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;               //OT
          pernocta    := ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Value2;         //Pernocta
          Compania    := ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2;         //Compania
          idCompania    := ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;      //Id Compania

          //Poner fondo blanco en las celdas de OT y Pernoctas(si hay un valor no valido, mas abajo lo pondra de rojo)
          ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Interior.Pattern := -4142;
          ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Interior.Pattern := -4142;
          ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Interior.Pattern := -4142;

          with connection.QryBusca2 do
          begin
            Active := False;
            SQL.Clear;
            SQL.Add('select sContrato from contratos where sContrato = :sContrato');
            Params.ParamByName('sContrato').AsString := OT;
            Open;
          end;

          with connection.qryBusca do
          begin
            Active := False;
            SQL.Clear;
            SQL.Add('select sIdCompania, sDescripcion from compersonal where sIdCompania = :sIdCompania and sDescripcion = :sDescripcion');
            Params.ParamByName('sIdCompania').AsString  := idCompania;
            Params.ParamByName('sDescripcion').AsString :=   Compania;
            Open;
          end;

          with connection.zcommand do
          begin
            Active := False;
            SQL.Clear;
            SQL.Add('select sIdPernocta from pernoctan where sIdPernocta = :sIdPernocta');
            Params.ParamByName('sIdPernocta').AsString := pernocta;
            Open;
          end;

          if connection.QryBusca2.RecordCount = 0 then           //Checar si la OT existe en la tabla contratos, pintarla de rojo
          begin
            datosExistenEnBD := False;
            ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Interior.ColorIndex := 3;
          end;

          if connection.QryBusca.RecordCount = 0 then           //Checar si la OT existe en la tabla contratos, pintarla de rojo
          begin
            datosExistenEnBD := False;
            ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Interior.ColorIndex := 3;
          end;

          if connection.zCommand.RecordCount = 0 then      //Checar si la Pernocta existe en la tabla pernoctan, pintarla de rojo
          begin
            datosExistenEnBD := False;
            ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Interior.ColorIndex := 3;
          end;
          Fila := Fila + 1;
          ficha    := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;      //Ficha Empleado
        end;

        if datosExistenEnBD = True then
        begin

          Fila := 2;

            //Procedemos a leer el archivo de Excel..
            ficha       := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;  //Ficha Empleado
          while (ficha <> '') do
          begin
            nombre      := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2; //Nombre Completo
            grupo       := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;           //Grupo
            idGrupo     := ExcelWorksheet1.Range['D' + Trim(IntToStr(Fila)), 'D' + Trim(IntToStr(Fila))].Value2;        //Id Grupo
            categoria   := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;       //Categoria
            idCategoria := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;    //id Categoria
            rfc         := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;             //RFC
            OT          := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;              //OT
            Compania    := ExcelWorksheet1.Range['I' + Trim(IntToStr(Fila)), 'I' + Trim(IntToStr(Fila))].Value2;        //Compania
            idCompania  := ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;     //id Compania
            noCabina    := ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2;      //No. Cabina
            tipoPernocta:= ExcelWorksheet1.Range['L' + Trim(IntToStr(Fila)), 'L' + Trim(IntToStr(Fila))].Value2;//Tipo de Pernocta
            idPernocta  := ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;  //Id de Pernocta
            pernocta    := ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Value2;        //Pernocta

            turno        :=       QrTurnos.FieldByName('sIdTurno').AsString;

            QryTriList.Active:=False;
            QryTriList.SQL.Clear;
            QryTriList.SQL.Add('select tl.*, CONCAT(sNombre, " ", sApellidoP, " ", sApellidoM) as nombre from tripulacion_listado tl where tl.sIdTripulacion=:trip');
            QryTriList.ParamByName('trip').AsString:=ficha;
            QryTriList.Open;

            with connection.QryBusca do
            begin
              Active := False;
              SQL.Clear;
              SQL.Add('select sNombre from tripulaciondiaria_listado where ' +
                      'sContrato = :sContrato and sIdTurno = :sIdTurno and ' +
                'dIdFecha = :dIdFecha and sIdCategoria = :sIdCategoria and ' +
                     'sIdTripulacion = :sIdTripulacion and sOrden = :sOrden');
              Params.ParamByName('sContrato').AsString:=global_contrato_barco;
              Params.ParamByName('sIdTurno').AsString       :=          turno;
              Params.ParamByName('dIdFecha').AsDateTime     :=           Date;
              Params.ParamByName('sIdCategoria').AsString   :=        idGrupo;
              Params.ParamByName('sIdTripulacion').AsString :=          ficha;
              Params.ParamByName('sOrden').AsString         :=             OT;
              Open;
            end;

            if connection.QryBusca.RecordCount > 0 then
              QTripulacion.Edit                     //Poner en modo de edicion
            else
              QTripulacion.Insert;                //Poner en modo de insercion

              sIdPersonal  :=  QryTriList.FieldByName('sIdPersonal').AsString;
            with QTripulacion do
            begin
              FieldByName('sContrato').AsString      := global_contrato_barco;
              FieldByName('sIdTurno').AsString       :=                 turno;
              FieldByName('dIdFecha').AsDateTime     :=                  Date;
              FieldByName('sIdPersonal').AsString    :=           idCategoria;
              FieldByName('sNacionalidad').AsString  :=                    '';
              FieldByName('iNacionales').AsInteger   :=                     1;
              FieldByName('iExtranjeros').AsInteger  :=                     0;
              FieldByName('sIdPernocta').AsString    :=                    '';
              FieldByName('lImprimeListado').AsString:=                  'Si';
              FieldByName('sidTripulacion').AsString :=                 ficha;
              FieldByName('sNombre').AsString        :=                nombre;
              FieldByName('grupo').AsString          :=                 grupo;
              FieldByName('sDescripcion').AsString   :=             categoria;
              FieldByName('sIdCategoria').AsString   :=               idGrupo;
              FieldByName('dCantidad').AsFloat       :=                   StrToFloat(rfc);
              FieldByName('sOrden').AsString         :=                    OT;
              FieldByName('sIdCompania').AsString    :=            idCompania;
              FieldByName('dSolicitado').AsFloat     :=              StrToFloat(noCabina);
              FieldByName('sIdCuenta').AsString      :=            idPernocta;
              FieldByName('pernocta2').AsString      :=              pernocta;
            end;

            QTripulacion.Post;
            QTripulacion.ApplyUpdates;

            if AvOChkBActualiza.Checked then
              ActualizarResumen(OT,QryTriList.FieldByName('sIdPersonal').AsString,'1',AvDtpFecha.Date);

            Fila        := Fila + 1;
            ficha       :=ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;   //Ficha Empleado
          end;
          TmpFecha:=IncDay(AvDtpFecha.Date,-1);
          AvDtpFecha.OnExit(sender);
          MessageDlg('Proceso Terminado con exito...', mtInformation, [mbOk], 0);
        end;
      finally
        QryTriList.Destroy;
      end;
    except
      on E: Exception do
      begin
        MessageDlg(e.Message, mtInformation, [mbOk], 0)
      end;
    end;
    //Desconectamos la aplicacion..
    ExcelApplication1.Disconnect;
  end;
end;

procedure TfrmListaPersonalV2.cmdImportarListaClick(Sender: TObject);
var
  x, y, i: Integer;

  sNombre, sDescripcion, sIdPersonal, Cantidad, sOrden, sIdCompania, Solicitado,
  sIdCuenta, sValue, sGrupo, sPernocta, sCadena: string;

  QryTriList   :TzReadOnlyQuery;
  sobrescribir, lNuevaCategoria : Boolean;

  //Variables para el archivo .sql que genera el proceso.
  archivo  : TextFile;
  eliminar : string;
  ruta     : string;

begin
  OpenXLS.Title := 'Inserta Archivo de Consulta';
  if OpenXLS.Execute then
  begin
    tsArchivo.Text := OpenXLS.FileName;

    try
      if (Not QTripulacion.Active) and (QTripulacion.RecordCount = 0) then
      begin
        MessageDlg('No existen datos, Filtrar por fecha.',mtWarning,[mbOK], 0);
        Exit;
      end;

      flcid := GetUserDefaultLCID;
      ExcelApplication1.Connect;
      ExcelApplication1.Visible[flcid] := true;
      ExcelApplication1.UserControl := true;

      ExcelWorkbook1.ConnectTo(ExcelApplication1.Workbooks.Open(tsArchivo.Text,
      emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam,
      emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, emptyParam, flcid));

      if (MessageDlg('Desea Reemplazar los Datos Existenes?', mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
      begin
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('delete from tripulaciondiaria_listado where sContrato =:Contrato And dIdFecha =:Fecha '+
                                    'and sIdTurno =:Turno and (:Categoria=2123 or (:Categoria<>2123 and sIdCategoria=:Categoria))');
        Connection.zCommand.Params.ParamByName('Contrato').DataType  := ftString;
        Connection.zCommand.Params.ParamByName('Contrato').Value     := global_contrato_Barco;
        Connection.zCommand.Params.ParamByName('Fecha').DataType     := ftDate;
        Connection.zCommand.Params.ParamByName('Fecha').Value        := AvDtpFecha.Date;
        Connection.zCommand.Params.ParamByName('Turno').DataType     := ftString;
        Connection.zCommand.Params.ParamByName('Turno').Value        := jDblCmbTurnos.KeyValue;
        Connection.zCommand.Params.ParamByName('Categoria').DataType := ftString;
        Connection.zCommand.Params.ParamByName('Categoria').Value    := jDblCmbCategoria.KeyValue;
        connection.zCommand.ExecSQL();
        sobrescribir:=True;
      end
      else
        sobrescribir:=False;

      ExcelWorksheet1.ConnectTo(ExcelWorkbook1.Sheets.Item[1] as ExcelWorkSheet);
      try
        QryTriList:=TzREadOnlyQuery.Create(nil);
        QryTriList.Connection:=connection.zConnection;

        //Creamos el archivos para importar los datos..
        eliminar := ExtractFilePath(Application.ExeName) + 'tripulacion.sql';
        if FileExists(eliminar) then
        DeleteFile(eliminar);

        AssignFile(archivo, ExtractFilePath(Application.ExeName) + 'tripulacion.sql');
        Rewrite(archivo);

        Fila    := 2;
        sCadena := '';

        Writeln(archivo, 'INSERT INTO tripulaciondiaria_listado ( sContrato , sIdTurno, dIdFecha, sIdCategoria, sIdTripulacion, sOrden, sNombre, sIdPersonal, sDescripcion, '+
                         'dSolicitado, sNacionalidad, iNacionales, iExtranjeros, sIdCompania, dCantidad, sIdCuenta, sIdPernocta) VALUES ');

        //Procedemos a leer el archivo de Excel..
        sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
        while (sValue <> '') do
        begin
            if sCadena <> '' then
               Writeln(archivo, ''+ Char(44));

            sNombre       := ExcelWorksheet1.Range['B' + Trim(IntToStr(Fila)), 'B' + Trim(IntToStr(Fila))].Value2;
            sGrupo        := ExcelWorksheet1.Range['C' + Trim(IntToStr(Fila)), 'C' + Trim(IntToStr(Fila))].Value2;
            sDescripcion  := ExcelWorksheet1.Range['E' + Trim(IntToStr(Fila)), 'E' + Trim(IntToStr(Fila))].Value2;
            sIdPersonal   := ExcelWorksheet1.Range['F' + Trim(IntToStr(Fila)), 'F' + Trim(IntToStr(Fila))].Value2;
            Cantidad      := ExcelWorksheet1.Range['G' + Trim(IntToStr(Fila)), 'G' + Trim(IntToStr(Fila))].Value2;
            sOrden        := ExcelWorksheet1.Range['H' + Trim(IntToStr(Fila)), 'H' + Trim(IntToStr(Fila))].Value2;
            sIdCompania   := ExcelWorksheet1.Range['J' + Trim(IntToStr(Fila)), 'J' + Trim(IntToStr(Fila))].Value2;
            Solicitado    := ExcelWorksheet1.Range['K' + Trim(IntToStr(Fila)), 'K' + Trim(IntToStr(Fila))].Value2;
            sIdCuenta     := ExcelWorksheet1.Range['M' + Trim(IntToStr(Fila)), 'M' + Trim(IntToStr(Fila))].Value2;
            sPernocta     := ExcelWorksheet1.Range['N' + Trim(IntToStr(Fila)), 'N' + Trim(IntToStr(Fila))].Value2;

            if sCadena <> sGrupo then
            begin
                procCategoria(sGrupo);
                sCadena := sGrupo;
            end;

            if not sobrescribir then
            begin
                connection.zCommand.Active := False;
                connection.zCommand.SQL.Clear;
                connection.zCommand.SQL.Add('delete from tripulaciondiaria_listado where sContrato =:Contrato And dIdFecha =:Fecha '+
                                            'and sIdTurno =:Turno and sIdTripulacion=:IdTripulacion');
                Connection.zCommand.Params.ParamByName('Contrato').DataType      := ftString;
                Connection.zCommand.Params.ParamByName('Contrato').Value         := global_contrato_Barco;
                Connection.zCommand.Params.ParamByName('Fecha').DataType         := ftDate;
                Connection.zCommand.Params.ParamByName('Fecha').Value            := AvDtpFecha.Date;
                Connection.zCommand.Params.ParamByName('Turno').DataType         := ftString;
                Connection.zCommand.Params.ParamByName('Turno').Value            := jDblCmbTurnos.KeyValue;
                Connection.zCommand.Params.ParamByName('IdTripulacion').DataType := ftString;
                Connection.zCommand.Params.ParamByName('IdTripulacion').Value    := sValue;
                connection.zCommand.ExecSQL();
            end;

            Writeln(archivo, '("'+ global_contrato_barco +'", "'+ QrTurnos.FieldByName('sIdTurno').AsString+'", "'+ formatoFecha(AvDtpFecha.Date) +
                             '", "'+ global_sIdCategoria +'", "'+ sValue +'", "' + sOrden + '", "' + sNombre + '", "'+ sIdPersonal + '", "' + sDescripcion +
                             '", '+ Solicitado +', "", ' + IntToStr(1) + ', ' + IntToStr(0) + ', "'+ sIdCompania + '", ' + Cantidad +', "' + sIdCuenta +'", "'+ sPernocta +'") ');
            Fila := Fila + 1;
            sValue := ExcelWorksheet1.Range['A' + Trim(IntToStr(Fila)), 'A' + Trim(IntToStr(Fila))].Value2;
        end;

        Writeln(archivo, ''+ Char(59));

        CloseFile(archivo);
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        ruta := ExtractFilePath(Application.ExeName) + 'tripulacion.sql';
        connection.zCommand.SQL.LoadFromFile(ruta);
        connection.zCommand.ExecSQL;

        //Pernoctas de PEP o Adicionales
        ProcPernoctas;

        //Se actualizan las nuevas categorias agregadas en la tripulacion..
        ProcCategoriasTripulacion;

        QrCategorias.Refresh;

        TmpFecha:=IncDay(TmpFecha,-1);
        AvDtpFechaExit(Sender);

        MessageDlg('Proceso Terminado con exito...', mtInformation, [mbOk], 0);
        
        eliminar := ExtractFilePath(Application.ExeName) + 'distribucion.sql';
        if FileExists(eliminar) then
           DeleteFile(eliminar);
      finally
        QryTriList.Destroy;
      end;
    except
      on E: Exception do
      begin
          MessageDlg(e.Message, mtInformation, [mbOk], 0)
      end;
    end;

    //Desconectamos la aplicacion..
    ExcelApplication1.Disconnect;
  end;
end;

procedure TfrmListaPersonalV2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:=caFree;
end;

procedure TfrmListaPersonalV2.FormShow(Sender: TObject);
var
  QrFecha:TzReadOnlyQuery;
begin
  QrTurnos.Active:=false;
  if connection.contrato.FieldByName('sTipoObra').AsString='BARCO' then
      QrTurnos.ParamByName('contrato').AsString:= global_contrato
  else
    QrTurnos.ParamByName('contrato').AsString:= global_contrato_Barco;

  QrTurnos.Open;
  if QrTurnos.RecordCount>0 then
    jDblCmbTurnos.KeyValue:=QrTurnos.FieldByName('sIdTurno').AsString;

  QrCategorias.Active:=false;
  QrCategorias.Open;
  jDblCmbCategoria.KeyValue:=null;
  AvDtpFecha.Date:=now;

  connection.zCommand.Active := False;
  connection.zCommand.SQL.Clear;
  connection.zCommand.SQL.Add('select dFechaVigencia from categorias where dFechaVigencia <=:fecha ');
  connection.zCommand.ParamByName('fecha').AsDate   := Date;
  connection.zCommand.Open;

  if connection.zCommand.RecordCount > 0 then
     dFechaVigencia := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;

  QrCompanias.Active:=false;
  QrCompanias.Open;

  QPernoctan.Active:=false;
  QPernoctan.Open;

  QrPernoctas.Active:=false;
  QrPernoctas.Open;
  QrFecha:=TzREadOnlyQuery.Create(nil);
  try
    QrFecha.Connection:=connection.zConnection;
    QrFecha.SQL.Text:='select max(didfecha) as didfecha from reportediario';
    QrFecha.Open;
    if QrFecha.RecordCount>0 then
      TmpFecha:=QrFecha.FieldByName('dIdFecha').AsDateTime
    else
      TmpFecha:=now;
    if AvDtpFecha.Date=tmpFecha then
      AvDtpFecha.Date:=incday(tmpFecha,-1);

  finally
    QrFecha.Destroy;
    AvDtpFechaExit(Sender);
  end;

end;

procedure TfrmListaPersonalV2.jDblCmbCategoriaExit(Sender: TObject);
begin
  TmpFecha:=IncDay(TmpFecha,-1);
  AvDtpFechaExit(Sender);
end;

procedure TfrmListaPersonalV2.jDblCmbTurnosExit(Sender: TObject);
begin
  TmpFecha:=IncDay(TmpFecha,-1);
  AvDtpFechaExit(Sender);
end;

procedure TfrmListaPersonalV2.QTripulacionAfterPost(DataSet: TDataSet);
begin
  //ActualizarResumen(ParamContrato: string; ParamIdPersonal: string;ParamCuenta:Integer; ParamFecha: TDate);
  global_sIdCategoria := QrCategorias.FieldByName('sIdCategoria').AsString;
  if AvOChkBActualiza.Checked then
    ActualizarResumen(QTripulacion.FieldByName('sOrden').AsString,QTripulacion.FieldByName('sIdPersonal').AsString,
                  QTripulacion.FieldByName('sIdCuenta').AsString,QTripulacion.FieldByName('dIdFecha').AsDateTime,
                  QTripulacion.FieldByName('lImprimeListado').AsString);

end;

procedure TfrmListaPersonalV2.QTripulacionBeforeDelete(DataSet: TDataSet);
begin
  LParamContrato:=QTripulacion.FieldByName('sOrden').AsString;
  LParamIdPersonal:=QTripulacion.FieldByName('sIdPersonal').AsString;
  LParamCuenta:=QTripulacion.FieldByName('sIdCuenta').AsString;
  lParamFecha:=QTripulacion.FieldByName('dIdFecha').AsDateTime;
end;

procedure TfrmListaPersonalV2.QTripulacionBeforePost(DataSet: TDataSet);
begin
  QTripulacion.fieldByname('Compania').asstring:='';
  QTripulacion.fieldByname('sidPernocta').asstring:='';
  QTripulacion.fieldByname('pernocta').asstring:='';
end;

procedure TfrmListaPersonalV2.raerPersonaldelDiaAnterios1Click(Sender: TObject);
begin
  TraerPersonal(IncDay(AvDtpFecha.Date,-1));
end;

procedure TfrmListaPersonalV2.raerPersonalxFecha1Click(Sender: TObject);

begin
  Application.CreateForm(TFrmEligeFecha,FrmEligeFecha);
  try
    FrmEligeFecha.FechaDef:=incDay(AvDtpFecha.Date,-1);
    if FrmEligeFecha.ShowModal=mrOk then
       TraerPersonal(FrmEligeFecha.DNFecha.Date);
  finally
    FrmEligeFecha.Destroy;
  
  end;
end;

procedure TfrmListaPersonalV2.ProcCategoria(sParamGrupo: string);
var
   dFechaVigencia : Tdate;
   MaxCategoria   : integer;
   sId            : string;
begin
    //Consultamos en la tabla de categorias si existe dependiendo de la fecha de la categoria..
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select sIdCategoria from categorias where sDescripcion = :grupo and dFechaVigencia <=:fecha ');
    connection.zCommand.ParamByName('grupo').AsString := sParamGrupo;
    connection.zCommand.ParamByName('fecha').AsDate   := AvDtpFecha.Date;
    connection.zCommand.Open;

    if connection.zCommand.RecordCount = 0 then
    begin
        //Buscamos el máximo Id de categoria..
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select max(cast(sIdCategoria as unsigned)) as maxima, dFechaVigencia from categorias where dFechaVigencia <=:fecha group by dFechaVigencia DESC ');
        connection.zCommand.ParamByName('fecha').AsDate   := AvDtpFecha.Date;
        connection.zCommand.Open;

        if connection.zCommand.RecordCount > 0 then
        begin
            sId := connection.zCommand.FieldByName('maxima').AsString;
            dFechaVigencia := connection.zCommand.FieldByName('dFechaVigencia').AsDateTime;
            try
               MaxCategoria := StrToInt(sId);
               inc(MaxCategoria);
               sId := IntToStr(MaxCategoria);
            Except
                 sId := '10';
            end;
        end
        else
        begin
            dFechaVigencia := date;
            sId := '10';
        end;
        //Sino se encuentra dentro de la vigencia insertamos la categoria.
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('insert into categorias (sIdCategoria, dFechaVigencia, sDescripcion, sMiGrupoResumen, lPersonalAnexo) '+
                                    'values (:categoria, :fecha, :descripcion, "PERSONAL ADMINISTRATIVO", "No")');
        connection.zCommand.ParamByName('categoria').AsString   := sId;
        connection.zCommand.ParamByName('fecha').AsDate         := dFechaVigencia;
        connection.zCommand.ParamByName('descripcion').AsString := sParamGrupo;
        connection.zCommand.ExecSQL;

        global_sIdCategoria := sId;

        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('select sIdTipoPersonal from tiposdepersonal where sIdTipoPersonal =:tipo ');
        connection.zCommand.ParamByName('tipo').AsString := sId;
        connection.zCommand.Open;

        if connection.zCommand.RecordCount = 0 then
        begin
            //Sino se encuentra dentro de la vigencia insertamos la categoria.
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            connection.zCommand.SQL.Add('insert into tiposdepersonal (sIdTipoPersonal, sDescripcion, sMascara, lPersonalEQ, lImprimeConcentrado, lPernocta) '+
                                        'values (:tipo, :descripcion, :mascara, "No", "No", "No")');
            connection.zCommand.ParamByName('tipo').AsString        := sId;
            connection.zCommand.ParamByName('descripcion').AsString := sParamGrupo;
            connection.zCommand.ParamByName('mascara').AsString     := copy(sParamGrupo, 0, 4) + '-';
            connection.zCommand.ExecSQL;
        end;    
    end
    else
       global_sIdCategoria := connection.zCommand.FieldByName('sIdCategoria').AsString;
end;

function TfrmListaPersonalV2.formatoFecha(fecha: TDate) : string;
var
  anio, mes, dia : Word;
  resultado : string;
begin
  DecodeDate(fecha, anio, mes, dia);
  resultado := IntToStr(anio) + '/' + IntToStr(mes) + '/' + IntToStr(dia);
  Result := resultado;
end;

procedure TfrmListaPersonalV2.ProcCategoriasTripulacion;
var
   dFechaVigencia : tDate;
   MaxCategoria   : integer;
   sId            : string;

   //Variables para el archivo .sql que genera el proceso.
   archivo, archivo2  : TextFile;
   eliminar : string;
   ruta     : string;
   sCadena  : string;
   sAnexoAgrupa, sPersonal : String;

   Datos: array[1..2, 1..2500] of string;
   numero, posicion : integer;
   lEncontrado : boolean;
begin
    //Consultamos en la tabla de Anexos cual es el anexo de Personal de abordo..
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('Select sAnexo from anexos where sTipo = "PERSONAL" and sTierra = "No"');
    connection.zCommand.Open;

    if connection.zCommand.RecordCount > 0 then
       sAnexoAgrupa := connection.zCommand.FieldByName('sAnexo').AsString
    else
       sAnexoAgrupa := '0';

    //Consultamos en la tabla de categorias si existe dependiendo de la fecha de la categoria..
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select sIdCategoria, dFechaVigencia from categorias where dFechaVigencia <=:fecha ');
    connection.zCommand.ParamByName('fecha').AsDate   := AvDtpFecha.Date;
    connection.zCommand.Open;

    sCadena := '';
    //Creamos el archivos para importar los datos..
    eliminar := ExtractFilePath(Application.ExeName) + 'tripulacion.sql';
    if FileExists(eliminar) then
    DeleteFile(eliminar);

    AssignFile(archivo, ExtractFilePath(Application.ExeName) + 'tripulacion.sql');
    Rewrite(archivo);
    Writeln(archivo, 'insert into tripulacion (sContrato, dFechaVigencia, sIdCategoria, sIdTripulacion, sDescripcion, iNacionales, iExtranjeros, iOrden, sIdTripulacionGrupo, sDescripcionGrupo) VALUES ');

    //Aqui igual se inserta el personal...
    eliminar := ExtractFilePath(Application.ExeName) + 'personal.sql';
    if FileExists(eliminar) then
    DeleteFile(eliminar);

    AssignFile(archivo2, ExtractFilePath(Application.ExeName) + 'personal.sql');
    Rewrite(archivo2);
    Writeln(archivo2, 'insert into personal (sContrato, sIdPersonal, iItemOrden, sDescripcion, sIdTipoPersonal, sMedida, dCantidad, lProrrateo, lCobro, lPernocta, sAgrupaPersonal, sAnexo) VALUES ');

    posicion := 0;
    While not connection.zCommand.Eof do
    begin
        //Consultamos desde categriri
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('select sIdCategoria, sIdPersonal, sDescripcion, sOrden, sIdCuenta, sIdPernocta from tripulaciondiaria_listado where sIdCategoria =:categoria and dIdFecha =:fecha group by sIdPernocta, sIdPersonal ');
        connection.QryBusca.ParamByName('categoria').AsString   := connection.zCommand.FieldByName('sIdCategoria').AsString;
        connection.QryBusca.ParamByName('fecha').AsDate         := AvDtpFecha.Date;
        connection.QryBusca.Open;

        while not connection.QryBusca.Eof do
        begin
            connection.QryBusca2.Active := False;
            connection.QryBusca2.SQL.Clear;
            connection.QryBusca2.SQL.Add('select sIdTripulacion from tripulacion where sIdCategoria =:categoria and sIdTripulacion =:tripulacion and dFechaVigencia =:fecha ');
            connection.QryBusca2.ParamByName('categoria').AsString   := connection.QryBusca.FieldByName('sIdCategoria').AsString;
            connection.QryBusca2.ParamByName('tripulacion').AsString := connection.QryBusca.FieldByName('sIdPersonal').AsString;
            connection.QryBusca2.ParamByName('fecha').AsDate         := connection.zcommand.FieldByName('dFechaVigencia').AsDateTime;
            connection.QryBusca2.Open;

            if connection.QryBusca2.RecordCount = 0 then
            begin
              connection.qryROProrrateos.Active := False;
              connection.qryROProrrateos.SQL.Clear;
              connection.qryROProrrateos.SQL.Add('select sIdTipoPersonal from personal where sContrato =:Contrato and sIdTipoPersonal =:Tipo and sIdPersonal =:Id ');
              connection.qryROProrrateos.ParamByName('contrato').AsString   := global_contrato_barco;
              connection.qryROProrrateos.ParamByName('Tipo').AsString       := connection.QryBusca.FieldByName('sIdCategoria').AsString;
              connection.qryROProrrateos.ParamByName('Id').AsString         := connection.QryBusca.FieldByName('sIdPersonal').AsString;
              connection.qryROProrrateos.Open;

               //Buscamos maximo ordenamiento...
               connection.qryBuscaTrx.Active := False;
               connection.qryBuscaTrx.SQL.Clear;
               connection.qryBuscaTrx.SQL.Add('select max(iOrden) as iOrden from tripulacion where sIdCategoria =:categoria and dFechaVigencia =:fecha group by sIdCategoria');
               connection.qryBuscaTrx.ParamByName('categoria').AsString   := connection.QryBusca.FieldByName('sIdCategoria').AsString;
               connection.qryBuscaTrx.ParamByName('fecha').AsDate         := connection.zcommand.FieldByName('dFechaVigencia').AsDateTime;
               connection.qryBuscaTrx.Open;

               if connection.qryBuscaTrx.RecordCount = 0 then
                  MaxCategoria := 1
               else
                  MaxCategoria := connection.qryBuscaTrx.FieldByName('iOrden').AsInteger + 1;

               if sCadena <> connection.QryBusca.FieldByName('sIdCategoria').AsString then
                  sCadena := connection.QryBusca.FieldByName('sIdCategoria').AsString;

                numero := 1;
                lEncontrado := False;
                while numero <= posicion do
                begin
                    if (Datos[1,numero] = connection.QryBusca.FieldByName('sIdCategoria').AsString) and
                       (Datos[2,numero] = connection.QryBusca.FieldByName('sIdPersonal').AsString)  then
                       lEncontrado := True;
                    inc(numero);
                end;

                if lEncontrado = False then
                begin
                    if posicion <> 0 then
                     begin
                        Writeln(archivo, ''+ Char(44));
                        if  (connection.qryROProrrateos.RecordCount = 0) and (sPersonal = 'Si') then
                            Writeln(archivo2, ''+ Char(44));
                     end;

                    inc(posicion);
                    Datos[1,posicion] := connection.QryBusca.FieldByName('sIdCategoria').AsString;
                    Datos[2,posicion] := connection.QryBusca.FieldByName('sIdPersonal').AsString;

                    //Tripulacion..
                    Writeln(archivo, '("'+ global_contrato_barco +'", "'+ formatoFecha(connection.zcommand.FieldByName('dFechaVigencia').AsDateTime) +
                                 '", "'+ connection.QryBusca.FieldByName('sIdCategoria').AsString +'", "'+ connection.QryBusca.FieldByName('sIdPersonal').AsString +
                                 '", "' + connection.QryBusca.FieldByName('sDescripcion').AsString + '", ' + IntToStr(1) + ', '+ IntToStr(0) + ', ' + IntToStr(MaxCategoria) +
                                 ', "'+ connection.QryBusca.FieldByName('sIdPersonal').AsString + '", "'+ connection.QryBusca.FieldByName('sDescripcion').AsString + '") ');

                  if  connection.qryROProrrateos.RecordCount = 0 then
                  begin
                    sPersonal := 'Si';
                    //Personal..
                    Writeln(archivo2, '("'+ global_contrato_barco +'", "'+ connection.QryBusca.FieldByName('sIdPersonal').AsString + '", '+IntToStr(MaxCategoria) +
                                 ', "'+ connection.QryBusca.FieldByName('sDescripcion').AsString +'", "'+ connection.QryBusca.FieldByName('sIdCategoria').AsString +
                                 '", "JOR", '+ IntToStr(0) + ', "No", "No", "No", "S/C", "'+sAnexoAgrupa+'") ');
                  end;
                end;
            end;

            global_sIdCategoria := connection.QryBusca.FieldByName('sIdCategoria').AsString;
            global_sIdPernocta  := connection.QryBusca.FieldByName('sIdPernocta').AsString;

            if AvOChkBActualiza.Checked then
               ActualizarResumen(connection.QryBusca.FieldByName('sOrden').AsString,connection.QryBusca.FieldByName('sIdPersonal').AsString,connection.QryBusca.FieldByName('sIdCuenta').AsString,AvDtpFecha.Date);

            connection.QryBusca.Next
        end;
        connection.zCommand.Next;
    end;

     Writeln(archivo, ''+ Char(59));
     CloseFile(archivo);

     Writeln(archivo2, ''+ Char(59));
     CloseFile(archivo2);

     if sCadena <> ''then
     begin
        //Tripulacion
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        ruta := ExtractFilePath(Application.ExeName) + 'tripulacion.sql';
        connection.zCommand.SQL.LoadFromFile(ruta);
        connection.zCommand.ExecSQL;

        if sPersonal = 'Si' then
        begin
            //Personal
            connection.zCommand.Active := False;
            connection.zCommand.SQL.Clear;
            ruta := ExtractFilePath(Application.ExeName) + 'personal.sql';
            connection.zCommand.SQL.LoadFromFile(ruta);
            connection.zCommand.ExecSQL;
        end;
     end;

     eliminar := ExtractFilePath(Application.ExeName) + 'tripulacion.sql';
     if FileExists(eliminar) then
        DeleteFile(eliminar);

      eliminar := ExtractFilePath(Application.ExeName) + 'personal.sql';
     if FileExists(eliminar) then
        DeleteFile(eliminar);
end;

procedure TfrmListaPersonalV2.ProcPernoctas;
var
  sFolio : string;
begin
    //Buscamos en la lista de personal, las categorias que adicionales para perncota..
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Clear;
    connection.zCommand.SQL.Add('select td.sIdCuenta,sum(td.iNacionales) as TotalPernocta from tripulaciondiaria_listado td '+
                                'inner join categorias c on (td.sIdCategoria = c.sIdCategoria and lPersonalAnexo = "Si") '+
                                'where td.dIdFecha = :fecha and sOrden =:Orden and td.sIdCuenta is not Null group by td.sIdcuenta ');
    connection.zCommand.ParamByName('Fecha').AsDate   := AvDtpFecha.Date;
    connection.zCommand.ParamByName('Orden').AsString := global_contrato;
    connection.zCommand.Open;

    //Eliminamos los movimientos de la bitacoradepernocta..
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('delete from bitacoradepernocta where sContrato =:Orden and dIdFecha =:fecha');
    connection.QryBusca.ParamByName('Orden').AsString := global_contrato;
    connection.QryBusca.ParamByName('Fecha').AsDate   := AvDtpFecha.Date;
    connection.QryBusca.ExecSQL;

    //Consultamos el folio donde se insertará la pernocta..
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('select sNumeroOrden from bitacoradepersonal where sContrato =:Orden and dIdFecha =:fecha');
    connection.QryBusca.ParamByName('Orden').AsString := global_contrato;
    connection.QryBusca.ParamByName('Fecha').AsDate   := AvDtpFecha.Date;
    connection.QryBusca.Open;

    if connection.QryBusca.RecordCount > 0 then
       sFolio := connection.QryBusca.FieldByName('sNumeroOrden').AsString;

    while not connection.zCommand.Eof do
    begin
        connection.QryBusca.Active := False;
        connection.QryBusca.SQL.Clear;
        connection.QryBusca.SQL.Add('insert into bitacoradepernocta (sContrato, dIdFecha, sNumeroOrden, sIdCuenta, dCantidad) '+
                                   'values (:contrato, :fecha, :orden, :cuenta, :cantidad)');
        connection.QryBusca.ParamByName('contrato').AsString := global_contrato;
        connection.QryBusca.ParamByName('fecha').AsDate      := Avdtpfecha.Date;
        connection.QryBusca.ParamByName('orden').AsString    := sFolio;
        connection.QryBusca.ParamByName('cuenta').AsString   := connection.zCommand.FieldByName('sIdCuenta').AsString;
        connection.QryBusca.ParamByName('cantidad').AsFloat  := connection.zCommand.FieldByName('TotalPernocta').AsFloat;
        connection.QryBusca.ExecSQL;

       connection.zCommand.Next;
    end;

end;

end.
