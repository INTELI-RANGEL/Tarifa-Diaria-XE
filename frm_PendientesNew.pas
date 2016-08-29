unit frm_PendientesNew;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,  RxGrdCpt, ComCtrls, StdCtrls, Buttons, frm_connection, global, DB,
  ADODB, frxClass, frxDBSet, RxMemDS, Mask,  masUtilerias, Utilerias,
  ZAbstractRODataset, ZDataset, rxToolEdit, rxCurrEdit, UnitExcepciones,
  UFunctionsGHH, cxGraphics, cxLookAndFeels, cxLookAndFeelPainters, Menus,
  dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint, dxSkinCaramel,
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
  dxSkinXmas2008Blue, cxButtons, cxControls, cxContainer, cxEdit, cxGroupBox,
  cxTextEdit, cxMemo, IdBaseComponent, IdComponent, IdCustomTCPServer,
  IdSocksServer, IdIOHandler, IdIOHandlerSocket, IdIOHandlerStack, IdSSL,
  IdSSLOpenSSL, ScktComp ;

type
  TfrmPendientesNew = class(TForm)
    tmDescripcion: TMemo;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    tdLaborado: TCurrencyEdit;
    tdTranscurrido: TCurrencyEdit;
    tdFechaInicio: TMaskEdit;
    tdFechaFinal: TMaskEdit;
    AvProyecto: TCurrencyEdit;
    Label5: TLabel;
    AvPendiente: TCurrencyEdit;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    avProgramado: TCurrencyEdit;
    AvReal: TCurrencyEdit;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    tdIdFecha: TDateTimePicker;
    tsTiempoEfectivo: TMaskEdit;
    tsTiempoInactivo: TMaskEdit;
    sReportes: TLabel;
    rptReportesDiarios: TfrxReport;
    dsReportesDiarios: TfrxDBDataset;
    QryReportesDiarios: TZReadOnlyQuery;
    QryTiempos: TZReadOnlyQuery;
    grpConsulta: TcxGroupBox;
    btnPrinter: TcxButton;
    BitBtn1: TcxButton;
    cxConectar: TcxButton;
    cxEstado: TcxMemo;
    cxTexto: TcxMemo;
    cxEnviar: TcxButton;
    ServerSocket1: TServerSocket;
    ClientSocket1: TClientSocket;
    procedure FormShow(Sender: TObject);
    procedure tdIdFechaChange(Sender: TObject);
    procedure btnPrinterClick(Sender: TObject);
    procedure rptReportesDiariosGetValue(const VarName: String;
      var Value: Variant);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure cmdOcultarClick(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure cxConectarClick(Sender: TObject);
    procedure cxEnviarClick(Sender: TObject);
    procedure ServerSocket1ClientConnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocket1ClientDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocket1ClientRead(Sender: TObject;
      Socket: TCustomWinSocket);
  private
  sMenuP, str: String;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmPendientesNew: TfrmPendientesNew;
implementation

{$R *.dfm}

procedure TfrmPendientesNew.FormShow(Sender: TObject);
begin
    sMenuP:=stMenu;
    global_PendientesOculto:=False;
    If global_contrato <> '' Then
    Begin
        tdIdFecha.Date := Date ;
        tmDescripcion.Text := Connection.contrato.FieldValues['mDescripcion'] ;
        Connection.qryBusca.Active := False ;
        Connection.qryBusca.SQL.Clear ;
        Connection.qryBusca.SQL.Add('Select dFechaInicio, dFechaFinal From convenios Where sContrato = :Contrato And sIdConvenio = :Convenio') ;
        Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString ;
        Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio ;
        Connection.qryBusca.Open ;
        avProyecto.Value := 0 ;
        If Connection.qryBusca.RecordCount > 0 Then
        Begin
            tdFechaInicio.Text := Connection.qryBusca.FieldValues['dFechaInicio'] ;
            If VarIsNull (Connection.qryBusca.FieldValues['dFechaFinal']) then
                MessageDlg('No Hay Fecha Final del Convenio !!', mtError, [mbOk], 0)
            Else
                tdFechaFinal.Text := Connection.qryBusca.FieldValues['dFechaFinal'] ;
            tdLaborado.Value := (Date - Connection.qryBusca.FieldValues['dFechaInicio']) + 1;
            tdTranscurrido.Value := Connection.qryBusca.FieldValues['dFechaFinal']  - Date ;
            If Date <= Connection.qryBusca.FieldValues['dFechaFinal'] Then
            Begin
                avProyecto.Value := (Connection.qryBusca.FieldValues['dFechaFinal'] - Connection.qryBusca.FieldValues['dFechaInicio']) + 1  ;
                avProyecto.Value := (tdLaborado.Value / avProyecto.Value ) * 100 ;
                avPendiente.Value := 100 - avProyecto.Value
            End
            Else
            Begin
                avProyecto.Value := 100 ;
                avPendiente.Value := 0 ;
          End
        End
        Else
        Begin
            tdFechaInicio.Text := DateToStr(Date) ;
            tdFechaFinal.Text := DateToStr(Date) ;
            tdLaborado.Value := 0 ;
            tdTranscurrido.Value := 0 ;
            avProyecto.Value := 0 ;
            avPendiente.Value := 0 ;
        End ;

        If tdTranscurrido.Value <= 10 Then
        Begin
            tdTranscurrido.Font.Style := [fsBold] ;
            tdTranscurrido.Font.Color := clRed;
            tdTranscurrido.Font.Size  := 9 ;
        End
        Else
        Begin
            tdTranscurrido.Font.Style := [] ;
            tdTranscurrido.Font.Color := clWindowText ;
            tdTranscurrido.Font.Size  := 8 ;
        End ;

        Connection.qryBusca.Active := False ;
        Connection.qryBusca.SQL.Clear ;
        Connection.qryBusca.SQL.Add('Select dAvancePonderadoGlobal From avancesglobales Where sContrato = :Contrato And sIdConvenio = :Convenio And dIdFecha = :Fecha And sNumeroOrden = ""') ;
        Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString ;
        Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio ;
        Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.qryBusca.Params.ParamByName('Fecha').Value := Date ;
        Connection.qryBusca.Open ;
        avProgramado.Value := 0 ;
        If Connection.qryBusca.RecordCount > 0 Then
            avProgramado.Value := Connection.qryBusca.FieldValues['dAvancePonderadoGlobal'] ;

        Connection.qryBusca.Active := False ;
        Connection.qryBusca.SQL.Clear ;
        Connection.qryBusca.SQL.Add('Select Sum(dAvance)  as dAvance From avancesglobalesxorden Where sContrato = :Contrato And sIdConvenio = :Convenio And dIdFecha <= :Fecha And sNumeroOrden = "" Group By sContrato') ;
        Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca.Params.ParamByName('Convenio').DataType := ftString ;
        Connection.qryBusca.Params.ParamByName('Convenio').Value := global_convenio ;
        Connection.qryBusca.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.qryBusca.Params.ParamByName('Fecha').Value := Date ;
        Connection.qryBusca.Open ;
        avReal.Value := 0 ;
        If Connection.qryBusca.RecordCount > 0 Then
            avReal.Value := Connection.qryBusca.FieldValues['dAvance'] ;
    End
end;

procedure TfrmPendientesNew.tdIdFechaChange(Sender: TObject);
Var
    iReportes : Byte ;
    iJornada  : Byte ;
begin
  try
      QryTiempos.Active := False ;
      QryTiempos.Params.ParamByName('Contrato').DataType := ftString ;
      QryTiempos.Params.ParamByName('Contrato').Value := global_contrato ;
      QryTiempos.Params.ParamByName('Fecha').DataType := ftDate ;
      QryTiempos.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
      QryTiempos.Open ;
      tsTiempoEfectivo.Text := '00:00' ;
      tsTiempoInactivo.Text := '00:00' ;
      iReportes := 0 ;
      While NOT QryTiempos.Eof Do
      Begin
          tsTiempoInactivo.Text := sfnSumaHoras (tsTiempoInactivo.Text , QryTiempos.FieldValues['sTiempoMuertoReal'] ) ;
          iReportes := iReportes + 1 ;
          QryTiempos.Next ;
      End ;

      iJornada := ifnJornadaDia (global_contrato, tdIdFecha.Date, frmPendientesNew) ;
      If iJornada < 10 Then
           tsTiempoEfectivo.Text := '0' + Trim(IntToStr(iJornada)) + ':00'
      Else
           tsTiempoEfectivo.Text := Trim(IntToStr(iJornada)) + ':00' ;
      tsTiempoEfectivo.Text := sfnRestaHoras (tsTiempoEfectivo.Text , tsTiempoInactivo.Text) ;
      sReportes.Caption := 'Numero de Reportes Autorizados : ' + IntToStr(iReportes) ;
  except
      on e : exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Pendientes', 'Al seleccionar fecha', 0);
      end
  end;
end;

procedure TfrmPendientesNew.BitBtn1Click(Sender: TObject);
begin
global_PendientesOculto:=True;
close
end;

procedure TfrmPendientesNew.btnPrinterClick(Sender: TObject);
begin
    QryReportesDiarios.Active := False ;
    QryReportesDiarios.Params.ParamByName('Contrato').DataType := ftString ;
    QryReportesDiarios.Params.ParamByName('Contrato').Value := global_contrato ;
    QryReportesDiarios.Params.ParamByName('Fecha').DataType := ftDate ;
    QryReportesDiarios.Params.ParamByName('Fecha').Value := tdIdFecha.Date ;
    QryReportesDiarios.Open ;
    //<ROJAS>
    rptReportesDiarios.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP))
    //
end;

procedure TfrmPendientesNew.rptReportesDiariosGetValue(
  const VarName: String; var Value: Variant);
begin
  If CompareText(VarName, 'TIEMPO_MUERTO') = 0 then
      Value := tsTiempoEfectivo.Text ;
  If CompareText(VarName, 'TIEMPO_EFECTIVO') = 0 then
      Value := tsTiempoInactivo.Text ;

end;

procedure TfrmPendientesNew.ServerSocket1ClientConnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
    Socket.SendText('Connected');//Sends a message to the client
    //If at least a client is connected to the server, then the server can communicate
    //Enables the Send button and the edit box
      cxEnviar.Enabled:=true;
      cxTexto.Enabled:=true;
end;

procedure TfrmPendientesNew.ServerSocket1ClientDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
begin
    //The server cannot send messages if there is no client connected to it
    if ServerSocket1.Socket.ActiveConnections-1=0 then
    begin
      cxEnviar.Enabled:=false;
      cxTexto.Enabled:=false;
    end;
end;

procedure TfrmPendientesNew.ServerSocket1ClientRead(Sender: TObject;
  Socket: TCustomWinSocket);
begin
    //Read the message received from the client and add it to the memo text
    // The client identifier appears in front of the message
    cxTexto.Lines.Add(cxTexto.Text+'Client'+IntToStr(Socket.SocketHandle)+' :'+Socket.ReceiveText);
end;

procedure TfrmPendientesNew.cmdOcultarClick(Sender: TObject);
begin
Visible:=False;
end;

procedure TfrmPendientesNew.cxConectarClick(Sender: TObject);
begin
   if(ServerSocket1.Active = False)//The button caption is ‘Start’
   then
   begin
      ServerSocket1.Active := True;//Activates the server socket
      cxEstado.Lines.Add('Server Started');
      cxConectar.Caption:='Stop';//Set the button caption
   end
   else//The button caption is ‘Stop’
   begin
      ServerSocket1.Active := False;//Stops the server socket
      cxEstado.Lines.Add('Server Stopped');
      cxConectar.Caption:='Start';
      //If the server is closed, then it cannot send any messages
      cxEnviar.Enabled := false;//Disables the “Send” cxEnviar
      cxTexto.Enabled  := false;//Disables the cxTexto box
   end;
end;

procedure TfrmPendientesNew.cxEnviarClick(Sender: TObject);
var
  i: integer;
begin
     Str:=cxTexto.Text;//Take the string (message) sent by the server
     cxEstado.Lines.Add('me: '+Str);//Adds the message to the memo box
     cxTexto.Text:='';//Clears the edit box
     //Sends the messages to all clients connected to the server
     for i:=0 to ServerSocket1.Socket.ActiveConnections-1  do
      ServerSocket1.Socket.Connections[i].SendText(str);//Sent
end;

procedure TfrmPendientesNew.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  action := cafree ;
end;

end.
