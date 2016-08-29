unit frm_seleccion2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, DBCtrls, Buttons, frm_connection, global, DB, ADODB,
  ComCtrls, ImgList, Newpanel, Mask, DateUtils, ExtCtrls, UnitExcepciones,
  ZAbstractRODataset, ZDataset, Math, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxContainer, cxEdit, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinBlueprint, dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom,
  dxSkinDarkSide, dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle, dxSkinFoggy,
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
  cxGroupBox, cxLabel, Menus, cxButtons, cxCheckBox, masUtilerias, cxImage,
  dxGDIPlusClasses, JvBackgrounds, jpeg;

type
  TfrmSeleccion2 = class(TForm)
    pbRegenera: TProgressBar;
    TreeObras: TTreeView;
    tmDescripcion: TMemo;
    tdFechaInicio: TMaskEdit;
    tdFechaFinal: TMaskEdit;
    ds_turnos: TDataSource;
    ImageList1: TImageList;
    ds_ordenesdetrabajo: TDataSource;
    QryDiferencia: TZReadOnlyQuery;
    QryDiferencia3: TZReadOnlyQuery;
    QryDiferencia2: TZReadOnlyQuery;
    Turnos: TZReadOnlyQuery;
    ordenesdetrabajo: TZReadOnlyQuery;
    QryResidencias: TZReadOnlyQuery;
    grpContratos: TcxGroupBox;
    grpInfContrato: TcxGroupBox;
    grpInteligent: TcxGroupBox;
    cxLabel1: TcxLabel;
    cxLabel2: TcxLabel;
    cxLabel3: TcxLabel;
    btnOk: TcxButton;
    chkContrato: TcxCheckBox;
    iconos: TcxImageList;
    cxLabel4: TcxLabel;
    lblOrden: TcxLabel;
    grpImagen: TcxGroupBox;
    imgContrato: TImage;
    imgBackground: TJvBackground;
    tsIdTurno: TDBLookupComboBox;
    tsNumeroOrden: TDBLookupComboBox;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnOkClick(Sender: TObject);
    procedure tsContratoKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdTurnoKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdTurnoEnter(Sender: TObject);
    procedure tsIdTurnoExit(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure TreeObrasClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TreeObrasKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure TreeObrasKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure refresh ;
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
     procedure LoadImageConf(contrato : string; var imagen : TImage);
  public
    { Public declarations }
    cerrar :boolean;
  end;

var
  frmSeleccion2: TfrmSeleccion2;
  MyTreeNode2,
  MyTreeNode3: TTreeNode;
implementation

uses frm_inteligent, frm_reportediarioturno;

{$R *.dfm}

procedure TfrmSeleccion2.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
    If global_contrato = '' Then
        treeObras.SetFocus;
    if (cerrar = False) and (not frmInteligent.adentro) then
        frmInteligent.Tiempo.Enabled := True;
end;

procedure TfrmSeleccion2.btnOkClick(Sender: TObject);
var
    MensajeBarco, filtro_check  : Boolean ;
begin
  try
    If TreeObras.Selected.Text <> '' then
        If TreeObras.Selected.getFirstChild = Nil Then
        Begin


      // Para localizar el usuario y determinar si puede Insertar, Eliminar, Editar, Imprimir
            Connection.QryBusca.Active := False ;
            Connection.QryBusca.SQL.Clear ;
            Connection.QryBusca.SQL.Text := 'select eInsertar, eEditar, eEliminar, eImprimir, eGrabar from usuarios where sIdUsuario = :Usuario';
            Connection.QryBusca.ParamByName('usuario').AsString := global_Usuario;
            Connection.QryBusca.Open;
            If Connection.QryBusca.RecordCount > 0 Then
              begin
                global_Insertar  := Connection.qryBusca.fieldValues['eInsertar'] ;
                global_Editar    := Connection.qryBusca.fieldValues['eEditar'] ;
                global_grabar    := Connection.qryBusca.fieldValues['eGrabar'] ;
                global_Eliminar  := Connection.qryBusca.fieldValues['eEliminar'] ;
                global_Imprimir  := Connection.qryBusca.fieldValues['eImprimir'] ;
              end;


            global_contrato := TreeObras.Selected.Text ;
            //inicialización de las variables globales para la etiqueta de personal y equipo
            Connection.QryBusca2.Close;
            Connection.QryBusca2.Active := False ;
            Connection.QryBusca2.SQL.Clear ;
            Connection.QryBusca2.SQL.Add('select * from anexos');
            Connection.QryBusca2.Open;
             While NOT Connection.QryBusca2.Eof Do
              begin
                if Connection.QryBusca2.FieldValues['sTipo']='PERSONAL' then
                  begin
                    global_labelPersonal     := Connection.QryBusca2.FieldValues['sAnexo'];
                    global_labelPersonalDesc := Connection.QryBusca2.FieldValues['sDescripcion'];
                  end;
                if Connection.QryBusca2.FieldValues['sTipo']='EQUIPO' then
                  begin
                  global_labelEquipo     := Connection.QryBusca2.FieldValues['sAnexo'];
                  global_labelEquipoDesc     := Connection.QryBusca2.FieldValues['sDescripcion'];
                  end;
                if (Connection.QryBusca2.FieldValues['sTipo']='MATERIAL') And (Connection.QryBusca2.FieldValues['sTierra']='Si') then
                  begin
                    global_Materialtierra     := Connection.QryBusca2.FieldValues['sAnexo'];
                    global_MaterialtierraDesc := Connection.QryBusca2.FieldValues['sDescripcion'];
                  end;
                if (Connection.QryBusca2.FieldValues['sTipo']='MATERIAL') And (Connection.QryBusca2.FieldValues['sTierra']='No') then
                   begin
                    global_Materialbordo     := Connection.QryBusca2.FieldValues['sAnexo'];
                    global_MaterialbordoDesc := Connection.QryBusca2.FieldValues['sDescripcion'];
                   end;
                if (Connection.QryBusca2.FieldValues['sTipo']='PERNOCTA') And (Connection.QryBusca2.FieldValues['sTierra']='No') Then
                  global_labelPernocta        := Connection.QryBusca2.FieldValues['sAnexo'];
                  global_labelPernoctaDesc        := Connection.QryBusca2.FieldValues['sDescripcion'];
                if Connection.QryBusca2.FieldValues['sTipo']='CONSUMIBLES' then
                  global_MaterialConsumible   := Connection.QryBusca2.FieldValues['sAnexo'];
                 Connection.QryBusca2.Next ;
              end;
            // Verificar si se encuentra dado de alta un contrato correspondiente a conceptos de embarcación
            Connection.QryBusca.Close;
            Connection.QryBusca.Active := False ;
            Connection.QryBusca.SQL.Clear ;
            Connection.QryBusca.SQL.Add('select scontrato from contratos where sTipoObra="BARCO" and sContrato = sCodigo '); //Q p2??
            Connection.QryBusca.Open;

            if Connection.QryBusca.RecordCount > 0 then
              begin
                 global_contrato_barco := Connection.QryBusca.FieldByName('sContrato').AsString ;
                 Connection.QryBusca2.Active := False ;
                 Connection.qryBusca2.SQL.Clear ;
                 Connection.qryBusca2.SQL.Add('Select sIdPernocta from configuracion Where sContrato = :Contrato ') ;
                 Connection.QryBusca2.ParamByName('contrato').Value := global_contrato_barco;
                 Connection.QryBusca2.Open;
                 if Connection.QryBusca2.RecordCount >0 Then
                   global_Pernocta  := Connection.QryBusca2.FieldValues['sIdPernocta'] ;

              end
            Else
            begin
                 global_contrato_barco := Connection.QryBusca.FieldByName('sContrato').AsString ;
                 Connection.QryBusca2.Active := False ;
                 Connection.qryBusca2.SQL.Clear ;
                 Connection.qryBusca2.SQL.Add('Select sIdPernocta from configuracion Where sContrato = :Contrato') ;
                 Connection.QryBusca2.ParamByName('contrato').Value := global_contrato_barco;
                 Connection.QryBusca2.Open;
                 if Connection.QryBusca2.RecordCount >0 Then
                   global_Pernocta  := Connection.QryBusca2.FieldValues['sIdPernocta'] ;
{              if MensajeBarco then
                showmessage('El contrato que ha seleccionado es un contrato que controla la información de una embarcación y debe estar ligado a un contrato principal, ' + 'algunas de las acciones del sistema no estarán disponibles para este contrato, se debe acceder al contrato principal para poder activarlas.' + #13 + #13 + 'El contrato principal de este contrato de embarcación es: ' + global_contrato_barco);
}
            end;

            (* Modificación para CICSA, debido a que existen programas que utilizan en su código y querys en su sentencia select
               el contrato 428238800, para evitar modificaciones
            if Global_Contrato = 'SIEM DORADO' then Global_Contrato := '428238800';*)

            If tsIdTurno.Text <> '' Then
            begin
                global_turno := tsIdTurno.KeyValue ;
                global_sturno := tsIdTurno.Text
            end ;

             // Actualizo Kardex del Sistema ....
            connection.zCommand.Active := False ;
            connection.zCommand.SQL.Clear ;
            connection.zCommand.SQL.Add ( 'Insert Into kardex_sistema (sContrato, sIdUsuario, dIdFecha, sHora, sDescripcion, lOrigen)' +
                                          'Values (:Contrato, :Usuario, :Fecha, :Hora, :Descripcion, :Origen)') ;
            connection.zCommand.Params.ParamByName('Contrato').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Contrato').Value := Global_Contrato ;
            connection.zCommand.Params.ParamByName('Usuario').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Usuario').Value := Global_Usuario ;
            connection.zCommand.Params.ParamByName('Fecha').DataType := ftDate ;
            connection.zCommand.Params.ParamByName('Fecha').Value := Date ;
            connection.zCommand.Params.ParamByName('Hora').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Hora').value := FormatDateTime('hh:mm:ss', Now) ;
            connection.zCommand.Params.ParamByName('Descripcion').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Descripcion').Value := 'Selección del Contrato [' + global_contrato + '] turno [' + global_turno + '] desde la dirección [' + global_ip + ']' ;
            connection.zCommand.Params.ParamByName('Origen').DataType := ftString ;
            connection.zCommand.Params.ParamByName('Origen').Value := 'Otros Movimientos' ;
            connection.zCommand.ExecSQL ;

            Connection.QryBusca2.Active := False ;
            Connection.QryBusca2.SQL.Clear ;
            Connection.QryBusca2.SQL.Add('Select lCobraPersonal, lJorPu, lCobraEquipo, sTipoObra From contratos Where sContrato = :Contrato') ;
            Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
            Connection.QryBusca2.Params.ParamByName('Contrato').Value    := Global_Contrato ;
            Connection.QryBusca2.Open ;
            if Connection.QryBusca2.RecordCount > 0 then
                Begin
                  Global_Optativa  := Connection.QryBusca2.FieldValues['sTipoObra'] ;
                  Global_Personal  := Connection.QryBusca2.FieldValues['lCobraPersonal'] ;
                  Global_Equipo    := Connection.QryBusca2.FieldValues['lCobraEquipo'] ;
                  Global_PuJor     := Connection.QryBusca2.FieldValues['lJorPu'] ;
                End;

            If global_usuario = 'INTEL-CODE' Then
            Begin
                Connection.ContratosxUsuario.Active := False ;
                Connection.ContratosxUsuario.SQL.Clear ;
                Connection.ContratosxUsuario.SQL.Add('Select stipoObra, sContrato, mDescripcion From contratos Order By sContrato') ;
                Connection.ContratosxUsuario.Open ;
            End
            Else
            Begin
                Connection.ContratosxUsuario.Active := False ;
                Connection.ContratosxUsuario.SQL.Clear ;
                Connection.ContratosxUsuario.SQL.Add('Select c.sTipoObra, c.sContrato, c.mDescripcion From contratosxusuario u ' +
                                                     'INNER JOIN contratos c ON (c.sContrato = u.sContrato and c.lStatus = "Activo") ' +
                                                     'Where u.sIdUsuario = :Usuario Order By c.sContrato') ;
                Connection.ContratosxUsuario.Params.ParamByName('Usuario').DataType := ftString ;
                Connection.ContratosxUsuario.Params.ParamByName('Usuario').Value := global_usuario ;
                Connection.ContratosxUsuario.Open ;
            End ;


            connection.contrato.Close;
            Connection.contrato.ParamByName('Contrato').AsString := global_contrato;
            connection.contrato.Open;

            try
                frmDiarioTurno.OrdenesdeTrabajo.Active := False ;
                frmDiarioTurno.OrdenesdeTrabajo.SQL.Clear ;
                if (global_grupo = 'INTEL-CODE') Then
                    frmDiarioTurno.OrdenesdeTrabajo.SQL.Add('Select sNumeroOrden, iJornada from ordenesdetrabajo where sContrato = :Contrato and ' +
                                         'cIdStatus = :status order by sNumeroOrden')
                Else
                frmDiarioTurno.OrdenesdeTrabajo.SQL.Add('Select ot.sNumeroOrden, ot.iJornada from ordenesdetrabajo ot ' +
                                         'INNER JOIN ordenesxusuario ou On (ou.sContrato=ot.sContrato '  +
                                         'And ou.sNumeroOrden=ot.sNumeroOrden) ' +
                                         'where ot.sContrato =:Contrato And ou.sDerechos<>"BLOQUEADO" ' +
                                         'And ou.sIdUsuario =:Usuario And ot.cIdStatus =:Status order by ot.sNumeroOrden') ;
                frmDiarioTurno.OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString ;
                frmDiarioTurno.OrdenesdeTrabajo.Params.ParamByName('Contrato').Value    := Global_Contrato ;
                frmDiarioTurno.OrdenesdeTrabajo.Params.ParamByName('status').DataType   := ftString ;
                frmDiarioTurno.OrdenesdeTrabajo.Params.ParamByName('status').Value      := connection.configuracion.FieldValues [ 'cStatusProceso' ];
                if global_grupo <> 'INTEL-CODE' Then
                Begin
                    frmDiarioTurno.OrdenesdeTrabajo.Params.ParamByName('Usuario').DataType  := ftString ;
                    frmDiarioTurno.OrdenesdeTrabajo.Params.ParamByName('Usuario').Value     := Global_Usuario ;
                end;
                frmDiarioTurno.OrdenesdeTrabajo.Open ;

                If frmDiarioTurno.OrdenesdeTrabajo.RecordCount > 0 Then
                Begin
                    tsNumeroOrden.KeyValue := frmDiarioTurno.OrdenesdeTrabajo.FieldValues['sNumeroOrden'] ;
                    frmDiarioTurno.ReporteDiario.Active := False ;
                    frmDiarioTurno.ReporteDiario.Params.ParamByName('contrato').DataType := ftString ;
                    frmDiarioTurno.ReporteDiario.Params.ParamByName('contrato').Value    := global_contrato ;
                    frmDiarioTurno.ReporteDiario.Params.ParamByName('orden').DataType    := ftString ;
                    frmDiarioTurno.ReporteDiario.Params.ParamByName('orden').Value       := frmDiarioTurno.tsNumeroOrden.Text ;
                    frmDiarioTurno.ReporteDiario.Open ;
                    frmDiarioTurno.Grid_Reportes.SetFocus
                End
                Else
                    frmDiarioTurno.tsNumeroOrden.SetFocus;
            Except

            end;

            Connection.EstimacionPeriodo.Active := False ;
            Connection.EstimacionPeriodo.Params.ParamByName('Contrato').DataType := ftString ;
            Connection.EstimacionPeriodo.Params.ParamByName('Contrato').Value := Global_Contrato ;
            Connection.EstimacionPeriodo.Open ;


            connection.configuracion.Active := False ;
            connection.configuracion.Params.ParamByName('Contrato').DataType:= ftString ;
            connection.configuracion.Params.ParamByName('Contrato').Value := global_contrato ;
            connection.configuracion.Open ;

            global_miReporte := connection.configuracion.FieldValues['sReportesCIA'];

            global_convenio := 'C' ;
            Global_Afectacion := 'Anexo' ;

            If upperCase(global_independiente) = 'SI' then
                global_orden_general := tsNumeroOrden.KeyValue
            else
                global_orden_general := '' ;

            If connection.configuracion.RecordCount = 0 then
                  application.MessageBox('Precaución: No se encontro el archivo principal de configuración, notifique al Administrador del Sistema' , 'Inteligent' ,  0 )
            Else
            Begin
                  If global_usuario = 'INTEL-CODE' then
                        global_depto := connection.configuracion.FieldValues['sIdDepartamento'] ;

                  Global_Convenio := connection.configuracion.FieldValues['sIdConvenio'] ;
                  Global_Afectacion := connection.configuracion.FieldValues['sPartidaEfectiva'] ;

                  filtro_check := False ;

                  If chkContrato.Checked Then
                      filtro_check := True 
                  Else
                      If (DayOf(Date) in [1,10,20,28]) OR (global_usuario = 'INTEL-CODE') Then
                            filtro_check := False ;
                  If filtro_check Then
                      If connection.configuracion.FieldValues['sTipoContrato'] = 'Precio Unitario' Then
                      Begin
                          QryDiferencia3.Active := False ;
                          QryDiferencia3.Params.ParamByName('Contrato').DataType := ftString ;
                          QryDiferencia3.Params.ParamByName('Contrato').Value := global_contrato ;
                          QryDiferencia3.Params.ParamByName('Convenio').DataType := ftString ;
                          QryDiferencia3.Params.ParamByName('Convenio').Value := global_convenio ;
                          QryDiferencia3.Open ;
                          If QryDiferencia3.RecordCount > 0 Then
                          Begin
                              QryDiferencia3.First ;
                              While NOT QryDiferencia3.Eof Do
                              Begin
                                  If (RoundTo(QryDiferencia3.FieldValues['dPonderado'],2) <> 100) Then
                                      MessageDlg('Precaución: La Orden de Trabajo No. ' + QryDiferencia3.FieldValues['sNumeroOrden'] +
                                                 ' en la sumatoria de ponderados, sumatoria resultante = ' + QryDiferencia3.fieldByName('dPonderado').asString +
                                                 ', verifique esta información.', mtWarning, [mbOk], 0);
                                  QryDiferencia3.Next
                              End
                          End ;

                          //*****************************************************************************************************
                          //Actualizacion omitida para funcionamiento de sincronizador..
                          //Diferencia entre lo almacenado en el registor orden.paquete.partida y la bitacora de actividades ...
                          {connection.zCommand.SQL.Clear ;
                          connection.zCommand.SQL.Add ( 'UPDATE actividadesxorden SET dInstalado = 0, dExcedente = 0 ' +
                                                        'where sContrato = :contrato And sIdConvenio = :Convenio') ;
                          connection.zCommand.Params.ParamByName('contrato').DataType := ftString ;
                          connection.zCommand.Params.ParamByName('contrato').value := global_contrato ;
                          connection.zCommand.Params.ParamByName('convenio').DataType := ftString ;
                          connection.zCommand.Params.ParamByName('Convenio').Value := global_convenio ;
                          connection.zCommand.ExecSQL ;   }

                          QryDiferencia.Active := False ;
                          QryDiferencia.Params.ParamByName('Contrato').DataType := ftString ;
                          QryDiferencia.Params.ParamByName('Contrato').Value := global_contrato ;
                          QryDiferencia.Params.ParamByName('Convenio').DataType := ftString ;
                          QryDiferencia.Params.ParamByName('Convenio').Value := global_convenio ;
                          QryDiferencia.Open ;
                          pbRegenera.Visible := False ;
                          If QryDiferencia.RecordCount > 0 Then
                          Begin
                              pbRegenera.Visible := True ;
                              pbRegenera.Min := 0 ;
                              pbRegenera.Max := QryDiferencia.RecordCount ;
                              pbRegenera.Position := 0 ;
                              QryDiferencia.First ;
                              While NOT QryDiferencia.Eof Do
                              Begin
                                  pbRegenera.Position := QryDiferencia.RecNo ;
                                  If (QryDiferencia.FieldValues['dInstalado'] - QryDiferencia.FieldValues['dReportado']) <> 0 Then
                                  Begin
                                      // Actualizamos las partidas ...
                                       connection.zCommand.Active := False ;
                                       connection.zCommand.SQL.Clear ;
                                       connection.zcommand.SQL.Add ( 'UPDATE actividadesxorden SET dInstalado = :Instalado, dExcedente = :Excedente ' +
                                                                     'where sContrato = :contrato And sIdConvenio = :Convenio And sNumeroOrden = :Orden And sWbs = :Wbs And sNumeroActividad = :Actividad') ;
                                       connection.zcommand.Params.ParamByName('contrato').DataType := ftString ;
                                       connection.zcommand.Params.ParamByName('contrato').value := global_contrato ;
                                       connection.zcommand.Params.ParamByName('convenio').DataType := ftString ;
                                       connection.zcommand.Params.ParamByName('convenio').value := global_convenio ;
                                       connection.zcommand.Params.ParamByName('Orden').DataType := ftString ;
                                       connection.zcommand.Params.ParamByName('Orden').value := QryDiferencia.FieldValues ['sNumeroOrden'] ;
                                       connection.zcommand.Params.ParamByName('Wbs').DataType := ftString ;
                                       connection.zcommand.Params.ParamByName('Wbs').value := QryDiferencia.FieldValues ['sWbs'] ;
                                       connection.zcommand.Params.ParamByName('Actividad').DataType := ftString ;
                                       connection.zcommand.Params.ParamByName('Actividad').value := QryDiferencia.FieldValues ['sNumeroActividad'] ;
                                       If ( QryDiferencia.FieldValues ['dReportado'] > QryDiferencia.FieldValues ['dCantidad']) Then
                                       Begin
                                           connection.zcommand.Params.ParamByName('Instalado').DataType := ftFloat ;
                                           connection.zcommand.Params.ParamByName('Instalado').value := QryDiferencia.FieldValues['dCantidad'] ;
                                           connection.zcommand.Params.ParamByName('Excedente').DataType := ftFloat ;
                                           connection.zcommand.Params.ParamByName('Excedente').value := QryDiferencia.FieldValues ['dReportado'] - QryDiferencia.FieldValues['dCantidad'] ;
                                       End
                                       Else
                                       Begin
                                           connection.zcommand.Params.ParamByName('Instalado').DataType := ftFloat ;
                                           connection.zcommand.Params.ParamByName('Instalado').value := QryDiferencia.FieldValues ['dReportado'] ;
                                           connection.zcommand.Params.ParamByName('Excedente').DataType := ftFloat ;
                                           connection.zcommand.Params.ParamByName('Excedente').value := 0 ;
                                       End ;
                                       connection.zCommand.ExecSQL ;
                                  End ;
                                  QryDiferencia.Next
                              End ;
                              pbRegenera.Visible := False ;
                          End ;

                          //*********************************************************************************************
                          //Actualizacion omitida para funcionamiento de sincronizador..
                          //Diferencia entre lo almacenado entre la partida anexo y el registro orden.paquete.partida ....
                          {connection.zCommand.SQL.Clear ;
                          connection.zCommand.SQL.Add ( 'UPDATE actividadesxanexo SET dInstalado = 0, dExcedente = 0 ' +
                                                        'where sContrato = :contrato And sIdConvenio = :Convenio') ;
                          connection.zCommand.Params.ParamByName('contrato').DataType := ftString ;
                          connection.zCommand.Params.ParamByName('contrato').value := global_contrato ;
                          connection.zCommand.Params.ParamByName('convenio').DataType := ftString ;
                          connection.zCommand.Params.ParamByName('Convenio').Value := global_convenio ;
                          connection.zCommand.ExecSQL ;  }

                          QryDiferencia2.Active := False ;
                          QryDiferencia2.Params.ParamByName('Contrato').DataType := ftString ;
                          QryDiferencia2.Params.ParamByName('Contrato').Value := global_contrato ;
                          QryDiferencia2.Params.ParamByName('Convenio').DataType := ftString ;
                          QryDiferencia2.Params.ParamByName('Convenio').Value := global_convenio ;
                          QryDiferencia2.Open ;
                          If QryDiferencia2.RecordCount > 0 Then
                          Begin
                              pbRegenera.Visible := True ;
                              pbRegenera.Min := 0 ;
                              pbRegenera.Max := QryDiferencia2.RecordCount ;
                              pbRegenera.Position := 0 ;
                              QryDiferencia2.First ;
                              While NOT QryDiferencia2.Eof Do
                              Begin
                                 pbRegenera.Position := QryDiferencia2.RecNo ;
                                 If (QryDiferencia2.FieldValues['dInstalado'] - QryDiferencia2.FieldValues['dReportado']) <> 0 Then
                                 Begin
                                       connection.zCommand.Active := False ;
                                       connection.zCommand.SQL.Clear ;
                                       connection.zcommand.SQL.Add ( 'UPDATE actividadesxanexo SET dInstalado = :Instalado, dExcedente = :Excedente ' +
                                                                           'where sContrato = :contrato And sIdConvenio = :Convenio And sNumeroActividad = :Actividad') ;
                                       connection.zcommand.Params.ParamByName('contrato').DataType := ftString ;
                                       connection.zcommand.Params.ParamByName('contrato').value := global_contrato ;
                                       connection.zcommand.Params.ParamByName('convenio').DataType := ftString ;
                                       connection.zcommand.Params.ParamByName('convenio').value := global_convenio ;
                                       connection.zcommand.Params.ParamByName('Actividad').DataType := ftString ;
                                       connection.zcommand.Params.ParamByName('Actividad').value := QryDiferencia2.FieldValues ['sNumeroActividad'] ;
                                       If (QryDiferencia2.FieldValues ['dReportado'] > QryDiferencia2.FieldValues ['dCantidadAnexo'])  Then
                                       Begin
                                           connection.zcommand.Params.ParamByName('Instalado').DataType := ftFloat ;
                                           connection.zcommand.Params.ParamByName('Instalado').value := QryDiferencia2.FieldValues['dCantidadAnexo'] ;
                                           connection.zcommand.Params.ParamByName('Excedente').DataType := ftFloat ;
                                           connection.zcommand.Params.ParamByName('Excedente').value := QryDiferencia2.FieldValues ['dReportado'] - QryDiferencia2.FieldValues['dCantidadAnexo'] ;
                                       End
                                       Else
                                       Begin
                                           connection.zcommand.Params.ParamByName('Instalado').DataType := ftFloat ;
                                           connection.zcommand.Params.ParamByName('Instalado').value := QryDiferencia2.FieldValues ['dReportado'] ;
                                           connection.zcommand.Params.ParamByName('Excedente').DataType := ftFloat ;
                                           connection.zcommand.Params.ParamByName('Excedente').value := 0 ;
                                       End ;
                                       connection.zCommand.ExecSQL ;
                                 End ;
                                 QryDiferencia2.Next
                              End ;
                              pbRegenera.Visible := False ;
                          End
                      End
               End ;
               cerrar := True;
            Close
        End;
  except
      on e : exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Seleccion de Contrato', 'Al seleccionar contrato', 0);
      end;
  end;
end;

procedure TfrmSeleccion2.tsContratoKeyPress(Sender: TObject; var Key: Char);
begin
  If Key = #13 then
     tsIdTurno.SetFocus
end;

procedure TfrmSeleccion2.tsIdTurnoKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
       btnOk.SetFocus;
end;

procedure TfrmSeleccion2.tsNumeroOrdenEnter(Sender: TObject);
begin
  tsnumeroorden.Color:= global_color_entrada;
end;

procedure TfrmSeleccion2.tsNumeroOrdenExit(Sender: TObject);
begin
  tsnumeroorden.Color:= global_color_salida;
end;

procedure TfrmSeleccion2.tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
begin
      If Key = #13 Then
        tsidturno.SetFocus;
end;

procedure TfrmSeleccion2.tsIdTurnoEnter(Sender: TObject);
begin
    tsIdTurno.Color := global_color_entrada
end;

procedure TfrmSeleccion2.tsIdTurnoExit(Sender: TObject);
begin
    tsIdTurno.Color :=global_color_salida
end;

procedure TfrmSeleccion2.FormActivate(Sender: TObject);
begin
    if not frmInteligent.adentro then
    begin
        ordenesdetrabajo.Active := False ;
        tmDescripcion.Text := '' ;
        turnos.Active := False ;
        global_contrato := '' ;
        chkContrato.Checked := False ;
        If global_grupo = 'INTEL-CODE' Then
          chkContrato.Visible := True
        Else
          chkContrato.Visible := False ;
    end
end;

procedure TfrmSeleccion2.refresh ;
begin
    tmDescripcion.Text := '' ;
    tdFechaInicio.Text := '' ;
    tdFechaFinal.Text := '' ;
    Turnos.Active := False ;
    If TreeObras.Selected.Text <> '' then
        If TreeObras.Selected.getFirstChild = Nil Then
        Begin
            If upperCase(global_independiente) = 'SI' then
            begin
                OrdenesdeTrabajo.Active := False ;
                OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString ;
                OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := TreeObras.Selected.Text ;
                OrdenesdeTrabajo.Open ;
                If OrdenesdeTrabajo.RecordCount > 0 Then
                    tsNumeroOrden.KeyValue := OrdenesdeTrabajo.FieldValues['sNumeroOrden'] ;
            End ;
            Connection.QryBusca.Active := False ;
            Connection.QryBusca.SQL.Clear ;
            Connection.QryBusca.SQL.Add('Select c.mDescripcion, 0 as dFechaInicio, 0 as dFechaFinal From contratos c ' +
                                             'INNER JOIN configuracion cnf ON (c.sContrato = cnf.sContrato) '  +
                                             'Where c.sContrato = :Contrato') ;
            Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
            Connection.QryBusca.Params.ParamByName('Contrato').Value := TreeObras.Selected.Text ;
            Connection.QryBusca.Open ;
            If Connection.QryBusca.RecordCount > 0 Then
            Begin

                LoadImageConf(TreeObras.Selected.Text, imgContrato);

                tmDescripcion.Text := Connection.QryBusca.FieldValues['mDescripcion'] ;
                tdFechaInicio.Text := Connection.QryBusca.FieldValues['dFechaInicio'] ;
                If VarIsNull(Connection.QryBusca.FieldValues['dFechaFinal']) Then
                     MessageDlg('No hay Fecha Final del Convenio  !!', mtError, [mbOk], 0)
                Else
                  tdFechaFinal.Text := Connection.QryBusca.FieldValues['dFechaFinal'] ;
                Turnos.Active := False ;
                Turnos.Params.ParamByName('Contrato').DataType := ftString ;
                Turnos.Params.ParamByName('Contrato').Value := TreeObras.Selected.Text ;
                Turnos.Open ;
                If Turnos.RecordCount > 0 Then
                Begin
                    Turnos.First ;
                    tsIdTurno.KeyValue := Turnos.FieldValues['sIdTurno']
                End
             End
        End
        Else
        Begin
             Turnos.Active := False ;
             Turnos.Params.ParamByName('Contrato').DataType := ftString ;
             Turnos.Params.ParamByName('Contrato').Value := '' ;
             Turnos.Open ;
             tmDescripcion.Text := '' ;
             tdFechaInicio.Text := '' ;
             tdFechaFinal.Text := '' ;
             tsIdTurno.KeyValue := '';
             imgContrato.Picture.Bitmap := nil;
        End
end;


procedure TfrmSeleccion2.TreeObrasClick(Sender: TObject);
begin
    refresh
end;

procedure TfrmSeleccion2.FormShow(Sender: TObject);
begin
  cerrar := False;
  If upperCase(global_independiente) = 'SI' then
  begin
      tsNumeroOrden.Visible := True ;
      lblOrden.Visible := True
  End
  Else
  begin
      tsNumeroOrden.Visible := False ;
      lblOrden.Visible := False
  End ;

//  {Verificamos el numeor de accesos Oceanografia...}
//  Connection.QryBusca.Active := False ;
//  Connection.QryBusca.SQL.Clear ;
//  connection.QryBusca.SQL.Add('select * from activos where iAccesos <= 200') ;
//  connection.QryBusca.Open ;
//
//  if connection.QryBusca.RecordCount = 0 then
//  begin
//      messageDLG('La Version de Inteligent 2014.0.2.18 ha Expirado, Actualice su version!', mtInformation, [mbOk], 0);
//      exit;
//  end
//  else
//  begin
//      Connection.QryBusca.Active := False ;
//      Connection.QryBusca.SQL.Clear ;
//      connection.QryBusca.SQL.Add('Update activos set iAccesos = iAccesos + 1 ') ;
//      connection.QryBusca.ExecSQL;
//  end;


  with TreeObras.Items do
  begin
    Clear;
    Connection.QryBusca.Active := False ;
    Connection.QryBusca.SQL.Clear ;
    Connection.QryBusca.SQL.Add('Select * From activos Order By sIdActivo') ;
    Connection.QryBusca.Open ;
    While NOT Connection.QryBusca.Eof Do
    Begin
        MyTreeNode2 := Add(nil,Connection.QryBusca.FieldValues['sDescripcion'] );
        // Selecciono las distintas residencias o municipios
        QryResidencias.Active := False ;
        qryResidencias.Params.ParamByName('activo').DataType := ftString ;
        qryResidencias.Params.ParamByName('activo').Value := Connection.QryBusca.FieldValues['sIdActivo'] ;
        qryResidencias.Open ;
        While NOT qryResidencias.Eof Do
        Begin
            MyTreeNode3 := AddChild(MyTreeNode2,qryResidencias.FieldValues['sDescripcion']);

            // Seleciono los contratos del municipio
            Connection.QryBusca2.Active := False ;
            Connection.QryBusca2.SQL.Clear ;
            If (global_usuario = 'INTEL-CODE') or (global_usuario = 'ADMIN') Then
            Begin
                Connection.QryBusca2.SQL.Add('select c.sContrato From contratos c Where c.sIdResidencia = :Residencia And c.lStatus = "Activo" Order By c.sContrato') ;
                Connection.QryBusca2.Params.ParamByName('Residencia').DataType := ftString ;
                Connection.QryBusca2.Params.ParamByName('Residencia').Value := qryResidencias.FieldValues['sIdResidencia'] ;
            End
            Else
            Begin
                Connection.QryBusca2.SQL.Add('select c.sContrato From contratos c INNER JOIN contratosxusuario u ON ' +
                                             '(c.sContrato = u.sContrato And u.sIdUsuario = :Usuario) ' +
                                             'Where c.sIdResidencia = :Residencia And c.sIdActivo = :Activo And c.lStatus = "Activo" Order By c.sContrato') ;
                Connection.QryBusca2.Params.ParamByName('Activo').DataType := ftString ;
                Connection.QryBusca2.Params.ParamByName('Activo').Value := Connection.QryBusca.FieldValues['sIdActivo'] ;
                Connection.QryBusca2.Params.ParamByName('Residencia').DataType := ftString ;
                Connection.QryBusca2.Params.ParamByName('Residencia').Value := qryResidencias.FieldValues['sIdResidencia'] ;
                Connection.QryBusca2.Params.ParamByName('Usuario').DataType := ftString ;
                Connection.QryBusca2.Params.ParamByName('Usuario').Value := global_usuario ;
            End ;
            Connection.QryBusca2.Open ;
            if connection.QryBusca2.RecordCount = 0 then
               TreeObras.Items.Delete(MyTreeNode3);
            While NOT Connection.QryBusca2.Eof Do
            Begin
                AddChild(MyTreeNode3,Connection.QryBusca2.FieldValues['sContrato']);
                Connection.QryBusca2.Next
            End ;
            QryResidencias.Next
        End ;
        Connection.QryBusca.Next
    End
  End ;
  treeObras.SetFocus
end;

procedure TfrmSeleccion2.TreeObrasKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    refresh
end;

procedure TfrmSeleccion2.TreeObrasKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    refresh
end;

procedure TfrmSeleccion2.LoadImageConf(contrato: string; var imagen: TImage);
var
  bS: TStream;
  Pic: TJpegImage;
  BlobField: tField;

  qrImg : TZReadOnlyQuery;

begin
  qrImg := TZReadOnlyQuery.Create(nil);
  qrImg.Connection := connection.zConnection;
  with qrImg do
  begin
    Active := False;
    SQL.Clear;
    SQL.Add('select conf.bImagen from contratos con '+
            'inner join configuracion conf '+
            'on (conf.sContrato = con.sContrato) '+
            'inner join convenios conv '+
            'on (conv.sContrato = con.sContrato and conv.sIdConvenio = conf.sIdConvenio) '+
            'where con.sContrato = :contrato');
    ParamByName('contrato').AsString := contrato;
    Open;
  end;

  if qrImg.RecordCount > 0 then
  begin
    BlobField := qrImg.FieldByName('bImagen');
    bS := qrImg.CreateBlobStream(BlobField, bmRead);

    if bS.Size > 1 then
    begin
      try
        Pic := TJPEGImage.Create;
        try
          Pic.LoadFromStream(bS);
          imagen.Picture.Graphic := Pic;
        finally
          Pic.Free;
        end;
      finally
        bS.Free;
      end;
    end
    else
      imagen.Picture.LoadFromFile('');
  end;
end;

end.
