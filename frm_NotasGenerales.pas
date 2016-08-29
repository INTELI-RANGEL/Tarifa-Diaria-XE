unit frm_NotasGenerales;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, frm_barra, StdCtrls, DBCtrls,
  Mask, ExtCtrls, DB, Global, Menus, frxClass, frxDBSet, UnitTarifa,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UdbGrid, utilerias,
  UnitExcepciones, unittbotonespermisos,
  ComCtrls, JvExComCtrls, JvDateTimePicker, StrUtils, DateUtils, ZSqlProcessor,
  AdvSmoothPanel;

type
  TfrmNotasGenerales = class(TForm)
    frmBarra1: TfrmBarra;
    DBPlataformas: TfrxDBDataset;
    rDiario: TfrxReport;
    ds_NotasGenerales: TDataSource;
    zq_NotasGenerales: TZQuery;
    tmNotas: TDBMemo;
    Label1: TLabel;
    DBEdit1: TDBEdit;
    PopupMenu1: TPopupMenu;
    ImportarNotasdeldadeAyer1: TMenuItem;
    PnlSuperior: TPanel;
    PnlInf: TPanel;
    PnlFecha: TPanel;
    tSeleccionarFecha: TJvDateTimePicker;
    lbl2: TLabel;
    Splitter1: TSplitter;
    grid_plataformas: TDBGrid;
    zqActualiza: TZQuery;
    Enumerar: TMenuItem;
    ScriplAplicaLibro: TZSQLProcessor;
    zq_NotasGeneralessNotaGeneral: TMemoField;
    zq_NotasGeneralesiId: TIntegerField;
    zq_NotasGeneralesiOrden: TIntegerField;
    zq_NotasGeneralessNotaCorta: TStringField;
    zq_NotasGeneralesdIdFecha: TDateField;
    zq_NotasGeneralessContrato: TStringField;
    ReporteDiario: TZReadOnlyQuery;
    pnlDiaAnterior: TAdvSmoothPanel;
    tdFecha: TDateTimePicker;
    cmdAceptar: TButton;
    cmdCancelar: TButton;
    Label3: TLabel;
    Label2: TLabel;
    dbAplicaResumenPersonal: TDBCheckBox;
    zq_NotasGeneraleseAplicaResumenPersonal: TStringField;
    DBCheckBox1: TDBCheckBox;
    zq_NotasGeneraleslAplicaConsumos: TStringField;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure grid_plataformasCellClick(Column: TColumn);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure grid_plataformasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_plataformasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_plataformasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure tSeleccionarFechaExit(Sender: TObject);
    procedure tsIdEquipoEnter(Sender: TObject);
    procedure tsIdEquipoExit(Sender: TObject);
    procedure tsIdEquipoKeyPress(Sender: TObject; var Key: Char);
    procedure zq_NotasGeneralesCalcFields(DataSet: TDataSet);
    procedure ImportarNotasdeldadeAyer1Click(Sender: TObject);
    procedure zq_NotasGeneralesAfterInsert(DataSet: TDataSet);
    procedure EnumerarClick(Sender: TObject);
    procedure tmNotasExit(Sender: TObject);
    procedure tmNotasEnter(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure tmNotasDblClick(Sender: TObject);
    procedure rDiarioGetValue(const VarName: string; var Value: Variant);
    procedure frmBarra1btnPrinterClick(Sender: TObject);
    Procedure GeneraReporteDiario_PDF(RTipo:FtTipo;RImpresion:FtSeccion);
    procedure cmdCancelarClick(Sender: TObject);
    procedure cmdAceptarClick(Sender: TObject);
    procedure tdFechaChange(Sender: TObject);
  private
  sMenuP: String;
  zq_NotasGeneraleslAplicaLibro: TStringField;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmNotasGenerales: TfrmNotasGenerales;
   UtGrid:TicDbGrid;
   botonpermiso: tbotonespermisos;
   sOpcion, lStatusOrig: string;
   FechaReporte: TDateTime;
   iOrdenamiento : integer;
implementation

uses {frm_EditorBitacoraDepartamental,} frm_ReporteDiarioTurno;

{$R *.dfm}

procedure TfrmNotasGenerales.FormShow(Sender: TObject);
begin

  sMenuP:=stMenu;
//  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cPlataformas', PopupPrincipal);
  OpcButton := '' ;
  frmbarra1.btnCancel.Click ;
  tSeleccionarFecha.DateTime := FechaReporte;
  zq_NotasGenerales.Active := False;
  zq_NotasGenerales.ParamByName('Fecha').AsDateTime := FechaReporte;
  zq_NotasGenerales.ParamByName('Contrato').AsString := Global_Contrato;

  zq_NotasGenerales.Open;
  if zq_NotasGenerales.FieldList.IndexOf('laplicalibro') < 0 then
  begin
    zq_NotasGenerales.Close;
    zq_NotasGeneraleslAplicaLibro.DataSet := zq_NotasGenerales;
  end;
  if zq_NotasGenerales.FieldDefs.IndexOf('lAplicaLibro') < 0 then
  begin
    if MessageDlg('El campo para el manejo de comentarios en el libro de proyecto "lAplicaLibro" no existe y es necesario.'+#10+'¿Quiere que el sistema lo autogenere con default "No" a los registros ya existentes?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      ScriplAplicaLibro.Execute;
      zq_NotasGenerales.Active := False;
      zq_NotasGenerales.ParamByName('Fecha').AsDateTime := FechaReporte;
      zq_NotasGenerales.ParamByName('Contrato').AsString := Global_Contrato;
      zq_NotasGenerales.FieldDefs.Add('lAplicaLibro',ftString,3,True);
      zq_NotasGenerales.Open;
    end
    else
    begin
      zq_NotasGenerales.Fields.Remove(zq_NotasGeneraleslAplicaLibro);
      zq_NotasGenerales.open;
    end;
  end
  else
    zq_NotasGenerales.open;

   ReporteDiario.Active := False ;
   ReporteDiario.Params.ParamByName('contrato').DataType := ftString ;
   ReporteDiario.Params.ParamByName('contrato').Value    := global_Contrato_Barco;
   ReporteDiario.Params.ParamByName('orden').DataType    := ftString ;
   ReporteDiario.Params.ParamByName('orden').Value       := param_global_contrato;
   ReporteDiario.Params.ParamByName('Fecha').DataType    := ftDate ;
   ReporteDiario.Params.ParamByName('Fecha').Value       := tSeleccionarFecha.Date;
   ReporteDiario.Open ;
  
end;



procedure TfrmNotasGenerales.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  zq_NotasGenerales.Cancel ;
  action := cafree ;
//  utgrid.Destroy;
//  botonpermiso.Free;
end;

procedure TfrmNotasGenerales.FormCreate(Sender: TObject);
begin
  if not Assigned(zq_NotasGeneraleslAplicaLibro) then
  begin
    zq_NotasGeneraleslAplicaLibro:= TStringField.Create(nil);
    zq_NotasGeneraleslAplicaLibro.Name := 'zq_NotasGeneraleslAplicaLibro';
    zq_NotasGeneraleslAplicaLibro.FieldName := 'lAplicaLibro';
    zq_NotasGeneraleslAplicaLibro.FieldKind := fkData;
  end;
end;

procedure TfrmNotasGenerales.tSeleccionarFechaExit(Sender: TObject);
begin
   zq_NotasGenerales.Active := False;
   zq_NotasGenerales.ParamByName('Fecha').AsDateTime  := tSeleccionarFecha.DateTime;
   zq_NotasGenerales.ParamByName('Contrato').AsString := param_global_contrato;
   zq_NotasGenerales.Open;

   ReporteDiario.Active := False ;
   ReporteDiario.Params.ParamByName('contrato').DataType := ftString ;
   ReporteDiario.Params.ParamByName('contrato').Value    := global_Contrato_Barco;
   ReporteDiario.Params.ParamByName('orden').DataType    := ftString ;
   ReporteDiario.Params.ParamByName('orden').Value       := param_global_contrato;
   ReporteDiario.Params.ParamByName('Fecha').DataType    := ftDate ;
   ReporteDiario.Params.ParamByName('Fecha').Value       := tSeleccionarFecha.Date;
   ReporteDiario.Open ;
end;

procedure TfrmNotasGenerales.tsIdEquipoEnter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmNotasGenerales.tsIdEquipoExit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmNotasGenerales.tsIdEquipoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tmNotas.SetFocus;
end;

procedure TfrmNotasGenerales.zq_NotasGeneralesAfterInsert(DataSet: TDataSet);
var
  serie : integer;
begin
  Connection.QryBusca.Active := False;
  Connection.QryBusca.SQL.Clear;
  Connection.QryBusca.SQL.Text := 'SELECT (MAX(iOrden) + 1) AS iNextOrden FROM notas_generales WHERE sContrato = :Contrato AND dIdFecha = :Fecha group by sContrato ';
  Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
  Connection.QryBusca.Params.ParamByName('Fecha').AsDateTime  := tSeleccionarFecha.DateTime;
  Connection.QryBusca.Open;

  if Connection.QryBusca.RecordCount = 0 then
     serie := 1
  else
      serie := Connection.QryBusca.FieldByName('iNextOrden').AsInteger;

  zq_NotasGenerales.FieldByName('sContrato').AsString  := Global_Contrato;
  zq_NotasGenerales.FieldByName('dIdFecha').AsDateTime := tSeleccionarFecha.DateTime;
  zq_NotasGenerales.FieldByName('iOrden').AsInteger    := serie;
 end;

procedure TfrmNotasGenerales.zq_NotasGeneralesCalcFields(DataSet: TDataSet);
begin
  zq_NotasGenerales.FieldByName('sNotaCorta').AsString := MidStr(zq_NotasGenerales.FieldByName('sNotaGeneral').AsString, 1, 250);
end;

procedure TfrmNotasGenerales.frmBarra1btnAddClick(Sender: TObject);
begin
  if ObtenerEstatusReporte(global_contrato, tSeleccionarFecha.Date) <> 'Pendiente' then
  begin
        MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
        exit;
  end;

  sOpcion := 'Insert';
  zq_NotasGenerales.Append;
  zq_NotasGenerales.FieldByName('lAplicaConsumos').AsString := 'No';
  zq_NotasGenerales.FieldByName('eAplicaResumenPersonal').AsString := 'No';
  frmBarra1.btnAddClick(Sender);
  grid_plataformas.Enabled := False;

  frmBarra1.btnAddClick(Sender);
  tmNotas.SetFocus;
  tmNotas.SelStart := Length(tmNotas.Text);
end;

procedure TfrmNotasGenerales.frmBarra1btnEditClick(Sender: TObject);
begin
   if ObtenerEstatusReporte(global_contrato, tSeleccionarFecha.Date) <> 'Pendiente' then
   begin
        MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
        exit;
   end;

   frmBarra1.btnEditClick(Sender);
   sOpcion := 'Edit';
   try
     zq_NotasGenerales.Edit;
     iOrdenamiento := zq_NotasGenerales.FieldByName('iOrden').AsInteger;

   except
     on e : exception do
     begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Notas_Generales', 'Al editar registro', 0);
        frmbarra1.btnCancel.Click ;
     end;
   end;
   tmNotas.SetFocus;
   grid_plataformas.Enabled := False;

end;

procedure TfrmNotasGenerales.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
  try
    if sOpcion = 'Insert' then
       tmNotas.Lines.Add('');

    zq_NotasGenerales.Post ;

    if  sOpcion = 'Edit' then
    begin
        if iOrdenamiento <> zq_NotasGenerales.FieldByName('iOrden').AsInteger then
         zq_NotasGenerales.refresh ;
    end;

  except
      on e : exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Notas Generales ..', 'Al salvar registro', 0);
          frmbarra1.btnCancel.Click ;
      end;
  end;
  if sOpcion = 'Edit' then
  begin
      grid_plataformas.Enabled := True;
      sOpcion := '';
  end
  else
  begin
     frmbarra1.btnCancel.Click ;
     frmBarra1.btnAdd.Click;
  end;
end;

procedure TfrmNotasGenerales.frmBarra1btnPrinterClick(Sender: TObject);
begin
    GeneraReporteDiario_PDF(FtAbordo,FtsAll)
end;

procedure TfrmNotasGenerales.frmBarra1btnCancelClick(Sender: TObject);
begin
  zq_NotasGenerales.Cancel ;
  frmBarra1.btnCancelClick(Sender);
  grid_plataformas.Enabled := True;
  grid_plataformas.SetFocus;
  sOpcion := '';
end;

procedure TfrmNotasGenerales.frmBarra1btnDeleteClick(Sender: TObject);
begin
  if ObtenerEstatusReporte(global_contrato, tSeleccionarFecha.Date) <> 'Pendiente' then
  begin
      MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
      exit;
  end;

  If zq_NotasGenerales.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        zq_NotasGenerales.Delete ;
      except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al eliminar registro', 0);
       end;
      end
    end
end;

procedure TfrmNotasGenerales.frmBarra1btnRefreshClick(Sender: TObject);
begin
  zq_NotasGenerales.refresh ;
end;

procedure TfrmNotasGenerales.frmBarra1btnExitClick(Sender: TObject);
begin
   frmBarra1.btnExitClick(Sender);
   close;
end;


procedure TfrmNotasGenerales.grid_plataformasCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
      frmbarra1.btnCancel.Click ;
end;

procedure TfrmNotasGenerales.grid_plataformasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
//    UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmNotasGenerales.grid_plataformasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
//  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmNotasGenerales.grid_plataformasTitleClick(Column: TColumn);
begin
//  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmNotasGenerales.ImportarNotasdeldadeAyer1Click(Sender: TObject);
begin
    tdFecha.Date :=  tSeleccionarFecha.DateTime - 1;
    pnlDiaAnterior.Visible := True;
    tdFecha.OnChange(sender);
end;

procedure TfrmNotasGenerales.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click;
end;

procedure TfrmNotasGenerales.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click 
end;

procedure TfrmNotasGenerales.Paste1Click(Sender: TObject);
begin
//   UtGrid.AddRowsFromClip;
end;

procedure TfrmNotasGenerales.rDiarioGetValue(const VarName: string;
  var Value: Variant);
begin


  If CompareText(VarName, 'SUPERINTENDENTE') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
          Value := sSuperIntendente
      Else
          Value := sSuperIntendentePatio ;

  If CompareText(VarName, 'SUPERVISOR') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
          Value := sSupervisor
      Else
          Value := sSupervisorPatio ;

  If CompareText(VarName, 'SUPERVISOR_TIERRA') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
          Value := sSupervisorTierra
      Else
          Value := sResidente ;

   If CompareText(VarName, 'PUESTO_SUPERVISOR_TIERRA') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
      begin
          if pos('#', sPuestoSuperIntendente) > 0 then
             Value := copy(sPuestoSuperIntendente,0, pos('#', sPuestoSuperIntendente)-1) +#13+ copy(sPuestoSuperIntendente,pos('#', sPuestoSuperIntendente)+1, length(sPuestoSuperIntendente))
          else
             Value := sPuestoSuperIntendente
      end
      Else
          Value := sPuestoSuperIntendentePatio ;

  If CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
      begin
          if pos('#', sPuestoSupervisor) > 0 then
             Value := copy(sPuestoSupervisor,0, pos('#', sPuestoSupervisor)-1) +#13+ copy(sPuestoSupervisor,pos('#', sPuestoSupervisor)+1, length(sPuestoSupervisor))
          else
             Value := sPuestoSupervisor
      end
      Else
          Value := sPuestoSupervisorPatio ;

  If CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
      If ReporteDiario.FieldValues['sOrigenTierra'] = 'No' Then
      begin
          if pos('#', sPuestoSupervisorTierra) > 0 then
             Value := copy(sPuestoSupervisorTierra,0, pos('#', sPuestoSupervisorTierra)-1) +#13+ copy(sPuestoSupervisorTierra,pos('#', sPuestoSupervisorTierra)+1, length(sPuestoSupervisorTierra))
          else
             Value := sPuestoSupervisorTierra
      end
      Else
          Value := sPuestoResidente ;


  If CompareText(VarName, 'DESCRIPCION_ORDEN') = 0 then
      Value := mDescripcionOrden  ;

  If CompareText(VarName, 'Oficio_Orden') = 0 then
      Value := sFolio  ;

  If CompareText(VarName, 'PLATAFORMA') = 0 then
      Value := sPlataformaOrden  ;

  If CompareText(VarName, 'JORNADAS_SUSPENDIDAS') = 0 then
      Value := sJornadasSuspendidas  ;

  If CompareText(VarName, 'TURNO') = 0 then
      Value := sDescripcionTurno ;

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
    If CompareText(VarName, 'SUMPERSONAL') = 0 then
      Value := SumaPersonal ;
  If CompareText(VarName, 'SUMEQUIPOS') = 0 then
      Value := SumaEquipos ;
end;

procedure TfrmNotasGenerales.cmdAceptarClick(Sender: TObject);
begin
    if Connection.QryBusca.RecordCount > 0 then
    begin
       while Not Connection.QryBusca.Eof do
       begin
          connection.zCommand.Active := False;
          Connection.zCommand.SQL.Text := 'INSERT INTO notas_generales (iOrden, sNotaGeneral, sContrato, dIdFecha) VALUES ' +
                                          '(:Orden, :NotaGeneral, :Contrato, :Fecha) ';
          Connection.zCommand.Params.ParamByName('Orden').AsInteger      := Connection.QryBusca.FieldByName('iOrden').AsInteger;
          Connection.zCommand.Params.ParamByName('NotaGeneral').AsString := Connection.QryBusca.FieldByName('sNotaGeneral').AsString;
          Connection.zCommand.Params.ParamByName('Contrato').AsString    := Connection.QryBusca.FieldByName('sContrato').AsString;
          Connection.zCommand.Params.ParamByName('Fecha').AsDateTime     := tSeleccionarFecha.DateTime;
          Connection.zCommand.ExecSQL;
          Connection.QryBusca.Next;
       end;
       zq_NotasGenerales.Refresh;
    end
    else
    begin
      ShowMessage('No se encontraron notas en la Fecha Seleccionada.');
    end;
end;

procedure TfrmNotasGenerales.cmdCancelarClick(Sender: TObject);
begin
    pnlDiaanterior.Visible := False;
end;

procedure TfrmNotasGenerales.Copy1Click(Sender: TObject);
begin
//    UtGrid.CopyRowsToClip;
end;

procedure TfrmNotasGenerales.tdFechaChange(Sender: TObject);
begin
    label3.Caption := 'Total 0 Notas';
    Connection.QryBusca.Active := False;
    Connection.QryBusca.SQL.Text := 'SELECT * FROM notas_generales WHERE sContrato = :Contrato AND dIdFecha = :Fecha ';
    Connection.QryBusca.Params.ParamByName('Contrato').AsString := Global_Contrato;
    connection.QryBusca.Params.ParamByName('Fecha').DataType    := ftDate;
    connection.QryBusca.Params.ParamByName('Fecha').Value       := tdFecha.Date;
    Connection.QryBusca.Open;

    if Connection.QryBusca.RecordCount > 0 then
       label3.Caption := 'Total '+ IntToStr(Connection.QryBusca.RecordCount)+ ' Notas';
end;

procedure TfrmNotasGenerales.tmNotasDblClick(Sender: TObject);
begin
   if ObtenerEstatusReporte(global_contrato, tSeleccionarFecha.Date) <> 'Pendiente' then
    begin
          MessageDlg('El Reporte ha sido Validado/Autorizado por lo tanto no puede modificarse.', mtInformation, [mbOk], 0);
          exit;
    end;

    if global_Editor <> 'Nuevo' then
    begin
        sTituloVentana := ' NOTAS GENERALES';
        {Application.CreateForm(TfrmEditorBitacoraDepartamental, frmEditorBitacoraDepartamental);
        frmEditorBitacoraDepartamental.ShowModal;}
    end;
end;

procedure TfrmNotasGenerales.tmNotasEnter(Sender: TObject);
begin
  //DBMemo1.Color := global_color_entrada;
end;

procedure TfrmNotasGenerales.tmNotasExit(Sender: TObject);
begin
  //DBMemo1.Color := global_color_salida;
end;

procedure TfrmNotasGenerales.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click
end;

procedure TfrmNotasGenerales.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click 
end;

procedure TfrmNotasGenerales.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmNotasGenerales.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click
end;

procedure TfrmNotasGenerales.EnumerarClick(Sender: TObject);
var
  id  : Integer;
begin

  if zq_NotasGenerales.RecordCount > 0 then
  begin
    zq_NotasGenerales.First;
    Id := 1;
    while not zq_NotasGenerales.Eof do
    begin
        zq_NotasGenerales.Edit;
        zq_NotasGenerales.FieldValues['iOrden'] := Id;
        zq_NotasGenerales.Post;
        inc(Id);
        zq_NotasGenerales.Next;
    end;
    zq_NotasGenerales.Refresh;
  end;
  frmBarra1.btnRefresh.Click;

end;

procedure TfrmNotasGenerales.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmNotasGenerales.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

Procedure TfrmNotasGenerales.GeneraReporteDiario_PDF(RTipo:FtTipo;RImpresion:FtSeccion);
var
   sSeccion: string;
begin
    Reportediario.Active := False;
    ReporteDiario.ParamByName('Contrato').AsString := global_contrato_barco;
    ReporteDiario.ParamByName('Orden').AsString    := global_contrato;
    ReporteDiario.ParamByName('Fecha').AsDate      := tSeleccionarFecha.Date;
    ReporteDiario.Open;

    EncabezadoPDF_Horizontal(ReporteDiario,rDiario,FtAbordo);
    FirmasPDF_Generales(ReporteDiario,     rDiario,FtAbordo);
    sSeccion := connection.configuracion.FieldByName('sSeccionImprime').AsString;
    {Clasificacion de secciones a Imprimir..}

   if pos('Notas Generales', sSeccion) > 0 then
       ReportePDF_NotasGenerales(ReporteDiario,        rDiario,RTipo,RImpresion)
   else
      ReportePDF_NotasGenerales(ReporteDiario,        rDiario,RTipo,ftsNone);

    rDiario.LoadFromFile(global_files + global_Mireporte + '_TDReporteDiarioNotas.fr3') ;
    rDiario.ShowReport();
    ReportePDF_ClearDataset(rDiario);
end;

end.

