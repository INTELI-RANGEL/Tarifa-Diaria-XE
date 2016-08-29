unit frm_compara;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, DBCtrls, StdCtrls, Grids, DBGrids, DB, global,
  Buttons, Mask, ExtCtrls, frxClass, frxDBSet, RXCtrls, frxDMPExport,
  frxCross, ComCtrls, TeEngine, Series, TeeProcs, Chart, DbChart, Newpanel,
  RxMemDS, ZAbstractRODataset, ZDataset, Menus, DateUtils, masUtilerias,
  udbgrid, unitexcepciones, UFunctionsGHH,UnitTBotonesPermisos;

type
  TfrmComparativo = class(TForm)
    ds_avancesglobales: TDataSource;
    GroupBox1: TGroupBox;
    grid_avances: TDBGrid;
    btnExit: TBitBtn;
    btnPrinter: TBitBtn;
    Bevel2: TBevel;
    Label6: TLabel;
    btnRefresh: TBitBtn;
    Avances: TfrxDBDataset;
    chImpresion: TCheckBox;
    frxAvances: TfrxReport;
    frxAvancesTotales: TfrxReport;
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
    chkAdicionales: TCheckBox;
    rxAvancesContrato: TRxMemoryData;
    rxAvancesContratodIdFecha: TDateField;
    rxAvancesContratodProgramadoDia: TFloatField;
    rxAvancesContratodProgramadoAcum: TFloatField;
    rxAvancesContratodFisicoDia: TFloatField;
    rxAvancesContratodFisicoAcumulado: TFloatField;
    SaveSql: TSaveDialog;
    popGraphics: TPopupMenu;
    Exportar1: TMenuItem;
    grGeneral: tNewGroupBox;
    dbGraphics: TDBChart;
    Series1: TFastLineSeries;
    Series2: TFastLineSeries;
    Series3: TFastLineSeries;
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
    procedure chkAdicionalesClick(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
    procedure grid_avancesMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_avancesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
   procedure grid_avancesTitleClick(Column: TColumn);
  private
  sMenuP: String;
    { Private declarations }
  public
    { Public declarations }
    procedure Refrescar;
  end;

var
  frmComparativo: TfrmComparativo;
  utgrid:ticdbgrid;
  BotonPermiso: TBotonesPermisos;
  BotonPermiso2: TBotonesPermisos;
implementation

{$R *.dfm}

procedure TfrmComparativo.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cAvContrato');
  BotonPermiso2 := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cAvContrato',popGraphics);
  BotonPermiso.permisosBotones2(nil, nil, nil, btnPrinter);
  BotonPermiso2.permisosBotones(nil);

  try
     sMenuP:=stMenu;
     UtGrid:=TicdbGrid.create(grid_Avances);
     Refrescar;
     dbGraphics.Title.Text.Clear ;
     dbGraphics.Title.Text.Add ('Avances Programado/Fisico/Financiero') ;
     dbGraphics.Title.Text.Add (Caption) ;
     dbGraphics.Title.Text.Add (connection.contrato.FieldValues['mDescripcion']) ;
     // Termino de Generar la Grafica ..
     // Termino de Actualizar los Datos principales del Contrato ...
     rxAvancesContrato.Locate('dIdFecha', Date() , [loPartialKey]) ;
     Grid_Avances.SetFocus
  except
   on e : exception do begin
   UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_compara', 'Al iniciar el funcionario', 0);
   end;
  end;
end;
//*******************************************************************************
procedure TfrmComparativo.grid_avancesMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid.dbGridMouseMoveCoord(x,y);
end;
procedure TfrmComparativo.grid_avancesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;
procedure TfrmComparativo.grid_avancesTitleClick(Column: TColumn);
begin
UtGrid.DbGridTitleClick(Column);
end;
//********************************************************************************
procedure TfrmComparativo.Refrescar;
var
    sFecha     : String ;
    iMiMes     : Byte ;
    dAcumulado : Currency ;
    dAcumuladoFisico : Currency ;
    dAvanceFisico : Currency ;
begin
  try
  // Primero Genera la Grafica ....

    Caption := connection.contrato.FieldValues['sContrato']  + '-' + connection.configuracion.FieldValues['sNombre'] + '. Monto del Contrato [' + connection.contrato.fieldByName('dMontoMN').AsString + ']' ;
    dMontoContrato := connection.contrato.FieldValues['dMontoMN'] ;

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

    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select a.dIdFecha, a.dAvancePonderadoDia, a.dAvancePonderadoGlobal, af.dAvance ' +
                                'From avancesglobales a left join avancesglobalesxorden af on (a.sContrato = af.sContrato and a.sIdConvenio = af.sIdConvenio and a.sNumeroOrden = af.sNumeroOrden and a.dIdFecha = af.dIdFecha) ' +
                                'Where a.sContrato = :Contrato And a.sIdConvenio = :Convenio And a.sNumeroOrden = :Orden order by a.dIdFecha' ) ;
    Connection.qryBusca.params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.params.ParamByName('Contrato').Value := global_contrato ;
    Connection.qryBusca.params.ParamByName('Convenio').DataType := ftString ;
    Connection.qryBusca.params.ParamByName('Convenio').Value := global_convenio ;
    Connection.qryBusca.params.ParamByName('Orden').DataType := ftString ;
    Connection.qryBusca.params.ParamByName('Orden').Value := '' ;
    Connection.qryBusca.Open ;
    If connection.QryBusca.RecordCount > 0 Then
        iMiMes := MonthOf(Connection.QryBusca.FieldValues['dIdFecha'])
    Else
        iMiMes := 0 ;
    dAcumuladoFisico := 0 ;
    While NOT Connection.qryBusca.Eof Do
    Begin
        If iMiMes <> MonthOf(Connection.QryBusca.FieldValues['dIdFecha']) Then
        Begin
            iMiMes := MonthOf(Connection.QryBusca.FieldValues['dIdFecha']) ;
            rxGraficaProgramado.Append ;
            rxGraficaProgramado.FieldValues['sDescripcion'] := global_contrato ;
            rxGraficaProgramado.FieldValues['dFecha'] := Connection.QryBusca.FieldValues['dIdFecha'] - 1 ;
            rxGraficaProgramado.FieldValues['dProgramado'] := dAcumulado ;
            rxGraficaProgramado.Post ;
            dAcumulado := 0 ;
        End ;
        dAvanceFisico := 0 ;
        If Connection.QryBusca.FieldValues['dAvance'] <> Null Then
        begin
            dAcumuladoFisico := dAcumuladoFisico + Connection.QryBusca.FieldValues['dAvance'] ;
            dAvanceFisico := Connection.QryBusca.FieldValues['dAvance']
        End ;
        rxAvancesContrato.Append ;
        rxAvancesContrato.FieldValues['dIdFecha'] := Connection.QryBusca.FieldValues['dIdFecha'] ;
        rxAvancesContrato.FieldValues['dProgramadoDia'] := Connection.QryBusca.FieldValues['dAvancePonderadoDia'] ;
        rxAvancesContrato.FieldValues['dProgramadoAcum'] := Connection.QryBusca.FieldValues['dAvancePonderadoGlobal'] ;
        rxAvancesContrato.FieldValues['dFisicoDia'] := dAvanceFisico ;
        rxAvancesContrato.FieldValues['dFisicoAcumulado'] := dAcumuladoFisico ;
        rxAvancesContrato.Post ;

        dAcumulado := Connection.QryBusca.FieldValues['dAvancePonderadoGlobal'] ;
        Connection.qryBusca.Next
    End ;

    If dAcumulado <> 0 Then
        With Connection.qryBusca DO
        begin
            sFecha :=  sfnFechaFinal(MonthOf(FieldValues['dIdFecha']),YearOf(FieldValues['dIdFecha'])) ;
            rxGraficaProgramado.Append ;
            rxGraficaProgramado.FieldValues['sDescripcion'] := global_contrato ;
            rxGraficaProgramado.FieldValues['dFecha'] := sFecha ;
            rxGraficaProgramado.FieldValues['dProgramado'] := dAcumulado ;
            rxGraficaProgramado.Post ;

      End ;
    // El resto del avance Fisico .....

    Connection.qryBusca2.Active := False ;
    Connection.qryBusca2.SQL.Clear ;
    Connection.qryBusca2.SQL.Add('Select dIdFecha, dAvance ' +
                                'From avancesglobalesxorden ' +
                                'Where sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroOrden = :Orden and dIdFecha > :fecha order by dIdFecha' ) ;
    Connection.qryBusca2.params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca2.params.ParamByName('Contrato').Value := global_contrato ;
    Connection.qryBusca2.params.ParamByName('Convenio').DataType := ftString ;
    Connection.qryBusca2.params.ParamByName('Convenio').Value := global_convenio ;
    Connection.qryBusca2.params.ParamByName('Orden').DataType := ftString ;
    Connection.qryBusca2.params.ParamByName('Orden').Value := '' ;
    Connection.qryBusca2.params.ParamByName('fecha').DataType := ftDate ;
    Connection.qryBusca2.params.ParamByName('fecha').Value := connection.QryBusca.FieldValues['dIdFecha'] ;
    Connection.qryBusca2.Open ;
    While NOT Connection.qryBusca2.Eof Do
    Begin
        dAvanceFisico := 0 ;
        If Connection.QryBusca2.FieldValues['dAvance'] <> Null Then
        begin
            dAcumuladoFisico := dAcumuladoFisico + Connection.QryBusca2.FieldValues['dAvance'] ;
            dAvanceFisico := Connection.QryBusca2.FieldValues['dAvance']
        End ;
        rxAvancesContrato.Append ;
        rxAvancesContrato.FieldValues['dIdFecha'] := Connection.QryBusca2.FieldValues['dIdFecha'] ;
        rxAvancesContrato.FieldValues['dProgramadoDia'] := 0 ;
        rxAvancesContrato.FieldValues['dProgramadoAcum'] := Connection.QryBusca.FieldValues['dAvancePonderadoGlobal'] ;
        rxAvancesContrato.FieldValues['dFisicoDia'] := dAvanceFisico ;
        rxAvancesContrato.FieldValues['dFisicoAcumulado'] := dAcumuladoFisico ;
        rxAvancesContrato.Post ;
        Connection.qryBusca2.Next
    End ;


    // Real ...
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select year(dIdFecha) as dAnno , month(dIdFecha) as dMes From avancesglobalesxorden Where ' +
                                'sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroOrden = :Orden Group By Year(dIdFecha), month(dIdFecha)' ) ;
    Connection.qryBusca.params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.params.ParamByName('Contrato').Value := global_contrato ;
    Connection.qryBusca.params.ParamByName('Convenio').DataType := ftString ;
    Connection.qryBusca.params.ParamByName('Convenio').Value := global_convenio ;
    Connection.qryBusca.params.ParamByName('Orden').DataType := ftString ;
    Connection.qryBusca.params.ParamByName('Orden').Value := '' ;
    Connection.qryBusca.Open ;
    While NOT Connection.qryBusca.Eof Do
    Begin
          sFecha :=  sfnFechaFinal (Connection.qryBusca.FieldValues['dMes'],Connection.qryBusca.FieldValues['dAnno']) ;
          Connection.qryBusca2.Active := False ;
          Connection.qryBusca2.SQL.Clear ;
          Connection.qryBusca2.SQL.Add('Select Sum(dAvance) as dMensual From avancesglobalesxorden Where ' +
                                       'sContrato = :Contrato And sIdConvenio = :Convenio And sNumeroOrden = :Orden And dIdFecha <= :Fecha Group By sContrato' ) ;
          Connection.qryBusca2.params.ParamByName('Contrato').DataType := ftString ;
          Connection.qryBusca2.params.ParamByName('Contrato').Value := global_contrato ;
          Connection.qryBusca2.params.ParamByName('Convenio').DataType := ftString ;
          Connection.qryBusca2.params.ParamByName('Convenio').Value := global_convenio ;
          Connection.qryBusca2.params.ParamByName('Orden').DataType := ftString ;
          Connection.qryBusca2.params.ParamByName('Orden').Value := '' ;
          Connection.qryBusca2.params.ParamByName('Fecha').DataType := ftDate ;
          Connection.qryBusca2.params.ParamByName('Fecha').Value := StrToDate(sFecha) ;
          Connection.qryBusca2.Open ;
          If Connection.qryBusca.RecordCount > 0 Then
          Begin
             rxGraficaFisico.Append ;
             rxGraficaFisico.FieldValues['sDescripcion'] := global_contrato ;
             rxGraficaFisico.FieldValues['dFecha'] := sFecha ;
             rxGraficaFisico.FieldValues['dFisico'] := Connection.qryBusca2.FieldValues['dMensual'] ;
             rxGraficaFisico.Post ;
          End ;
          Connection.qryBusca.Next
   End ;

     // Financiero ....
     Connection.qryBusca.Active := False ;
     Connection.qryBusca.SQL.Clear ;
     Connection.qryBusca.SQL.Add('Select year(dFechaFinal) as dAnno , month(dFechaFinal) as dMes From estimaciones Where ' +
                                 'sContrato = :Contrato Group By Year(dFechaFinal), month(dFechaFinal)' ) ;
     Connection.qryBusca.params.ParamByName('Contrato').DataType := ftString ;
     Connection.qryBusca.params.ParamByName('Contrato').Value := global_contrato ;
     Connection.qryBusca.Open ;
     While NOT Connection.qryBusca.Eof Do
     Begin
         sFecha :=  sfnFechaFinal (Connection.qryBusca.FieldValues['dMes'],Connection.qryBusca.FieldValues['dAnno']) ;
         Connection.qryBusca2.Active := False ;
         Connection.qryBusca2.SQL.Clear ;
         If chkAdicionales.Checked Then
              Connection.qryBusca2.SQL.Add('Select Sum(dMontoMN) as dReal From estimaciones ' +
                                           'Where sContrato = :Contrato And dFechaFinal <= :Fecha Group By sContrato' )
         Else
              Connection.qryBusca2.SQL.Add('Select Sum(dMontoMN) as dReal From estimaciones ' +
                                           'Where sContrato = :Contrato And dFechaFinal <= :Fecha And sNumeroGenerador NOT Like "%A%" Group By sContrato' ) ;

         Connection.qryBusca2.params.ParamByName('Contrato').DataType := ftString ;
         Connection.qryBusca2.params.ParamByName('Contrato').Value := global_contrato ;
         Connection.qryBusca2.params.ParamByName('Fecha').DataType := ftDate ;
         Connection.qryBusca2.params.ParamByName('Fecha').Value := strToDate(sFecha) ;
         Connection.qryBusca2.Open ;
         If Connection.qryBusca2.RecordCount > 0 Then
         Begin
            try
              rxGraficaFinanciero.Append ;
              rxGraficaFinanciero.FieldValues['sDescripcion'] := global_contrato ;
              rxGraficaFinanciero.FieldValues['dFecha'] := sFecha ;
              if dMontoContrato = 0 then
                  rxGraficaFinanciero.FieldValues['dFinanciero'] := 0
              else
                  rxGraficaFinanciero.FieldValues['dFinanciero'] := (Connection.qryBusca2.FieldValues['dReal'] / dMontoContrato) * 100 ;
              rxGraficaFinanciero.Post ;
            Except
                 rxGraficaFinanciero.FieldValues['dFinanciero'] := 0;
            end;
         End ;
         Connection.qryBusca.Next
     End ;
  except
   on e : exception do begin
   UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_compara', 'Al generar la gráfica', 0);
   end;
  end;

end;

procedure TfrmComparativo.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  BotonPermiso.Free;
  BotonPermiso2.Free;
  action := cafree ;
  utGrid.Destroy;
end;

procedure TfrmComparativo.btnExitClick(Sender: TObject);
begin
      close
end;

procedure TfrmComparativo.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        grid_avances.SetFocus
end;

procedure TfrmComparativo.btnPrinterClick(Sender: TObject);
begin
if grid_avances.DataSource.DataSet.IsEmpty=false then
 begin
 try
  If chImpresion.Checked = True Then
      If MessageDlg( 'Desea imprimir el comparativo fisico-programado',  mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          frxAvancesTotales.ShowReport    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP))
      Else
          frxAvances.ShowReport    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP))
  Else
      catalogo_001.ShowReport;    //(connection.configuracion.FieldByName('sFormatos').AsString, PermisosExportar(connection.zConnection, global_grupo, sMenuP))
 except
on e : exception do begin
UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_compara', 'Al imprimir', 0);
end;
 end
 end
 else
  showmessage('No existen registros para imprimir');
end;

procedure TfrmComparativo.Exportar1Click(Sender: TObject);
begin
  try
    SaveSql.Title := 'Guardar Grafica';
    If SaveSql.Execute Then
         dbGraphics.SaveToBitmapFile(SaveSql.FileName) ;
  except
   on e : exception do begin
   UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_compara', 'Al exportar gráfica', 0);
   end;
  end;
end;

procedure TfrmComparativo.up3DChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
    dbGraphics.Chart3DPercent := Up3D.Position ;

end;

procedure TfrmComparativo.chk3DClick(Sender: TObject);
begin
    dbGraphics.View3D := chk3d.Checked ;
end;

procedure TfrmComparativo.chkLeyendasClick(Sender: TObject);
begin
    dbGraphics.Series[0].Marks.Visible := chkLeyendas.Checked ;
    dbGraphics.Series[1].Marks.Visible := chkLeyendas.Checked ;
    dbGraphics.Series[2].Marks.Visible := chkLeyendas.Checked ;
end;

procedure TfrmComparativo.chkEjesClick(Sender: TObject);
begin
    dbGraphics.LeftAxis.Visible := chkEjes.Checked
end;

procedure TfrmComparativo.chkProgramadoClick(Sender: TObject);
begin
    dbGraphics.Series[0].Active := chkProgramado.Checked ;
end;

procedure TfrmComparativo.chkFisicoClick(Sender: TObject);
begin
    dbGraphics.Series[1].Active := chkFisico.Checked ;
end;

procedure TfrmComparativo.chkFinancieroClick(Sender: TObject);
begin
    dbGraphics.Series[2].Active := chkFinanciero.Checked ;
end;

procedure TfrmComparativo.chkAdicionalesClick(Sender: TObject);
var
   sFecha : String ;
begin
  try
    If rxGraficaFinanciero.RecordCount > 0 then
        rxGraficaFinanciero.EmptyTable  ;

    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select year(dFechaFinal) as dAnno , month(dFechaFinal) as dMes From estimaciones Where ' +
                                'sContrato = :Contrato Group By Year(dFechaFinal), month(dFechaFinal)' ) ;
    Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.qryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
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
        If chkAdicionales.Checked Then
            Connection.qryBusca2.SQL.Add('Select Sum(dMontoMN) as dReal From estimaciones ' +
                                         'Where sContrato = :Contrato And dFechaFinal <= :Fecha Group By sContrato' )
        Else
            Connection.qryBusca2.SQL.Add('Select Sum(dMontoMN) as dReal From estimaciones ' +
                                         'Where sContrato = :Contrato And dFechaFinal <= :Fecha And sNumeroGenerador NOT Like "%A%" Group By sContrato' ) ;
        Connection.qryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.qryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.qryBusca2.Params.ParamByName('Fecha').DataType := ftDate ;
        Connection.qryBusca2.Params.ParamByName('Fecha').Value := strToDate(sFecha) ;
        Connection.qryBusca2.Open ;
        If Connection.qryBusca2.RecordCount > 0 Then
        Begin
           rxGraficaFinanciero.Append ;
           rxGraficaFinanciero.FieldValues['sDescripcion'] := global_contrato ;
           rxGraficaFinanciero.FieldValues['dFecha'] := sFecha ;
           rxGraficaFinanciero.FieldValues['dFinanciero'] := (Connection.qryBusca2.FieldValues['dReal'] / dMontoContrato) * 100 ;
           rxGraficaFinanciero.Post ;
        End ;
        Connection.qryBusca.Next
    End
  except
   on e : exception do begin
   UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_compara', 'Al consultar la tabla estimaciones ', 0);
   end;
  end;

end;

procedure TfrmComparativo.btnRefreshClick(Sender: TObject);
begin
   try
     Refrescar;
     rxAvancesContrato.Locate('dIdFecha', Date() , [loPartialKey])
   except
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_compara', 'Al actualizar Grid', 0);
    end;
   end;
end;

end.
