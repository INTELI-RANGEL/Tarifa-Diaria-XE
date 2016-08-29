unit frm_ReprogramacionFolio;

interface

uses



  frm_connection, UnitMetodos, global, DateUtils, Utilerias,

  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_barra, cxGraphics, cxControls, cxLookAndFeels,
  cxLookAndFeelPainters, cxStyles, dxSkinsCore, dxSkinDevExpressStyle,
  dxSkinFoggy, dxSkinscxPCPainter, cxCustomData, cxFilter, cxData,
  cxDataStorage, cxEdit, cxNavigator, DB, cxDBData, cxGridLevel, cxClasses,
  cxGridCustomView, cxGridCustomTableView, cxGridTableView, cxGridDBTableView,
  cxGrid, ZAbstractRODataset, ZAbstractDataset, ZDataset, cxContainer, cxLabel,
  cxTextEdit, cxDBEdit, cxMaskEdit, cxDropDownEdit, cxCalc, cxCalendar,
  cxLookupEdit, cxDBLookupEdit, cxDBLookupComboBox, cxGroupBox, cxProgressBar,
  Menus, StdCtrls, cxButtons, ImgList, DBClient, cxGridChartView,
  cxGridDBChartView, cxMemo;

type
  TfrmReprogramacionFolio = class(TForm)
    TfrmBarra1: TfrmBarra;
    gridDbReprogramaciones: TcxGridDBTableView;
    cxFoliosLvl: TcxGridLevel;
    gridReprogramaciones: TcxGrid;
    ZReprogramaciones: TZQuery;
    dsReprogramaciones: TDataSource;
    ZPartidas: TZReadOnlyQuery;
    dsPartidas: TDataSource;
    gridDbReprogramacionesColumn1: TcxGridDBColumn;
    gridDbReprogramacionesColumn3: TcxGridDBColumn;
    ZProgramas: TZReadOnlyQuery;
    dsProgramas: TDataSource;
    grpCaptura: TcxGroupBox;
    dbIdReprogramacion: TcxDBTextEdit;
    dbFolio: TcxDBLookupComboBox;
    dbFechaInicio: TcxDBDateEdit;
    dbDias: TcxDBCalcEdit;
    cxLabel4: TcxLabel;
    cxLabel3: TcxLabel;
    cxLabel2: TcxLabel;
    cxLabel1: TcxLabel;
    prgProgreso: TcxProgressBar;
    cximgReprogramaciones: TcxImageList;
    popPrincipal: TPopupMenu;
    RecalcularReprogramaciones1: TMenuItem;
    dbGerencialInicio: TcxDBLookupComboBox;
    cxLabel5: TcxLabel;
    dsGerenciales: TDataSource;
    prgProgresoPartidas: TcxProgressBar;
    ZGerenciales: TZReadOnlyQuery;
    gridDbReprogramacionesColumn2: TcxGridDBColumn;
    cxLabel6: TcxLabel;
    grpGrafica: TcxGroupBox;
    ZGrafica: TZReadOnlyQuery;
    dsGrafica: TDataSource;
    gridGrafica: TcxGrid;
    gridDbGrafica: TcxGridDBTableView;
    cxGraficaLvl: TcxGridLevel;
    GridChartGrafica: TcxGridDBChartView;
    GridChartGraficaSeries1: TcxGridDBChartSeries;
    Representacindeavances1: TMenuItem;
    GridChartGraficaSeries2: TcxGridDBChartSeries;
    GridChartGraficaSeries3: TcxGridDBChartSeries;
    dbFechaFinal: TcxDBDateEdit;
    cxLabel7: TcxLabel;
    cxLabel8: TcxLabel;
    dbDescripcion: TcxDBMemo;
    cxLabel9: TcxLabel;
    dbGerencialFinal: TcxDBLookupComboBox;
    gridDbReprogramacionesColumn4: TcxGridDBColumn;
    cbbFolios: TcxComboBox;
    procedure FormCreate(Sender: TObject);
    procedure TfrmBarra1btnExitClick(Sender: TObject);
    procedure TfrmBarra1btnAddClick(Sender: TObject);
    procedure TfrmBarra1btnRefreshClick(Sender: TObject);
    procedure TfrmBarra1btnDeleteClick(Sender: TObject);
    procedure TfrmBarra1btnCancelClick(Sender: TObject);
    procedure TfrmBarra1btnPostClick(Sender: TObject);
    procedure TfrmBarra1btnEditClick(Sender: TObject);
    procedure GlobalKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure RecalcularReprogramaciones1Click(Sender: TObject);
    procedure ZReprogramacionesAfterScroll(DataSet: TDataSet);
    procedure Representacindeavances1Click(Sender: TObject);
    procedure dbFechaFinalKeyPress(Sender: TObject; var Key: Char);
    procedure cbbFoliosPropertiesCloseUp(Sender: TObject);
    procedure GlobalExit(Sender: TObject);
    procedure GlobalEnter(Sender: TObject);
  private
    { Private declarations }
    HorarioAnterior : string;
  public
    { Public declarations }
  end;

var
  frmReprogramacionFolio: TfrmReprogramacionFolio;

implementation

{$R *.dfm}

procedure TfrmReprogramacionFolio.cbbFoliosPropertiesCloseUp(Sender: TObject);
begin
  ZReprogramaciones.Active := False;
  ZReprogramaciones.ParamByName( 'orden' ).AsString := global_contrato;
  ZReprogramaciones.ParamByName( 'folio' ).AsString := cbbFolios.Text;
  ZReprogramaciones.Open;
end;

procedure TfrmReprogramacionFolio.dbFechaFinalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = Char( VK_RETURN ) then
  begin
    ZGerenciales.Active := False;
    ZGerenciales.Open; 
  end;
end;

procedure TfrmReprogramacionFolio.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TfrmReprogramacionFolio.FormCreate(Sender: TObject);
const
  HORARIOS : array[ 0..2 ] of string = ( '05:00', '17:00', '24:00' );
var
  Item : Integer;
begin

  ZProgramas.Active := False;
  ZProgramas.SQL.Text := 'select * from ordenesdetrabajo where sContrato =:Orden ';
  ZProgramas.ParamByName( 'orden' ).AsString := global_contrato;
  ZProgramas.Open;
  ZProgramas.First;

  CbbFolios.Properties.Items.Clear;
  while not ZProgramas.Eof do
  begin
    CbbFolios.Properties.Items.Add( ZProgramas.FieldByName( 'sNumeroOrden' ).AsString );
    ZProgramas.Next;
  end;
  ZProgramas.First;

  cbbFolios.ItemIndex := 0;
  Application.ProcessMessages;

  ZGerenciales.Active := False;
  ZGerenciales.SQL.Text := ObtenerSentencia( 'horarios_gerenciales', 'sql_horarios_gerenciales', ftCatalogo );
  ZGerenciales.ParamByName( 'todos' ).AsInteger := Integer( True );
  ZGerenciales.ParamByName( 'principales' ).AsInteger := Integer( False );
  ZGerenciales.Open;

  ZReprogramaciones.Active := False;
  ZReprogramaciones.SQL.Text := ObtenerSentencia( 'convenios', 'sql_reprogramaciones_folios', ftCatalogo );
  ZReprogramaciones.ParamByName( 'orden' ).AsString := global_contrato;
  ZReprogramaciones.ParamByName( 'folio' ).AsString := cbbFolios.Text;
  ZReprogramaciones.Open;
end;

procedure TfrmReprogramacionFolio.GlobalEnter(Sender: TObject);
begin
  if ( Sender is TcxDBTextEdit ) then
    ( Sender as TcxDBTextEdit ).Style.Color := global_color_entrada;

  if ( Sender is TcxDBCalcEdit ) then
    ( Sender as TcxDBCalcEdit ).Style.Color := global_color_entrada;

  if ( Sender is TcxDBDateEdit ) then
    ( Sender as TcxDBDateEdit ).Style.Color := global_color_entrada;

  if ( Sender is TcxDBLookupComboBox ) then
    ( Sender as TcxDBLookupComboBox ).Style.Color := global_color_entrada;

  if ( Sender is TcxDBMemo ) then
    ( Sender as TcxDBMemo ).Style.Color := global_color_entrada;
end;

procedure TfrmReprogramacionFolio.GlobalExit(Sender: TObject);
begin

  if ( Sender is TcxDBTextEdit ) then
    ( Sender as TcxDBTextEdit ).Style.Color := global_color_salida;

  if ( Sender is TcxDBCalcEdit ) then
    ( Sender as TcxDBCalcEdit ).Style.Color := global_color_salida;

  if ( Sender is TcxDBDateEdit ) then
    ( Sender as TcxDBDateEdit ).Style.Color := global_color_salida;

  if ( Sender is TcxDBLookupComboBox ) then
    ( Sender as TcxDBLookupComboBox ).Style.Color := global_color_salida;

  if ( Sender is TcxDBMemo ) then
    ( Sender as TcxDBMemo ).Style.Color := global_color_salida;
  
end;

procedure TfrmReprogramacionFolio.GlobalKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    Perform( CM_DIALOGKEY, VK_TAB, 0 );
    Key := 0
  end;
end;

procedure TfrmReprogramacionFolio.RecalcularReprogramaciones1Click( Sender: TObject);
var
  Fecha : string;
  Error : Boolean;
  ZCrearAvances : TZQuery;
begin
  try
    try
      connection.zConnection.StartTransaction;
      prgProgreso.Visible := True;
      Application.ProcessMessages;
      InicializarZDataSet( ZCrearAvances );
      ZCrearAvances.SQL.Text := ObtenerSentencia( 'avancesglobales', 'sql_generar_avances', ftInsert );
      ZCrearAvances.ParamByName( 'orden' ).AsString := global_contrato;
      ZCrearAvances.ParamByName( 'folio' ).AsString := cbbFolios.Text;
      ZCrearAvances.ParamByName( 'reprogramacion' ).AsString := ZReprogramaciones.FieldByName( 'sIdConvenio' ).AsString;
      ZCrearAvances.ExecSQL;
      Application.ProcessMessages;
      ReProgramarFolio( cbbFolios.Text, zReprogramaciones.FieldByName('sIdConvenio').AsString, prgProgreso, prgProgresoPartidas );
      connection.zConnection.Commit;
      prgProgreso.Visible := False;
      Error := False;
    except
      on e:Exception do
      begin
        Error := True;
        connection.zConnection.Rollback;
        TaskMessageDlg( 'Ha ocurrido un error', e.Message + #10+#13 + 'Se revertirán los cambios.', mtInformation, [ mbOK ], 0 );
      end;
    end;
  finally
    if not Error then    
      TaskMessageDlg( 'Listo', Format( 'Se ha reprogramado el folio %s ', [ ZReprogramaciones.FieldByName( 'sNumeroOrden' ).AsString ] ), mtInformation, [ mbOK ], 0 );
  end;
end;

procedure TfrmReprogramacionFolio.Representacindeavances1Click(Sender: TObject);
var
  Form : TForm;
begin
  try
    try
      ZGrafica.Active := False;
      ZGrafica.SQL.Text := ObtenerSentencia( 'avancesglobales', 'sql_proyeccion_reprogramacion', ftConsulta );
      ZGrafica.ParamByName( 'orden' ).AsString := global_contrato;
      ZGrafica.ParamByName( 'folio' ).AsString := cbbFolios.Text;
      ZGrafica.ParamByName( 'idreprogramacion' ).AsString := zReprogramaciones.FieldByName('sIdConvenio').AsString;
      ZGrafica.Open;
      ZGrafica.First;

      if ZGrafica.RecordCount = 0 then
        raise Exception.Create( 'No se hallaron datos.' );

      InicializarForm( Form, grpGrafica, 1000, 500 );
      Form.BorderStyle := bsSizeable;
      Form.WindowState := wsMaximized;

      Form.ShowModal;
      grpGrafica.Visible := False;
      grpGrafica.Align := alNone;
      grpGrafica.Width := 0;
      grpGrafica.Height := 0;
      grpGrafica.Left := 0;
      grpGrafica.Top := 0;
      grpGrafica.Parent := Self;

    except
      on e:Exception do
        TaskMessageDlg( 'Mensaje', e.Message, mtInformation, [ mbOK ], 0 );
    end;

  finally
    Form.Free;
  end;
end;

procedure TfrmReprogramacionFolio.TfrmBarra1btnAddClick(Sender: TObject);
begin
  TfrmBarra1.btnAddClick(Sender);
  ZReprogramaciones.Append;
  ZReprogramaciones.FieldByName( 'sNumeroOrden' ).AsString := cbbFolios.Text;
  grpCaptura.Enabled := True;
  gridReprogramaciones.Enabled := False; 
  dbIdReprogramacion.SetFocus;
end;

procedure TfrmReprogramacionFolio.TfrmBarra1btnCancelClick(Sender: TObject);
begin
  TfrmBarra1.btnCancelClick(Sender);
  ZReprogramaciones.Cancel;
  grpCaptura.Enabled := False;
  gridReprogramaciones.Enabled := True;
end;

procedure TfrmReprogramacionFolio.TfrmBarra1btnDeleteClick(Sender: TObject);
begin
  if ( ZReprogramaciones.RecordCount > 0 ) and ( TaskMessageDlg( 'Confirmación', '¿Desea relamente eliminar el registro activo?', mtConfirmation, [ mbYes, mbCancel ], 0 ) = mrYes ) then
    ZReprogramaciones.Delete;    
end;

procedure TfrmReprogramacionFolio.TfrmBarra1btnEditClick(Sender: TObject);
begin
  if ZReprogramaciones.RecordCount > 0 then
  begin
    TfrmBarra1.btnEditClick(Sender);
    ZReprogramaciones.Edit;
    grpCaptura.Enabled := True;
    gridReprogramaciones.Enabled := False;
    dbIdReprogramacion.SetFocus;
  end;
end;

procedure TfrmReprogramacionFolio.TfrmBarra1btnExitClick(Sender: TObject);
begin
  TfrmBarra1.btnExitClick(Sender);
  Close;
end;

procedure TfrmReprogramacionFolio.TfrmBarra1btnPostClick(Sender: TObject);
var
  ZActualiza : TZQuery;
  GenerarAvances : Boolean;
begin
  try
    try
      OpcButton := EmptyStr;

      GenerarAvances := ZReprogramaciones.State = dsInsert;

      Application.ProcessMessages;
      TfrmBarra1.btnPostClick(Sender);
      ZReprogramaciones.FieldByName( 'sContrato' ).AsString := global_contrato;
      ZReprogramaciones.FieldByName( 'sNumeroOrden' ).AsString := dbFolio.Text;
      ZReprogramaciones.FieldByName( 'dFecha' ).AsDateTime := dbFechaInicio.Date;
      ZReprogramaciones.FieldByName( 'dFechaInicio' ).AsDateTime := dbFechaInicio.Date;
      ZReprogramaciones.FieldByName( 'dFechaFinal' ).AsDateTime := dbFechaFinal.Date;
      ZReprogramaciones.FieldByName( 'dDuracion' ).AsFloat := dbDias.Value;
      ZGerenciales.Locate( 'Horario', dbGerencialInicio.Text, [] );
      ZReprogramaciones.FieldByName( 'iGerencialInicio' ).AsInteger := ZGerenciales.FieldByName( 'IdHorarioGerencial' ).AsInteger;
      ZGerenciales.Locate( 'Horario', dbGerencialFinal.Text, [] );
      ZReprogramaciones.FieldByName( 'iGerencialFinal' ).AsInteger := ZGerenciales.FieldByName( 'IdHorarioGerencial' ).AsInteger;
      ZReprogramaciones.FieldByName( 'sHorarioInicio' ).AsString := dbGerencialInicio.Text;
      ZReprogramaciones.FieldByName( 'sHorarioFinal' ).AsString := dbGerencialFinal.Text;
      ZReprogramaciones.Post;

      //Se comenta esta parte para subsea 7
//      InicializarZDataSet( ZActualiza );
//      ZActualiza.SQL.Text := ObtenerSentencia( 'avancesglobales', 'sql_actualiza_id_reprogramacion', ftUpdate );
//      ZActualiza.ParamByName( 'orden' ).AsString := global_contrato;
//      ZActualiza.ParamByName( 'folio' ).AsString := dbFolio.Text;
//      ZActualiza.ExecSQL;

      gridReprogramaciones.Enabled := True;
    except
      on e:Exception do
        TaskMessageDlg( 'Mensaje', e.Message, mtInformation, [ mbOK ], 0 );    
    end;
  finally
    ZActualiza.Free;
  end;
end;

procedure TfrmReprogramacionFolio.TfrmBarra1btnRefreshClick(Sender: TObject);
begin
  ZReprogramaciones.Active := False;
  ZReprogramaciones.Open;
end;

procedure TfrmReprogramacionFolio.ZReprogramacionesAfterScroll(
  DataSet: TDataSet);
begin
  ZGerenciales.Active := False;
  ZGerenciales.Open;
end;

end.
