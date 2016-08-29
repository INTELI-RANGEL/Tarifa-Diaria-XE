unit frm_Calidad_Rir;

interface

uses

  {$region 'Uses'}

  frm_connection, global,

  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, cxLookAndFeels, cxLookAndFeelPainters, Menus,
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
  dxSkinXmas2008Blue, StdCtrls, cxButtons, ImgList, ExtCtrls, cxControls,
  cxStyles, dxSkinscxPCPainter, cxCustomData, cxFilter, cxData, cxDataStorage,
  cxEdit, cxNavigator, DB, cxDBData, cxGridLevel, cxClasses, cxGridCustomView,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxGrid,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, cxContainer, cxGroupBox,
  cxLabel, cxTextEdit, cxMaskEdit, cxDropDownEdit, cxCalendar, cxDBEdit,
  cxLookupEdit, cxDBLookupEdit, cxDBExtLookupComboBox, cxCalc, JvComponentBase,
  JvValidators, cxCheckBox, cxDBLookupComboBox, frxClass, frxDBSet, UnitTarifa_Calidad;

  {$endregion}

type
  TfrmCalidad_Rir = class(TForm)

    {$region 'Componentes'}

    Panel1: TPanel;
    cxImgAcciones: TcxImageList;
    btnInsertar: TcxButton;
    btnEditar: TcxButton;
    btnEliminar: TcxButton;
    btnGuardar: TcxButton;
    btnCancelar: TcxButton;
    btnImprimir: TcxButton;
    cxGridDatosVista: TcxGridDBTableView;
    gridDatosLevel1: TcxGridLevel;
    gridDatos: TcxGrid;
    zCalidad_Rir: TZQuery;
    zCalidad_Riridregistro: TIntegerField;
    zCalidad_Rirfecha: TDateField;
    zCalidad_Rirnumero_oc: TStringField;
    zCalidad_Rirnumero_aviso_emb: TStringField;
    zCalidad_RirsContrato: TStringField;
    zCalidad_RirsNumeroReporte: TStringField;
    zCalidad_RirsIdPlataforma: TStringField;
    zCalidad_RirsIdEmbarcacion: TStringField;
    zCalidad_RirsIdProveedor: TStringField;
    zCalidad_RirdCantidad: TFloatField;
    zCalidad_RirsIdInsumo: TStringField;
    zCalidad_RirsReferencia: TStringField;
    zCalidad_RirsObservaciones: TStringField;
    zCalidad_RirsInstrumento: TStringField;
    dsCalidad_rir: TDataSource;
    cxGridDatosVistafecha: TcxGridDBColumn;
    cxGridDatosVistanumero_oc: TcxGridDBColumn;
    cxGridDatosVistanumero_aviso_emb: TcxGridDBColumn;
    cxGridDatosVistasNumeroReporte: TcxGridDBColumn;
    cxGridDatosVistasIdPlataforma: TcxGridDBColumn;
    cxGridDatosVistasIdEmbarcacion: TcxGridDBColumn;
    cxGridDatosVistadCantidad: TcxGridDBColumn;
    cxGridDatosVistasIdInsumo: TcxGridDBColumn;
    cxGridDatosVistasInstrumento: TcxGridDBColumn;
    grpCaptura: TcxGroupBox;
    dbFecha: TcxDBDateEdit;
    dbOrdenCompra: TcxDBTextEdit;
    dbAvEmbarque: TcxDBTextEdit;
    dbReporte: TcxDBTextEdit;
    cbbPlataforma: TcxDBExtLookupComboBox;
    cbbEmbarcacion: TcxDBExtLookupComboBox;
    cbbProveedor: TcxDBExtLookupComboBox;
    dbCantidad: TcxDBCalcEdit;
    cbbInsumo: TcxDBExtLookupComboBox;
    dbReferencia: TcxDBTextEdit;
    dbObservaciones: TcxDBTextEdit;
    dbInstrumento: TcxDBTextEdit;
    zPlataformas: TZQuery;
    zEmbarcaciones: TZQuery;
    zProveedores: TZQuery;
    zInsumos: TZQuery;
    dsPlataformas: TDataSource;
    dsEmbarcaciones: TDataSource;
    dsProveedores: TDataSource;
    dsInsumos: TDataSource;
    dxGrids: TcxGridViewRepository;
    zPlataformassIdPlataforma: TStringField;
    zPlataformassDescripcion: TStringField;
    zEmbarcacionessIdEmbarcacion: TStringField;
    zEmbarcacionessDescripcion: TStringField;
    zProveedoressIdProveedor: TStringField;
    zProveedoressRazon: TStringField;
    cxGridPlataformas: TcxGridDBTableView;
    cxGridPlataformassIdPlataforma: TcxGridDBColumn;
    cxGridPlataformassDescripcion: TcxGridDBColumn;
    cxGridEmbarcaciones: TcxGridDBTableView;
    cxGridEmbarcacionessIdEmbarcacion: TcxGridDBColumn;
    cxGridEmbarcacionessDescripcion: TcxGridDBColumn;
    cxGridProveedores: TcxGridDBTableView;
    cxGridProveedoressIdProveedor: TcxGridDBColumn;
    cxGridProveedoressRazon: TcxGridDBColumn;
    zInsumossIdInsumo: TStringField;
    zInsumosmDescripcion: TMemoField;
    cxGridInsumos: TcxGridDBTableView;
    cxGridInsumossIdInsumo: TcxGridDBColumn;
    cxGridInsumosmDescripcion: TcxGridDBColumn;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    btnConceptos: TcxButton;
    grpConceptos: TcxGroupBox;
    Panel2: TPanel;
    btnInsertarConcepto: TcxButton;
    btnEditarConcepto: TcxButton;
    btnEliminarConcepto: TcxButton;
    btnGuardarConcepto: TcxButton;
    btnCancelarConcepto: TcxButton;
    cxGrid1DBTableView1: TcxGridDBTableView;
    cxGrid1Level1: TcxGridLevel;
    cxGrid1: TcxGrid;
    Label13: TLabel;
    dbDescripcionConcepto: TcxDBTextEdit;
    chkSubConcepto: TcxCheckBox;
    cbbPadre: TcxDBLookupComboBox;
    zCalidad_Conceptos: TZQuery;
    zSrcCalidad_Conceptos: TZQuery;
    dsCalidad_Conceptos: TDataSource;
    dsSrc_Calidad_Conceptos: TDataSource;
    dbResultado: TcxDBTextEdit;
    Label14: TLabel;
    cxGrid1DBTableView1sDescripcion: TcxGridDBColumn;
    cxGrid1DBTableView1sresultado: TcxGridDBColumn;
    lblPaquete: TLabel;
    frxTitulo: TfrxDBDataset;
    frxCuerpo: TfrxDBDataset;
    frxDetalle: TfrxDBDataset;
    frxFirmantes: TfrxDBDataset;
    zqryTitulo: TZQuery;
    zqryCuerpo: TZQuery;
    zqryDetalle: TZQuery;
    zqryFirmante: TZQuery;
    reporteRIR: TfrxReport;
    procedure FormCreate(Sender: TObject);
    procedure btnInsertarClick(Sender: TObject);
    procedure btnEditarClick(Sender: TObject);
    procedure btnGuardarClick(Sender: TObject);
    procedure btnCancelarClick(Sender: TObject);
    procedure btnEliminarClick(Sender: TObject);
    procedure GlobalKeyUp(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure chkSubConceptoClick(Sender: TObject);
    procedure btnInsertarConceptoClick(Sender: TObject);
    procedure btnEditarConceptoClick(Sender: TObject);
    procedure btnGuardarConceptoClick(Sender: TObject);
    procedure btnCancelarConceptoClick(Sender: TObject);
    procedure btnEliminarConceptoClick(Sender: TObject);
    procedure btnConceptosClick(Sender: TObject);
    procedure imprimirRIR();
    procedure btnImprimirClick(Sender: TObject);
    procedure reporteRIRGetValue(const VarName: string; var Value: Variant);
    {$endregion}

  private
    { Private declarations }
    inciso : Integer;

    procedure BloquearBotonesRIR( Lock_Actions : Boolean );
    procedure BloquearBotonesRIR_Conceptos( Lock_Actions : Boolean );
    procedure CamposRIRListos();
    function ValidaCamposRIR():Boolean;
    function ExisteReporte( Fecha : TDate; Reporte : string ):Boolean;
  public
    { Public declarations }
  end;

const

  SQL_VALIDA_CALIDAD_RIR : string = 'select idregistro '+
                                    'from calidad_rir '+
                                    'where fecha = :fecha '+
                                    'and sContrato = :contrato '+
                                    'and sNumeroReporte = :num_rep limit 1; ';

var
  frmCalidad_Rir: TfrmCalidad_Rir;

implementation

{$R *.dfm}

procedure TfrmCalidad_Rir.BloquearBotonesRIR( Lock_Actions : Boolean );
begin
  btnInsertar.Enabled := not Lock_Actions;
  btnEditar.Enabled := not Lock_Actions;
  btnEliminar.Enabled :=  not Lock_Actions;
  btnImprimir.Enabled := not Lock_Actions;

  btnGuardar.Enabled := Lock_Actions;
  btnCancelar.Enabled := Lock_Actions;
  grpCaptura.Enabled := Lock_Actions;

end;

procedure TfrmCalidad_Rir.btnCancelarClick(Sender: TObject);
begin
  BloquearBotonesRIR( False );
  zCalidad_Rir.Cancel;
  CamposRIRListos;
end;

procedure TfrmCalidad_Rir.btnCancelarConceptoClick(Sender: TObject);
begin
  BloquearBotonesRIR_Conceptos( False );
  zCalidad_Conceptos.Cancel;
end;

procedure TfrmCalidad_Rir.btnConceptosClick(Sender: TObject);
var
  Form : TForm;
begin
  Form := TForm.Create( nil );
  Form.Width := 720;
  Form.Height := 400;
  Form.BorderStyle := bsDialog;
  Form.Position := poScreenCenter;

  grpConceptos.Visible := True;
  grpConceptos.Parent := Form;
  grpConceptos.Align := alClient;

  zCalidad_Conceptos.Active := False;
  zCalidad_Conceptos.ParamByName( 'id_rir' ).AsInteger := zCalidad_Rir.FieldByName( 'idregistro' ).AsInteger;
  zCalidad_Conceptos.Open;

  Form.ShowModal;

  //grpConceptos.Visible := False;
  grpConceptos.Parent := Self;
  grpConceptos.Align := alNone;
  grpConceptos.Left := 0;
  grpConceptos.Top := 0;
  grpConceptos.Width := 0;
  grpConceptos.Height := 0;

end;

procedure TfrmCalidad_Rir.btnEditarClick(Sender: TObject);
begin
  BloquearBotonesRIR( True );
  zCalidad_Rir.Edit;
end;

procedure TfrmCalidad_Rir.btnEditarConceptoClick(Sender: TObject);
begin
  BloquearBotonesRIR_Conceptos( True );
  zCalidad_Conceptos.Edit;
end;

procedure TfrmCalidad_Rir.btnEliminarClick(Sender: TObject);
begin
  if MessageDlg( '¿Desea eliminar el registro activo?', mtConfirmation, [ mbYes, mbCancel ], 0 ) = mrYes then
    zCalidad_Rir.Delete;
end;

procedure TfrmCalidad_Rir.btnEliminarConceptoClick(Sender: TObject);
var
  zAcciones : TZQuery;
begin
  if MessageDlg( '¿Desea eliminar el registro activo?', mtConfirmation, [ mbYes, mbCancel ], 0 ) = mrYes then
  begin
    try
      zAcciones := TZQuery.Create( nil );
      zAcciones.Connection := connection.zConnection;
      zAcciones.SQL.Text := 'delete from calidad_conceptos_rir where id_registro_rir = :id_rir and padre = :id_concepto ';
      zAcciones.ParamByName( 'id_rir' ).AsInteger := zCalidad_Rir.FieldByName( 'idregistro' ).AsInteger;
      zAcciones.ParamByName( 'id_concepto' ).AsInteger := zCalidad_Conceptos.FieldByName( 'idregistro' ).AsInteger;
      zAcciones.ExecSQL;

      zCalidad_Conceptos.Delete;
    finally
      zAcciones.Free;
    end;
  end;
end;

procedure TfrmCalidad_Rir.btnGuardarClick(Sender: TObject);
var
  f : integer;
begin

  try

    if ( zCalidad_Rir.State in [ dsInsert ] )
       and  ( ValidaCamposRIR() )
       and ( not ExisteReporte( zCalidad_Rir.FieldByName( 'fecha' ).AsDateTime, zCalidad_Rir.FieldByName( 'sNumeroReporte' ).AsString ) ) then
      zCalidad_Rir.Post
    else
    begin
      if ( zCalidad_Rir.State in [ dsEdit ] ) and ( ValidaCamposRIR() ) then
        zCalidad_Rir.Post;
    end;

    if zCalidad_Rir.State in [ dsBrowse ] then
    begin
      BloquearBotonesRIR( False );
      CamposRIRListos;
    end;

  except
    on e:Exception do
      MessageDlg( e.Message, mtInformation, [ mbOK ], 0 )
  end;
end;

procedure TfrmCalidad_Rir.btnGuardarConceptoClick(Sender: TObject);
begin
  BloquearBotonesRIR_Conceptos( False );
  if ( Length( Trim( dbDescripcionConcepto.Text ) ) > 0 ) then
  begin
    zCalidad_Conceptos.Post;
    Application.ProcessMessages;

    if not chkSubConcepto.Checked then
    begin
      zCalidad_Conceptos.Edit;
      zCalidad_Conceptos.FieldByName( 'padre' ).AsInteger := zCalidad_Conceptos.FieldByName( 'idregistro' ).AsInteger;  
      zCalidad_Conceptos.Post;
      
    end;

    zCalidad_Conceptos.Refresh;
  end
  else
    ShowMessage( 'Hay campos vacios.' );
end;

procedure TfrmCalidad_Rir.btnImprimirClick(Sender: TObject);
begin
  inciso := 97;
  imprimirRIR;
end;

procedure TfrmCalidad_Rir.btnInsertarClick(Sender: TObject);
begin
  BloquearBotonesRIR( True );
  zCalidad_Rir.Append;
  zCalidad_Rir.FieldByName( 'idregistro' ).AsInteger := 0;
  zCalidad_Rir.FieldByName( 'scontrato' ).AsString := global_contrato;

  dbFecha.SetFocus;
end;

procedure TfrmCalidad_Rir.btnInsertarConceptoClick(Sender: TObject);
begin
  BloquearBotonesRIR_Conceptos( True );
  zCalidad_Conceptos.Append;
  zCalidad_Conceptos.FieldByName( 'id_registro_rir' ).AsInteger := zCalidad_Rir.FieldByName( 'idregistro' ).AsInteger;
  chkSubConcepto.Checked := False;

  dbDescripcionConcepto.SetFocus;
end;

procedure TfrmCalidad_Rir.FormCreate(Sender: TObject);
begin
  zPlataformas.Active := False;
  zPlataformas.Open;

  zEmbarcaciones.Active := False;
  zEmbarcaciones.ParamByName( 'contrato' ).AsString := global_Contrato_Barco;
  zEmbarcaciones.Open;

  zProveedores.Active := False;
  zProveedores.Open;

  zInsumos.Active := False;
  zInsumos.ParamByName( 'contrato' ).AsString := global_Contrato_barco;
  zInsumos.Open;

  zCalidad_Rir.Active := False;
  zCalidad_Rir.ParamByName( 'contrato' ).AsString := global_contrato;
  zCalidad_Rir.Open;

end;

procedure TfrmCalidad_Rir.GlobalKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_RETURN then
  begin
    Perform( CM_DIALOGKEY, VK_TAB, 0 );
    Key := 0
  end;
end;

procedure TfrmCalidad_Rir.imprimirRIR;
begin
  //Imprime el reporte RIR

  //Abrir la consulta que trae los datos de la empresa(Imagen, Nombre de la Compania, direccion)
  with zqryTitulo do
  begin
    Active := False;
    Params.ParamByName('idregistro').AsInteger := zCalidad_Rir.FieldByName('idregistro').AsInteger;
    Open;
  end;

  with zqryCuerpo do
  begin
    Active := False;
    Params.ParamByName('idregistro').AsInteger := zCalidad_Rir.FieldByName('idregistro').AsInteger;
    Open;
  end;

  with zqryDetalle do
  begin
    Active := False;
    Params.ParamByName('idregistro').AsInteger := zCalidad_Rir.FieldByName('idregistro').AsInteger;
    Open;
  end;

  //FirmasPDF_Generales(zqryCuerpo, reporteRIR, FtAbordo);

  reporteRIR.LoadFromFile (global_files + 'CAMSA_Calidad_RIR.fr3');
  reporteRIR.ShowReport();
end;

procedure TfrmCalidad_Rir.reporteRIRGetValue(const VarName: string;
  var Value: Variant);
var
  Dia, Mes, Ano : Word;

begin

  //Mandar el dia, mes y Año por separado para las variables del reporte
  if VarName = 'Dia' then
  begin
    DecodeDate(dbFecha.Date, Ano, Mes, Dia);
    Value := IntToStr(Dia);
  end

  else if VarName = 'Mes' then
  begin
    DecodeDate(dbFecha.Date, Ano, Mes, Dia);
    Value := Mes;
  end

  else if VarName = 'Ano' then
  begin
    DecodeDate(dbFecha.Date, Ano, Mes, Dia);
    Value := Ano;
  end

  else if VarName = 'PadreHijo' then
  begin
    if zqryDetalle.FieldByName('PadreHijo').AsString = 'HIJO' then
    begin
      Value := '                         ' + zqryDetalle.FieldByName('nomenclatura').AsString;
    end

    else
    begin
      Value := Chr(inciso) + ') ' + zqryDetalle.FieldByName('nomenclatura').AsString;
      inciso := inciso + 1;
    end;
  end

  else if VarName = 'Firmante' then
  begin
    if sFirmante_CalidadRIR = null then
      Value := ''
    else
      Value := sFirmante_CalidadRIR;
    

  end

  else if VarName = 'PuestoFirmante' then
  begin
    Value := sPuesto_CalidadRIR;
  end;
end;

function TfrmCalidad_Rir.ExisteReporte(Fecha: TDate; Reporte: string):Boolean;
var
  zBusca : TZReadOnlyQuery;
begin

  zBusca := TZReadOnlyQuery.Create( nil );
  zBusca.Connection := connection.zConnection;
  zBusca.Active := False;
  zBusca.SQL.Text := SQL_VALIDA_CALIDAD_RIR;
  zBusca.ParamByName( 'contrato' ).AsString := global_contrato;
  zBusca.ParamByName( 'fecha' ).AsString := FormatDateTime( 'YYYY-MM-DD', Fecha );
  zBusca.ParamByName( 'num_rep' ).AsString := Reporte;
  zBusca.Open;

  Result := zBusca.RecordCount = 1;

end;

function TfrmCalidad_Rir.ValidaCamposRIR():Boolean;
var
  c : Integer;

const
  COLOR_OK : Integer = $0060DFCC;
  COLOR_ERROR : Integer = $008080FF;

begin
  Result := True;
  for c := 0 to grpCaptura.ControlCount - 1 do
  begin

    //Date
    if ( not ( grpCaptura.Controls[ c ].ClassType = TLabel ) )
       and ( grpCaptura.Controls[ c ].ClassType = TcxDBDateEdit ) then
    begin

      ( grpCaptura.Controls[ c ] as TcxDBDateEdit ).Style.Color := COLOR_OK;
      if Length( Trim( ( grpCaptura.Controls[ c ] as TcxDBDateEdit ).Text ) ) = 0 then
      begin
        ( grpCaptura.Controls[ c ] as TcxDBDateEdit ).Style.Color := COLOR_ERROR;
        Result := False;
      end;

    end;

    //Text
    if ( not ( grpCaptura.Controls[ c ].ClassType = TLabel ) )
       and ( grpCaptura.Controls[ c ].ClassType = TcxDBTextEdit ) then
    begin

      ( grpCaptura.Controls[ c ] as TcxDBTextEdit ).Style.Color := COLOR_OK;
      if Length( Trim( ( grpCaptura.Controls[ c ] as TcxDBTextEdit ).Text ) ) = 0 then
      begin
        ( grpCaptura.Controls[ c ] as TcxDBTextEdit ).Style.Color := COLOR_ERROR;
        Result := False;
      end;

    end;

    //Lookup
    if ( not ( grpCaptura.Controls[ c ].ClassType = TLabel ) )
       and ( grpCaptura.Controls[ c ].ClassType = TcxDBExtLookupComboBox ) then
    begin

      ( grpCaptura.Controls[ c ] as TcxDBExtLookupComboBox ).Style.Color := COLOR_OK;
      if Length( Trim( ( grpCaptura.Controls[ c ] as TcxDBExtLookupComboBox ).Text ) ) = 0 then
      begin
        ( grpCaptura.Controls[ c ] as TcxDBExtLookupComboBox ).Style.Color := COLOR_ERROR;
        Result := False;
      end;

    end;

    //Calc
    if ( not ( grpCaptura.Controls[ c ].ClassType = TLabel ) )
       and ( grpCaptura.Controls[ c ].ClassType = TcxDBCalcEdit ) then
    begin

      ( grpCaptura.Controls[ c ] as TcxDBCalcEdit ).Style.Color := COLOR_OK;
      if Length( Trim( ( grpCaptura.Controls[ c ] as TcxDBCalcEdit ).Text ) ) = 0 then
      begin
        ( grpCaptura.Controls[ c ] as TcxDBCalcEdit ).Style.Color := COLOR_ERROR;
        Result := False;
      end;

    end;

  end;

end;

procedure TfrmCalidad_Rir.CamposRIRListos;
var
  c : Integer;

begin
  for c := 0 to grpCaptura.ControlCount - 1 do
  begin

    //Date
    if ( not ( grpCaptura.Controls[ c ].ClassType = TLabel ) )
       and ( grpCaptura.Controls[ c ].ClassType = TcxDBDateEdit ) then
    begin
      ( grpCaptura.Controls[ c ] as TcxDBDateEdit ).Style.Color := clWindow;
    end;

    //Text
    if ( not ( grpCaptura.Controls[ c ].ClassType = TLabel ) )
       and ( grpCaptura.Controls[ c ].ClassType = TcxDBTextEdit ) then
    begin
      ( grpCaptura.Controls[ c ] as TcxDBTextEdit ).Style.Color := clWindow;
    end;

    //Lookup
    if ( not ( grpCaptura.Controls[ c ].ClassType = TLabel ) )
       and ( grpCaptura.Controls[ c ].ClassType = TcxDBExtLookupComboBox ) then
    begin
      ( grpCaptura.Controls[ c ] as TcxDBExtLookupComboBox ).Style.Color := clWindow;
    end;

    //Calc
    if ( not ( grpCaptura.Controls[ c ].ClassType = TLabel ) )
       and ( grpCaptura.Controls[ c ].ClassType = TcxDBCalcEdit ) then
    begin
      ( grpCaptura.Controls[ c ] as TcxDBCalcEdit ).Style.Color := clWindow;
    end;

  end;

end;

procedure TfrmCalidad_Rir.chkSubConceptoClick(Sender: TObject);
begin
  cbbPadre.Visible := chkSubConcepto.Checked;
  lblPaquete.Visible := chkSubConcepto.Checked;

  if chkSubConcepto.Checked then
  begin
    zSrcCalidad_Conceptos.Active := False;
    zSrcCalidad_Conceptos.ParamByName( 'id_rir' ).AsInteger := zCalidad_Rir.FieldByName( 'idregistro' ).AsInteger;
    zSrcCalidad_Conceptos.Open;

    cbbPadre.SetFocus;
  end;
end;

procedure TfrmCalidad_Rir.BloquearBotonesRIR_Conceptos(Lock_Actions: Boolean);
begin
  btnInsertarConcepto.Enabled := not Lock_Actions;
  btnEditarConcepto.Enabled := not Lock_Actions;
  btnEliminarConcepto.Enabled := not Lock_Actions;

  btnGuardarConcepto.Enabled := Lock_Actions;
  btnCancelarConcepto.Enabled := Lock_Actions;
  chkSubConcepto.Enabled := Lock_Actions;

end;

end.
