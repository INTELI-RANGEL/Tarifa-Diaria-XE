unit frm_detalledeinstalacion;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, StdCtrls, ComObj, UnitExcel,
  DBCtrls, DB, Menus, ADODB, frxClass, frxDBSet, ComCtrls,
  Buttons, ExtCtrls, jpeg, RxMemDS, utilerias, ZAbstractRODataset, ZDataset,
  DateUtils, unitexcepciones, Global, UnitTBotonesPermisos, DBDateTimePicker,
  AdvGroupBox, UnitPatrick;


type
  TfrmDetalledeInstalacion = class(TForm)
    ds_ordenesdetrabajo: TDataSource;
    dbResumen: TfrxDBDataset;
    btnReport2: TBitBtn;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    tsNumeroOrden: TDBLookupComboBox;
    GroupBox2: TGroupBox;
    rbConceptosTurnos: TRadioButton;
    rbSubActividadesdia: TRadioButton;
    rbSubActividadTurno: TRadioButton;
    rxAnexoGenerado: TRxMemoryData;
    StringField6: TStringField;
    StringField7: TStringField;
    MemoField1: TMemoField;
    StringField10: TStringField;
    FloatField1: TFloatField;
    rxAnexoGeneradodGenerado: TFloatField;
    rxAnexoGeneradodAnterior: TFloatField;
    rxAnexoGeneradodInstalado: TFloatField;
    rxAnexoGeneradodAcumulado: TFloatField;
    rxAnexoGeneradodVentaMN: TCurrencyField;
    rxAnexoGeneradodVentaDLL: TCurrencyField;
    rxAnexoGeneradodGeneradoAcumulado: TFloatField;
    OrdenesdeTrabajo: TZReadOnlyQuery;
    Detalle: TZReadOnlyQuery;
    frxDetalle: TfrxReport;
    ActividadesxOrden: TZReadOnlyQuery;
    ActividadesxOrdensContrato: TStringField;
    ActividadesxOrdensNumeroOrden: TStringField;
    ActividadesxOrdeniNivel: TIntegerField;
    ActividadesxOrdensSimbolo: TStringField;
    ActividadesxOrdensWbs: TStringField;
    ActividadesxOrdensWbsAnterior: TStringField;
    ActividadesxOrdensNumeroActividad: TStringField;
    ActividadesxOrdensTipoActividad: TStringField;
    ActividadesxOrdenmDescripcion: TMemoField;
    ActividadesxOrdendFechaInicio: TDateField;
    ActividadesxOrdendFechaFinal: TDateField;
    ActividadesxOrdendPonderado: TFloatField;
    ActividadesxOrdendVentaMN: TFloatField;
    ActividadesxOrdendVentaDLL: TFloatField;
    ActividadesxOrdeniColor: TIntegerField;
    ActividadesxOrdensIdConvenio: TStringField;
    ActividadesxOrdensWbsSpace: TStringField;
    ActividadesxOrdendAcumuladoAnterior: TFloatField;
    ActividadesxOrdendCantidadPeriodo: TFloatField;
    ActividadesxOrdendAcumulado: TFloatField;
    ActividadesxOrdensMedida: TStringField;
    ActividadesxOrdendCantidad: TFloatField;
    ActividadesxOrdendTotal: TFloatField;
    ActividadesxOrdendTotalAcumulado: TCurrencyField;
    ActividadesxOrdeniItemOrden: TStringField;
    dsActividadesxOrden: TfrxDBDataset;
    tdFechaInicial: TDBDateTimePicker;
    tdFechaFinal: TDBDateTimePicker;
    SaveDialog1: TSaveDialog;
    rbConceptosdia: TRadioButton;
    rbVolumetria: TRadioButton;
    rbConMateriales: TRadioButton;
    qryTitulo2: TZReadOnlyQuery;
    qryTitulo1: TZReadOnlyQuery;
    qryBitacoradeMateriales: TZReadOnlyQuery;
    rbPartidasSubActividades: TRadioButton;
    GrpBxPartidas: TAdvGroupBox;
    DescrL: TLabel;
    OptReportadas: TRadioButton;
    opcPartidas: TRadioButton;
    EditPartidas: TEdit;
    rbMaterialesInstalados: TRadioButton;
    rbPersonalReportado: TRadioButton;
    rbEquipoReportado: TRadioButton;
    ZQBuscaFechas: TZReadOnlyQuery;
    Label4: TLabel;
    grp1: TGroupBox;
    FechaInicio: TDateTimePicker;
    FechaTermino: TDateTimePicker;
    lbl1: TLabel;
    lbl2: TLabel;
    btn1: TButton;
    procedure ActualizaFirmas (dFecha : TDateTime ) ;
    procedure btnReport2Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tsNumeroOrdenKeyPress(Sender: TObject; var Key: Char);
    procedure tsNumeroOrdenEnter(Sender: TObject);
    procedure tsNumeroOrdenExit(Sender: TObject);
    procedure tdFechaInicialEnter(Sender: TObject);
    procedure tdFechaInicialExit(Sender: TObject);
    procedure tdFechaInicialKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaFinalEnter(Sender: TObject);
    procedure tdFechaFinalExit(Sender: TObject);
    procedure tdFechaFinalKeyPress(Sender: TObject; var Key: Char);
    procedure formatoEncabezado();
    Procedure RecursosReportados(ParamPartidas: string = ''; iTipo: Integer = 1);
    Procedure MaterialesInstalados(ParamPartidas: String = '');
    Procedure DatossubActividadesturnos(ParamPartidas: string = '');
    Procedure DatossubActividadesDias(ParamPartidas: string = '');
    Procedure DatosVolumetrias(ParamPartidas: string = '');
    Procedure DatosVolumetriasconmateriales(ParamPartidas: string = '');
    Procedure DatosPartidasPaquetes(ParamPartidas: string = '');
    procedure frxDetalleGetValue(const VarName: String;
      var Value: Variant);
    procedure ActividadesxOrdenCalcFields(DataSet: TDataSet);
    procedure tdFechaFinalChange(Sender: TObject);
    procedure tdFechaInicialChange(Sender: TObject);
    procedure EditPartidasEnter(Sender: TObject);

    procedure CabeceraExcel(Texto: String = '');
    procedure OrdenesdeTrabajoAfterScroll(DataSet: TDataSet);
    procedure tsNumeroOrdenClick(Sender: TObject);

    procedure AcumuladoDePartidas;
    function FormatearFecha(Fecha: TDate) : string;
    function VolverAInicioDeMes(Fecha : TDate) : TDate;
    procedure btn1Click(Sender: TObject);
  private
  sMenuP: String;
    { Private declarations }
    procedure AjustarTexto(var rangoE:Variant;TotalR:Integer;AddHeight:Extended);
  public
    { Public declarations }
  end;

var
  frmDetalledeInstalacion: TfrmDetalledeInstalacion;
  sSuperintendente, sSupervisor : String ;
  sPuestoSuperintendente, sPuestoSupervisor : String ;
  sTipoReporte : String ;
  dMontoMN, dMontoDLL : Currency ;
  BotonPermiso: TBotonesPermisos;
  columnas: array[1..1400] of string;
  Excel, Libro, Hoja: Variant;

implementation

uses masUtilerias;

{$R *.dfm}


procedure TfrmDetalledeInstalacion.AjustarTexto(var rangoE:Variant;TotalR:Integer;AddHeight:Extended);
var
  sngAnchoTotal,sngAnchoCelda,sngAlto:Extended;
  n:Integer;
begin
  sngAnchoTotal:=0;
  For n := 1 To TotalR do
    sngAnchoTotal := sngAnchoTotal + rangoE.columns.columns[n].ColumnWidth;

  sngAnchoCelda :=rangoE.columns.columns[1].ColumnWidth;
  rangoE.HorizontalAlignment := xlJustify;
  rangoE.VerticalAlignment := xlcenter;
  rangoE.MergeCells := False;

  if sngAnchoTotal>255 then
    rangoE.columns.columns[1].ColumnWidth :=255
  else
    rangoE.columns.columns[1].ColumnWidth := sngAnchoTotal;

  rangoE.parent.rows[rangoE.row].Autofit;
  sngAlto :=rangoE.RowHeight;

  rangoE.Merge;
  rangoE.Columns[1].EntireColumn.ColumnWidth := sngAnchoCelda;
  rangoE.Columns[1].RowHeight := sngAlto+AddHeight;
end;


procedure TfrmDetalledeInstalacion.btn1Click(Sender: TObject);
begin
  AcumuladoDePartidas;
end;

procedure TfrmDetalledeInstalacion.btnReport2Click(Sender: TObject);
Var
    dGenerado          : Real ;
    dGeneradoAcumulado : Real ;
    dReporteAnterior   : Real ;
    dInstalado         : Real ;
    CadError, OrdenVigencia: String;
    Resultado: Boolean;
    sPartidas:string;
Begin
  TamFont:=11;
  try
    Excel := CreateOleObject('Excel.Application');
  except
    FreeAndNil(Excel);
    showmessage('No es posible generar el ambiente de EXCEL, informe de esto al administrador del sistema.');
    Exit;
  end;

  Excel.Visible := True;
  Excel.DisplayAlerts := False;
  Excel.ScreenUpdating := True;

  Libro := Excel.Workbooks.Add;    // Crear el libro sobre el que se ha de trabajar

  // Verificar si cuenta con las hojas necesarias
  while Libro.Sheets.Count < 2 do
    Libro.Sheets.Add;

  // Verificar si se pasa de hojas necesarias
  Libro.Sheets[1].Select;
  while Libro.Sheets.Count > 1 do
    Excel.ActiveWindow.SelectedSheets.Delete;

  // Proceder a generar la hoja REPORTE
  CadError := '' ;
  Resultado := True;

  Hoja := Libro.Sheets[1];
  Hoja.Select;
  try
     Hoja.Name := 'Sub Actividades '+ global_contrato;
  Except
     Hoja.Name := 'Sub Actividades '+ global_contrato;
  end;

  sPartidas:='';
  if opcPartidas.Checked then
    sPartidas:=EditPartidas.Text;


   // Generar el ambiente de excel
  if rbSubActividadTurno.Checked= True then
  begin
    DatossubActividadesturnos(sPartidas);
   // Grabar el archivo de excel con el nombre dado
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end ;

  if rbConceptosTurnos.Checked= True then
  begin
    DatossubActividadesTurnos(sPartidas);
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end ;

  if rbSubactividadesDia.Checked= True then
  begin
    DatossubActividadesDias(sPartidas);
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end ;

  if rbConceptosDia.Checked= True then
  begin
    DatossubActividadesDias(sPartidas);
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end ;

  if rbVolumetria.Checked= True then
  begin
    DatosVolumetrias(sPartidas);
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end ;

  if rbConMateriales.Checked= True then
  begin
    DatosVolumetriasConmateriales(sPartidas);
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end;

  if rbPartidasSubActividades.Checked= True then
  begin
    DatosPartidasPaquetes(sPartidas);
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end;

  if rbMaterialesInstalados.Checked then begin
    MaterialesInstalados(sPartidas);
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end;

  if rbPersonalReportado.Checked then begin
    RecursosReportados(sPartidas, 1);
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end;

  if rbEquipoReportado.Checked then begin
    RecursosReportados(sPartidas, 2);
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;
    Excel := '';
    if CadError <> '' then
      showmessage(CadError);
  end;

End;

Procedure TfrmDetalledeInstalacion.CabeceraExcel(Texto: String = '');
Var
  iFila, iColumna: Integer;
  FechaInicial, FechaFinal: TDateTime;

  TmpName: String;
  TempPath: array [0..MAX_PATH-1] of Char;
  Fs: TStream;
  Pic : TJpegImage;
  imgAux: TImage;
begin
  Excel.Columns[ColumnaNombre(1) + ':' + ColumnaNombre(55)].ColumnWidth := 1;
  {$REGION 'IMAGENES DE CABECERA'}
  //Imagen Izquierda
  Try
    TmpName := '';
    imgAux := TImage.Create(nil);
    if TmpName='' then begin
      GetTempPath(SizeOf(TempPath), TempPath);
      TmpName:=TempPath +'imgtempSln1'+formatdatetime('dddddd hhnnss',now)+'.jpg';
      fs := Connection.contrato.CreateBlobStream(Connection.contrato.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
      If fs.Size > 1 Then Begin
        try
          Pic:=TJpegImage.Create;
          try
            Pic.LoadFromStream(fs);
            imgAux.Picture.Graphic := Pic;
          finally
            Pic.Free;
          end;
        finally
          fs.Free;
        End;
        imgAux.Picture.SaveToFile(TmpName);
      End;
    end;
  Finally
    imgAux.Free;
  End;
  Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 7, 7, 70, 35);
  //Imagen Derecha
  Try
    TmpName := '';
    imgAux := TImage.Create(nil);
    if TmpName='' then begin
      GetTempPath(SizeOf(TempPath), TempPath);
      TmpName:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';
      fs := Connection.configuracion.CreateBlobStream(Connection.configuracion.FieldByName('bImagen'), bmRead); //QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagenpep'), bmRead);
      If fs.Size > 1 Then Begin
        try
          Pic:=TJpegImage.Create;
          try
            Pic.LoadFromStream(fs);
            imgAux.Picture.Graphic := Pic;
          finally
            Pic.Free;
          end;
        finally
          fs.Free;
        End;
        imgAux.Picture.SaveToFile(TmpName);
      End;
    end;
  Finally
    imgAux.Free;
  End;
  Excel.ActiveSheet.Shapes.AddPicture(TmpName, True, True, 415, 7, 70, 35);
  Excel.Range['A2:BC2'].Select;
  PFormatosExcel_H2(Excel, 0, True, 10);
  Excel.Selection.Value := Connection.configuracion.FieldByName('sNombre').AsString;

  Excel.Range['A3:BC3'].Select;
  PFormatosExcel_H2(Excel, 0, True, 10);
  Excel.Selection.Value := Texto;
  {$ENDREGION}

  iFila := 5;
  iColumna := 1;

  FechaInicial := tdFechaInicial.date;
  FechaFinal := tdFechaFinal.Date;

  Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(20)+IntToStr(iFila)].Select;
  PFormatosExcel_H2(Excel, 20, True, 10);
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.Value := 'FOLIO';

  Excel.Range[ColumnaNombre(21)+IntToStr(iFila)+':'+ColumnaNombre(34)+IntToStr(iFila)].Select;
  PFormatosExcel_H2(Excel, 20, False, 10);
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.Value := Trim(tsNumeroOrden.Text);
  Inc(iFila);

  Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(20)+IntToStr(iFila)].Select;
  PFormatosExcel_H2(Excel, 20, True, 10);
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.Value := 'PERIODO DE CONSULTA';

  Excel.Range[ColumnaNombre(21)+IntToStr(iFila)+':'+ColumnaNombre(34)+IntToStr(iFila)].Select;
  PFormatosExcel_H2(Excel, 20, False, 10);
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.Value := FormatDateTime('yyyy-mm-dd', FechaInicial) + ' - ' + FormatDateTime('yyyy-mm-dd', FechaFinal);

end;

Procedure TfrmDetalledeInstalacion.RecursosReportados(ParamPartidas: string = ''; iTipo: Integer = 1);
Var
  AuxCadena: String;
  i, iFila, iColumna: Integer;
  FechaInicial, FechaFinal: TDateTime;
  QryConsulta: TZQuery;
Const
  TipoConsulta: Array[1..2, 1..4] Of String=(
                                              ('bitacoradepersonal','sIdPersonal','personal','PERSONAL'),
                                              ('bitacoradeequipos','sIdEquipo','equipos','EQUIPO')
                                            );
begin
  Excel.Columns[ColumnaNombre(1) + ':' + ColumnaNombre(70)].ColumnWidth := 1;
  Try
    CabeceraExcel('RECURSOS POR ACTIVIDAD - ' + TipoConsulta[iTipo, 4]);
    FechaInicial := tdFechaInicial.DateTime;
    FechaFinal := tdFechaFinal.DateTime;
    AuxCadena:='';
    for I := 1 to NumItems(ParamPartidas,',') do begin
      if AuxCadena='' then
        AuxCadena:=' and (ao.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
      else
        AuxCadena:=AuxCadena + ' or ao.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
    end;
    if AuxCadena <> '' then begin
      AuxCadena := AuxCadena + ') ';
    end;

    iFila := 8;

    {$REGION 'CABECERAS'}
    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'FECHA';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(14)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'FOLIO';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(15)+IntToStr(iFila)+':'+ColumnaNombre(19)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'ACTIVIDAD';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(20)+IntToStr(iFila)+':'+ColumnaNombre(23)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'CLAS.';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(24)+IntToStr(iFila)+':'+ColumnaNombre(29)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := '% ACUMULADO';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(30)+IntToStr(iFila)+':'+ColumnaNombre(34)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'HR. DE INICIO';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(35)+IntToStr(iFila)+':'+ColumnaNombre(39)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'HR. DE TERMINO';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(40)+IntToStr(iFila)+':'+ColumnaNombre(44)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'PARTIDA';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(45)+IntToStr(iFila)+':'+ColumnaNombre(49)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'CANT.';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(50)+IntToStr(iFila)+':'+ColumnaNombre(54)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'CANT. HH.';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(55)+IntToStr(iFila)+':'+ColumnaNombre(65)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'CATEGORIA';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);
    Inc(iFila);
    {$ENDREGION}

    {$REGION 'CONSULTA'}
    QryConsulta := TZQuery.Create(Self);
    QryConsulta.Connection := Connection.zConnection;
    QryConsulta.SQL.Text := '' +
                            'SELECT ' +
                            '	bp.dIdFecha,  ' +
                            '	bp.sNumeroOrden,  ' +
                            '	ba.sNumeroActividad,  ' +
                            ' IFNULL( ' +
                            '	( ' +
                            '		SELECT sIdClasificacion  ' +
                            '		FROM bitacoradeactividades AS bact ' +
                            '		WHERE bact.sContrato = bp.sContrato ' +
                            '		AND bact.sNumeroOrden = bp.sNumeroOrden  ' +
                            '		AND bact.dIdFecha = bp.dIdFecha ' +
                            '		AND bact.sNumeroActividad = ba.sNumeroActividad  ' +
                            '		AND bact.sIdTipoMovimiento = "ED"  ' +
                            '		AND (TIME(bact.sHoraInicio) <= TIME(bp.sHoraInicio) AND TIME(bact.sHoraFinal) >= TIME(bp.sHoraFinal))  ' +
                            '		LIMIT 1  ' +
                            '	), ' +
                            ' ( ' +
                            '		SELECT sIdClasificacion  ' +
                            '		FROM bitacoradeactividades AS bact ' +
                            '		WHERE bact.sContrato = bp.sContrato ' +
                            '		AND bact.sNumeroOrden = bp.sNumeroOrden  ' +
                            '		AND bact.dIdFecha = bp.dIdFecha ' +
                            '		AND bact.sNumeroActividad = ba.sNumeroActividad  ' +
                            '		AND bact.sIdTipoMovimiento = "ED"  ' +
                            '		AND ( TIME(bact.sHoraInicio) <= TIME(bp.sHoraInicio) ) ' +
                            '		LIMIT 1 ' +
                            ' ) ' +
                            ' ) AS sIdClasificacion, ' +
                            '	( ' +
                            '		SELECT  ' +
                            '			(IFNULL(SUM(bac.dAvance), 0))  ' +
                            '		FROM  ' +
                            '			bitacoradeactividades AS bac  ' +
                            '		WHERE  ' +
                            '		bac.sContrato = ba.sContrato  ' +
                            '		AND bac.sNumeroOrden = ba.sNumeroOrden  ' +
                            '		AND bac.sNumeroActividad = ba.sNumeroActividad  ' +
                            '		AND bac.sIdTipoMovimiento = "ED"  ' +
                            '		AND (  ' +
                            '			bac.didfecha < ba.didfecha  ' +
                            '			OR (  ' +
                            '				bac.didfecha = ba.didfecha  ' +
                            '				AND cast(bac.sHoraInicio AS Time) < cast(ba.sHoraInicio AS Time)  ' +
                            '			)  ' +
                            '		)  ' +
                            '	) AS dAvanceAnteriorPorPartida,  ' +
                            '	ba.dAvance,  ' +
                            '	bp.sHoraInicio,  ' +
                            '	bp.sHoraFinal,  ' +
                            '	bp.'+TipoConsulta[iTipo, 2]+',  ' +
                            '	bp.dCantidad,  ' +
                            '	bp.dCantHH,  ' +
                            '	p.sDescripcion  ' +
                            'FROM ' +
                            '	'+TipoConsulta[iTipo, 1]+' AS bp  ' +
                            '	INNER JOIN bitacoradeactividades AS ba 	 ' +
                            '		ON(ba.dIdFecha = bp.dIdFecha AND ba.sNumeroOrden = bp.sNumeroOrden AND ba.sContrato = bp.sContrato AND ba.iIdDiario = bp.iIdDiario)  ' +
                            '	INNER JOIN '+TipoConsulta[iTipo, 3]+' AS p ' +
                            '		ON(p.'+TipoConsulta[iTipo, 2]+' = bp.'+TipoConsulta[iTipo, 2]+' AND p.sContrato = :ContratoBarco)  ' +
                            '	INNER JOIN actividadesxorden AS ao ' +
                            '		ON(ao.sNumeroActividad = ba.sNumeroActividad AND ao.sContrato = ba.sContrato AND ao.sNumeroOrden = ba.sNumeroOrden)  ' +
                            'WHERE  ' +
                            ' bp.sContrato = :Contrato ' +
                            '	AND ba.sNumeroOrden = :Folio ' +
                            '	AND bp.dIdFecha >= :FechaInicial  ' +
                            '	AND bp.dIdFecha <= :FechaFinal  ' +
                            '	AND sTipoObra = "PU" ' +
                            'ORDER BY ' +
                            ' bp.dIdFecha, ' +
                            ' TIME(bp.sHoraFinal) ';
    QryConsulta.ParamByName('Contrato').AsString := Global_Contrato;
    QryConsulta.ParamByName('Folio').AsString := Trim(tsNumeroOrden.Text);
    QryConsulta.ParamByName('FechaInicial').AsDate := FechaInicial;
    QryConsulta.ParamByName('FechaFinal').AsDate := FechaFinal;
    QryConsulta.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
    QryConsulta.Open;
    {$ENDREGION}

    while Not QryConsulta.Eof do begin

      {$REGION 'REGISTROS'}
      Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(6)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 20, False, 7, 0, 'Arial', 'aaaa-mm-dd');
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName('dIdFecha').AsDateTime;

      Excel.Range[ColumnaNombre(7)+IntToStr(iFila)+':'+ColumnaNombre(14)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 20, False, 7);
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName('sNumeroOrden').AsString;

      Excel.Range[ColumnaNombre(15)+IntToStr(iFila)+':'+ColumnaNombre(19)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 20, False, 7, 0, 'Arial', '@');
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName('sNumeroActividad').AsString;

      Excel.Range[ColumnaNombre(20)+IntToStr(iFila)+':'+ColumnaNombre(23)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 0, False, 7);
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName('sIdClasificacion').AsString;

      Excel.Range[ColumnaNombre(24)+IntToStr(iFila)+':'+ColumnaNombre(29)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 0, False, 7, 0, 'Arial', '% #00.00');
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName('dAvanceAnteriorPorPartida').AsFloat + QryConsulta.FieldByName('dAvance').AsFloat;

      Excel.Range[ColumnaNombre(30)+IntToStr(iFila)+':'+ColumnaNombre(34)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 20, False, 7, 0, 'Arial', '@');
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName('sHoraInicio').AsString;

      Excel.Range[ColumnaNombre(35)+IntToStr(iFila)+':'+ColumnaNombre(39)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 20, False, 7, 0, 'Arial', '@');
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName('sHoraFinal').AsString;

      Excel.Range[ColumnaNombre(40)+IntToStr(iFila)+':'+ColumnaNombre(44)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 20, False, 7, 0, 'Arial', '@');
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName(TipoConsulta[iTipo, 2]).AsString;

      Excel.Range[ColumnaNombre(45)+IntToStr(iFila)+':'+ColumnaNombre(49)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 0, False, 7, 0, 'Arial', '0.00');
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName('dCantidad').AsFloat;

      Excel.Range[ColumnaNombre(50)+IntToStr(iFila)+':'+ColumnaNombre(54)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 0, False, 7, 0, 'Arial', '0.00000000');
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.Value := QryConsulta.FieldByName('dCantHH').AsFloat;

      Excel.Range[ColumnaNombre(55)+IntToStr(iFila)+':'+ColumnaNombre(65)+IntToStr(iFila)].Select;
      PFormatosExcel_H2(Excel, 20, False, 7);
      PFormatosExcel_Bordes(Excel);
      Excel.Selection.HorizontalAlignment := xlLeft;
      Excel.Selection.Value := QryConsulta.FieldByName('sDescripcion').AsString;
      PFormatosExcel_AjustarFila(Excel);
      {$ENDREGION}

      Inc(iFila);

      QryConsulta.Next;
    end;    
  Finally
    QryConsulta.Free;
  End;
end;

Procedure TfrmDetalledeInstalacion.MaterialesInstalados(ParamPartidas: string = '');
Var
  AuxCadena: String;
  i, iFila, iColumna: Integer;
  FechaInicial, FechaFinal: TDateTime;
  QryConsulta: TZQuery;
begin
  AuxCadena:='';
  for I := 1 to NumItems(ParamPartidas,',') do begin
    if AuxCadena='' then
      AuxCadena:=' and (ao.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
    else
      AuxCadena:=AuxCadena + ' or ao.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
  end;
  if AuxCadena <> '' then begin
    AuxCadena := AuxCadena + ') ';
  end;

  CabeceraExcel('MATERIALES INSTALADOS POR FOLIO');

  Excel.ActiveWindow.Zoom := 135;
  FechaInicial := tdFechaInicial.date;
  FechaFinal := tdFechaFinal.Date;

  iFila := 8;

  Try
    
    {$REGION 'CABECERAS'}
    Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'TRAZABILIDAD';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(16)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'ACTIVIDAD';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(17)+IntToStr(iFila)+':'+ColumnaNombre(22)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'FECHA';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(23)+IntToStr(iFila)+':'+ColumnaNombre(41)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'DESCRIPCIÓN';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(42)+IntToStr(iFila)+':'+ColumnaNombre(47)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'TOTAL';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);

    Excel.Range[ColumnaNombre(48)+IntToStr(iFila)+':'+ColumnaNombre(54)+IntToStr(iFila)].Select;
    PFormatosExcel_H2(Excel, 20, True, 7);
    PFormatosExcel_Bordes(Excel);
    Excel.Selection.Value := 'MEDIDA';
    PFormatosExcel_Rellenar(Excel, $00FFA04A);
    {$ENDREGION}

    {$REGION 'CONSULTA'}
    QryConsulta := TZQuery.Create(Self);
    QryConsulta.Connection := Connection.zConnection;

    QryConsulta.SQL.Add('' +
                        'SELECT ' +
                        '	ba.sWbs, ' +
                        '	ba.sNumeroActividad, ' +
                        ' bm.dIdFecha, ' +
                        '	bm.sIdMaterial, ' +
                        '	m.mDescripcion, ' +
                        '	bm.sTrazabilidad, ' +
                        '	bm.dCantidad, ' +
                        '	m.sMedida, ' +
                        '	bm.dCantidad AS dTotalInstalado ' +
                        'FROM ' +
                        '	bitacorademateriales AS bm ' +
                        '	INNER JOIN bitacoradeactividades AS ba ' +
                        '		ON (ba.sContrato = bm.sContrato AND ba.dIdFecha = bm.dIdFecha AND bm.sWbs = ba.sWbs and bm.iIdDiario = ba.iIdDiario) ' +
                        '	INNER JOIN insumos AS m ' +
                        '		ON (m.sIdInsumo = bm.sIdMaterial AND m.sContrato = bm.sContrato) ' +
                        '	INNER JOIN ordenesdetrabajo AS ot ' +
                        '		ON (ot.sContrato = bm.sContrato AND ot.sNumeroOrden = ba.sNumeroOrden) ' +
                        '	INNER JOIN actividadesxorden AS ao ' +
                        '		ON (ao.sContrato = ot.sContrato AND ao.sWbs = ba.sWbs) ' +
                        'WHERE ' +
                        ' bm.sContrato = :Contrato ' +
                        '	AND ba.sNumeroOrden = :Folio ' +
                        '	AND bm.dIdFecha >= :FechaInicio ' +
                        '	AND bm.dIdFecha <= :FechaFinal ' +
                        AuxCadena +
                        'GROUP BY ' +
                        ' bm.dIdFecha, bm.sIdMaterial ' +
                        '');
    QryConsulta.ParamByName('Contrato').AsString := Global_Contrato;
    QryConsulta.ParamByName('FechaInicio').AsDate := FechaInicial;
    QryConsulta.ParamByName('FechaFinal').AsDate := FechaFinal;
    QryConsulta.ParamByName('Folio').AsString := Trim(tsNumeroOrden.Text);
    QryConsulta.Open;
    {$ENDREGION}

    Inc(iFila);

    if QryConsulta.RecordCount > 0 then begin
      while Not QryConsulta.Eof do begin

        {$REGION 'REGISTROS'}
        Excel.Range[ColumnaNombre(2)+IntToStr(iFila)+':'+ColumnaNombre(11)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 20, False, 7);
        Excel.Selection.Value := QryConsulta.FieldByName('sTrazabilidad').AsString;

        Excel.Range[ColumnaNombre(12)+IntToStr(iFila)+':'+ColumnaNombre(16)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 7);
        Excel.Selection.Value := QryConsulta.FieldByName('sNumeroActividad').AsString;

        Excel.Range[ColumnaNombre(17)+IntToStr(iFila)+':'+ColumnaNombre(22)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 7, 0, 'Arial', 'dd-mm-aaaa');
        Excel.Selection.Value := QryConsulta.FieldByName('dIdFecha').AsDateTime;

        Excel.Range[ColumnaNombre(23)+IntToStr(iFila)+':'+ColumnaNombre(41)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 7);
        Excel.Selection.Value := QryConsulta.FieldByName('mDescripcion').AsString;

        Excel.Range[ColumnaNombre(42)+IntToStr(iFila)+':'+ColumnaNombre(47)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 7);
        Excel.Selection.Value := QryConsulta.FieldByName('dCantidad').AsFloat;

        Excel.Range[ColumnaNombre(48)+IntToStr(iFila)+':'+ColumnaNombre(54)+IntToStr(iFila)].Select;
        PFormatosExcel_H2(Excel, 0, False, 7);
        Excel.Selection.Value := QryConsulta.FieldByName('sMedida').AsString;;
        {$ENDREGION}

        Inc(iFila);
        QryConsulta.Next;
      end;
    end;

  Finally
    Hoja.Cells[2,2].Select;
    QryConsulta.Free;
  End;

end;


procedure TfrmDetalledeInstalacion.OrdenesdeTrabajoAfterScroll(
  DataSet: TDataSet);
begin
{  if ordenesdetrabajo.State = dsOpening then
    ShowMessage('abriendo')
  else
    ShowMessage('otro scroll'); }

end;

//////////////////////////
Procedure TfrmDetalledeInstalacion.DatossubActividadesturnos(ParamPartidas: string = '');
Var
  cadenasql, CadFecha, tmpNombre, cadena : String;
  fs: tStream;
  Alto : Extended;
  MiFechaI, MiFechaF, MiFecha: tDate;
  FechaDia : tDateTime;
  dAvanceTotal: double;
  Ren, nivel, fila, indice, i, total : integer;
  AuxCadena:string;
Begin
  AuxCadena:='';
  for I := 1 to NumItems(ParamPartidas,',') do
  begin
    if AuxCadena='' then
      AuxCadena:=' and (aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
    else
      AuxCadena:=AuxCadena + ' or aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
  end;

  if AuxCadena<>'' then
    AuxCadena:=AuxCadena + ') ';

  Ren := 2;
  // Realizar los ajustes visuales y de formato de hoja
  Excel.ActiveWindow.Zoom := 100;
  Excel.Columns['A:A'].ColumnWidth := 10;
  Excel.Columns['B:B'].ColumnWidth := 20;
  Excel.Columns['C:C'].ColumnWidth := 58;
  Excel.Columns['D:D'].ColumnWidth := 20;
  Excel.Columns['E:E'].ColumnWidth := 11;
  Excel.Columns['F:F'].ColumnWidth := 11;
  MiFecha := tdFechaInicial.date;
  MiFechaI := tdFechaInicial.date;
  MiFechaF := tdFechaFinal.Date;
  for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
  begin
      Hoja.Cells[Ren, 6 + i].Select;
      Excel.Selection.NumberFormat := '@';
      Excel.Selection.Value := DateToStr(MiFecha);
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment := xlCenter;
      Excel.Selection.Font.Color := clWhite;
      Excel.Selection.Font.Bold := True;
      Excel.Selection.Interior.ColorIndex := 41;
      MiFecha := IncDay(MiFecha);
  end;
  total := i;

  Hoja.Cells[Ren, 6 + i].Select;
  Excel.Selection.Value := 'Total';
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 16;
  // Colocar los encabezados de la plantilla...
  Hoja.Range['A2:A2'].Select;
  Excel.Selection.Value := 'Concepto';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Hoja.Range['B2:B2'].Select;
  Excel.Selection.Value := 'SubActividad';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Hoja.Range['C2:C2'].Select;
  Excel.Selection.Value := 'Descripcion';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Hoja.Range['D2:D2'].Select;
  Excel.Selection.Value := 'Turno';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;


  Hoja.Range['E2:E2'].Select;
  If rbConceptosTurnos.Checked = True  then
    Excel.Selection.Value := 'Id. Nota' ;
  If rbsubActividadTurno.Checked = True then
    Excel.Selection.Value := 'Hora Inicio';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Hoja.Range['F2:F2'].Select;
  If rbsubActividadTurno.Checked = True then
    Excel.Selection.Value := 'Hora Final';
  If rbConceptosTurnos.Checked = True  then
    Excel.Selection.Value := '' ;
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  connection.QryBusca2.Active := False ;
  Connection.QryBusca2.Filtered := False;
  connection.QryBusca2.SQL.Clear ;
  connection.QryBusca2.SQL.Add('select a.dIdFecha, a.sNumeroActividad, t.sDescripcion as turno, a.dCantidad, a.dAvance, a.sHoraInicio, a.sHoraFinal, b.sNumeroActividadSub, ' +
                               'a.iIdDiarioNota,  b.sDescripcion, a.dAvanceActividad from bitacoradealcances a ' +
                               'inner join actividadesxorden ao on(ao.scontrato=a.scontrato and ao.snumeroorden=a.snumeroorden and ao.swbs=a.swbs '+
                               'and ao.snumeroactividad=a.snumeroactividad) ' +
                               'inner join actividadesxanexo aa on(aa.scontrato=ao.scontrato and aa.sidconvenio=ao.sidconvenio and aa.swbs=ao.swbscontrato ' +
                               'and aa.snumeroactividad=ao.snumeroactividad) ' +
                               'Inner Join alcancesxactividad b On (b.sContrato=aa.sContrato And b.sIdConvenio=aa.sidconvenio And b.swbs=aa.swbs and b.sNumeroActividad=aa.sNumeroActividad And a.iFase=b.iFase) ' +
                               'Inner Join turnos_horas t On (a.sIdTurnoHora=t.sIdTurnoHora) ' +
                               'where a.sContrato = :Contrato And ao.sIdConvenio =:Convenio And a.dIdFecha >=:FechaInicial ' +
                               ' And a.dIdFecha <=:FechaFinal And a.lConceptoEjecutado=:Concepto And a.sPaquete="" '+
                               AuxCadena + ' and ao.snumeroorden=:orden Order by a.dIdFecha, ' +
                               'a.sNumeroActividad, b.sNumeroActividadSub, a.sIdTurnoHora') ;

  connection.QryBusca2.Params.ParamByName('Contrato').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Contrato').Value         := global_contrato;
  connection.QryBusca2.Params.ParamByName('Convenio').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Convenio').Value         := global_convenio;
  connection.QryBusca2.Params.ParamByName('FechaInicial').DataType  := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaInicial').Value     := tdFechaInicial.DateTime;
  connection.QryBusca2.Params.ParamByName('Concepto').DataType      := ftString;
  If rbConceptosTurnos.Checked = True  then
    connection.QryBusca2.Params.ParamByName('Concepto').Value := 'Si';
  If rbsubActividadTurno.Checked = True then
    connection.QryBusca2.Params.ParamByName('Concepto').Value := 'No';
  connection.QryBusca2.Params.ParamByName('FechaFinal').DataType    := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaFinal').Value       := tdFechaFinal.DateTime;
  Connection.QryBusca2.ParamByName('Orden').AsString:=Tsnumeroorden.KeyValue;
  Connection.QryBusca2.Open ;

  if connection.QryBusca2.RecordCount > 0 then
  begin
    while not connection.QryBusca2.Eof do
    begin
      Hoja.Cells[Ren+1,1].Select;
      Excel.Selection.Value := Connection.QryBusca2.FieldValues['sNumeroActividad'];
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Size := 11;
      Excel.Selection.Font.Bold := False;
      Excel.Selection.Font.Name := 'Calibri';

      Hoja.Cells[Ren+1,2].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['sNumeroActividadSub'];
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;

      Hoja.Cells[Ren+1,3].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['sDescripcion'];
      Excel.Selection.HorizontalAlignment   := xlLeft;
      Excel.Selection.HorizontalAlignment   := xlLeft;
      Excel.selection.WrapText := True    ;
      Excel.selection.Orientation := 0    ;
      Excel.selection.AddIndent := False  ;
      Excel.selection.IndentLevel := 0    ;
      Excel.selection.ShrinkToFit := False  ;
      Excel.selection.MergeCells := False  ;

      Hoja.Cells[Ren+1,4].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['turno'];
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;

      Hoja.Cells[Ren+1,5].Select;
      If rbConceptosTurnos.Checked = True  then
        Excel.Selection.Value := connection.QryBusca2.FieldValues['iIdDiarioNota'] ;
      If rbsubActividadTurno.Checked = True then
        Excel.Selection.Value := connection.QryBusca2.FieldValues['sHoraInicio'];
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;

      Hoja.Cells[Ren+1,6].Select;
      If rbsubActividadTurno.Checked = True then
        Excel.Selection.Value := connection.QryBusca2.FieldValues['sHoraFinal'];
      If rbConceptosTurnos.Checked = True  then
        Excel.Selection.Value := '' ;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      MiFecha := tdFechaInicial.date;
      dAvanceTotal := 0 ;
      for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
      begin
        cadena:= formatDateTime('dd/mm/yyyy',Mifecha);
        if connection.QryBusca2.FieldValues['dIdFecha'] = StrToDate(cadena) then
        begin
          Hoja.Cells[Ren+1, 6 + i].Select;
          Excel.Selection.NumberFormat := '#0.00';
          If rbSubActividadTurno.Checked = True  then
            Excel.Selection.Value := connection.QryBusca2.FieldValues['dAvanceActividad'] ;

          If rbConceptosTurnos.Checked = True  then
          begin
            Excel.Selection.Value := connection.QryBusca2.FieldValues['dCantidad'];
            Excel.Selection.NumberFormat := '#0.0000';
          end;
          Excel.Selection.HorizontalAlignment := xlRight;
          Excel.Selection.VerticalAlignment   := xlCenter;

          If rbSubActividadTurno.Checked = True  then
            dAvanceTotal := dAvanceTotal + connection.QryBusca2.FieldValues['dAvanceActividad'];
          If rbConceptosTurnos.Checked = True  then
            dAvanceTotal := dAvanceTotal + connection.QryBusca2.FieldValues['dCantidad'];
        end;
        Mifecha := IncDay(Mifecha);
      end ;
      Hoja.Cells[Ren+1, 6 + i].Select;
      Excel.Selection.NumberFormat := '#0.00';
      Excel.Selection.Value := FloatToStrF(dAvanceTotal, ffNumber, 4, 2 ) ;
      Excel.Selection.HorizontalAlignment := xlRight;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      connection.QryBusca2.Next;
      Inc(Ren);
    end;
  end;
  Hoja.Cells[2,2].Select;

end;
/////////////////////////



//////////////////////////
Procedure TfrmDetalledeInstalacion.DatossubActividadesDias(ParamPartidas: string = '');
Var
  cadenasql, CadFecha, tmpNombre, cadena,AuxCadena : String;
  fs: tStream;
  Alto : Extended;
  MiFechaI, MiFechaF, MiFecha: tDate;
  FechaDia : tDateTime;
  dAvanceTotal: double;
  Ren, nivel, fila, indice, i, total : integer;
Begin
  Ren := 2;
  AuxCadena:='';
  for I := 1 to NumItems(ParamPartidas,',') do
  begin
    if AuxCadena='' then
      AuxCadena:=' and (aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
    else
      AuxCadena:=AuxCadena + ' or aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
  end;

  if AuxCadena<>'' then
    AuxCadena:=AuxCadena + ') ';
  // Realizar los ajustes visuales y de formato de hoja
  Excel.ActiveWindow.Zoom := 100;
  Excel.Columns['A:A'].ColumnWidth := 10;
  Excel.Columns['B:B'].ColumnWidth := 20;
  Excel.Columns['C:C'].ColumnWidth := 58;

  MiFecha := tdFechaInicial.date;
  MiFechaI := tdFechaInicial.date;
  MiFechaF := tdFechaFinal.Date;
  for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
  begin
    Hoja.Cells[Ren, 3 + i].Select;
    Excel.Selection.NumberFormat := '@';
    Excel.Selection.Value := DateToStr(MiFecha);
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.VerticalAlignment := xlCenter;
    Excel.Selection.Font.Color := clWhite;
    Excel.Selection.Font.Bold := True;
    Excel.Selection.Interior.ColorIndex := 41;
    MiFecha := IncDay(MiFecha);
  end;
  total := i;

  Hoja.Cells[Ren, 3 + i].Select;
  Excel.Selection.Value := 'Total';
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 16;
  // Colocar los encabezados de la plantilla...
  Excel.ActiveSheet.Range['A3:L20'].Select;
  Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlEdgeLeft].Weight := xlThin;
  Excel.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlEdgeTop].Weight := xlThin;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;
  Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlEdgeRight].Weight := xlThin;
  Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlInsideVertical].Weight := xlThin;
  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlInsideHorizontal].Weight := xlThin;

  Hoja.Range['A2:A2'].Select;
  Excel.Selection.Value := 'Concepto';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Hoja.Range['B2:B2'].Select;
  Excel.Selection.Value := 'SubActividad';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Hoja.Range['C2:C2'].Select;
  Excel.Selection.Value := 'Descripcion';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  connection.QryBusca2.Active := False ;
  connection.QryBusca2.Filtered := False;
  connection.QryBusca2.SQL.Clear ;
  connection.QryBusca2.SQL.Add('select b.iFase,ao.swbs as swbsFrente, b.sNumeroActividad, b.dCantidad, b.dAvance, b.sHoraInicio, b.sHoraFinal, '+
                               'a.sNumeroActividadSub, b.iIdDiarioNota,  a.sDescripcion from bitacoradealcances b ' +
                               'inner join actividadesxorden ao on(ao.scontrato=b.scontrato and ao.snumeroorden=b.snumeroorden '+
                               'and ao.swbs=b.swbs and ao.snumeroactividad=b.snumeroactividad) '+
                               'inner join actividadesxanexo aa on(aa.scontrato=ao.scontrato and aa.sidconvenio=ao.sidconvenio and '+
                               'aa.swbs=ao.swbscontrato and aa.snumeroactividad=ao.snumeroactividad) ' +
                               'Inner Join alcancesxactividad a On (a.sContrato=aa.sContrato And a.sIdConvenio=aa.sidconvenio and a.swbs=aa.swbs And a.sNumeroActividad=aa.sNumeroActividad And a.iFase=b.iFase) ' +
                               'where aa.sContrato = :Contrato And aa.sIdConvenio =:Convenio And b.dIdFecha >=:FechaInicial ' +
                               auxCadena + ' and ao.snumeroorden=:orden ' +
                               'And b.dIdFecha <=:FechaFinal And lConceptoEjecutado=:Concepto And b.sPaquete="" ' +
                               'Group by b.sNumeroActividad, a.sNumeroActividadSub Order by  b.sNumeroActividad ') ;

  connection.QryBusca2.Params.ParamByName('Contrato').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Contrato').Value         := global_contrato;
  connection.QryBusca2.Params.ParamByName('Convenio').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Convenio').Value         := global_convenio;
  connection.QryBusca2.Params.ParamByName('FechaInicial').DataType  := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaInicial').Value     := tdFechaInicial.DateTime;
  connection.QryBusca2.Params.ParamByName('Concepto').DataType     := ftString;
  If rbsubActividadesdia.Checked = True then
        connection.QryBusca2.Params.ParamByName('Concepto').Value    := 'No'  ;
  If rbConceptosDia.Checked = True  then
       connection.QryBusca2.Params.ParamByName('Concepto').Value    := 'Si'  ;
  connection.QryBusca2.Params.ParamByName('FechaFinal').DataType    := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaFinal').Value       := tdFechaFinal.DateTime;
  connection.QryBusca2.ParamByName('orden').AsString:=tsnumeroorden.KeyValue;
  Connection.QryBusca2.Open ;

  if connection.QryBusca2.RecordCount > 0 then
  begin
    while not connection.QryBusca2.Eof do
    begin
      Hoja.Cells[Ren+1,1].Select;
      Excel.Selection.Value := Connection.QryBusca2.FieldValues['sNumeroActividad'];
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Size := 11;
      Excel.Selection.Font.Bold := False;
      Excel.Selection.Font.Name := 'Calibri';

      Hoja.Cells[Ren+1,2].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['sNumeroActividadSub'];
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;

      Hoja.Cells[Ren+1,3].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['sDescripcion'];
      Excel.Selection.HorizontalAlignment   := xlLeft;
      Excel.selection.WrapText := True    ;
      Excel.selection.Orientation := 0    ;
      Excel.selection.AddIndent := False  ;
      Excel.selection.IndentLevel := 0    ;
      Excel.selection.ShrinkToFit := False  ;
      Excel.selection.MergeCells := False  ;

      MiFecha := tdFechaInicial.date;
      dAvanceTotal := 0 ;
      for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
      begin
        cadena:= formatDateTime('dd/mm/yyyy',Mifecha);
        connection.QryBusca.Active := False ;
        connection.QryBusca.SQL.Clear ;

        connection.QryBusca.SQL.Add('select b.dIdFecha, b.iFase, b.sNumeroActividad, sum(b.dCantidad) as sumacantidad, Sum(b.dAvanceActividad) as sumaactividad, ' +
                                    'b.sHoraInicio, b.sHoraFinal, b.iIdDiarioNota from bitacoradealcances b  ' +
                                    'inner join actividadesxorden ao on(ao.scontrato=b.scontrato and ao.snumeroorden=b.snumeroorden '+
                                    'and ao.swbs=b.swbs and ao.snumeroactividad=b.snumeroactividad) '+
                                    'inner join actividadesxanexo aa on(aa.scontrato=ao.scontrato and aa.sidconvenio=ao.sidconvenio and '+
                                    'aa.swbs=ao.swbscontrato and aa.snumeroactividad=ao.snumeroactividad) ' +
                                    'Inner join alcancesxactividad a On (a.sContrato=aa.sContrato And a.sIdConvenio=aa.sidconvenio and a.swbs=aa.swbs '+
                                    'And a.sNumeroActividad=aa.sNumeroActividad And a.iFase=b.iFase ) '+
                                    'Where b.sContrato=:Contrato and aa.sIdConvenio=:Convenio and ao.snumeroorden=:orden and b.sPaquete=""  and lConceptoEjecutado=:Concepto ' +
                                    'And b.dIdFecha=:Fecha and ao.swbs=:wbs and b.sNumeroActividad=:Actividad and a.sNumeroActividadSub=:SubActividad ' +
                                    'Group By a.sNumeroActividadSub') ;

        connection.QryBusca.Params.ParamByName('Contrato').DataType     := ftString;
        connection.QryBusca.Params.ParamByName('Contrato').Value        := global_contrato;
        connection.QryBusca.Params.ParamByName('Convenio').DataType     := ftString;
        connection.QryBusca.Params.ParamByName('Convenio').Value        := global_convenio;
        connection.QryBusca.Params.ParamByName('Fecha').DataType        := ftDate;
        connection.QryBusca.Params.ParamByName('Fecha').Value           := MiFecha ;
        connection.QryBusca.Params.ParamByName('Actividad').DataType    := ftString;
        connection.QryBusca.Params.ParamByName('Actividad').Value       := connection.QryBusca2.FieldValues['sNumeroActividad'] ;
        connection.QryBusca.Params.ParamByName('Concepto').DataType     := ftString;
        If rbsubActividadesdia.Checked = True then
           connection.QryBusca.Params.ParamByName('Concepto').Value    := 'No'  ;
        If rbConceptosDia.Checked = True  then
           connection.QryBusca.Params.ParamByName('Concepto').Value    := 'Si'  ;
        connection.QryBusca.Params.ParamByName('SubActividad').DataType := ftString;
        connection.QryBusca.Params.ParamByName('SubActividad').Value    := connection.QryBusca2.FieldValues['sNumeroActividadSub'] ;
        connection.QryBusca.Params.ParamByName('wbs').AsString    := connection.QryBusca2.FieldValues['swbsFrente'] ;
        connection.QryBusca.Params.ParamByName('orden').AsString:=tsNumeroOrden.KeyValue;
        connection.QryBusca.Open ;
        if connection.QryBusca.RecordCount > 0 then
          if connection.QryBusca.FieldValues['dIdFecha'] = StrToDate(cadena) then
          begin
            Hoja.Cells[Ren+1, 3 + i].Select;
            Excel.Selection.NumberFormat := '#0.00';
            If rbSubActividadesDia.Checked = True  then
              Excel.Selection.Value := connection.QryBusca.FieldValues['sumaactividad'] ;
            If rbConceptosDia.Checked = True  then
            begin
              Excel.Selection.Value := connection.QryBusca.FieldValues['sumacantidad'];
              Excel.Selection.NumberFormat := '#0.0000';
            end;
            Excel.Selection.HorizontalAlignment := xlRight;
            Excel.Selection.VerticalAlignment   := xlCenter;
            If rbSubActividadesDia.Checked = True  then
              dAvanceTotal := dAvanceTotal + connection.QryBusca.FieldValues['sumaActividad'];
            If rbConceptosDia.Checked = True  then
              dAvanceTotal := dAvanceTotal + connection.QryBusca.FieldValues['sumacantidad'];
          end;
        connection.QryBusca.Next ;
        Mifecha := IncDay(Mifecha);
      end ;
      Hoja.Cells[Ren+1, 3 + i].Select;
      Excel.Selection.NumberFormat := '#0.00';
      Excel.Selection.Value := FloatToStrF(dAvanceTotal, ffNumber, 4, 2 ) ;
      Excel.Selection.HorizontalAlignment := xlRight;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      connection.QryBusca2.Next;
      Inc(Ren);
    end;
  end;
  Hoja.Cells[2,2].Select;

end;

///////////////////////////////////////////////
Procedure TfrmDetalledeInstalacion.DatosVolumetrias(ParamPartidas: string = '');
Var
  cadenasql, CadFecha, tmpNombre, cadena : String;
  fs: tStream;
  Alto : Extended;
  MiFechaI, MiFechaF, MiFecha: tDate;
  FechaDia : tDateTime;
  dAvanceTotal: double;
  Ren, Ren2, nivel, fila, indice, i, total : integer;
  NombreArchivo, CadPeriodo: String;
  Imagen: TField;
  Altura, Margen: Extended;
  AuxCadena:string;
Begin
  Ren  := 10;
  Ren2 := 11 ;
  // Realizar los ajustes visuales y de formato de hoja
  Excel.ActiveWindow.Zoom := 100;
  Excel.Columns['A:A'].ColumnWidth := 16.43;
  Excel.Columns['B:B'].ColumnWidth := 68;
  Excel.Columns['C:C'].ColumnWidth := 9.57;
  Excel.Columns['D:D'].ColumnWidth := 11;

  AuxCadena:='';
  for I := 1 to NumItems(ParamPartidas,',') do
  begin
    if AuxCadena='' then
      AuxCadena:=' and (aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
    else
     AuxCadena:=AuxCadena + ' or aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
  end;

  if AuxCadena<>'' then
    AuxCadena:=AuxCadena + ') ';

  MiFecha := tdFechaInicial.date;
  MiFechaI := tdFechaInicial.date;
  MiFechaF := tdFechaFinal.Date;
  for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
  begin
    Hoja.Cells[Ren, 4 + i].Select;
    Excel.Selection.NumberFormat := '@';
    Excel.Selection.Value := DateToStr(MiFecha);
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.VerticalAlignment := xlCenter;
    Excel.Selection.Font.Color := clWhite;
    Excel.Selection.Font.Bold := True;
    Excel.Selection.Interior.ColorIndex := 41;
    MiFecha := IncDay(MiFecha);
  end;
  total := i;

  Hoja.Cells[Ren, 4 + i].Select;
  Excel.Selection.Value := 'Total';
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 16;

    {*****************************************************************************
     ** Colocar el logotipo de la empresa                                        }
  NombreArchivo := NombreArchivoTemporal;   // Obtener un nombre de archivo temporal
  Imagen := Connection.configuracion.FieldByName('bImagen');
  fs := Connection.configuracion.CreateBlobStream(Imagen, bmRead);
  if fs.size > 0 then
  Begin
    try
      fs.Seek(0, soFromBeginning);
      with TFileStream.Create(NombreArchivo, fmCreate) do
      try
        CopyFrom(fs, fs.Size)
      finally
        Free
      end;
    finally
      fs.Free
    end;
    Excel.ActiveSheet.Pictures.Insert(NombreArchivo).Select;
    // Determinar el tamaño real de la imagen
    Altura := Excel.Rows[1].Height * 0.7;
    Margen := (Excel.Rows[1].Height - Altura) / 2 ;
    Excel.Selection.ShapeRange.Left := Excel.Columns['A:A'].Width ;
    Excel.Selection.ShapeRange.Top := Margen;
    SysUtils.DeleteFile(NombreArchivo); // Borrar el archivo temporal
  End;
    {** Termina Colocar el logotipo de la empresa
     *****************************************************************************}

  Hoja.Range['C5:L5'].Select;
  Excel.Selection.Value := 'VOLUMETRIAS DE ANEXO CON SISTEMAS / PARTIDA ANEXO C';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Excel.Selection.HorizontalAlignment := xlLeft ;
  Excel.Selection.Merge ;
  Excel.Selection.WrapText := False   ;
  Excel.Selection.Orientation := 0     ;
  Excel.Selection.AddIndent := False   ;
  Excel.Selection.IndentLevel := 0     ;
  Excel.Selection.ShrinkToFit := False  ;
  Excel.Selection.MergeCells := False     ;

  Hoja.Range['A10:A10'].Select;
  Excel.Selection.Value := 'PARTIDA ANEXO C';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Hoja.Range['B10:B10'].Select;
  Excel.Selection.Value := 'DESCRIPCION';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Hoja.Range['C10:C10'].Select;
  Excel.Selection.Value := 'UNIDAD';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Hoja.Range['D10:D10'].Select;
  Excel.Selection.Value := 'CANTIDAD';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  connection.QryBusca2.Active := False ;
  connection.QryBusca2.Filtered := False;
  connection.QryBusca2.SQL.Clear ;
  connection.QryBusca2.SQL.Add('select b.iFase,ao.swbs as swbsFrente, b.sNumeroActividad, aa.mDescripcion, aa.sMedida, aa.dCantidadAnexo, b.dCantidad, b.dAvance, '+
                               'b.iIdDiarioNota, aa.sWbsAnterior from bitacoradealcances b ' +
                               'inner join actividadesxorden ao on(ao.scontrato=b.scontrato and ao.snumeroorden=b.snumeroorden and '  +
                               'ao.swbs=b.swbs and ao.snumeroactividad=b.snumeroactividad) ' +
                               'Inner Join actividadesxanexo  aa On (aa.sContrato=ao.sContrato and aa.sIdConvenio=ao.sidconvenio ' +
                               'and aa.swbs=ao.swbscontrato and aa.sNumeroActividad=ao.sNumeroActividad) ' +
                               'where b.sContrato = :Contrato And aa.sIdConvenio =:Convenio And b.dIdFecha >=:FechaInicial ' +
                               auxcadena + ' and ao.snumeroorden=:orden ' +
                               'And b.dIdFecha <=:FechaFinal And b.lConceptoEjecutado=:Concepto And b.sPaquete="" ' +
                               'Group by b.sNumeroActividad Order by  b.sNumeroActividad And b.dCantidad > 0') ;

  connection.QryBusca2.Params.ParamByName('Contrato').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Contrato').Value         := global_contrato;
  connection.QryBusca2.Params.ParamByName('Convenio').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Convenio').Value         := global_convenio;
  connection.QryBusca2.Params.ParamByName('FechaInicial').DataType  := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaInicial').Value     := tdFechaInicial.DateTime;
  connection.QryBusca2.Params.ParamByName('Concepto').DataType     := ftString;
  connection.QryBusca2.Params.ParamByName('Concepto').Value    := 'No'  ;
  connection.QryBusca2.Params.ParamByName('FechaFinal').DataType    := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaFinal').Value       := tdFechaFinal.DateTime;
  Connection.QryBusca2.ParamByName('orden').AsString:=tsnumeroorden.KeyValue;
  Connection.QryBusca2.Open ;

  if connection.QryBusca2.RecordCount > 0 then
  begin
    while not connection.QryBusca2.Eof do
    begin
      Ren2 := Ren2 + 3 ;
      Hoja.Cells[Ren2,1].Select;
      Excel.Selection.Value := Connection.QryBusca2.FieldValues['sNumeroActividad'];
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Size := 11;
      Excel.Selection.Font.Bold := False;
      Excel.Selection.Font.Name := 'Calibri';

      Hoja.Cells[Ren2,2].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['mDescripcion'];
      Excel.Selection.HorizontalAlignment   := xlLeft;
      Excel.selection.WrapText := True    ;
      Excel.selection.Orientation := 0    ;
      Excel.selection.AddIndent := False  ;
      Excel.selection.IndentLevel := 0    ;
      Excel.selection.ShrinkToFit := False  ;
      Excel.selection.MergeCells := False  ;

      Hoja.Cells[Ren2,3].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['sMedida'];
      Excel.Selection.HorizontalAlignment   := xlLeft;
      Excel.Selection.VerticalAlignment   := xlCenter;

      Hoja.Cells[Ren2,4].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['dCantidadAnexo'];
      Excel.Selection.HorizontalAlignment   := xlLeft;
      Excel.selection.WrapText := True    ;
      Excel.selection.Orientation := 0    ;
      Excel.selection.AddIndent := False  ;
      Excel.selection.IndentLevel := 0    ;
      Excel.selection.ShrinkToFit := False  ;
      Excel.selection.MergeCells := False  ;

      qryTitulo1.Active := False ;
      qryTitulo1.SQL.Clear;
      qryTitulo1.SQL.Add('select sEspecificacion, mDescripcion, sWbsAnterior, sWbs, sNumeroActividad, sWbsAnterior from actividadesxanexo where sContrato =:Contrato ' +
                   'and sIdConvenio=:Convenio And sWbs=:Wbs and sTipoActividad="Paquete"') ;
      qryTitulo1.Params.ParamByName('Contrato').DataType      := ftString;
      qryTitulo1.Params.ParamByName('Contrato').Value         := global_contrato;
      qryTitulo1.Params.ParamByName('Convenio').DataType      := ftString;
      qryTitulo1.Params.ParamByName('Convenio').Value         := global_convenio;
      qryTitulo1.Params.ParamByName('Wbs').DataType           := ftString;
      qryTitulo1.Params.ParamByName('Wbs').Value              := connection.QryBusca2.FieldValues['sWbsAnterior'];
      qryTitulo1.Open ;
      if qryTitulo1.RecordCount > 0 then
      begin
        Hoja.Cells[Ren2-1,2].Select;
        Excel.Selection.Value := qryTitulo1.FieldValues['mDescripcion'];
        Excel.Selection.HorizontalAlignment   := xlLeft;
        Excel.selection.WrapText := True    ;
        Excel.selection.Orientation := 0    ;
        Excel.selection.AddIndent := False  ;
        Excel.selection.IndentLevel := 0    ;
        Excel.selection.ShrinkToFit := False  ;
        Excel.selection.MergeCells := False  ;

        Hoja.Cells[Ren2-1,1].Select;
        Excel.Selection.Value := qryTitulo1.FieldValues['sEspecificacion'];
        Excel.Selection.HorizontalAlignment   := xlLeft;
        Excel.selection.WrapText := True    ;
        Excel.selection.Orientation := 0    ;
        Excel.selection.AddIndent := False  ;
        Excel.selection.IndentLevel := 0    ;
        Excel.selection.ShrinkToFit := False  ;
        Excel.selection.MergeCells := False  ;
      end;

      qryTitulo2.Active := False ;
      qryTitulo2.SQL.Clear;
      qryTitulo2.SQL.Add('select sEspecificacion, mDescripcion, sWbs, sNumeroActividad, sWbsAnterior from actividadesxanexo where sContrato =:Contrato ' +
                   'and sIdConvenio=:Convenio And sWbs=:Wbs and sTipoActividad="Paquete"') ;
      qryTitulo2.Params.ParamByName('Contrato').DataType      := ftString;
      qryTitulo2.Params.ParamByName('Contrato').Value         := global_contrato;
      qryTitulo2.Params.ParamByName('Convenio').DataType      := ftString;
      qryTitulo2.Params.ParamByName('Convenio').Value         := global_convenio;
      qryTitulo2.Params.ParamByName('Wbs').DataType           := ftString;
      qryTitulo2.Params.ParamByName('Wbs').Value              := qryTitulo1.FieldValues['sWbsAnterior'];
      qryTitulo2.Open ;
      if qryTitulo2.RecordCount > 0 then
      begin
        Hoja.Cells[Ren2-2,2].Select;
        Excel.Selection.Value := qryTitulo2.FieldValues['mDescripcion'];
        Excel.Selection.HorizontalAlignment   := xlLeft;
        Excel.selection.WrapText := True    ;
        Excel.selection.Orientation := 0    ;
        Excel.selection.AddIndent := False  ;
        Excel.selection.IndentLevel := 0    ;
        Excel.selection.ShrinkToFit := False  ;
        Excel.selection.MergeCells := False  ;

        Hoja.Cells[Ren2-2,1].Select;
        Excel.Selection.Value := qryTitulo2.FieldValues['sEspecificacion'];
        Excel.Selection.HorizontalAlignment   := xlLeft;
        Excel.selection.WrapText := True    ;
        Excel.selection.Orientation := 0    ;
        Excel.selection.AddIndent := False  ;
        Excel.selection.IndentLevel := 0    ;
        Excel.selection.ShrinkToFit := False  ;
        Excel.selection.MergeCells := False  ;
      end;

      MiFecha := tdFechaInicial.date;
      dAvanceTotal := 0 ;
      for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
      begin
        cadena:= formatDateTime('dd/mm/yyyy',Mifecha);
        connection.QryBusca.Active := False ;
        connection.QryBusca.SQL.Clear ;
        connection.QryBusca.SQL.Add('select b.dIdFecha, b.iFase, b.sNumeroActividad, sum(b.dCantidad) as sumacantidad, ' +
                                    'b.iIdDiarioNota from bitacoradealcances b  ' +
                                    'Where b.sContrato=:Contrato and b.sPaquete="" and b.lConceptoEjecutado=:Concepto ' +
                                    'And b.dIdFecha=:Fecha and b.swbs=:wbs and b.snumeroorden=:orden and b.sNumeroActividad=:Actividad Group By b.sNumeroActividad') ;
        connection.QryBusca.Params.ParamByName('Contrato').DataType     := ftString;
        connection.QryBusca.Params.ParamByName('Contrato').Value        := global_contrato;
        connection.QryBusca.Params.ParamByName('Fecha').DataType        := ftDate;
        connection.QryBusca.Params.ParamByName('Fecha').Value           := MiFecha ;
        connection.QryBusca.Params.ParamByName('Actividad').DataType    := ftString;
        connection.QryBusca.Params.ParamByName('Actividad').Value       := connection.QryBusca2.FieldValues['sNumeroActividad'] ;
        connection.QryBusca.Params.ParamByName('wbs').AsString       := connection.QryBusca2.FieldValues['swbsFrente'] ;
        connection.QryBusca.Params.ParamByName('Concepto').DataType     := ftString;
        connection.QryBusca.Params.ParamByName('Concepto').Value        := 'Si'  ;
        connection.QryBusca.Params.ParamByName('orden').AsString       :=  tsNumeroOrden.KeyValue;
        connection.QryBusca.Open ;
        if connection.QryBusca.RecordCount > 0 then
          if connection.QryBusca.FieldValues['dIdFecha'] = StrToDate(cadena) then
          begin
            Hoja.Cells[Ren2, 4 + i].Select;
            Excel.Selection.NumberFormat := '#0.00';
            Excel.Selection.Value := connection.QryBusca.FieldValues['sumacantidad'];
            Excel.Selection.NumberFormat := '#0.0000';
            Excel.Selection.HorizontalAlignment := xlRight;
            Excel.Selection.VerticalAlignment   := xlCenter;
            dAvanceTotal := dAvanceTotal + connection.QryBusca.FieldValues['sumacantidad'];
          end;
        connection.QryBusca.Next ;
        Mifecha := IncDay(Mifecha);
      end ;
      Hoja.Cells[Ren2, 4 + i].Select;
      Excel.Selection.NumberFormat := '#0.0000';
      Excel.Selection.Value := FloatToStrF(dAvanceTotal, ffNumber, 4, 2 ) ;
      Excel.Selection.HorizontalAlignment := xlRight;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Bold := True;
      connection.QryBusca2.Next;
      Inc(Ren);
      Inc(Ren2);
    end;
  end;
  Hoja.Cells[2,2].Select;
end;


///////////////////////////////////////////////
Procedure TfrmDetalledeInstalacion.DatosVolumetriasconmateriales(ParamPartidas: string = '');
{$REGION 'Declaracion de variables'}
Var
  cadenasql, CadFecha, tmpNombre, cadena : String;
  fs: tStream;
  Alto : Extended;
  MiFechaI, MiFechaF, MiFecha,AuxFecha: tDate;
  FechaDia : tDateTime;
  dAvanceTotal: double;
  Ren, Ren1, Ren2, Ren3, nivel, fila, indice, i, total : integer;
  NombreArchivo, CadPeriodo: String;
  Imagen: TField;
  contadormateriales : Byte ;
  Altura, Margen: Extended;
  AuxCadena:string;
  slMesFin:TstringList;
  myYear, myMonth, myDay ,AuxMonth: Word;
  iCol,ColInicio,AuxColInicio:Integer;
  RenInicio:Integer;
  imgAux:TImage;
  Pic : TJpegImage;
  TempPath: array [0..MAX_PATH-1] of Char;
  FNombre1,FNombre2:TFileName;
  CNombre1,CNombre2:TFileName;
  QrConfiguracion,QrAnexo:TZReadOnlyQuery;
  rangoE:Variant;
  Salir:Boolean;
  sAuxWbsAnt,sDescAnexo:string;
  QryBuscarFirmas:TZReadOnlyQuery;
{$ENDREGION}
Begin
  {$REGION 'Crear e Inicializar'}
  Ren  := 10;
  Ren1 := 10 ;
  Ren2 := 11 ;
  Ren3 := 12;
  ColInicio:=7;
  imgAux:=TImage.Create(nil);
  sLMesFin:=TStringList.Create;
  AuxMonth:=0;
  AuxCadena:='';
  QrAnexo:=TZReadOnlyQuery.Create(nil);
  QrAnexo.Connection:=connection.zConnection;

  for I := 1 to NumItems(ParamPartidas,',') do
  begin
    if AuxCadena='' then
      AuxCadena:=' and (aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
    else
       AuxCadena:=AuxCadena + ' or aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
  end;

  if AuxCadena<>'' then
    AuxCadena:=AuxCadena + ') ';
  {$ENDREGION}

  {$REGION 'Encabezado de Reporte'}
  // Realizar los ajustes visuales y de formato de hoja
  Excel.ActiveWindow.Zoom :=85;
  Excel.Columns['A:A'].ColumnWidth := 16.43;
  Excel.Columns['B:B'].ColumnWidth := 13;
  Excel.Columns['C:C'].ColumnWidth := 27;
  Excel.Columns['D:D'].ColumnWidth := 30;
  Excel.Columns['E:E'].ColumnWidth := 10;
  Excel.Columns['F:F'].ColumnWidth := 11  ;

  Excel.Rows['1:1'].RowHeight := 45;
  Excel.Rows['2:2'].RowHeight := 26.25;
  Excel.Rows['4:5'].RowHeight := 30;
  Excel.Rows['6:6'].RowHeight := 36;
  Excel.Rows['7:7'].RowHeight := 31.5;
  Excel.Rows['8:8'].RowHeight := 13.5;
  Excel.Rows['9:9'].RowHeight := 14.25;
 // Excel.Rows['10:10'].RowHeight := 32.62;

  Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren+1)].RowHeight :=27;
  Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren+1)].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Wraptext:=True;
  Excel.Selection.Value := 'PARTIDA'+#13+#10+ 'ANEXO C';


  Hoja.Range['B'+inttostr(ren)+':D' +inttostr(ren+1)].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Value := 'DESCRIPCION';


  Hoja.Range['E'+inttostr(ren)+':E'+inttostr(ren+1)].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Value := 'UNIDAD';


  Hoja.Range['F'+inttostr(ren)+':F'+inttostr(ren+1)].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Value := 'CANTIDAD';

  MiFecha := tdFechaInicial.date;
  MiFechaI := tdFechaInicial.date;
  MiFechaF := tdFechaFinal.Date;
  iCol:= ColInicio;
  AuxColInicio:=iCol;
  Inc(Ren);
  for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
  begin

   // Selection.NumberFormat = "General"
    DecodeDate(MiFecha, myYear, myMonth, myDay);
    if AuxMonth<>myMonth then
    begin
      if AuxMonth<>0 then
      begin
        if  MiFecha=MiFechaF then
          if (iCol-ColInicio)< 14 then
          begin
            while (iCol-ColInicio)< 14 do
            begin
              Excel.Columns[ColumnaNombre(icol)+ ':' + ColumnaNombre(icol)].ColumnWidth :=10;
              Hoja.Cells[Ren,iCol].Select;
              //Excel.Selection.NumberFormat := 'General';
              AuxFecha:=MiFecha;
              Excel.Selection.Value :='';
              Excel.Selection.HorizontalAlignment := xlCenter;
              Excel.Selection.VerticalAlignment := xlCenter;
              Excel.Selection.Font.Name := 'Arial';
              Excel.Selection.Font.Size :=TamFont;
              //Excel.Selection.Font.Color := clWhite;
              Excel.Selection.Font.Bold := True;
              Inc(iCol);
            end;
          end;


        Hoja.Range[ColumnaNombre(AuxColInicio) + IntToStr(Ren-1) + ':' + ColumnaNombre(iCol-1) + IntToStr(Ren-1)].Select;
        Excel.Selection.MergeCells := True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := True;
        Excel.Selection.Value :=Uppercase(FormatDateTime('mmmm',AuxFecha)) + ' DE ' + inttostr(myYear);
        sLMesFin.Add(FormatDateTime('mmmm',AuxFecha)+'='+inttostr(iCol-1));
        AuxColInicio:=iCol;
      end;
      AuxMonth:=myMonth;
    end;
    Excel.Columns[ColumnaNombre(icol)+ ':' + ColumnaNombre(icol)].ColumnWidth :=10;
    Hoja.Cells[Ren,iCol].Select;
    //Excel.Selection.NumberFormat := 'General';
    AuxFecha:=MiFecha;
    Excel.Selection.Value :=myDay;
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.VerticalAlignment := xlCenter;
    Excel.Selection.Font.Name := 'Arial';
    Excel.Selection.Font.Size :=TamFont;
    //Excel.Selection.Font.Color := clWhite;
    Excel.Selection.Font.Bold := True;
    //Excel.Selection.Interior.ColorIndex := 41;
    MiFecha := IncDay(MiFecha);
    Inc(iCol);
  end;

  if AuxMonth<>0 then
  begin
    if (iCol-ColInicio)< 14 then
    begin
      while (iCol-ColInicio)< 14 do
      begin
        Excel.Columns[ColumnaNombre(icol)+ ':' + ColumnaNombre(icol)].ColumnWidth :=10;
        Hoja.Cells[Ren,iCol].Select;
        //Excel.Selection.NumberFormat := 'General';
        AuxFecha:=MiFecha;
        Excel.Selection.Value :='';
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        //Excel.Selection.Font.Color := clWhite;
        Excel.Selection.Font.Bold := True;
        Inc(iCol);
      end;
    end;

    Hoja.Range[ColumnaNombre(AuxColInicio) + IntToStr(Ren-1) + ':' + ColumnaNombre(iCol-1) + IntToStr(Ren-1)].Select;
    Excel.Selection.MergeCells := True;
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.VerticalAlignment := xlCenter;
    Excel.Selection.Font.Name := 'Arial';
    Excel.Selection.Font.Size :=TamFont;
    Excel.Selection.Font.Bold := True;
    Excel.Selection.Value :=Uppercase(FormatDateTime('mmmm',AuxFecha)) + ' DE ' + inttostr(myYear);
    sLMesFin.Add(FormatDateTime('mmmm',AuxFecha)+'='+inttostr(iCol-1));
    AuxColInicio:=iCol;
  end;
    //total := i;
  Excel.Columns[ColumnaNombre(icol)+ ':' + ColumnaNombre(icol)].ColumnWidth :=10.30;
  Hoja.Range[ColumnaNombre(iCol) + IntToStr(Ren-1) + ':' + ColumnaNombre(iCol) + IntToStr(Ren)].Select;
   Excel.Selection.MergeCells := True;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Value := 'ANTERIOR';
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;

  Inc(iCol);
  Excel.Columns[ColumnaNombre(icol)+ ':' + ColumnaNombre(icol)].ColumnWidth :=11.71;
  Hoja.Range[ColumnaNombre(iCol) + IntToStr(Ren-1) + ':' + ColumnaNombre(iCol) + IntToStr(Ren)].Select;
   Excel.Selection.MergeCells := True;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=True;
  Excel.Selection.Value :='ESTE' + #13 + #10 +'PERIODO';
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;

  Inc(iCol);
  Excel.Columns[ColumnaNombre(icol)+ ':' + ColumnaNombre(icol+1)].ColumnWidth :=6.86;
  Hoja.Range[ColumnaNombre(iCol) + IntToStr(Ren-1) + ':' + ColumnaNombre(iCol+1) + IntToStr(Ren)].Select;
  Excel.Selection.MergeCells := True;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=True;
  Excel.Selection.Value :='ACUMULADO';
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := True;
  Inc(iCol);
  Inc(Ren);
  RenInicio:=Ren;

  //Excel.Selection.Interior.ColorIndex := 16;
      // Colocar los encabezados de la plantilla...
  {*****************************************************************************
     ** Colocar el logotipo de la empresa                                        }
  GetTempPath(SizeOf(TempPath), TempPath);
  FNombre1:=TempPath +'imgtempAby'+formatdatetime('dddddd hhnnss',now)+'.jpg';

  fs := Connection.configuracion.CreateBlobStream(Connection.configuracion.FieldByName('bImagen'), bmRead) ;
  If fs.Size > 1 Then
  Begin
      try
          Pic:=TJpegImage.Create;
          try
             Pic.LoadFromStream(fs);
             imgAux.Picture.Graphic := Pic;
          finally
             Pic.Free;
          end;
      finally
          fs.Free;
      End;
    imgAux.Picture.SaveToFile(FNombre1);
  End;

  if FileExists(FNombre1) then
  begin
    Hoja.Cells[1,1].Select;
    Excel.ActiveSheet.Pictures.Insert(FNombre1).Select;
    // Determinar el tamaño real de la imagen
    Altura := Excel.Rows[1].Height * 0.7;
    Margen := (Excel.Rows[1].Height - Altura) / 2 ;
    Excel.Selection.ShapeRange.Left := Excel.Columns['A:A'].Width ;
    Excel.Selection.ShapeRange.Top := Margen;
  end;

  GetTempPath(SizeOf(TempPath), TempPath);
  FNombre2:=TempPath +'imgtempAby2'+formatdatetime('dddddd hhnnss',now)+'.jpg';

  fs :=Connection.contrato.CreateBlobStream(Connection.contrato.FieldByName('bImagen'), bmRead) ;
  If fs.Size > 1 Then
  Begin
    try
      Pic:=TJpegImage.Create;
      try
        Pic.LoadFromStream(fs);
        imgAux.Picture.Graphic := Pic;
      finally
        Pic.Free;
      end;
    finally
      fs.Free;
    End;
    imgAux.Picture.SaveToFile(FNombre2);
  End ;

  if FileExists(FNombre2) then
  begin
    // Agregar Imagen Cliente a la hoja de excel
    Hoja.Cells[1,1].Select;
    Excel.ActiveSheet.Pictures.Insert(FNombre2).Select;
    Excel.Selection.Cut;
    Hoja.Cells[2,iCol].Select;
    Hoja.Paste;

    // Determinar el tamaño real de la imagen
    Altura := (Excel.Rows[1].Height + Excel.Rows[2].Height ) ;   // * 0.7;
   // Margen := 0;  //(Excel.Rows[1].Height + Excel.Rows[2].Height + Excel.Rows[3].Height + Excel.Rows[4].Height + Excel.Rows[5].Height - Altura) / 2;
    Excel.Selection.ShapeRange.ScaleWidth((Altura) / (Excel.Selection.ShapeRange.Height), msoFalse, msoScaleFromTopLeft); //msoScaleFromBottomRight);
    Excel.Selection.ShapeRange.ScaleHeight(Altura/ Excel.Selection.ShapeRange.Height, msoFalse, msoScaleFromTopLeft);

    Excel.Selection.ShapeRange.IncrementLeft(Excel.Selection.ShapeRange.Width * -1);
    //Excel.Selection.ShapeRange.Left := -50;  //Margen;    //Excel.Columns['A:A'].Width + Margen;
    Excel.Selection.ShapeRange.Top :=Margen;   //Margen;

  end;


  {** Termina Colocar el logotipo de la empresa
   *****************************************************************************}
  Ren:=2;
  Hoja.Range['A2:'+Columnanombre(iCol)+'2'].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+6;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.MergeCells := True;
  Excel.Selection.Value := 'V O L U M E N   D E    O B R A   E J E C U T A D A';

  QrConfiguracion := TZReadOnlyQuery.Create(Nil);
  QrConfiguracion.Connection := Connection.zConnection;
  QrConfiguracion.SQL.Text:='select c.TamFuente,c.iFirmasReportes,ot.bcostafuera,c.sMostrarAvances,c.iFirmas, c.sOrdenPerEq, c.sTipoPartida, c.sImprimePEP, ' +
      ' (select sContrato from contratos where sTipoObra = "BARCO" ) as sContratoBarco, ' +
      ' (select mDescripcion from contratos where sTipoObra = "BARCO" ) as mDescripcionBarco, ' +
      ' (select sLocalizacion from condicionesclimatologicas where sContrato =:ContratoBarco and dIdFecha =:fecha group by sContrato Order by sHorario DESC ) as Localizacion, '+
      'c.sPartidaBarco, c.sClaveSeguridad, c.cStatusProceso, c.sOrdenExtraordinaria, c.lLicencia, c.sReportesCIA, c.sLeyenda1, c.sLeyenda2, c.sLeyenda3,' +
      'ot.bAvanceFrente, ot.bAvanceContrato, ot.bComentarios, ot.bPermisos, ot.lMostrarAvanceProgramado, ot.lImprimePersonalTM, ot.lPersonalxPartida, ot.lEquipoxPartida, ' +
      'c.bImagen, c.sContrato, c.sNombre, c2.sCodigo, c2.sProrrateoBarco, c.sPiePagina, c.sEmail, c.sWeb, c.sSlogan, c.sFirmasElectronicas, c.lImprimeExtraordinario, ' +
      'c2.mDescripcion, c2.sTitulo, c2.mCliente, c2.bImagen as bImagenPEP, ot.lImprimeFases, cv.dFechaInicio, cv.dfechaFinal ' +
      'From contratos c2 INNER JOIN configuracion c ON (c.sContrato = c2.sContrato) ' +
      'inner join ordenesdetrabajo ot on (ot.sContrato = c2.sContrato and ot.sNumeroOrden =:Orden ) ' +
      'inner join convenios cv on (cv.sContrato = c2.sContrato and cv.sIdConvenio =:convenio) '+
      'Where c2.sContrato = :Contrato';
  QrConfiguracion.ParamByName('contrato').AsString:= global_Contrato;
  QrConfiguracion.ParamByName('convenio').AsString:= Global_Convenio;
  QrConfiguracion.ParamByName('Orden').AsString:= tsNumeroOrden.KeyValue;
  QrConfiguracion.ParamByName('ContratoBarco').AsString:= global_Contrato_barco;
  QrConfiguracion.ParamByName('Fecha').AsDate:= tdFechaInicial.Date;
  QrConfiguracion.Open;

  inc(ren,2);
  Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=false;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment   := xlBottom;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := false;
  Excel.Selection.MergeCells := True;
  Excel.Selection.Value := 'CONTRATISTA:';

  Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=false;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment   := xlBottom;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := true;
  Excel.Selection.MergeCells := True;
  if QrConfiguracion.RecordCount=1 then
    Excel.Selection.Value :=QrConfiguracion.FieldByName('sNombre').AsString;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := 0;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;

  Hoja.Range['P'+inttostr(ren)+':R'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=True;
  Excel.Selection.HorizontalAlignment := xlRight;
  Excel.Selection.VerticalAlignment   := xlBottom;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := false;
  Excel.Selection.MergeCells := True;
  Excel.Selection.Value := 'PLATAFORMA:';

  Hoja.Range['T'+inttostr(ren)+':W'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=false;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment   := xlBottom;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := true;
  Excel.Selection.MergeCells := True;
  if QrConfiguracion.RecordCount=1 then
    Excel.Selection.Value :=QrConfiguracion.FieldByName('Localizacion').AsString;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := 0;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;

  inc(Ren);
  Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=false;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment   := xlBottom;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := false;
  Excel.Selection.MergeCells := True;
  Excel.Selection.Value := 'CONTRATO No.:';

  Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment   := xlBottom;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := true;
  Excel.Selection.MergeCells := True;
  if QrConfiguracion.RecordCount=1 then
    Excel.Selection.Value :=trim(QrConfiguracion.FieldByName('sProrrateoBarco').AsString);
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := 0;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;

  inc(Ren);
  Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment   := xlCenter;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := false;
  Excel.Selection.MergeCells := True;
  Excel.Selection.Value := 'OBRA:';

  Hoja.Range['B'+inttostr(ren)+':P'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment   := xlTop;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := true;
  Excel.Selection.MergeCells := True;
  if QrConfiguracion.RecordCount=1 then
    Excel.Selection.Value :=Trim(QrConfiguracion.FieldByName('mDescripcionBarco').AsString);
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := 0;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;

  inc(Ren);
  Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment   := xlBottom;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := false;
  Excel.Selection.MergeCells := True;
  Excel.Selection.Value :='PERIODO' + #13 + #10 + 'EJECUTADO :';

  Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=True;
  Excel.Selection.HorizontalAlignment := xlLeft;
  Excel.Selection.VerticalAlignment   := xlBottom;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont+1;
  Excel.Selection.Font.Bold := true;
  Excel.Selection.MergeCells := True;
  if FormatDateTime('mmmm',TdFechaInicial.Date)=FormatDateTime('mmmm',TdFechaFinal.Date) then
    Excel.Selection.Value :='DEL ' + FormatDateTime('dd',TdFechaInicial.Date) + ' AL ' + FormatDateTime('dd',TdFechaFinal.Date)+
                            ' DE ' + uppercase(FormatDateTime('mmmm',TdFechaInicial.Date))+' DEL ' + FormatDateTime('yyyy',TdFechaInicial.Date)
  else
    Excel.Selection.Value :=  'DEL ' + FormatDateTime('dd',TdFechaInicial.Date) + ' DE ' + uppercase(FormatDateTime('mmmm',TdFechaInicial.Date))+
                              ' AL ' + FormatDateTime('dd',TdFechaFinal.Date)+
                              ' DE ' + uppercase(FormatDateTime('mmmm',TdFechaFinal.Date))+' DEL ' + FormatDateTime('yyyy',TdFechaInicial.Date);


  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := 0;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;

  Hoja.Range['A4:'+Columnanombre(iCol)+'8'].Select;
  Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeLeft].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeLeft].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeTop].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeTop].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeTop].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeRight].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeRight].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeRight].Weight := xlThick;

  Ren:=RenInicio-2;
  Hoja.Range['A'+inttostr(Ren) + ':F'+inttostr(Ren+1)].Select;
  Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeLeft].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeLeft].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeTop].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeTop].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeTop].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeRight].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeRight].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeRight].Weight := xlThick;
  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlInsideHorizontal].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideHorizontal].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideHorizontal].Weight := xlThick;
  Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideVertical].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideVertical].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideVertical].Weight := xlThin;
  Excel.Selection.Interior.Pattern := xlSolid;
  Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
  Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
  Excel.Selection.Interior.TintAndShade := -0.149998474074526;
  Excel.Selection.Interior.PatternTintAndShade := 0;


  Hoja.Range['G'+inttostr(Ren) + ':'+Columnanombre(iCol-5)+inttostr(Ren+1)].Select;
  Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeLeft].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeLeft].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeTop].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeTop].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeTop].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeRight].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeRight].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeRight].Weight := xlThick;
  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlInsideHorizontal].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideHorizontal].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideHorizontal].Weight := xlThick;
  Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideVertical].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideVertical].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideVertical].Weight := xlThin;
    Excel.Selection.Interior.Pattern := xlSolid;
  Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
  Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
  Excel.Selection.Interior.TintAndShade := -0.149998474074526;
  Excel.Selection.Interior.PatternTintAndShade := 0;

  Hoja.Range[Columnanombre(iCol-3)+inttostr(Ren) + ':'+Columnanombre(iCol)+inttostr(Ren+1)].Select;
  Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeLeft].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeLeft].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeTop].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeTop].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeTop].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeRight].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeRight].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeRight].Weight := xlThick;
  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlInsideHorizontal].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideHorizontal].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideHorizontal].Weight := xlThick;
  Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideVertical].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideVertical].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideVertical].Weight := xlThin;
    Excel.Selection.Interior.Pattern := xlSolid;
  Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
  Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
  Excel.Selection.Interior.TintAndShade := -0.149998474074526;
  Excel.Selection.Interior.PatternTintAndShade := 0;





  {$ENDREGION}

  {$REGION 'Counsulta de Informacion'}
  contadormateriales := 0 ;
  connection.QryBusca2.Active := False ;
  connection.QryBusca2.Filtered := False;
  connection.QryBusca2.SQL.Clear ;
  connection.QryBusca2.SQL.Add('select b.iFase,ao.swbs as swbsFrente, b.sNumeroActividad, aa.mDescripcion, aa.sMedida, aa.dCantidadAnexo, b.dCantidad, b.dAvance, '+
                               'b.iIdDiarioNota, aa.sWbsAnterior from bitacoradealcances b ' +
                               'inner join actividadesxorden ao on(ao.scontrato=b.scontrato and ao.snumeroorden=b.snumeroorden and ' +
                               'ao.swbs=b.swbs and ao.snumeroactividad=b.snumeroactividad) ' +
                               'Inner Join actividadesxanexo  aa On (aa.sContrato and ao.sContrato and aa.sIdConvenio=ao.sidconvenio ' +
                               'And aa.swbs=ao.swbscontrato and aa.sNumeroActividad=ao.sNumeroActividad) ' +
                               'where b.sContrato = :Contrato And ao.sIdConvenio =:Convenio And b.dIdFecha >=:FechaInicial ' +
                               'And b.dIdFecha <=:FechaFinal And b.lConceptoEjecutado=:Concepto And b.sPaquete="" ' +
                                auxcadena +' and ao.snumeroorden=:orden ' +
                               'And b.dCantidad > 0 Group by b.sNumeroActividad Order by  b.sNumeroActividad ') ;

  connection.QryBusca2.Params.ParamByName('Contrato').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Contrato').Value         := global_contrato;
  connection.QryBusca2.Params.ParamByName('Convenio').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Convenio').Value         := global_convenio;
  connection.QryBusca2.Params.ParamByName('FechaInicial').DataType  := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaInicial').Value     := tdFechaInicial.DateTime;
  connection.QryBusca2.Params.ParamByName('Concepto').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Concepto').Value         := 'Si'  ;
  connection.QryBusca2.Params.ParamByName('FechaFinal').DataType    := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaFinal').Value       := tdFechaFinal.DateTime;
  connection.QryBusca2.ParamByName('orden').AsString:=tsnumeroorden.KeyValue;
  Connection.QryBusca2.Open ;
  {$ENDREGION}

  Ren:=RenInicio;
  if connection.QryBusca2.RecordCount > 0 then
  begin
    while not connection.QryBusca2.Eof do
    begin
      {$REGION 'Paquetes de la Actividad'}
      qryTitulo1.Active := False ;
      qryTitulo1.SQL.Clear;
      qryTitulo1.SQL.Add('select sEspecificacion, mDescripcion, sWbsAnterior, sWbs, sNumeroActividad, sWbsAnterior from actividadesxanexo where sContrato =:Contrato ' +
                         'and sIdConvenio=:Convenio And sWbs=:Wbs and sTipoActividad="Paquete"') ;
      qryTitulo1.Params.ParamByName('Contrato').DataType      := ftString;
      qryTitulo1.Params.ParamByName('Contrato').Value         := global_contrato;
      qryTitulo1.Params.ParamByName('Convenio').DataType      := ftString;
      qryTitulo1.Params.ParamByName('Convenio').Value         := global_convenio;
      qryTitulo1.Params.ParamByName('Wbs').DataType           := ftString;
      qryTitulo1.Params.ParamByName('Wbs').Value              := connection.QryBusca2.FieldValues['sWbsAnterior'];
      qryTitulo1.Open ;
      if qryTitulo1.RecordCount > 0 then
      begin
        qryTitulo2.Active := False ;
        qryTitulo2.SQL.Clear;
        qryTitulo2.SQL.Add('select sEspecificacion, mDescripcion, sWbs, sNumeroActividad, sWbsAnterior from actividadesxanexo where sContrato =:Contrato ' +
                           'and sIdConvenio=:Convenio And sWbs=:Wbs and sTipoActividad="Paquete"') ;
        qryTitulo2.Params.ParamByName('Contrato').DataType      := ftString;
        qryTitulo2.Params.ParamByName('Contrato').Value         := global_contrato;
        qryTitulo2.Params.ParamByName('Convenio').DataType      := ftString;
        qryTitulo2.Params.ParamByName('Convenio').Value         := global_convenio;
        qryTitulo2.Params.ParamByName('Wbs').DataType           := ftString;
        qryTitulo2.Params.ParamByName('Wbs').Value              := qryTitulo1.FieldValues['sWbsAnterior'];
        qryTitulo2.Open ;
        if qryTitulo2.RecordCount > 0 then
        begin
          Salir:=False;
          sAuxWbsAnt:=qryTitulo2.FieldByName('swbsAnterior').AsString;
          sDescAnexo:='';
          while not Salir do
          begin
            QrAnexo.Active:=False;
            QrAnexo.SQL.Text:='select iNivel,sEspecificacion, mDescripcion, sWbs, sNumeroActividad, sWbsAnterior from actividadesxanexo where sContrato =:Contrato ' +
                           'and sIdConvenio=:Convenio And sWbs=:Wbs and sTipoActividad="Paquete"';
            QrAnexo.ParamByName('Contrato').AsString:=global_contrato;
            QrAnexo.ParamByName('Contrato').AsString:=global_convenio;
            QrAnexo.ParamByName('Contrato').AsString:=sAuxWbsAnt;
            QrAnexo.Open;
            if QrAnexo.RecordCount=0 then
              Salir:=True
            else
              if QrAnexo.FieldByName('iNivel').AsInteger=1 then
              begin
                Salir:=True;
                sDescAnexo:=QrAnexo.FieldByName('mDescripcion').AsString;
              end
              else
                sAuxWbsAnt:=QrAnexo.FieldByName('sWbsAnterior').AsString;
          end;

          if sDescAnexo<>'' then
          begin
            Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
            Excel.Selection.NumberFormat :='@';
            Excel.Selection.Wraptext:=True;
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;
            Excel.Selection.Font.Name := 'Arial';
            Excel.Selection.Font.Size :=TamFont;
            Excel.Selection.Font.Bold := true;
            Excel.Selection.MergeCells := True;
            Excel.Selection.Value := sDescAnexo;
            rangoE:=Hoja.Range['B' + IntToStr(Ren) + ':D' + IntToStr(Ren)];
            AjustarTexto(rangoE,3,10);

            Hoja.Range[ColumnaNombre(iCol-1)+inttostr(Ren)+':'+ColumnaNombre(ICol)+inttostr(Ren)].Select;
            Excel.Selection.NumberFormat :='#,##0.0000';
            Excel.Selection.Wraptext:=True;
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;
            Excel.Selection.Font.Name := 'Arial';
            Excel.Selection.Font.Size :=TamFont;
            Excel.Selection.Font.Bold := false;
            Excel.Selection.MergeCells := True;
            Excel.Selection.Value :='';




            Inc(ren);
          end;

          Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
          Excel.Selection.NumberFormat :='@';
          Excel.Selection.Wraptext:=True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Font.Size :=TamFont;
          Excel.Selection.Font.Bold := true;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Value := qryTitulo2.FieldValues['mDescripcion'];
          rangoE:=Hoja.Range['B' + IntToStr(Ren) + ':D' + IntToStr(Ren)];
          AjustarTexto(rangoE,3,10);

          Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren)].Select;
          Excel.Selection.NumberFormat :='@';
          Excel.Selection.Wraptext:=True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Font.Size :=TamFont;
          Excel.Selection.Font.Bold := true;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Value := qryTitulo2.FieldValues['sEspecificacion'];

          Hoja.Range[ColumnaNombre(iCol-1)+inttostr(Ren)+':'+ColumnaNombre(ICol)+inttostr(Ren)].Select;
          Excel.Selection.NumberFormat :='#,##0.0000';
          Excel.Selection.Wraptext:=True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Font.Size :=TamFont;
          Excel.Selection.Font.Bold := false;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Value :='';

          Inc(ren);
        end;
        
        Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
        Excel.Selection.NumberFormat :='@';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := true;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value := qryTitulo1.FieldValues['mDescripcion'];
        rangoE:=Hoja.Range['B' + IntToStr(Ren) + ':D' + IntToStr(Ren)];
        AjustarTexto(rangoE,3,10);

        Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren)].Select;
        Excel.Selection.NumberFormat :='@';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := true;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value := qryTitulo1.FieldValues['sEspecificacion'];

        Hoja.Range[ColumnaNombre(iCol-1)+inttostr(Ren)+':'+ColumnaNombre(ICol)+inttostr(Ren)].Select;
        Excel.Selection.NumberFormat :='#,##0.0000';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := false;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value :='';
        Inc(ren);

      end;
    
      {$ENDREGION}

      {$REGION 'Materiales'}
      Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren)].Select;
      Excel.Selection.NumberFormat :='@';
      Excel.Selection.Wraptext:=True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.Font.Size :=TamFont;
      Excel.Selection.Font.Bold := true;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Value := Connection.QryBusca2.FieldValues['sNumeroActividad'];
      Excel.Selection.Interior.Pattern := xlSolid;
      Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
      Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
      Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
      Excel.Selection.Interior.PatternTintAndShade := 0;

      Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['mDescripcion'];
      Excel.Selection.NumberFormat :='@';
      Excel.Selection.Wraptext:=True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.Font.Size :=TamFont;
      Excel.Selection.Font.Bold := true;
      Excel.Selection.MergeCells := True;
      rangoE:=Hoja.Range['B' + IntToStr(Ren) + ':D' + IntToStr(Ren)];
      AjustarTexto(rangoE,3,20);
      Excel.Selection.Interior.Pattern := xlSolid;
      Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
      Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
      Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
      Excel.Selection.Interior.PatternTintAndShade := 0;

      Hoja.Range['E'+inttostr(ren)+':E'+inttostr(ren)].Select;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['sMedida'];
      Excel.Selection.NumberFormat :='@';
      Excel.Selection.Wraptext:=True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.Font.Size :=TamFont;
      Excel.Selection.Font.Bold := true;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Interior.Pattern := xlSolid;
      Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
      Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
      Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
      Excel.Selection.Interior.PatternTintAndShade := 0;

      Hoja.Range['F'+inttostr(ren)+':F'+inttostr(ren)].Select;
      Excel.Selection.NumberFormat :='#,##0.0000';
      Excel.Selection.Wraptext:=True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.Font.Size :=TamFont;
      Excel.Selection.Font.Bold := true;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Value := connection.QryBusca2.FieldValues['dCantidadAnexo'];
      Excel.Selection.Interior.Pattern := xlSolid;
      Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
      Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
      Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
      Excel.Selection.Interior.PatternTintAndShade := 0;

      Ren1:=Ren;
      Inc(ren);

      Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
      Excel.Selection.NumberFormat :='@';
      Excel.Selection.Wraptext:=True;
      Excel.Selection.HorizontalAlignment := xlLeft;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.Font.Size :=TamFont;
      Excel.Selection.Font.Bold := true;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Value := 'LISTADO DE MATERIALES' ;
      rangoE:=Hoja.Range['B' + IntToStr(Ren) + ':D' + IntToStr(Ren)];
      AjustarTexto(rangoE,3,10);

      Hoja.Range[ColumnaNombre(iCol-1)+inttostr(Ren)+':'+ColumnaNombre(ICol)+inttostr(Ren)].Select;
      Excel.Selection.NumberFormat :='#,##0.0000';
      Excel.Selection.Wraptext:=True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.Font.Size :=TamFont;
      Excel.Selection.Font.Bold := false;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Value :='';


      contadorMateriales := 0 ;
      MiFecha := tdFechaInicial.date;
      dAvanceTotal := 0 ;
      AuxColInicio:=ColInicio;
      for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
      begin
        //cadena:= formatDateTime('dd/mm/yyyy',Mifecha);
        connection.QryBusca.Active := False ;
        connection.QryBusca.SQL.Clear ;
        connection.QryBusca.SQL.Add('select b.dIdFecha, b.iFase, b.sNumeroActividad, sum(b.dCantidad) as sumacantidad, ' +
                    'b.iIdDiarioNota from bitacoradealcances b  ' +
                    'Where b.sContrato=:Contrato and b.sPaquete="" and b.lConceptoEjecutado=:Concepto ' +
                    'And b.dIdFecha=:Fecha and b.swbs=:wbs and b.sNumeroActividad=:Actividad Group By b.sNumeroActividad') ;
        connection.QryBusca.Params.ParamByName('Contrato').DataType     := ftString;
        connection.QryBusca.Params.ParamByName('Contrato').Value        := global_contrato;
        connection.QryBusca.Params.ParamByName('Fecha').DataType        := ftDate;
        connection.QryBusca.Params.ParamByName('Fecha').Value           := MiFecha ;
        connection.QryBusca.Params.ParamByName('Actividad').DataType    := ftString;
        connection.QryBusca.Params.ParamByName('Actividad').Value       := connection.QryBusca2.FieldValues['sNumeroActividad'] ;
        connection.QryBusca.Params.ParamByName('wbs').AsString       := connection.QryBusca2.FieldValues['swbsFrente'] ;
        connection.QryBusca.Params.ParamByName('Concepto').DataType     := ftString;
        connection.QryBusca.Params.ParamByName('Concepto').Value        := 'Si'  ;
        connection.QryBusca.Open ;
        if connection.QryBusca.RecordCount =1 then
        begin
          //Hoja.Cells[Ren,AuxColInicio].Select;
          Hoja.Range[ColumnaNombre(AuxColInicio)+inttostr(Ren1)+':'+ColumnaNombre(AuxColInicio)+inttostr(Ren1)].Select;
          Excel.Selection.NumberFormat :='#,##0.0000';
          Excel.Selection.Wraptext:=True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Font.Size :=TamFont;
          Excel.Selection.Font.Bold := true;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Value := connection.QryBusca.FieldByname('sumacantidad').AsFloat;
          dAvanceTotal := dAvanceTotal + connection.QryBusca.FieldByname('sumacantidad').AsFloat;
          Excel.Selection.Interior.Pattern := xlSolid;
          Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
          Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
          Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
          Excel.Selection.Interior.PatternTintAndShade := 0;
        end
        else
        begin
          Hoja.Range[ColumnaNombre(AuxColInicio)+inttostr(Ren1)+':'+ColumnaNombre(AuxColInicio)+inttostr(Ren1)].Select;
          Excel.Selection.NumberFormat :='#,##0.0000';
          Excel.Selection.Wraptext:=True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Font.Size :=TamFont;
          Excel.Selection.Font.Bold := true;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Value :='-';

          Excel.Selection.Interior.Pattern := xlSolid;
          Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
          Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
          Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
          Excel.Selection.Interior.PatternTintAndShade := 0;
        end;
        Mifecha := IncDay(Mifecha);
        Inc(AuxColInicio);
      end ;

      while (AuxColInicio<iCol-3) do
      begin
        Hoja.Range[ColumnaNombre(AuxColInicio)+inttostr(Ren1)+':'+ColumnaNombre(AuxColInicio)+inttostr(Ren1)].Select;
        Excel.Selection.NumberFormat :='#,##0.0000';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := true;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value :='-';

        Excel.Selection.Interior.Pattern := xlSolid;
        Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
        Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
        Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
        Excel.Selection.Interior.PatternTintAndShade := 0;
        Inc(AuxColInicio);
      end;

      
      connection.QryBusca.Active := False ;
      connection.QryBusca.SQL.Clear ;
      connection.QryBusca.SQL.Add('select b.dIdFecha, b.iFase, b.sNumeroActividad, sum(b.dCantidad) as sumacantidad, ' +
                  'b.iIdDiarioNota from bitacoradealcances b  ' +
                  'Where b.sContrato=:Contrato and b.sPaquete="" and b.lConceptoEjecutado=:Concepto ' +
                  'And b.dIdFecha<:Fecha and b.swbs=:wbs and b.sNumeroActividad=:Actividad Group By b.sNumeroActividad') ;
      connection.QryBusca.Params.ParamByName('Contrato').DataType     := ftString;
      connection.QryBusca.Params.ParamByName('Contrato').Value        := global_contrato;
      connection.QryBusca.Params.ParamByName('Fecha').DataType        := ftDate;
      connection.QryBusca.Params.ParamByName('Fecha').Value           := tdFechaInicial.date ;
      connection.QryBusca.Params.ParamByName('Actividad').DataType    := ftString;
      connection.QryBusca.Params.ParamByName('Actividad').Value       := connection.QryBusca2.FieldValues['sNumeroActividad'] ;
      connection.QryBusca.Params.ParamByName('wbs').AsString       := connection.QryBusca2.FieldValues['swbsFrente'] ;
      connection.QryBusca.Params.ParamByName('Concepto').DataType     := ftString;
      connection.QryBusca.Params.ParamByName('Concepto').Value        := 'Si'  ;
      connection.QryBusca.Open ;
      if connection.QryBusca.RecordCount =1 then
      begin
        Hoja.Range[ColumnaNombre(iCol-3)+inttostr(Ren1)+':'+ColumnaNombre(ICol-3)+inttostr(Ren1)].Select;
        Excel.Selection.NumberFormat :='#,##0.0000';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := true;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value := connection.QryBusca.FieldByname('sumacantidad').AsFloat;
                Excel.Selection.Interior.Pattern := xlSolid;
        Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
        Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
        Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
        Excel.Selection.Interior.PatternTintAndShade := 0;

        Hoja.Range[ColumnaNombre(iCol-1)+inttostr(Ren1)+':'+ColumnaNombre(ICol)+inttostr(Ren1)].Select;
        Excel.Selection.NumberFormat :='#,##0.0000';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := true;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value := connection.QryBusca.FieldByname('sumacantidad').AsFloat+dAvanceTotal;
                Excel.Selection.Interior.Pattern := xlSolid;
        Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
        Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
        Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
        Excel.Selection.Interior.PatternTintAndShade := 0;


      end
      else
      begin
        Hoja.Range[ColumnaNombre(iCol-3)+inttostr(Ren1)+':'+ColumnaNombre(ICol-3)+inttostr(Ren1)].Select;
        Excel.Selection.NumberFormat :='#,##0.0000';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := true;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value := 0;
                Excel.Selection.Interior.Pattern := xlSolid;
        Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
        Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
        Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
        Excel.Selection.Interior.PatternTintAndShade := 0;

        Hoja.Range[ColumnaNombre(iCol-1)+inttostr(Ren1)+':'+ColumnaNombre(ICol)+inttostr(Ren1)].Select;
        Excel.Selection.NumberFormat :='#,##0.0000';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := true;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value := dAvanceTotal;
                Excel.Selection.Interior.Pattern := xlSolid;
        Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
        Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
        Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
        Excel.Selection.Interior.PatternTintAndShade := 0;
      end;


      Hoja.Range[ColumnaNombre(iCol-2)+inttostr(Ren1)+':'+ColumnaNombre(ICol-2)+inttostr(Ren1)].Select;
      Excel.Selection.NumberFormat :='#,##0.0000';
      Excel.Selection.Wraptext:=True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.Font.Size :=TamFont;
      Excel.Selection.Font.Bold := true;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Value := dAvanceTotal;
              Excel.Selection.Interior.Pattern := xlSolid;
        Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
        Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
        Excel.Selection.Interior.TintAndShade := -4.99893185216834E-02;
        Excel.Selection.Interior.PatternTintAndShade := 0;


      inc(Ren);
      connection.QryBusca.Active := False ;
      connection.QryBusca.SQL.Clear ;
      connection.QryBusca.SQL.Text:='select i.sIdInsumo, i.sMedida, i.mDescripcion, sum(m.dCantidad) as sumaCantidad, b.sNumeroActividad from bitacoradeactividades b ' +
                                    'inner join bitacorademateriales m On (b.sContrato=m.sContrato and b.iIdDiario=m.iIdDiario ' +
                                    'and b.dIdFecha=m.dIdFecha) ' +
                                    'Inner Join insumos i On (i.sContrato=m.sContrato and i.sIdInsumo=m.sIdMaterial) '  +
                                    'Where b.sContrato=:Contrato And b.dIdFecha between :Fecha and :FechaF ' +
                                    'and b.swbs=:wbs and b.sNumeroActividad=:Actividad and b.snumeroorden=:orden and m.lanexoH="Si" and b.sIdTipoMovimiento="E" Group by i.sIdInsumo ';
      connection.QryBusca.Params.ParamByName('Contrato').DataType     := ftString;
      connection.QryBusca.Params.ParamByName('Contrato').Value        := global_contrato;
      connection.QryBusca.Params.ParamByName('Fecha').DataType        := ftDate;
      connection.QryBusca.Params.ParamByName('Fecha').Value           := tdFechaInicial.Date ;
      connection.QryBusca.Params.ParamByName('FechaF').AsDate           := tdFechaFinal.Date ;
      connection.QryBusca.Params.ParamByName('Actividad').DataType    := ftString;
      connection.QryBusca.Params.ParamByName('Actividad').Value       := Connection.QryBusca2.FieldValues['sNumeroActividad'] ;
      connection.QryBusca.Params.ParamByName('wbs').AsString       := connection.QryBusca2.FieldValues['swbsFrente'] ;
      connection.QryBusca.ParamByName('orden').AsString:=tsNumeroOrden.KeyValue;
      connection.QryBusca.Open ;

      while not connection.QryBusca.Eof do
      begin
        Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren)].Select;
        Excel.Selection.NumberFormat :='@';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := false;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value := Connection.QryBusca.FieldValues['sIdInsumo'];


        Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
        Excel.Selection.Value := connection.QryBusca.FieldValues['mDescripcion'];
        Excel.Selection.NumberFormat :='@';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := false;
        Excel.Selection.MergeCells := True;
        rangoE:=Hoja.Range['B' + IntToStr(Ren) + ':D' + IntToStr(Ren)];
        AjustarTexto(rangoE,3,10);




        Hoja.Range['E'+inttostr(ren)+':E'+inttostr(ren)].Select;
        Excel.Selection.Value := connection.QryBusca.FieldValues['sMedida'];
        Excel.Selection.NumberFormat :='@';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := false;
        Excel.Selection.MergeCells := True;


        MiFecha := tdFechaInicial.date;
        dAvanceTotal := 0 ;
        AuxColInicio:=ColInicio;
        for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
        begin
          qryBitacoradeMateriales.Active := False ;
          qrybitacoradeMateriales.SQL.Clear ;
          qrybitacoradeMateriales.SQL.Add('select i.sIdInsumo, i.sMedida, i.mDescripcion, sum(m.dCantidad) as sumaCantidad, b.sNumeroActividad from bitacoradeactividades b ' +
                                      'inner join bitacorademateriales m On (b.sContrato=m.sContrato and b.iIdDiario=m.iIdDiario ' +
                                      'and b.dIdFecha=m.dIdFecha) ' +
                                      'Inner Join insumos i On (i.sContrato=m.sContrato and i.sIdInsumo=m.sIdMaterial) '  +
                                      'Where b.sContrato=:Contrato and i.sIdInsumo=:Insumo And b.dIdFecha=:Fecha  ' +
                                      'and b.swbs=:wbs and b.sNumeroActividad=:Actividad and b.snumeroorden=:orden and m.lanexoH="Si" and b.sIdTipoMovimiento="E" Group by i.sIdInsumo ');
          qrybitacoradeMateriales.Params.ParamByName('Contrato').DataType     := ftString;
          qrybitacoradeMateriales.Params.ParamByName('Contrato').Value        := global_contrato;
          qrybitacoradeMateriales.Params.ParamByName('Fecha').DataType        := ftDate;
          qrybitacoradeMateriales.Params.ParamByName('Fecha').Value           := MiFecha ;
          qrybitacoradeMateriales.Params.ParamByName('Actividad').DataType    := ftString;
          qrybitacoradeMateriales.Params.ParamByName('Actividad').Value       := Connection.QryBusca2.FieldValues['sNumeroActividad'] ;
          qrybitacoradeMateriales.Params.ParamByName('wbs').AsString       := connection.QryBusca2.FieldValues['swbsFrente'] ;
          qrybitacoradeMateriales.Params.ParamByName('Insumo').AsString :=  Connection.QryBusca.FieldByName('sIdInsumo').AsString;
          qrybitacoradeMateriales.ParamByName('orden').AsString:=tsNumeroOrden.KeyValue;
          qrybitacoradeMateriales.Open ;
          if qrybitacoradeMateriales.RecordCount =1 then
          begin
            //Hoja.Cells[Ren,AuxColInicio].Select;
            Hoja.Range[ColumnaNombre(AuxColInicio)+inttostr(Ren)+':'+ColumnaNombre(AuxColInicio)+inttostr(Ren)].Select;
            Excel.Selection.NumberFormat :='#,##0.0000';
            Excel.Selection.Wraptext:=True;
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;
            Excel.Selection.Font.Name := 'Arial';
            Excel.Selection.Font.Size :=TamFont;
            Excel.Selection.Font.Bold := false;
            Excel.Selection.MergeCells := True;
            Excel.Selection.Value := qrybitacoradeMateriales.FieldByname('sumacantidad').AsFloat;

            //dAvanceTotal := dAvanceTotal + connection.QryBusca.FieldByname('sumacantidad').AsFloat;
          end
          else
          begin
             Hoja.Range[ColumnaNombre(AuxColInicio)+inttostr(Ren)+':'+ColumnaNombre(AuxColInicio)+inttostr(Ren)].Select;
            Excel.Selection.NumberFormat :='#,##0.0000';
            Excel.Selection.Wraptext:=True;
            Excel.Selection.HorizontalAlignment := xlCenter;
            Excel.Selection.VerticalAlignment   := xlCenter;
            Excel.Selection.Font.Name := 'Arial';
            Excel.Selection.Font.Size :=TamFont;
            Excel.Selection.Font.Bold := false;
            Excel.Selection.MergeCells := True;
            Excel.Selection.Value := '-';

          end;
          Mifecha := IncDay(Mifecha);
          Inc(AuxColInicio);
        end ;

        while (AuxColInicio<iCol-3) do
        begin
          Hoja.Range[ColumnaNombre(AuxColInicio)+inttostr(Ren)+':'+ColumnaNombre(AuxColInicio)+inttostr(Ren)].Select;
          Excel.Selection.NumberFormat :='#,##0.0000';
          Excel.Selection.Wraptext:=True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Font.Size :=TamFont;
          Excel.Selection.Font.Bold := true;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Value :='-';
          Inc(AuxColInicio);
        end;
        
        qryBitacoradeMateriales.Active := False ;
        qrybitacoradeMateriales.SQL.Clear ;
        qrybitacoradeMateriales.SQL.Add('select i.sIdInsumo, i.sMedida, i.mDescripcion, sum(m.dCantidad) as sumaCantidad, b.sNumeroActividad from bitacoradeactividades b ' +
                                    'inner join bitacorademateriales m On (b.sContrato=m.sContrato and b.iIdDiario=m.iIdDiario ' +
                                    'and b.dIdFecha=m.dIdFecha) ' +
                                    'Inner Join insumos i On (i.sContrato=m.sContrato and i.sIdInsumo=m.sIdMaterial) '  +
                                    'Where b.sContrato=:Contrato and i.sIdInsumo=:Insumo And b.dIdFecha<:Fecha  ' +
                                    'and b.swbs=:wbs and b.sNumeroActividad=:Actividad and b.snumeroorden=:orden and m.lanexoH="Si" and b.sIdTipoMovimiento="E" Group by i.sIdInsumo ');
        qrybitacoradeMateriales.Params.ParamByName('Contrato').DataType     := ftString;
        qrybitacoradeMateriales.Params.ParamByName('Contrato').Value        := global_contrato;
        qrybitacoradeMateriales.Params.ParamByName('Fecha').DataType        := ftDate;
        qrybitacoradeMateriales.Params.ParamByName('Fecha').Value           := tdFechaInicial.Date;
        qrybitacoradeMateriales.Params.ParamByName('Actividad').DataType    := ftString;
        qrybitacoradeMateriales.Params.ParamByName('Actividad').Value       := Connection.QryBusca2.FieldValues['sNumeroActividad'] ;
        qrybitacoradeMateriales.Params.ParamByName('wbs').AsString       := connection.QryBusca2.FieldValues['swbsFrente'] ;
        qrybitacoradeMateriales.Params.ParamByName('Insumo').AsString :=  Connection.QryBusca.FieldByName('sIdInsumo').AsString;
        qrybitacoradeMateriales.ParamByName('orden').AsString:=tsNumeroOrden.KeyValue;
        qrybitacoradeMateriales.Open ;
        if qrybitacoradeMateriales.RecordCount =1 then
        begin
          Hoja.Range[ColumnaNombre(iCol-3)+inttostr(Ren)+':'+ColumnaNombre(ICol-3)+inttostr(Ren)].Select;
          Excel.Selection.NumberFormat :='#,##0.0000';
          Excel.Selection.Wraptext:=True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Font.Size :=TamFont;
          Excel.Selection.Font.Bold := false;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Value := qrybitacoradeMateriales.FieldByname('sumacantidad').AsFloat;

          Hoja.Range[ColumnaNombre(iCol-1)+inttostr(Ren)+':'+ColumnaNombre(ICol)+inttostr(Ren)].Select;
          Excel.Selection.NumberFormat :='#,##0.0000';
          Excel.Selection.Wraptext:=True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Font.Size :=TamFont;
          Excel.Selection.Font.Bold := false;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Value := qrybitacoradeMateriales.FieldByname('sumacantidad').AsFloat+Connection.QryBusca.FieldByname('sumacantidad').AsFloat;
        end
        else
        begin
          Hoja.Range[ColumnaNombre(iCol-1)+inttostr(Ren)+':'+ColumnaNombre(ICol)+inttostr(Ren)].Select;
          Excel.Selection.NumberFormat :='#,##0.0000';
          Excel.Selection.Wraptext:=True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Name := 'Arial';
          Excel.Selection.Font.Size :=TamFont;
          Excel.Selection.Font.Bold := false;
          Excel.Selection.MergeCells := True;
          Excel.Selection.Value := Connection.QryBusca.FieldByname('sumacantidad').AsFloat;
        end;

        Hoja.Range[ColumnaNombre(iCol-2)+inttostr(Ren)+':'+ColumnaNombre(ICol-2)+inttostr(Ren)].Select;
        Excel.Selection.NumberFormat :='#,##0.0000';
        Excel.Selection.Wraptext:=True;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Name := 'Arial';
        Excel.Selection.Font.Size :=TamFont;
        Excel.Selection.Font.Bold := false;
        Excel.Selection.MergeCells := True;
        Excel.Selection.Value :=Connection.QryBusca.FieldByname('sumacantidad').AsFloat;
        Inc(ren);
        connection.QryBusca.Next;
      end;

      connection.QryBusca2.Next;
       {$ENDREGION}
    end;

  end;

  {$REGION 'Formato de celdas'}
  if Ren<25 then
  begin
    while Ren<25 do
    begin
      Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight :=58;
      Hoja.Range['B'+inttostr(ren)+':D'+inttostr(ren)].Select;
      Excel.Selection.NumberFormat :='@';
      Excel.Selection.Wraptext:=True;
      Excel.Selection.HorizontalAlignment := xlLeft;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.Font.Size :=TamFont;
      Excel.Selection.Font.Bold := true;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Value := '' ;

      Hoja.Range[ColumnaNombre(iCol-1)+inttostr(Ren)+':'+ColumnaNombre(ICol)+inttostr(Ren)].Select;
      Excel.Selection.NumberFormat :='#,##0.0000';
      Excel.Selection.Wraptext:=True;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment   := xlCenter;
      Excel.Selection.Font.Name := 'Arial';
      Excel.Selection.Font.Size :=TamFont;
      Excel.Selection.Font.Bold := false;
      Excel.Selection.MergeCells := True;
      Excel.Selection.Value :='';
      inc(ren);
    end;
  end;

  Hoja.Range['A'+inttostr(RenInicio) + ':F'+inttostr(Ren-1)].Select;
  Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeLeft].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeLeft].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeTop].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeTop].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeTop].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeRight].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeRight].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeRight].Weight := xlThick;
  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideHorizontal].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideHorizontal].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideHorizontal].Weight := xlHairline;
  Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideVertical].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideVertical].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideVertical].Weight := xlThin;


  Hoja.Range['G'+inttostr(RenInicio) + ':'+Columnanombre(iCol-4)+inttostr(Ren-1)].Select;
  Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeLeft].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeLeft].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeTop].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeTop].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeTop].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeRight].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeRight].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeRight].Weight := xlThick;
  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideHorizontal].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideHorizontal].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideHorizontal].Weight := xlHairline;
  Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideVertical].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideVertical].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideVertical].Weight := xlThin;

  Hoja.Range[Columnanombre(iCol-3)+inttostr(RenInicio) + ':'+Columnanombre(iCol)+inttostr(Ren-1)].Select;
  Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeLeft].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeLeft].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeTop].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeTop].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeTop].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThick;
  Excel.Selection.Borders[xlEdgeRight].LineStyle := xlDouble ;
  Excel.Selection.Borders[xlEdgeRight].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeRight].Weight := xlThick;
  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideHorizontal].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideHorizontal].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideHorizontal].Weight := xlHairline;
  Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous ;
  Excel.Selection.Borders[xlInsideVertical].ColorIndex := xlAutomatic;
  Excel.Selection.Borders[xlInsideVertical].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideVertical].Weight := xlThin;

  Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight :=40;
  Hoja.Range['A'+inttostr(ren)+':A'+inttostr(ren)].Select;
  Excel.Selection.NumberFormat :='@';
  Excel.Selection.Wraptext:=false;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlBottom;
  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size :=TamFont;
  Excel.Selection.Font.Bold := false;
  Excel.Selection.MergeCells := True;
  Excel.Selection.Value :='OBSERVACIONES:';
  {$ENDREGION}

  {$REGION 'Pie de Pagina'}
  QryBuscarFirmas:=TZReadOnlyQuery.Create(nil);
  QryBuscarFirmas.Connection := connection.zconnection;
  QryBuscarFirmas.SQL.Add('Select ImgFirma1,ImgFirma5 from firmas where sContrato = :contrato and sIdTurno =:Turno and sNumeroOrden = :Orden And dIdFecha = :fecha');
  QryBuscarFirmas.Params.ParamByName('Orden').DataType := ftString;
  QryBuscarFirmas.Params.ParamByName('Orden').Value := tsNumeroOrden.KeyValue;
  QryBuscarFirmas.Params.ParamByName('Contrato').DataType := ftString;
  QryBuscarFirmas.Params.ParamByName('Contrato').Value :=Global_contrato;
  QryBuscarFirmas.Params.ParamByName('Turno').DataType := ftString;
  QryBuscarFirmas.Params.ParamByName('Turno').Value := Global_Turno;
  QryBuscarFirmas.Params.ParamByName('fecha').DataType := ftDate;
  QryBuscarFirmas.Params.ParamByName('fecha').Value := TdFechaFinal.Date;
  QryBuscarFirmas.Open;
  if QryBuscarFirmas.RecordCount=0 then
  begin
    QryBuscarFirmas.Active := False;
    QryBuscarFirmas.SQL.Clear;
    QryBuscarFirmas.SQL.Add('Select ImgFirma1,ImgFirma5 from firmas where sContrato = :contrato and sNumeroOrden = :Orden and sIdTurno =:Turno And dIdFecha <= :fecha Order By dIdFecha DESC');
    QryBuscarFirmas.Params.ParamByName('Orden').DataType := ftString;
    QryBuscarFirmas.Params.ParamByName('Orden').Value := tsNumeroOrden.KeyValue;
    QryBuscarFirmas.Params.ParamByName('Contrato').DataType := ftString;
    QryBuscarFirmas.Params.ParamByName('Contrato').Value :=Global_contrato;
    QryBuscarFirmas.Params.ParamByName('Turno').DataType := ftString;
    QryBuscarFirmas.Params.ParamByName('Turno').Value := Global_Turno;
    QryBuscarFirmas.Params.ParamByName('fecha').DataType := ftDate;
    QryBuscarFirmas.Params.ParamByName('fecha').Value := TdFechaFinal.Date;
    QryBuscarFirmas.Open;

    if QryBuscarFirmas.RecordCount > 0 then
    begin
      GetTempPath(SizeOf(TempPath), TempPath);
      CNombre1:=TempPath +'imgtempSln'+formatdatetime('dddddd hhnnss',now)+'.jpg';

      fs := QryBuscarFirmas.CreateBlobStream(QryBuscarFirmas.FieldByName('ImgFirma1'), bmRead) ;
      // fs := QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagen'), bmRead);
      If fs.Size > 1 Then
      Begin
          try
              Pic:=TJpegImage.Create;
              try
                 Pic.LoadFromStream(fs);
                 imgAux.Picture.Graphic := Pic;
              finally
                 Pic.Free;
              end;
          finally
              fs.Free;
          End;
        imgAux.Picture.SaveToFile(CNombre1);
      End;

      GetTempPath(SizeOf(TempPath), TempPath);
      CNombre2:=TempPath +'imgtempSln2'+formatdatetime('dddddd hhnnss',now)+'.jpg';

      fs := QryBuscarFirmas.CreateBlobStream(QryBuscarFirmas.FieldByName('ImgFirma5'), bmRead) ;
      // fs := QrDatos.CreateBlobStream(QrDatos.FieldByName('bImagen'), bmRead);
      If fs.Size > 1 Then
      Begin
          try
              Pic:=TJpegImage.Create;
              try
                 Pic.LoadFromStream(fs);
                 imgAux.Picture.Graphic := Pic;
              finally
                 Pic.Free;
              end;
          finally
              fs.Free;
          End;
        imgAux.Picture.SaveToFile(CNombre2);
      End
    
    end;

  end;

  Excel.ActiveSheet.PageSetup.PaperSize := xlPaperLetter;
  Excel.ActiveWindow.View :=xlPageLayoutView;
  Excel.ActiveSheet.PageSetup.PrintTitleRows := '$1:$11';
  Excel.ActiveSheet.PageSetUp.CenterFooter :='';
  Excel.ActiveSheet.PageSetUp.LeftFooter := '';
  Excel.ActiveSheet.PageSetUp.RightFooter := '';
  Excel.ActiveSheet.PageSetUp.LeftMargin     := 0;
  Excel.ActiveSheet.PageSetUp.RightMargin    := 0;
  Excel.ActiveSheet.PageSetUp.TopMargin      := 14;
  Excel.ActiveSheet.PageSetUp.BottomMargin   := 120;
  Excel.ActiveSheet.PageSetUp.HeaderMargin   := 0;
  Excel.ActiveSheet.PageSetUp.FooterMargin   :=Excel.InchesToPoints(0.777777777777778); //56;
  Excel.ActiveSheet.PageSetUp.Zoom := 32;

  Excel.ActiveSheet.PageSetUp.ScaleWithDocHeaderFooter := True;
  Excel.ActiveSheet.PageSetUp.AlignMarginsHeaderFooter := False;
  Excel.ActiveSheet.PageSetUp.EvenPage.LeftFooter.Text := '';
  Excel.ActiveSheet.PageSetUp.EvenPage.CenterFooter.Text := '';
  Excel.ActiveSheet.PageSetUp.EvenPage.RightFooter.Text := '';
  Excel.ActiveSheet.PageSetUp.FirstPage.LeftFooter.Text := '';
  Excel.ActiveSheet.PageSetUp.FirstPage.CenterFooter.Text := '';
  Excel.ActiveSheet.PageSetUp.FirstPage.RightFooter.Text := '';
  Excel.ActiveSheet.PageSetup.Orientation := xlLandscape;
  Excel.ActiveSheet.PageSetUp.Zoom           := False;
  Excel.ActiveSheet.PageSetUp.FitToPagesWide := 1;
  Excel.ActiveSheet.PageSetUp.FitToPagesTall := False;
  if (Excel.Application.version >= 14) then
  begin
    Excel.PrintCommunication := True;
    Excel.PrintCommunication := False;
  end;
  Excel.ActiveSheet.PageSetup.LeftHeader := '';
  Excel.ActiveSheet.PageSetup.CenterHeader := '';
  Excel.ActiveSheet.PageSetup.RightHeader := '';
  Excel.ActiveSheet.PageSetup.LeftFooter := '';
  Excel.ActiveSheet.PageSetup.CenterFooter := '';
  Excel.ActiveSheet.PageSetup.RightFooter := '';
  Excel.ActiveSheet.PageSetup.LeftMargin := Excel.InchesToPoints(0);
  Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0);
  Excel.ActiveSheet.PageSetup.TopMargin := Excel.InchesToPoints(0.194444444444444);
  Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(1.66666666666667);
  Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0);
  Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0.777777777777778);
  Excel.ActiveSheet.PageSetup.Zoom := 32;
  Excel.ActiveSheet.PageSetup.PrintErrors := xlPrintErrorsDisplayed;
  Excel.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter := False;
  Excel.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter := False;
  Excel.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
  Excel.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := False;
  Excel.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.RightHeader.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.RightFooter.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.RightHeader.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.RightFooter.Text := '';
  if (Excel.Application.version >= 14) then
    Excel.PrintCommunication := True;

  if QryBuscarFirmas.RecordCount > 0 then
  begin
    if FileExists(CNombre1) then
    begin
      Excel.ActiveSheet.PageSetup.LeftFooterPicture.Filename :=CNombre1;
      Excel.ActiveSheet.PageSetup.LeftFooter :='&G';
      //Excel.ActiveSheet.PageSetup.LeftFooterPicture.Height := 216;
      //Excel.ActiveSheet.PageSetup.LeftFooterPicture.Width := 558;
    end;

    if FileExists(CNombre2) then
    begin
      Excel.ActiveSheet.PageSetup.RightFooterPicture.Filename :=CNombre2;
       Excel.ActiveSheet.PageSetup.RightFooter :='&G';

      //Excel.ActiveSheet.PageSetup.RightFooterPicture.Height := 216;
      //Excel.ActiveSheet.PageSetup.RightFooterPicture.Width := 558;
    end;
  end;

  if (Excel.Application.version >= 14) then
    Excel.PrintCommunication := False;

  Excel.ActiveSheet.PageSetup.PrintTitleRows := '$1:$11';
  Excel.ActiveSheet.PageSetup.PrintTitleColumns := '';

  if (Excel.Application.version >= 14) then
    Excel.PrintCommunication := True;
  Excel.ActiveSheet.PageSetup.PrintArea := '';

  if (Excel.Application.version >= 14) then
    Excel.PrintCommunication := False;


  Excel.ActiveSheet.PageSetup.LeftHeader := '';
  Excel.ActiveSheet.PageSetup.CenterHeader := '';
  Excel.ActiveSheet.PageSetup.RightHeader := '';

  Excel.ActiveSheet.PageSetup.CenterFooter :='&"Arial,Normal"&'+inttostr(TamFont)+'&P de &#';//'&Z&G&P de &#&D&G'; //'&P de &N';

  Excel.ActiveSheet.PageSetup.LeftMargin := Excel.InchesToPoints(0);
  Excel.ActiveSheet.PageSetup.RightMargin := Excel.InchesToPoints(0);
  Excel.ActiveSheet.PageSetup.TopMargin := Excel.InchesToPoints(0.196850393700787);
  Excel.ActiveSheet.PageSetup.BottomMargin := Excel.InchesToPoints(1.65354330708661);
  Excel.ActiveSheet.PageSetup.HeaderMargin := Excel.InchesToPoints(0);
  Excel.ActiveSheet.PageSetup.FooterMargin := Excel.InchesToPoints(0.78740157480315);
  Excel.ActiveSheet.PageSetup.PrintHeadings := False;
  Excel.ActiveSheet.PageSetup.PrintGridlines := False;
 // Excel.ActiveSheet.PageSetup.PrintComments := xlPrintNoComments;  //
  Excel.ActiveSheet.PageSetup.PrintQuality := 600;
  Excel.ActiveSheet.PageSetup.CenterHorizontally := False;
  Excel.ActiveSheet.PageSetup.CenterVertically := False;
  Excel.ActiveSheet.PageSetup.Orientation := xlLandscape;
  Excel.ActiveSheet.PageSetup.Draft := False;
  Excel.ActiveSheet.PageSetup.FirstPageNumber := xlAutomatic;
  Excel.ActiveSheet.PageSetup.Order := xlDownThenOver;
  Excel.ActiveSheet.PageSetup.BlackAndWhite := False;
  Excel.ActiveSheet.PageSetup.Zoom := 32;
  Excel.ActiveSheet.PageSetup.PrintErrors := xlPrintErrorsDisplayed;
  Excel.ActiveSheet.PageSetup.OddAndEvenPagesHeaderFooter := False;
  Excel.ActiveSheet.PageSetup.DifferentFirstPageHeaderFooter := False;
  Excel.ActiveSheet.PageSetup.ScaleWithDocHeaderFooter := True;
  Excel.ActiveSheet.PageSetup.AlignMarginsHeaderFooter := False;
  Excel.ActiveSheet.PageSetup.EvenPage.LeftHeader.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.CenterHeader.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.RightHeader.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.LeftFooter.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.CenterFooter.Text := '';
  Excel.ActiveSheet.PageSetup.EvenPage.RightFooter.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.LeftHeader.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.CenterHeader.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.RightHeader.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.LeftFooter.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.CenterFooter.Text := '';
  Excel.ActiveSheet.PageSetup.FirstPage.RightFooter.Text := '';
   Excel.ActiveSheet.PageSetUp.Zoom           := False;
  Excel.ActiveSheet.PageSetUp.FitToPagesWide := 1;
  Excel.ActiveSheet.PageSetUp.FitToPagesTall := False;
  if (Excel.Application.version >= 14) then
    Excel.PrintCommunication := True;

  Excel.ActiveWindow.View := xlNormalView;
  Excel.ActiveWindow.Zoom := 73;
  {$ENDREGION}

   Hoja.Range['A1:A1'].Select;
end;



procedure TfrmDetalledeInstalacion.EditPartidasEnter(Sender: TObject);
begin
  opcPartidas.Checked := True;
end;

///////////////////////////////////////////////
Procedure TfrmDetalledeInstalacion.DatosPartidasPaquetes(ParamPartidas: string = '');
Var
  cadenasql, CadFecha, tmpNombre, cadena : String;
  fs: tStream;
  Alto : Extended;
  MiFechaI, MiFechaF, MiFecha: tDate;
  FechaDia : tDateTime;
  Pondera, AvanceActual, AvanceAnterior, AvanceActualP, AvanceAnteriorP, dAvanceTotal: double;
  contadorm, cuenta, Ren, Ren1, Ren2, nivel, fila, indice, i1, i2,i3, total : Integer;
  yatermino, partida, partida1, partida2, partida3, NombreArchivo, CadPriodo, entroalpadre : String;
  Imagen: TField;
  Altura, Margen: Extended;
  ArrPonderado  : array[1..5] of string  ;
  actividades, Ponderado, ejecutadas, Anterior, Actual   : TZReadOnlyQuery;
  AuxCadena:string;
  i:Integer;
Begin

  Ponderado             := TZReadOnlyQuery.Create(Self);
  Ponderado.Connection  := connection.zConnection;
  actividades           := TZReadOnlyQuery.Create(Self);
  actividades.Connection  := connection.zConnection;
  Anterior              := TZReadOnlyQuery.Create(Self);
  Anterior.Connection   := connection.zConnection;
  Actual                := TZReadOnlyQuery.Create(Self);
  Actual.Connection     := connection.zConnection;
  Ejecutadas            := TZReadOnlyQuery.Create(Self);
  Ejecutadas.Connection := connection.zConnection;

  AuxCadena:='';
  for I := 1 to NumItems(ParamPartidas,',') do
  begin
    if AuxCadena='' then
      AuxCadena:=' and (aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
    else
       AuxCadena:=AuxCadena + ' or aa.snumeroactividad=' + QuotedStr(TraerItem(ParamPartidas,',',i))
  end;

  if AuxCadena<>'' then
    AuxCadena:=AuxCadena + ') ';



  Ren  := 10;
  Ren1 := 10 ;
  Ren2 := 8 ;

  // Realizar los ajustes visuales y de formato de hoja
  Excel.ActiveWindow.Zoom := 100;
  Excel.Columns['A:A'].ColumnWidth := 16;
  Excel.Columns['B:B'].ColumnWidth := 20;
  Excel.Columns['C:C'].ColumnWidth := 71;
  Excel.Columns['D:D'].ColumnWidth := 16;
  Excel.Columns['E:E'].ColumnWidth := 16;
  Excel.Columns['F:F'].ColumnWidth := 16;

  MiFecha  := tdFechaInicial.date;
  MiFechaI := tdFechaInicial.date;
  MiFechaF := tdFechaFinal.Date ;

      // Colocar los encabezados de la plantilla...

    {*****************************************************************************
     ** Colocar el logotipo de la empresa                                        }
  NombreArchivo := NombreArchivoTemporal;   // Obtener un nombre de archivo temporal
  Imagen := Connection.configuracion.FieldByName('bImagen');
  fs := Connection.configuracion.CreateBlobStream(Imagen, bmRead);
  if fs.size > 0 then
  Begin
    try
      fs.Seek(0, soFromBeginning);
      with TFileStream.Create(NombreArchivo, fmCreate) do
        try
          CopyFrom(fs, fs.Size)
        finally
          Free
        end;
    finally
      fs.Free
    end;
    Excel.ActiveSheet.Pictures.Insert(NombreArchivo).Select;
    // Determinar el tamaño real de la imagen
    Altura := Excel.Rows[1].Height * 0.7;
    Margen := (Excel.Rows[1].Height - Altura) / 2 ;
    Excel.Selection.ShapeRange.Left := Excel.Columns['A:A'].Width ;
    Excel.Selection.ShapeRange.Top := Margen;
    SysUtils.DeleteFile(NombreArchivo); // Borrar el archivo temporal
  End;
    {** Termina Colocar el logotipo de la empresa
     *****************************************************************************}

  Hoja.Range['C2:F3'].Select;
  Excel.Selection.Value := 'VOLUMETRIAS DE ANEXO CON SISTEMAS Y MATERIALES ';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;

  Excel.Selection.HorizontalAlignment := xlLeft ;
  Excel.Selection.Merge ;
  Excel.Selection.WrapText := False   ;
  Excel.Selection.Orientation := 0     ;
  Excel.Selection.AddIndent := False   ;
  Excel.Selection.IndentLevel := 0     ;
  Excel.Selection.ShrinkToFit := False  ;
  Excel.Selection.MergeCells := False     ;

  Hoja.Range['A7:A7'].Select;
  Excel.Selection.Value := 'FECHA';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;

  Hoja.Range['B7:B7'].Select;
  Excel.Selection.Value := 'PARTIDA/SUBPARTIDA';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;

  Hoja.Range['C7:C7'].Select;
  Excel.Selection.Value := 'DESCRIPCION';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;

  Hoja.Range['D7:D7'].Select;
  Excel.Selection.Value := 'ANTERIOR';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;

  Hoja.Range['E7:E7'].Select;
  Excel.Selection.Value := 'ACTUAL';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;

  Hoja.Range['F7:F7'].Select;
  Excel.Selection.Value := 'ACUMULADO';
  Excel.Selection.Font.Color := clWhite;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.Interior.ColorIndex := 41;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment   := xlCenter;

  contadorm := 0 ;
  connection.QryBusca2.Active := False ;
  connection.QryBusca2.Filtered := False;
  connection.QryBusca2.SQL.Clear ;
  connection.QryBusca2.SQL.Add( 'select b.iFase, b.sNumeroActividad, b.dCantidad, b.dAvance from bitacoradealcances b ' +
                                'inner join actividadesxorden ao on(ao.scontrato=b.scontrato and ao.snumeroorden=b.snumeroorden ' +
                                'and ao.swbs=b.swbs and ao.snumeroactividad=b.snumeroactividad) ' +
                                'inner join actividadesxanexo aa on(aa.scontrato=ao.scontrato and aa.sidconvenio=ao.sidconvenio '+
                                'and aa.swbs=ao.swbsContrato and aa.snumeroactividad=ao.snumeroactividad) ' +
                                'where ao.sContrato =:Contrato and ao.sidconvenio=:Convenio and ao.snumeroorden=:Orden ' +
                                'And b.dIdFecha >=:FechaInicial And b.dIdFecha <=:FechaFinal ' +
                                Auxcadena +' And b.lConceptoEjecutado="No" And b.sPaquete="" Group by b.sNumeroActividad Order by  b.sNumeroActividad ') ;
  connection.QryBusca2.Params.ParamByName('Contrato').DataType      := ftString;
  connection.QryBusca2.Params.ParamByName('Contrato').Value         := global_contrato;
  Connection.QryBusca2.ParamByName('Convenio').AsString:=global_Convenio;
  connection.QryBusca2.ParamByName('orden').AsString:=tsnumeroorden.KeyValue;
  connection.QryBusca2.Params.ParamByName('FechaInicial').DataType  := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaInicial').Value     := tdFechaInicial.DateTime;
  connection.QryBusca2.Params.ParamByName('FechaFinal').DataType    := ftDate;
  connection.QryBusca2.Params.ParamByName('FechaFinal').Value       := tdFechaFinal.DateTime;
  Connection.QryBusca2.Open ;
  cuenta := connection.QryBusca2.RecordCount ;
  
  while not connection.QryBusca2.Eof do
  Begin
    partida1 := connection.qryBusca2.fieldByname('sNumeroActividad').AsString;
    partida := partida1;

    ///Primero ya tengo la partida Identificar si es Hermana y depende d e algun padre.....
    {  if pos('.', partida1) > 0 then
    Begin}
      //De aqui saco los alcancesxactividad solamente los ponderados por cada hijo .....
    entroalpadre := 'Si' ;
    partida := copy(partida1, 1, pos('.', partida1) -1) ;
    partida := partida1;
    Fila := 1;
    actividades.Active := False;
    actividades.Filtered := False;
    actividades.SQL.Clear;
    actividades.SQL.Add('select aa.swbs,b.iFase, b.sNumeroActividad From bitacoradealcances b ' +
                        'inner join actividadesxorden ao on(ao.scontrato=b.scontrato and ao.snumeroorden=b.snumeroorden ' +
                        'and ao.swbs=b.swbs and ao.snumeroactividad=b.snumeroactividad) ' +
                        'inner join actividadesxanexo aa on(aa.scontrato=ao.scontrato and aa.sidconvenio=ao.sidconvenio '+
                        'and aa.swbs=ao.swbsContrato and aa.snumeroactividad=ao.snumeroactividad) ' +
                        'Where b.sContrato =:Contrato And aa.sidconvenio=:Convenio and ao.snumeroorden=:Orden ' +
                        'and b.sNumeroActividad like :Actividad ' +
                        'and b.dIdFecha >=:FechaInicial and b.dIdFecha <=:FechaFinal Group by b.sNumeroActividad ') ;
    actividades.Params.ParamByName('Contrato').DataType     := ftString;
    actividades.Params.ParamByName('Contrato').Value        := global_contrato;
    actividades.ParamByName('convenio').AsString:=global_convenio;
    actividades.ParamByName('orden').AsString:=tsNumeroOrden.KeyValue;
    actividades.Params.ParamByName('Actividad').DataType    := ftString;
    actividades.Params.ParamByName('Actividad').Value       := partida + '%';
    actividades.Params.ParamByName('FechaInicial').DataType := ftDate;
    actividades.Params.ParamByName('FechaInicial').Value    := tdFechaInicial.date;
    actividades.Params.ParamByName('FechaFinal').DataType   := ftDate;
    actividades.Params.ParamByName('FechaFinal').Value      := tdFechaFinal.date;
    actividades.Open;
    pondera := 0 ;
    if actividades.RecordCount > 0 Then
    begin
      for i1 := 1 to actividades.RecordCount  do
      begin
        Ponderado.Active := False;
        Ponderado.Filtered := False;
        Ponderado.SQL.Clear;
        Ponderado.SQL.Add('select sDescripcion, iFase, sNumeroActividad, dPonderado From alcancesxactividad Where sContrato =:Contrato And ' +
                          'sidconvenio=:Convenio and swbs=:wbs and sNumeroActividad =:Actividad and lPrincipal="Si" ');
        Ponderado.Params.ParamByName('Contrato').DataType     := ftString;
        Ponderado.Params.ParamByName('Contrato').Value        := global_contrato;
        Ponderado.Params.ParamByName('Actividad').DataType    := ftString;
        Ponderado.Params.ParamByName('Actividad').Value       := actividades.fieldByname('sNumeroActividad').AsString;
        Ponderado.Params.ParamByName('wbs').AsString       := actividades.fieldByname('swbs').AsString;
        Ponderado.ParamByName('convenio').AsString:=global_convenio;
        Ponderado.Open;
        arrPonderado[1] := Partida ;
        Pondera         := Ponderado.fieldByname('dPonderado').asfloat + Pondera;
        arrPonderado[3] := Ponderado.fieldByname('sDescripcion').AsString ;

        Ponderado.next ;
        actividades.Next ;
      end;
      arrPonderado[2] := FloatToStr(Pondera);
    end  ;
     {     End
       Else
          begin
            partida := partida1;
            entroalpadre := 'No' ;
          end;
      }

    cadena:= formatDateTime('dd/mm/yyyy',tdFechaFinal.date);
    Ejecutadas.Active := False;
    Ejecutadas.Filtered := False;
    Ejecutadas.SQL.Clear;
    Ejecutadas.SQL.Add( 'select a.dPonderado,ao.swbs,ao.snumeroorden, b.dIdFecha, a.iFase, a.sNumeroActividad,b.spaquete, a.sNumeroActividadsub, a.sDescripcion, sum(b.dAvanceActividad) as dAvanceActividadAnterior, ' +
                        'sum(b.dAvance) as dAvance from bitacoradealcances b  ' +
                        'inner join actividadesxorden ao on(ao.scontrato=b.scontrato and ao.snumeroorden=b.snumeroorden ' +
                        'and ao.swbs=b.swbs and ao.snumeroactividad=b.snumeroactividad) ' +
                        'inner join actividadesxanexo aa on(aa.scontrato=ao.scontrato and aa.sidconvenio=ao.sidconvenio '+
                        'and aa.swbs=ao.swbsContrato and aa.snumeroactividad=ao.snumeroactividad) ' +
                        'inner join alcancesxactividad a On (a.sContrato=b.sContrato And a.sIdConvenio=aa.sidconvenio and aa.swbs=a.swbs And a.sNumeroActividad=aa.sNumeroActividad and a.iFase=b.iFase) ' +
                        'where b.sContrato =:Contrato and ao.sidconvenio=:convenio and b.sNumeroActividad =:Actividad and ao.snumeroorden=:orden and b.dIdFecha >=:FechaInicial And b.dIdFecha <=:FechaFinal ' +
                        'group by b.dIdFecha, a.sNumeroActividadSub, b.iFase') ;
    Ejecutadas.Params.ParamByName('Contrato').DataType    := ftString;
    Ejecutadas.Params.ParamByName('Contrato').Value       := global_contrato;
    Ejecutadas.ParamByName('convenio').AsString:=global_convenio;
    Ejecutadas.ParamByName('orden').AsString:=tsNumeroOrden.KeyValue;
    Ejecutadas.Params.ParamByName('Actividad').DataType   := ftString;
    Ejecutadas.Params.ParamByName('Actividad').Value      := partida1;
    Ejecutadas.Params.ParamByName('FechaInicial').DataType := ftDate;
    Ejecutadas.Params.ParamByName('FechaInicial').Value    := tdFechaInicial.Date ;
    Ejecutadas.Params.ParamByName('FechaFinal').DataType   := ftDate;
    Ejecutadas.Params.ParamByName('FechaFinal').Value      := tdFechaFinal.Date ;
    Ejecutadas.Open;
    Ren:=1 ;
    MiFecha   := Ejecutadas.FieldBYname('dIdFecha').AsDateTime ;
    yatermino := 'No' ;

    if Ejecutadas.RecordCount > 0 Then
    while Not ejecutadas.Eof do
    begin
      MiFecha   := Ejecutadas.FieldByName('dIdFecha').AsDateTime ;      //Lobo
      contadorm := DaysBetween(MiFechaF, MiFechaI) ;
      for i2 := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
      begin
        if yatermino = 'Si' Then
          break ;
        MiFecha := Ejecutadas.FieldByname('dIdFecha').AsDateTime ;
        cadena:= formatDateTime('dd/mm/yyyy',Mifecha);

                          //checar las padres para poderlas imprimir
        Ren2 := Ren2 + 1 ;

        Anterior.Active := False ;
        Anterior.Filtered := False;
        Anterior.SQL.Clear ;
        Anterior.SQL.Add('select b.iFase, b.sNumeroActividad, sum(b.dAvanceActividad) as dAvanceActividadAnterior, ' +
                         'sum(b.dAvance) as dAvance from bitacoradealcances b  ' +
                         'where b.sContrato =:Contrato and b.snumeroorden=:orden and b.swbs=:wbs ' +
                         'and b.sNumeroActividad =:Actividad and b.dIdFecha <:FechaAnterior ' +
                         'and iFase=:Fase Group by b.sNumeroActividad, b.iFase Order by b.iFase') ;
        Anterior.Params.ParamByName('Contrato').DataType      := ftString;
        Anterior.Params.ParamByName('Contrato').Value         := global_contrato;
        Anterior.Params.ParamByName('Actividad').DataType     := ftString;
        Anterior.Params.ParamByName('Actividad').Value        := Ejecutadas.FieldByname('sNumeroActividad').AsString;
        Anterior.Params.ParamByName('FechaAnterior').DataType := ftDate;
        Anterior.Params.ParamByName('FechaAnterior').Value    := Ejecutadas.FieldByname('dIdFecha').AsDateTime ;
        Anterior.Params.ParamByName('Fase').DataType          := ftInteger;
        Anterior.Params.ParamByName('Fase').Value             := Ejecutadas.FieldByname('iFase').AsInteger;
        Anterior.Params.ParamByName('Orden').AsString        := Ejecutadas.FieldByname('snumeroorden').AsString;
        Anterior.Params.ParamByName('wbs').AsString        := Ejecutadas.FieldByname('swbs').AsString;
        Anterior.Open ;
        if Anterior.RecordCount = 0 then
          AvanceAnterior := 0
        else
          AvanceAnterior := Anterior.fieldByname('dAvanceActividadAnterior').asfloat ;

        Actual.Active := False ;
        Actual.Filtered := False;
        Actual.SQL.Clear ;
        Actual.SQL.Add('select b.dIdFecha, b.iFase, sum(b.dAvanceActividad) as dAvanceActividadActual, ' +
                       'sum(b.dAvance) as dAvance from bitacoradealcances b  ' +
                       'where b.sContrato =:Contrato and b.snumeroorden=:orden and b.swbs=:wbs '+
                       'and b.sNumeroActividad =:Actividad and b.dIdFecha =:FechaActual ' +
                       'and iFase=:Fase Group by b.sNumeroActividad, b.iFase Order by b.iFase') ;
        Actual.Params.ParamByName('Contrato').DataType    := ftString;
        Actual.Params.ParamByName('Contrato').Value       := global_contrato;
        Actual.Params.ParamByName('Actividad').DataType   := ftString;
        Actual.Params.ParamByName('Actividad').Value      := Ejecutadas.FieldBYname('sNumeroActividad').AsString;
        Actual.Params.ParamByName('FechaActual').DataType := ftDate;
        Actual.Params.ParamByName('FechaActual').Value    := Ejecutadas.FieldByname('dIdFecha').AsDateTime ;
        Actual.Params.ParamByName('Fase').DataType        := ftInteger;
        Actual.Params.ParamByName('Fase').Value           := Ejecutadas.FieldByname('iFase').AsInteger;
        Actual.Params.ParamByName('Orden').AsString        := Ejecutadas.FieldByname('snumeroorden').AsString;
        Actual.Params.ParamByName('wbs').AsString        := Ejecutadas.FieldByname('swbs').AsString;
        Actual.Open ;

        if actual.RecordCount = 0 then
          AvanceActual := 0
        else
          AvanceActual := Actual.fieldByname('dAvanceActividadActual').asfloat ;

        ////////////////////////////////////////////////////
        If (entroalpadre = 'Si')  And (ejecutadas.FieldByname('spaquete').AsString<>'')  Then
        begin
          Hoja.Cells[Ren2,1].Select;
          Excel.Selection.Interior.ColorIndex := 35;
          Excel.Selection.Font.Color := clBlack;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Value := Ejecutadas.FieldByname('dIdFecha').AsDateTime ;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Font.Size := 11;
          Excel.Selection.Font.Bold := False;
          Excel.Selection.Font.Name := 'Calibri';

          Hoja.Cells[Ren2,2].Select;
          Excel.Selection.Value         := arrPonderado[1] ;
          Excel.Selection.Interior.ColorIndex := 35;
          Excel.Selection.Font.Color := clBlack;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.HorizontalAlignment  := xlLeft;
          Excel.Selection.VerticalAlignment    := xlCenter;


          Hoja.Cells[Ren2,3].Select;
          Excel.Selection.Interior.ColorIndex := 35;
          Excel.Selection.Font.Color := clBlack;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Excel.Selection.Value := arrPonderado[3];
          Excel.Selection.HorizontalAlignment   := xlLeft;
          Excel.selection.WrapText := True    ;
          Excel.selection.Orientation := 0    ;
          Excel.selection.AddIndent := False  ;
          Excel.selection.IndentLevel := 0    ;
          Excel.selection.ShrinkToFit := False  ;
          Excel.selection.MergeCells := False  ;

          Hoja.Cells[Ren2,4].Select;
          Excel.Selection.Interior.ColorIndex := 35;
          Excel.Selection.Font.Color := clBlack;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          AvanceanteriorP := (AvanceAnterior*Pondera)/100  ;
          Excel.Selection.Value               := AvanceAnteriorP ;
          Excel.Selection.NumberFormat        := '#0.00';
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment   := xlCenter;

          Hoja.Cells[Ren2,5].Select;
          Excel.Selection.Interior.ColorIndex := 35;
          Excel.Selection.Font.Color := clBlack;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
          AvanceactualP  := (AvanceActual*Pondera)/100 ;
          Excel.Selection.Value               := AvanceActualP ;
          Excel.Selection.NumberFormat        := '#0.00';
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment   := xlCenter;

          Hoja.Cells[Ren2,6].Select;
          Excel.Selection.Interior.ColorIndex := 35;
          Excel.Selection.Font.Color := clBlack;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;

          Excel.Selection.Value               := AvanceanteriorP + AvanceActualP ;
          Excel.Selection.NumberFormat        := '#0.00';
          Excel.Selection.HorizontalAlignment := xlLeft;
          Excel.Selection.VerticalAlignment   := xlCenter;
          Ren2 := Ren2 + 1 ;
        end;
                          /// /////////////////////////////////////////

        Hoja.Cells[Ren2,1].Select;
        if ejecutadas.FieldByname('dPonderado').asfloat=100 Then
        begin
          Excel.Selection.Interior.ColorIndex := 42;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
        end ;
        Excel.Selection.Value := Ejecutadas.FieldByname('dIdFecha').AsDateTime ;
        Excel.Selection.HorizontalAlignment := xlCenter;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Excel.Selection.Font.Size := 11;
        Excel.Selection.Font.Bold := False;
        Excel.Selection.Font.Name := 'Calibri';

        Hoja.Cells[Ren2,2].Select;
        if ejecutadas.FieldBYName('dPonderado').asfloat=100 Then
        begin
          Excel.Selection.Value         := Ejecutadas.FieldByname('sNumeroActividad').AsString ;
          Excel.Selection.Interior.ColorIndex := 42;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
        end
        else
          Excel.Selection.Value         := Ejecutadas.FieldByname('sNumeroActividadsub').AsString;

        Excel.Selection.HorizontalAlignment  := xlLeft;
        Excel.Selection.VerticalAlignment    := xlCenter;


        Hoja.Cells[Ren2,3].Select;
        if ejecutadas.FieldByname('dPonderado').asfloat=100 Then
        begin
          Excel.Selection.Interior.ColorIndex := 42;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
        end ;
        Excel.Selection.Value := Ejecutadas.FieldByname('sDescripcion').AsString;
        Excel.Selection.HorizontalAlignment   := xlLeft;
        Excel.selection.WrapText := True    ;
        Excel.selection.Orientation := 0    ;
        Excel.selection.AddIndent := False  ;
        Excel.selection.IndentLevel := 0    ;
        Excel.selection.ShrinkToFit := False  ;
        Excel.selection.MergeCells := False  ;

        Hoja.Cells[Ren2,4].Select;
        if ejecutadas.FieldByname('dPonderado').asfloat=100 Then
        begin
          Excel.Selection.Interior.ColorIndex := 42;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
        end ;
        
        Excel.Selection.Value               := Avanceanterior;
        Excel.Selection.NumberFormat        := '#0.00';
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.VerticalAlignment   := xlCenter;

        Hoja.Cells[Ren2,5].Select;
        if ejecutadas.FieldByName('dPonderado').asfloat=100 Then
        begin
          Excel.Selection.Interior.ColorIndex := 42;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
        end  ;
        Excel.Selection.Value               := AvanceActual ;
        Excel.Selection.NumberFormat        := '#0.00';
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.VerticalAlignment   := xlCenter;

        Hoja.Cells[Ren2,6].Select;
        if ejecutadas.FieldByname('dPonderado').asfloat=100 Then
        begin
          Excel.Selection.Interior.ColorIndex := 42;
          Excel.Selection.Font.Color := clWhite;
          Excel.Selection.Font.Bold := True;
          Excel.Selection.HorizontalAlignment := xlCenter;
          Excel.Selection.VerticalAlignment   := xlCenter;
        end ;
        Excel.Selection.Value               := Avanceanterior + AvanceActual ;
        Excel.Selection.NumberFormat        := '#0.00';
        Excel.Selection.HorizontalAlignment := xlLeft;
        Excel.Selection.VerticalAlignment   := xlCenter;
        Ren := Ren + 1 ;

        Ejecutadas.Next ;
        if ejecutadas.eof= true Then
        begin
          //connection.QryBusca2.Next ;
          yatermino := 'Si'  ;
        end;
      end;
    end ;
    connection.QryBusca2.Next ;
  end;
  Ponderado.Destroy ;
  Actual.Destroy ;
  Anterior.Destroy ;
  Ejecutadas.Destroy ;
End;




procedure TfrmDetalledeInstalacion.FormShow(Sender: TObject);
begin
  TRY
    sMenuP:=stMenu;
    BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'rInstalado');
  	BotonPermiso.permisosBotones2(nil,nil,nil,btnReport2);
    OrdenesdeTrabajo.Active := False ;
    OrdenesdeTrabajo.Params.ParamByName('Contrato').DataType := ftString ;
    OrdenesdeTrabajo.Params.ParamByName('Contrato').Value := global_contrato ;
    OrdenesdeTrabajo.Open ;

    If OrdenesdeTrabajo.RecordCount > 0 Then
        tsNumeroOrden.KeyValue := OrdenesdeTrabajo.FieldValues['sNumeroOrden'] ;

    tdFechaInicial.Date := Date ;
    tdFechaFinal.Date := Date ;
    tsNumeroOrden.SetFocus;

    FechaInicio.Date := VolverAInicioDeMes(FechaInicio.Date);
    FechaTermino.Date := Date;

  EXCEPT
    on e : exception do begin
    UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_detalledeinstalacion', 'Al iniciar el formulario', 0);
    end;
  END;
end;

procedure TfrmDetalledeInstalacion.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  BotonPermiso.Free;
  action := cafree ;
end;

procedure TfrmDetalledeInstalacion.tsNumeroOrdenKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tdFechaInicial.SetFocus
end;

procedure TfrmDetalledeInstalacion.tsNumeroOrdenClick(Sender: TObject);
begin
  ZQBuscaFechas.Active := False;
  ZQBuscaFechas.ParamByName('folio').AsString := tsNumeroOrden.Text;
  ZQBuscaFechas.Open;
  if ZQBuscaFechas.RecordCount = 1 then
  begin
    tdFechaInicial.date := zqbuscafechas.FieldByName('fechai').AsDateTime;
    tdFechaFinal.Date := zqbuscafechas.FieldByName('fechaf').AsDateTime;
  end;

end;

procedure TfrmDetalledeInstalacion.tsNumeroOrdenEnter(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_entrada
end;

procedure TfrmDetalledeInstalacion.tsNumeroOrdenExit(Sender: TObject);
begin
    tsNumeroOrden.Color := global_color_salida
end;

procedure TfrmDetalledeInstalacion.tdFechaInicialChange(Sender: TObject);
begin
  tdFechaFinal.Date:=tdFechainicial.Date;
end;

procedure TfrmDetalledeInstalacion.tdFechaInicialEnter(Sender: TObject);
begin
    tdFechaInicial.Color := global_color_entrada
end;

procedure TfrmDetalledeInstalacion.tdFechaInicialExit(Sender: TObject);
begin
    tdFechaInicial.Color := global_color_salida
end;

procedure TfrmDetalledeInstalacion.tdFechaInicialKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tdFechaFinal.SetFocus 
end;

procedure TfrmDetalledeInstalacion.tdFechaFinalChange(Sender: TObject);
begin
//  tdFechaFinal.MinDate:=tdFechainicial.Date;
end;

procedure TfrmDetalledeInstalacion.tdFechaFinalEnter(Sender: TObject);
begin
    tdFechaFinal.Color := global_color_entrada
end;

procedure TfrmDetalledeInstalacion.tdFechaFinalExit(Sender: TObject);
begin
    tdFechaFinal.Color := global_color_salida
end;

procedure TfrmDetalledeInstalacion.tdFechaFinalKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        btnReport2.SetFocus
end;        


procedure TfrmDetalledeInstalacion.ActividadesxOrdenCalcFields(
  DataSet: TDataSet);
begin
 try
   If ActividadesxOrden.FieldValues['sWbs'] <> Null Then
         ActividadesxOrdensWbsSpace.Text := espaces (ActividadesxOrden.FieldValues['iNivel']) + ActividadesxOrden.FieldValues['sWbs'] ;

     If ActividadesxOrden.FieldValues['sTipoActividad'] = 'Actividad' Then
     Begin
          Connection.qryBusca2.Active := False ;
          Connection.qryBusca2.SQL.Clear ;
          Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad) as Instalado From bitacoradeactividades b ' +
                                        'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.dIdFecha < :Fecha And b.sWbs = :Wbs And b.sNumeroActividad = :Actividad ' +
                                        'Group By b.sWbs, b.sNumeroActividad') ;
          Connection.QryBusca2.Params.ParamByName('Contrato').DataType  := ftString ;
          Connection.QryBusca2.Params.ParamByName('Contrato').Value     := global_Contrato ;
          Connection.QryBusca2.Params.ParamByName('Orden').DataType     := ftString ;
          Connection.QryBusca2.Params.ParamByName('Orden').Value        := tsNumeroOrden.Text ;
          Connection.QryBusca2.Params.ParamByName('Fecha').DataType     := ftDate ;
          Connection.QryBusca2.Params.ParamByName('Fecha').Value        := tdFechaInicial.Date ;
          Connection.QryBusca2.Params.ParamByName('Wbs').DataType       := ftString ;
          Connection.QryBusca2.Params.ParamByName('Wbs').Value          := ActividadesxOrden.FieldValues['sWbs'] ;
          Connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Actividad').Value    := ActividadesxOrden.FieldValues['sNumeroActividad'] ;
          Connection.qryBusca2.Open ;
          If Connection.qryBusca2.RecordCount > 0 then
               ActividadesxOrdendAcumuladoAnterior.Value := Connection.qryBusca2.FieldValues['Instalado']
          Else
               ActividadesxOrdendAcumuladoAnterior.Value := 0 ;

          Connection.qryBusca2.Active := False ;
          Connection.qryBusca2.SQL.Clear ;
          Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad) as Instalado From bitacoradeactividades b ' +
                                        'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.dIdFecha >= :FechaI And b.dIdFecha <= :FechaF And b.sWbs = :Wbs And b.sNumeroActividad = :Actividad ' +
                                        'Group By b.sWbs, b.sNumeroActividad') ;
          Connection.QryBusca2.Params.ParamByName('Contrato').DataType  := ftString ;
          Connection.QryBusca2.Params.ParamByName('Contrato').Value     := global_Contrato ;
          Connection.QryBusca2.Params.ParamByName('Orden').DataType     := ftString ;
          Connection.QryBusca2.Params.ParamByName('Orden').Value        := tsNumeroOrden.Text ;
          Connection.QryBusca2.Params.ParamByName('FechaI').DataType    := ftDate ;
          Connection.QryBusca2.Params.ParamByName('FechaI').Value       := tdFechaInicial.Date ;
          Connection.QryBusca2.Params.ParamByName('FechaF').DataType    := ftDate ;
          Connection.QryBusca2.Params.ParamByName('FechaF').Value       := tdFechaFinal.Date ;
          Connection.QryBusca2.Params.ParamByName('Wbs').DataType       := ftString ;
          Connection.QryBusca2.Params.ParamByName('Wbs').Value          := ActividadesxOrden.FieldValues['sWbs'] ;
          Connection.QryBusca2.Params.ParamByName('Actividad').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Actividad').Value    := ActividadesxOrden.FieldValues['sNumeroActividad'] ;
          Connection.qryBusca2.Open ;
          If Connection.qryBusca2.RecordCount > 0 then
               ActividadesxOrdendCantidadPeriodo.Value := Connection.qryBusca2.FieldValues['Instalado']
          Else
               ActividadesxOrdendCantidadPeriodo.Value := 0 ;


         ActividadesxOrdendAcumulado.Value := ActividadesxOrdendCantidadPeriodo.Value + ActividadesxOrdendAcumuladoAnterior.Value ;
         ActividadesxOrdendTotal.Value := ActividadesxOrdendCantidadPeriodo.Value * ActividadesxOrdendVentaMN.Value ;
         ActividadesxOrdendTotalAcumulado.Value := ActividadesxOrdendAcumulado.Value * ActividadesxOrdendVentaMN.Value ;
     End
     Else
     Begin
         // Es Paquete ...

          Connection.qryBusca2.Active := False ;
          Connection.qryBusca2.SQL.Clear ;
          Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad * a.dVentaMN) as dTotal From bitacoradeactividades b ' +
                                       'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sNumeroOrden = b.sNumeroOrden and a.sWbs = b.sWbs And a.sNumeroActividad = b.sNumeroActividad and a.sIdConvenio = :Convenio) ' +
                                       'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.dIdFecha >= :FechaI and b.dIdFecha <= :FechaF And b.sWbs Like :Wbs ' +
                                       'Group By b.sContrato') ;
          Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Contrato').Value    := global_Contrato ;
          Connection.QryBusca2.Params.ParamByName('Orden').DataType    := ftString ;
          Connection.QryBusca2.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text ;
          Connection.QryBusca2.Params.ParamByName('convenio').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('convenio').Value    := global_convenio ;
          Connection.QryBusca2.Params.ParamByName('FechaI').DataType   := ftDate ;
          Connection.QryBusca2.Params.ParamByName('FechaI').Value      := tdFechaInicial.Date ;
          Connection.QryBusca2.Params.ParamByName('FechaF').DataType   := ftDate ;
          Connection.QryBusca2.Params.ParamByName('FechaF').Value      := tdFechaFinal.Date ;
          Connection.QryBusca2.Params.ParamByName('Wbs').DataType      := ftString ;
          Connection.QryBusca2.Params.ParamByName('Wbs').Value         := Trim(ActividadesxOrden.FieldValues['sWbs']) + '.%' ;
          Connection.qryBusca2.Open ;
          If Connection.qryBusca2.RecordCount > 0 then
               ActividadesxOrdendTotal.Value := connection.QryBusca2.FieldValues['dTotal']
          Else
               ActividadesxOrdendTotal.Value := 0 ;

          Connection.qryBusca2.Active := False ;
          Connection.qryBusca2.SQL.Clear ;
          Connection.qryBusca2.SQL.Add('Select Sum(b.dCantidad * a.dVentaMN) as dTotal From bitacoradeactividades b ' +
                                       'inner join actividadesxorden a on (a.sContrato = b.sContrato and a.sNumeroOrden = b.sNumeroOrden and a.sWbs = b.sWbs And a.sNumeroActividad = b.sNumeroActividad and a.sIdConvenio = :Convenio) ' +
                                       'Where b.sContrato = :Contrato And b.sNumeroOrden = :Orden And b.dIdFecha <= :Fecha And b.sWbs Like :Wbs ' +
                                       'Group By b.sContrato') ;
          Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('Contrato').Value    := global_Contrato ;
          Connection.QryBusca2.Params.ParamByName('Orden').DataType    := ftString ;
          Connection.QryBusca2.Params.ParamByName('Orden').Value       := tsNumeroOrden.Text ;
          Connection.QryBusca2.Params.ParamByName('convenio').DataType := ftString ;
          Connection.QryBusca2.Params.ParamByName('convenio').Value    := global_convenio ;
          Connection.QryBusca2.Params.ParamByName('Fecha').DataType    := ftDate ;
          Connection.QryBusca2.Params.ParamByName('Fecha').Value       := tdFechaFinal.Date ;
          Connection.QryBusca2.Params.ParamByName('Wbs').DataType      := ftString ;
          Connection.QryBusca2.Params.ParamByName('Wbs').Value         := Trim(ActividadesxOrden.FieldValues['sWbs']) + '.%' ;
          Connection.qryBusca2.Open ;
          If Connection.qryBusca2.RecordCount > 0 then
               ActividadesxOrdendTotalAcumulado.Value := connection.QryBusca2.FieldValues['dTotal']
          Else
               ActividadesxOrdendTotalAcumulado.Value := 0 ;

     End ;

 except
        on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_detalledeinstalacion', 'En el proceso actividadesxordencalcfields', 0);
        end;
 end;
end;

procedure TfrmDetalledeInstalacion.ActualizaFirmas( dFecha: TDateTime );
Begin
  try
    sSuperIntendente := '' ;
    sSupervisor := '' ;
    sPuestoSuperintendente := '' ;
    sPuestoSupervisor := '' ;
    connection.qryBusca2.Active := False ;
    connection.qryBusca2.SQL.Clear ;
    connection.qryBusca2.SQL.Add('Select * from firmas where sContrato = :contrato and dIdFecha = :fecha') ;
    Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
    Connection.QryBusca2.Params.ParamByName('fecha').DataType := ftDate ;
    Connection.QryBusca2.Params.ParamByName('fecha').Value := dFecha ;
    connection.qryBusca2.Open ;
    If connection.qryBusca2.RecordCount > 0 then
    Begin
        sSuperintendente := connection.qryBusca2.FieldValues['sFirmante1'] ;
        sSupervisor := connection.qryBusca2.FieldValues['sFirmante3'] ;
        sPuestoSuperintendente := connection.qryBusca2.FieldValues['sPuesto1'] ;
        sPuestoSupervisor := connection.qryBusca2.FieldValues['sPuesto3'] ;
    End
    Else
    Begin
        connection.qryBusca2.Active := False ;
        connection.qryBusca2.SQL.Clear ;
        connection.qryBusca2.SQL.Add('Select * from firmas where sContrato = :contrato and dIdFecha <= :fecha Order By dIdFecha DESC') ;
        Connection.QryBusca2.Params.ParamByName('Contrato').DataType := ftString ;
        Connection.QryBusca2.Params.ParamByName('Contrato').Value := global_contrato ;
        Connection.QryBusca2.Params.ParamByName('fecha').DataType := ftDate ;
        Connection.QryBusca2.Params.ParamByName('fecha').Value := dFecha ;
        connection.qryBusca2.Open ;
        If connection.qryBusca2.RecordCount > 0 then
        Begin
            sSuperintendente := connection.qryBusca2.FieldValues['sFirmante1'] ;
            sSupervisor := connection.qryBusca2.FieldValues['sFirmante3'] ;
            sPuestoSuperintendente := connection.qryBusca2.FieldValues['sPuesto1'] ;
            sPuestoSupervisor := connection.qryBusca2.FieldValues['sPuesto3'] ;
        End
    End ;
  except
        on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_detalledeinstalacion', 'En el proceso Actualizar firmas', 0);
        end;
  end;
End ;

{
procedure TfrmDetalledeInstalacion.btnReport1Click(Sender: TObject);
begin
    messageDLg('No se puede Imprimir!, Existe una version mejorada del Reporde de Detalle de Movimientos en el Modulo de Reporte Periodo (Icono de la Camra), Checalo!', mtInformation, [mbOk], 0);
    exit;
    Connection.qryBusca.Active := False ;
    Connection.qryBusca.SQL.Clear ;
    Connection.qryBusca.SQL.Add('Select dFechaInicio, dFechaFinal, dMontoMN, dMontoDLL From convenios Where sContrato = :Contrato And sIdConvenio = :Convenio') ;
    Connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
    Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato ;
    Connection.QryBusca.Params.ParamByName('Convenio').DataType := ftString ;
    Connection.QryBusca.Params.ParamByName('Convenio').Value := global_convenio ;
    Connection.qryBusca.Open ;

    Detalle.Active := False ;
    Detalle.Params.ParamByName('Contrato').DataType := ftString ;
    Detalle.Params.ParamByName('Contrato').Value := global_contrato ;
    Detalle.Params.ParamByName('Convenio').DataType := ftString ;
    Detalle.Params.ParamByName('Convenio').Value := global_convenio ;
    Detalle.Params.ParamByName('Orden').DataType := ftString ;
    Detalle.Params.ParamByName('Orden').Value := tsNumeroOrden.Text ;
    Detalle.Params.ParamByName('Inicio').DataType := ftDate ;
    Detalle.Params.ParamByName('Inicio').Value := tdFechaInicial.Date ;
    Detalle.Params.ParamByName('Final').DataType := ftDate ;
    Detalle.Params.ParamByName('Final').Value := tdFechaFinal.Date ;
    Detalle.Open ;

    rDiarioFirmas (global_contrato, tsNumeroOrden.Text, tdFechaFinal.Date , frmDetalledeInstalacion ) ;
    frxDetalle.PreviewOptions.MDIChild := True ;
    frxDetalle.PreviewOptions.Modal := False ;
    frxDetalle.PreviewOptions.Maximized := lCheckMaximized () ;
    frxDetalle.PreviewOptions.ShowCaptions := False ;
    frxDetalle.Previewoptions.ZoomMode := zmPageWidth ;
    frxDetalle.LoadFromFile (global_files + 'DetalleInstalacion.fr3') ;
    frxDetalle.ShowReport ;

end;  }

procedure TfrmDetalledeInstalacion.frxDetalleGetValue(
  const VarName: String; var Value: Variant);
begin
  If CompareText(VarName, 'SUPERINTENDENTE') = 0 then
      Value := sSuperIntendente ;
  If CompareText(VarName, 'SUPERVISOR') = 0 then
      Value := sSupervisor ;
  If CompareText(VarName, 'PUESTO_SUPERINTENDENTE') = 0 then
      Value := sPuestoSuperIntendente ;
  If CompareText(VarName, 'PUESTO_SUPERVISOR') = 0 then
      Value := sPuestoSupervisor ;
  If CompareText(VarName, 'PERIODO_CONTRATO') = 0 then
      Value := 'DEL ' + DateToStr ( Connection.qryBusca.FieldValues ['dFechaInicio']) + ' AL ' + DateToStr ( Connection.qryBusca.FieldValues ['dFechaFinal']) ;
  If CompareText(VarName, 'PERIODO') = 0 then
      Value := 'Del ' + FormatDateTime('d "de" mmmm "del" yyyy' , tdFechaInicial.Date) + ' al ' + FormatDateTime('d "de" mmmm "del" yyyy' , tdFechaFinal.Date);
  If CompareText(VarName, 'MONTOMN') = 0 then
      Value := Connection.qryBusca.FieldValues ['dMontoMN'] ;
  If CompareText(VarName, 'MONTODLL') = 0 then
      Value := Connection.qryBusca.FieldValues ['dMontoDLL'] ;
  If CompareText(VarName, 'INICIO') = 0 then
      Value := DateToStr (tdFechaInicial.Date) ;
  If CompareText(VarName, 'FINAL') = 0 then
      Value := DateToStr (tdFechaFinal.Date) ;
  If CompareText(VarName, 'ORDEN') = 0 then
      Value := tsNumeroOrden.Text  ;
  If CompareText(VarName, 'DESCRIPCION') = 0 then
      Value := Connection.contrato.FieldValues['mDescripcion'] ;
end;

procedure TfrmDetalledeInstalacion.formatoEncabezado;
begin
  Excel.Selection.MergeCells := False;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.Font.Size := 12;
  Excel.Selection.Font.Bold := False;
  Excel.Selection.Font.Name := 'Calibri';
end;

procedure TfrmDetalledeInstalacion.AcumuladoDePartidas;
var
  iFila : Integer;
  qrAcumulado : TZQuery;
  sInicioTabla : string;
  cambioDeColor : Boolean;

  {$REGION 'PROCEDIMIENTOS PARA FORMATO'}

  procedure contornoPunteado;
  begin
    Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeLeft].ColorIndex := xlAutomatic;
    Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
    Excel.Selection.Borders[xlEdgeLeft].Weight := xlThin;

    Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeBottom].ColorIndex := xlAutomatic;
    Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
    Excel.Selection.Borders[xlEdgeBottom].Weight := xlHairline;

    Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeRight].ColorIndex := xlAutomatic;
    Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
    Excel.Selection.Borders[xlEdgeRight].Weight := xlThin;
  end;

  procedure EstablecerContornos;
  begin
    Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeLeft].Weight := xlThin;
    Excel.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeTop].Weight := xlThin;
    Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;
    Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
    Excel.Selection.Borders[xlEdgeRight].Weight := xlThin;
  end;

  procedure DarFormato(combinar : Boolean ; alinear : string ; Negritas : Boolean ; contornos : Boolean);
  begin
    if combinar then
    begin
      Excel.Selection.MergeCells := True;
    end;

    if alinear = 'centro' then
    begin
      Excel.Selection.HorizontalAlignment := xlCenter;
    end

    else if alinear = 'der' then
    begin
      Excel.Selection.HorizontalAlignment := xlRight;
    end

    else if alinear = 'izq' then
    begin
      Excel.Selection.HorizontalAlignment  := xlLeft;
    end;

    if Negritas then
    begin
      Excel.Selection.Font.Bold := True;
    end;

    if contornos then
    begin
      EstablecerContornos;
    end;

    Excel.Selection.VerticalAlignment := xlCenter;
    Excel.Selection.Font.Size := 7;
  end;

  {$ENDREGION}

begin

  {$REGION 'CREAR EXCEL'}
  try
    Excel := CreateOleObject('Excel.Application');
  except
    FreeAndNil(Excel);
    showmessage('No es posible generar el ambiente de EXCEL, informe de esto al administrador del sistema.');
    Exit;
  end;

  Excel.Visible := True;
  Excel.DisplayAlerts := False;
  Excel.ScreenUpdating := True;
  Libro := Excel.Workbooks.Add;

  while Libro.Sheets.Count > 1 do
      Excel.ActiveWindow.SelectedSheets.Delete;

  Hoja := Libro.Sheets[1];
  Hoja.Select;

  try
     Hoja.Name := 'ACUMULADO DE PARTIDAS';
  Except
    ;
  end;

  {$ENDREGION}

  {$REGION 'FORMATO'}

  Excel.Columns['A:A'].ColumnWidth := 20;
  Excel.Columns['B:B'].ColumnWidth := 20;
  Excel.Columns['C:C'].ColumnWidth := 8;
  Excel.Columns['D:D'].ColumnWidth := 8;
  Excel.Columns['E:E'].ColumnWidth := 8;
  Excel.Columns['F:F'].ColumnWidth := 8;
  Excel.Columns['G:G'].ColumnWidth := 8;
  Excel.Columns['H:H'].ColumnWidth := 9;
  Excel.Columns['I:I'].ColumnWidth := 8;
  Excel.Columns['J:J'].ColumnWidth := 85;

  iFila := 7;
  cambioDeColor := True;

  {$ENDREGION}

  {$REGION 'ENCABEZADO'}

  Excel.Range['A2:J3'].Select;
  DarFormato(True, 'centro', False, False);
  Excel.Selection.Font.Size           := 20;
  Excel.Selection.Value               := 'A C U M U L A D O    D E    P A R T I D A S';

  Excel.Rows[6].RowHeight := 30;
  Excel.Rows[4].RowHeight := 0;

  Excel.Range['A6:J6'].Select;
  Excel.Selection.Font.Size            := 11;
  Excel.Selection.Font.Bold            := True;
  Excel.Selection.Font.Name            := 'Arial';
  Excel.Selection.HorizontalAlignment  := xlCenter;
  Excel.Selection.VerticalAlignment    := xlCenter;
  Excel.Selection.Interior.ColorIndex  := 33;

  Excel.Range['A6'].Select;
  Excel.Selection.Value := 'FECHA';
  EstablecerContornos;

  Excel.Range['B6'].Select;
  Excel.Selection.Value := 'FOLIO';
  EstablecerContornos;

  Excel.Range['C6'].Select;
  Excel.Selection.Value := 'PDA';
  EstablecerContornos;

  Excel.Range['D6'].Select;
  Excel.Selection.Value := 'CLAS';
  EstablecerContornos;

  Excel.Range['E6'].Select;
  Excel.Selection.Value := 'INI.';
  EstablecerContornos;

  Excel.Range['F6'].Select;
  Excel.Selection.Value := 'TERM.';
  EstablecerContornos;

  Excel.Range['G6'].Select;
  Excel.Selection.Value := 'ANT';
  EstablecerContornos;

  Excel.Range['H6'].Select;
  Excel.Selection.Value := 'AVANCE';
  EstablecerContornos;

  Excel.Range['I6'].Select;
  Excel.Selection.Value := 'ACUM';
  EstablecerContornos;

  Excel.Range['J6'].Select;
  Excel.Selection.Value := 'DESCRIPCIÓN DEL TRABAJO';
  EstablecerContornos;

  {$ENDREGION}

  {$REGION 'CONSULTA'}

  qrAcumulado := TZQuery.Create(nil);
  qrAcumulado.Connection := connection.zConnection;
  qrAcumulado.Active := False;
  qrAcumulado.SQL.Clear;
  qrAcumulado.SQL.Add('SELECT dIdFecha, sNumeroOrden, sNumeroActividad, sIdClasificacion, sHoraInicio, sHoraFinal, '+
                     '(SELECT IFNULL(sum(b2.dAvance),0) '+
                     'FROM bitacoradeactividades b2 '+
                     'WHERE b2.sContrato= :contrato AND b2.sNumeroOrden=b1.sNumeroOrden AND b2.sNumeroActividad=b1.sNumeroActividad '+
                     'AND b2.sIdClasificacion=b1.sIdClasificacion AND b1.sWbs=b2.sWbs '+
                     'AND (b2.dIdFecha<b1.dIdFecha OR (b2.dIdFecha=b1.dIdFecha AND b2.sHoraInicio<b1.sHoraInicio)) '+
                     'AND b2.sIdTipoMovimiento = "ED") as dAnt, '+
                     'dAvance, mDescripcion '+
                     'FROM bitacoradeactividades b1 '+
                     'WHERE (b1.dIdFecha BETWEEN :fechaInicio AND :fechaFinal) and b1.sIdTipoMovimiento = "ED" '+
                     'ORDER BY b1.dIdFecha, b1.sNumeroOrden, b1.sNumeroActividad, b1.sHoraInicio');

  qrAcumulado.ParamByName('contrato').AsString    := global_contrato;
  qrAcumulado.ParamByName('fechaInicio').AsString := FormatearFecha(FechaInicio.Date);
  qrAcumulado.ParamByName('fechaFinal').AsString  := FormatearFecha(FechaTermino.Date);
  qrAcumulado.Open;

  if qrAcumulado.RecordCount > 0 then
  begin
    qrAcumulado.First;
    while not qrAcumulado.Eof do
    begin
      Excel.Range['A' + IntToStr(iFila)].Select;
      Excel.Selection.Value := Char(39) + qrAcumulado.FieldByName('dIdFecha').AsString;

      Excel.Range['B' + IntToStr(iFila)].Select;
      Excel.Selection.Value := qrAcumulado.FieldByName('sNumeroOrden').AsString;

      Excel.Range['C' + IntToStr(iFila)].Select;
      Excel.Selection.Value := qrAcumulado.FieldByName('sNumeroActividad').AsString;

      Excel.Range['D' + IntToStr(iFila)].Select;
      Excel.Selection.Value := qrAcumulado.FieldByName('sIdClasificacion').AsString;

      Excel.Range['E' + IntToStr(iFila)].Select;
      Excel.Selection.Value := Char(39) +  qrAcumulado.FieldByName('sHoraInicio').AsString;

      Excel.Range['F' + IntToStr(iFila)].Select;
      Excel.Selection.Value := Char(39) +  qrAcumulado.FieldByName('sHoraFinal').AsString;

      Excel.Range['G' + IntToStr(iFila)].Select;
      Excel.Selection.Value := qrAcumulado.FieldByName('dAnt').AsFloat;
      Excel.Selection.NumberFormat := '00.00%';

      Excel.Range['H' + IntToStr(iFila)].Select;
      Excel.Selection.Value := qrAcumulado.FieldByName('dAvance').AsString;
      Excel.Selection.NumberFormat := '00.00%';

      Excel.Range['I' + IntToStr(iFila)].Select;
      Excel.Selection.Formula := '=G' + IntToStr(iFila) + '+H' + IntToStr(iFila);
      Excel.Selection.NumberFormat := '00.00%';

      Excel.Range['J' + IntToStr(iFila)].Select;
      Excel.Selection.Value := qrAcumulado.FieldByName('mDescripcion').AsString;
      Excel.Selection.WrapText := False;      

      if cambioDeColor then
      begin
        Excel.Range['A' + IntToStr(iFila) + ':J' + IntToStr(iFila)].Select;
        Excel.Selection.Interior.ColorIndex := 37;
        cambioDeColor := False;
      end
      else
      begin
        cambioDeColor := True;
      end;
      qrAcumulado.Next;
      Inc(iFila);
    end;

    Excel.Range['A7:I' + IntToStr(iFila)].Select;
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.VerticalAlignment   := xlCenter;
    Excel.Selection.Font.Name := 'Arial Narrow';
    Excel.Selection.Font.Size := 11;

    Excel.Range['J7:J' + IntToStr(iFila)].Select;
    Excel.Selection.HorizontalAlignment := xlLeft;
    Excel.Selection.VerticalAlignment := xlCenter;

    Excel.ActiveWindow.DisplayGridlines := False;

  end;

  {$ENDREGION}

  qrAcumulado.Destroy;
  ShowMessage('Proceso terminado con exito.');

end;

function TfrmDetalledeInstalacion.FormatearFecha(Fecha: TDate) : string;
var
   DiaI, MesI, AnoI: Word;
   i : Integer;
   valor : string;
begin
  DecodeDate( Fecha, AnoI, MesI, DiaI);
  valor := '';
  valor := IntToStr(AnoI) + '/';

  if MesI < 9 then
  begin
    valor := valor + '0' + IntToStr(MesI) + '/';
  end
  else
  begin
    valor := valor + IntToStr(MesI) + '/';
  end;

  if DiaI < 9 then
  begin
    valor := valor + '0' + IntToStr(DiaI);
  end
  else
  begin
    valor := valor + IntToStr(DiaI);
  end;
  Result := valor;
end;

function TfrmDetalledeInstalacion.VolverAInicioDeMes(Fecha: TDate) : TDate;
var
   DiaI, MesI, AnoI: Word;
   i : Integer;
   valor : string;
begin
  DecodeDate( Fecha, AnoI, MesI, DiaI);
  DiaI := 1;
  valor := '0' + IntToStr(DiaI);
  if MesI <= 9 then
  begin
    Valor := Valor + '/0'
  end
  else
  begin
    Valor := valor + '/'
  end;
  Valor := valor + IntToStr(MesI) + '/' + IntToStr(AnoI);
  Result := StrToDate(valor);
end;

end.
