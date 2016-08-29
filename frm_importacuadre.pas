unit frm_importacuadre;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, cxGraphics, frm_connection, global, UnitExcel, ComObj,
  dateutils,


  cxControls, cxLookAndFeels, cxLookAndFeelPainters,
  cxContainer, cxEdit, dxSkinsCore, dxSkinBlack, dxSkinBlue, dxSkinBlueprint,
  dxSkinCaramel, dxSkinCoffee, dxSkinDarkRoom, dxSkinDarkSide,
  dxSkinDevExpressDarkStyle, dxSkinDevExpressStyle, dxSkinFoggy,
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
  cxGroupBox, Menus, StdCtrls, cxButtons, ImgList, cxCheckBox, cxCheckListBox,
  DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, cxProgressBar, cxTextEdit,
  JvComponentBase, cxMemo,
  cxListBox, ExtCtrls;

type
  TfrmImportaCuadre = class(TForm)
    grpPlantilla: TcxGroupBox;
    btnPlantilla: TcxButton;
    img48: TcxImageList;
    cxCheckBox1: TcxCheckBox;
    chkPersonal: TcxCheckBox;
    chkEquipo: TcxCheckBox;
    chklstFolios: TcxCheckListBox;
    qrFolios: TZQuery;
    qrMoePersonal: TZQuery;
    qrMoeEquipo: TZQuery;
    qrMoe: TZQuery;
    grpImportar: TcxGroupBox;
    chkImportarPersonal: TcxCheckBox;
    chkReemplazar: TcxCheckBox;
    cxGroupBox1: TcxGroupBox;
    prgImporta: TcxProgressBar;
    grpArchivo: TcxGroupBox;
    txtFile: TcxMemo;
    btnFile: TcxButton;
    dlgOpenExcel: TOpenDialog;
    img32: TcxImageList;
    btnImportar: TcxButton;
    img80: TcxImageList;
    dlgSaveExcel: TSaveDialog;
    lstLog: TcxListBox;
    tmrExcel: TTimer;
    qrActividades: TZQuery;
    qrMovtos: TZQuery;
    chkImportarEquipo: TcxCheckBox;
    qrPernoctan: TZQuery;
    qrPlataformas: TZQuery;
    qrPernoctas: TZQuery;
    Label1: TLabel;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnPlantillaClick(Sender: TObject);
    procedure btnFileClick(Sender: TObject);
    procedure btnImportarClick(Sender: TObject);
  private
    { Private declarations }
    meses: array[1..12] of string;

    procedure GenerarPlantilla(Personal, Equipo : Boolean; Folios : TcxCheckListBox);
    procedure ImportaCuadre(Personal, Equipo : Boolean; Folios : TcxCheckListBox);
    procedure GeneraCuadreAnterior(Fecha : TDateTime; Personal, Equipo : Boolean; Folios : TcxCheckListBox );
  public
    { Public declarations }
    param_contrato : string;
    param_fecha : TDateTime;
  end;

var
  frmImportaCuadre: TfrmImportaCuadre;

implementation

{$R *.dfm}

procedure TfrmImportaCuadre.btnFileClick(Sender: TObject);
begin
  if dlgOpenExcel.Execute then
    txtFile.Text := dlgOpenExcel.FileName;
end;

procedure TfrmImportaCuadre.btnImportarClick(Sender: TObject);
begin
  if Trim(txtFile.Lines.Text) = '' then
  begin
    MessageDlg('Especifique un archivo valido', mtInformation, [mbOk], 0);
    Exit;
  end;

  if not FileExists( Trim( txtFile.Text ) ) then
  begin
    MessageDlg('El archivo: '+Trim( txtFile.Text ) + ', No existe', mtInformation, [mbOk], 0);
    exit;
  end;

  Importacuadre(chkImportarPersonal.Checked, chkImportarEquipo.Checked, chklstFolios);

end;

procedure TfrmImportaCuadre.btnPlantillaClick(Sender: TObject);
begin
  dlgSaveExcel.FileName := 'Cuadre '+
                           param_contrato +' - ' +
                           inttostr( dayof( param_fecha ) ) + ' ' +
                           meses[ monthof( param_fecha ) ] + ' ' +
                           inttostr( yearof( param_fecha ) );
  if dlgSaveExcel.Execute then
  begin
    txtFile.Text := dlgSaveExcel.FileName;
    if chkPersonal.Checked or chkEquipo.Checked then
      GenerarPlantilla(chkPersonal.Checked, chkEquipo.Checked, chklstFolios);
  end;
end;

procedure TfrmImportaCuadre.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree
end;

procedure TfrmImportaCuadre.FormCreate(Sender: TObject);
begin
  chklstfolios.Items.Clear;
end;

procedure TfrmImportaCuadre.FormShow(Sender: TObject);
var
  cxCheck : TcxCheckListBoxItem;
begin
  meses[1]  := 'Enero';
  meses[2]  := 'Febrero';
  meses[3]  := 'Marzo';
  meses[4]  := 'Abril';
  meses[5]  := 'Mayo';
  meses[6]  := 'Junio';
  meses[7]  := 'Julio';
  meses[8]  := 'Agosto';
  meses[9]  := 'Septiembre';
  meses[10] := 'Octubre';
  meses[11] := 'Noviembre';
  meses[12] := 'Diciembre';

  {CONSULTA PERNOCTAS Y PLATAFORMAS}
  {$REGION 'PERNOCTAS Y PLATAFORMAS'}

  qrPernoctan.Active := false;
  qrPernoctan.open;

  qrPlataformas.Active := false;
  qrPlataformas.Open;

  qrPernoctas.Active := False;
  qrPernoctas.open;

  {$ENDREGION}

  {CONSULTA Y GENERA LOS FOLIOS EN chklstFolios}
  {$REGION 'CONSULTA FOLIOS'}

  with qrFolios do
  begin
    active := false;
    sql.Text := 'select ots.sNumeroOrden, ots.iJornada '+
                'from ordenesdetrabajo ots '+
                'where sContrato = :contrato';
    parambyname('contrato').asstring := param_contrato;
    open;

    if recordcount > 0 then
    begin
      chklstFolios.Items.Clear;
      first;
      while not eof do
      begin
        cxCheck := chklstFolios.Items.Add;
        cxCheck.Text := fieldbyname('snumeroorden').asstring;
        next;
      end;
    end;
  end;

  {$ENDREGION}

  {CONSULTA MOE, PERSONAL Y EQUIPO VIGENTE}
  {$REGION 'MOE VIGENTE, PERSONAL Y EQUIPO SOLICITADO'}

  try

    {$REGION 'MOE'}
    
    with qrMoe do
    begin
      active := false;
      sql.Text := 'select iIdMoe '+
                  'from moe '+
                  'where sContrato = :contrato '+
                  'and dIdFecha <= :fecha';
      parambyname('contrato').asstring := param_contrato;
      parambyname('fecha').asdatetime := param_fecha;
      open;
    end;

    if qrMoe.RecordCount = 0 then
      raise exception.Create('No esta dado de alta un moe vigente hasta la fecha '+datetostr(param_fecha)+' en la OT: '+ param_contrato+
                             'verifique su información.');
    {$ENDREGION}

    {$REGION 'PERSONAL'}

    with qrMoePersonal do
    begin
      active := false;
      sql.Text := 'select p.sIdPersonal '+
                  'from moerecursos mr '+

                  'inner join moe m '+
                    'on ( m.iIdMoe = mr.iIdMoe ) '+

                  'inner join personal p '+
                    'on ( p.sContrato = :contratoBarco and p.sIdPersonal = mr.sIdRecurso ) '+

                  'inner join tiposdepersonal tp '+
                    'on ( tp.sIdTipoPersonal = p.sIdTipoPersonal ) '+

                  'where m.sContrato = :contrato '+
                  'and m.iIdMoe = :idmoe '+
                  'and mr.eTipoRecurso = "Personal" ';
      parambyname('contratoBarco').asstring := global_contrato_barco;
      parambyname('contrato').asstring      := param_contrato;
      parambyname('idmoe').asinteger        := qrMoe.FieldByName('iIdMoe').AsInteger;
      open;
    end;

    {$ENDREGION}

    {$REGION 'EQUIPO'}

    with qrMoeEquipo do
    begin
      active := false;
      sql.Text := 'select e.sIdEquipo '+
                 'from moerecursos mr '+
                 'inner join moe m '+
                   'on ( m.iIdMoe = mr.iIdMoe ) '+

                 'inner join equipos e '+
                   'on ( e.sContrato = :contratoBarco and e.sIdEquipo = mr.sIdRecurso ) '+

                 'inner join tiposdeequipo te '+
                   'on ( te.sIdTipoEquipo = e.sIdTipoEquipo ) '+

                 'where m.sContrato = :contrato '+
                 'and m.iIdMoe = :idmoe '+
                 'and mr.eTipoRecurso = "Equipo" ';
      parambyname('contratoBarco').asstring := global_contrato_barco;
      parambyname('contrato').asstring      := param_contrato;
      parambyname('idmoe').asinteger        := qrMoe.FieldByName('iIdMoe').AsInteger;
      open;
    end;

    {$ENDREGION}

  except
    on e:exception do
    begin
      MessageDlg('No se puede iniciar el modulo de importación debido lo siguiente: '+#10+#10+ e.Message, mtInformation, [mbOk], 0);
      Close;
    end;
  end;

  {$ENDREGION}

  {CONSULTA MOVIMIENTOS}
  {$REGION 'MOVIMIENTOS'}

  with qrMovtos do
  begin
    active := false;
    sql.text := 'select * from tiposdemovimiento where sClasificacion <> "movimiento de barco" and sContrato = :contrato order by iorden';
    parambyname('contrato').asstring := param_contrato;
    open;
  end;

  {$ENDREGION}

end;

procedure TfrmImportaCuadre.GenerarPlantilla(Personal: Boolean; Equipo: Boolean; Folios: TcxCheckListBox);
var
  Excel,
  Libro,
  Hoja,
  rango : Variant;

  iFila,
  iColumna,
  iFolios,
  iHoja : Integer;

const
  COLUMNAS : array[0..11] of string = (
    'ACTIVIDAD',
    'INICIO',
    'FIN',
    'PARTIDA',
    'CANT.',
    'DURACION',
    'H.H',
    'TIPO OBRA',
    'PERNOCTA',
    'PLATAFORMA',
    'TIPO PERNOCTA',
    'TAREA'
  );

  FORMATOS : array[0..11] of string = (
    '@',
    '@',
    '@',
    '@',
    '0',
    '0.00000000',
    '0.00',
    '@',
    '@',
    '@',
    '@',
    '@'
  );
begin

  {INICIALIZAR EXCEL}
  {$REGION 'EXCEL'}

  try
    excel := CreateOleObject('Excel.Application');
    excel.visible := True;
    excel.DisplayAlerts:= False;

    libro := excel.workbooks.add;

    while libro.sheets.count > 1 do
      libro.activesheet.delete;

    hoja := libro.sheets[1].select;
//    excel.range['A1'].value := 'Esperar';
  except
    on e:exception do
    begin
      MessageDlg('No se puede iniciar excel, verifique tener instalada la Suite de Micrsoft Office', mtInformation, [mbOk], 0);
    end;
  end;

  ifila := 2;

  {$ENDREGION}

  {$REGION 'CREA TABLAS'}

  if Personal and Equipo then
    libro.sheets.add;

  if Personal then
    Libro.sheets[1].name := 'PERSONAL';
  if Equipo and Personal then
    Libro.sheets[2].name := 'EQUIPOS';
  if Equipo and (not Personal) then
    Libro.sheets[1].name := 'EQUIPOS';

  for iHoja := 1 to Libro.Sheets.Count do
  begin
    iFila := 2;

    libro.sheets[iHoja].select;

    if Personal or Equipo then
    begin
      for iFolios := 0 to Folios.Items.Count - 1 do
      begin
        if not Folios.Items.Items[iFolios].Checked then
          continue;

        rango := excel.range['B'+inttostr(ifila)];
        rango.value := 'FOLIO';
        rango.interior.colorindex := 15;

        rango := excel.range['F'+inttostr(ifila)];
        rango.value := 'INICIO';
        rango.interior.colorindex := 15;

        rango := excel.range['H'+inttostr(ifila)];
        rango.value := 'TERMINO';
        rango.interior.colorindex := 15;


        rango := excel.range['C'+inttostr(ifila)+':D'+inttostr(ifila)];
        rango.numberformat := '@';
        rango.mergecells := true;
        rango.value := Folios.Items.Items[iFolios].Text;
        excel.range['B'+inttostr(ifila)+':I'+inttostr(ifila)].borders.color := clBlack;

        inc(ifila);
        for iColumna := 0 to length(COLUMNAS) - 1 do
        begin
          excel.range[columnanombre(icolumna+2)+inttostr(ifila)].value := COLUMNAS[icolumna];
        end;

        rango := excel.range['B'+inttostr(ifila)+':'+columnanombre(icolumna+1)+inttostr(ifila)];
        rango.borders.color := clBlack;
        rango.interior.colorindex := 15;

        inc(ifila);
        for iColumna := 0 to length(FORMATOS) - 1 do
          excel.range[columnanombre(icolumna+2)+inttostr(ifila)].numberformat := FORMATOS[icolumna];
//        excel.range['B'+inttostr(ifila)+':'+columnanombre(icolumna+1)+inttostr(ifila)].borders.color := clBlack;

        excel.range['G'+inttostr(ifila)].formula := '=D'+inttostr(ifila)+'-C'+inttostr(ifila);
        if iHoja = 1 then
          excel.range['H'+inttostr(ifila)].formula := '=F'+inttostr(ifila)+'* G'+inttostr(ifila) + '* 2'
        else
          excel.range['H'+inttostr(ifila)].formula := '=F'+inttostr(ifila)+'* G'+inttostr(ifila);

        Inc(ifila, 4);
      end;
    end;
  end;


  {$ENDREGION}

  Libro.SaveAs(trim(txtfile.Text));
//  TmrExcel.Enabled := True;
end;

procedure TfrmImportaCuadre.ImportaCuadre(Personal, Equipo : Boolean; Folios : TcxCheckListBox);
var
  Excel,
  Libro,
  Hoja,
  Rango : Variant;

  iRecorridos,
  iFila,
  iColumna,
  iFolios,
  iHoja,
  iLeer,
  iPestania,
  iVacias,
  iFilasP,
  iHojaInicio,
  iIdDiario,
  iFilaCab,
  iMaxPgr,
  iTarea,
  iRegistros,
  iJornada : Integer;

  sTipo,
  sSQL_Recurso,
  sFolio,
  sHInicio,
  sHTermino,
  sActividad,
  sRecurso,
  sPernocta,
  sPlataforma,
  sTipoPernocta,
  sObra,
  sId,
  sHPdaInicio,
  sHPdaFin : string;

  dCantidad,
  dCantidadHH : Double;
  
  bError : Boolean;
  qrRecurso,
  qrNotas,
  qrTObra,
  qrTiempoExtra : TZReadOnlyQuery;

  Function BuscaHoja(Nomb:String):Integer;
  var ResBuscaHoja,x:Integer;
  begin
    ResBuscaHoja := -1;
    try
      x := 1;
      while (x <= Libro.Sheets.Count) and (ResBuscaHoja = -1) do
      begin
        if Libro.Sheets[x].Visible then
          if LowerCase(trim(Libro.workSheets[x].Name)) = LowerCase(Nomb) then
            ResBuscaHoja := x;
        Inc(x);
      end;
    finally
      Result := ResBuscaHoja;
    end;
  end;


begin

  {$REGION 'VALIDA TIPO DE IMPORTACION'}
  iLeer := 1;

  if ( not personal ) and ( not equipo ) then
    iLeer := 0;

  if personal and equipo then
    iLeer := 2;

  if ( personal ) and ( not equipo ) then
    iLeer := 1;

  if ( not personal ) and ( equipo ) then
    iLeer := 2;

  bError := False;

  if iLeer = 0 then
  begin
    MessageDlg('No se ha especificado que desea importar', mtInformation, [mbOk], 0);
    exit;
  end;
  {$ENDREGION}

  {$REGION 'CREA EXCEL'}

  excel := CreateOleObject('Excel.Application');
  excel.visible := False;
  excel.DisplayAlerts:= False;

  excel.workbooks.open(trim(txtfile.text));
  libro := excel.workbooks[1];
  hoja := libro.sheets[1];
  hoja.select;

  {$ENDREGION}

  {$REGION 'CREA QUERY"S'}

  qrRecurso := TZReadOnlyQuery.Create(nil);
  qrRecurso.Connection := connection.zConnection;
  qrNotas := TZReadOnlyQuery.Create(nil);
  qrNotas.connection := connection.zconnection;
  qrTObra := TZReadOnlyQuery.Create(nil);
  qrTObra.Connection := connection.zConnection;
  qrRecurso.Active := false;
  qrTiempoExtra := TZReadOnlyQuery.Create(nil);
  qrTiempoExtra.connection := connection.zconnection;

  {$ENDREGION}
  try
    iRegistros := 0;
    iHojaInicio := 1;
    if Equipo then
      iHojaInicio := 2;

    if Personal then
      iHojaInicio := 1;

    {$REGION 'TIEMPO EXTRA DE PERSONAL'}

    qrTiempoExtra.Active := false;
    qrTiempoExtra.SQL.Text := 'select sIdPersonal from personal where sIdTipoPersonal = "EXT" ';
    qrTiempoExtra.Open;

    {$ENDREGION}
      
    for iHoja := iHojaInicio to iLeer do
    begin

      {$REGION 'TIPO DE RECURSO'}

      if iHoja = 1 then
      begin

        {$REGION 'CONSULTAS PERSONAL'}

        sTipo := 'PERSONAL';
        sId := 'sidpersonal';
        qrRecurso.SQL.Text := 'select mr.sIdRecurso, mr.sDescripcion '+
                        'from moerecursos mr '+
                        'inner join moe m '+
                          'on ( m.iIdMoe = mr.iIdMoe ) '+

                        'inner join personal p '+
                          'on ( p.sContrato = :contratoBarco and p.sIdPersonal = mr.sIdRecurso ) '+

                        'inner join tiposdepersonal tp '+
                          'on ( tp.sIdTipoPersonal = p.sIdTipoPersonal ) '+

                        'where m.sContrato = :contrato '+
                        'and m.iIdMoe = :idmoe '+
                        'and mr.eTipoRecurso = "Personal" ';

        qrTObra.Active := false;
        qrTObra.SQL.Text := 'select sIdTipoPersonal as sTipoObra from tiposdepersonal ';
        qrTObra.Open;

        {$ENDREGION}

      end;
      if iHoja = 2 then
      begin

        {$REGION 'CONSULTAS EQUIPO'}

        sTipo := 'EQUIPOS';
        sId := 'sidequipo';
        qrRecurso.SQL.Text := 'select mr.sIdRecurso, mr.sDescripcion '+
                 'from moerecursos mr '+
                 'inner join moe m '+
                   'on ( m.iIdMoe = mr.iIdMoe ) '+

                 'inner join equipos e '+
                   'on ( e.sContrato = :contratoBarco and e.sIdEquipo = mr.sIdRecurso ) '+

                 'inner join tiposdeequipo te '+
                   'on ( te.sIdTipoEquipo = e.sIdTipoEquipo ) '+

                 'where m.sContrato = :contrato '+
                 'and m.iIdMoe = :idmoe '+
                 'and mr.eTipoRecurso = "Equipo" ';

        qrTObra.Active := false;
        qrTObra.SQL.Text := 'select sIdTipoEquipo as sTipoObra from tiposdeequipo  ';
        qrTObra.Open;

        {$ENDREGION}
        
      end;

      qrRecurso.ParamByName('contratoBarco').asstring := global_contrato_barco;
      qrRecurso.ParamByName('contrato').asstring      := param_contrato;
      qrRecurso.ParamByName('idmoe').asinteger        := qrMoe.FieldByName('iidmoe').asinteger;
      qrRecurso.Open;

      {$ENDREGION}

      {$REGION 'BUSCA HOJAS VALIDAS'}

      Libro.Sheets[iHoja].Select;
      Hoja := Libro.Sheets[iHoja];
      iPestania := BuscaHoja(sTipo);
      if iPestania = -1 then
        raise exception.Create('No se encontro la pestaña correspondiente el cuadre de ' + lowercase( sTipo ));
      Hoja := Libro.Sheets[iPestania];
      Hoja.Select;

      ifila := 2;

      {$ENDREGION}

      {$REGION 'EN CASO DE QUERER REEMPLAZAR LA INFORMACIÓN EXISTENTE'}

       if chkReemplazar.Checked then
       begin
         connection.zCommand.Active := false;
         connection.zcommand.sql.Text := 'delete from bitacorade'+lowercase(stipo)+' where sContrato = :contrato and didfecha = :fecha ';
         connection.zCommand.parambyname('contrato').asstring := param_contrato;
         connection.zCommand.Parambyname('fecha').asdatetime := param_fecha;
         connection.zCommand.ExecSQL;
       end;

      {$ENDREGION}

      try
        label1.Visible := True;
        while (Trim(Excel.cells[ifila,3].text) <> '') and (iVacias < 10) do
        begin
          sFolio := Trim(Excel.cells[ifila,3].text);
          Inc(ifila,2);
          iFilaCab := iFila - 2;
          ifilasp := ifila;

          label1.Caption := 'Leyendo registros: '+ IntToStr(ifila);
          label1.Refresh;

          {$REGION 'VALIDA FOLIO'}

          if not qrFolios.Locate('sNumeroorden', sFolio,  []) then
          begin
            excel.range['C'+inttostr(ifila-2)].addcomment('Folio no valido');
            excel.range['C'+inttostr(ifila-2)].comment.visible := true;

            raise exception.create('next f');
          end;

          {$ENDREGION}

          {$REGION 'VALIDA TABLA'}


          {$ENDREGION}

          {$REGION 'VALIDA LONGITUD DE HORA INICIO DE LA ACTIVIDAD'}

          sHPdaInicio := excel.cells[ifilacab, 7].text;

          if length(trim(sHPdaInicio)) = 0 then
          begin
            excel.cells[ifilacab, 5].interior.colorindex := 44;
            raise exception.Create('next f');
          end;

          {$ENDREGION}

          {$REGION 'VALIDA FORMATO DE LA HORA DE INICIO'}

          try
            if sHPdaInicio <> '24:00' then
              StrToTime(sHPdaInicio);
          except
            on e:exception do
            begin
              excel.cells[ifilacab, 3].interior.colorindex := 38;
              raise exception.Create('next f');
            end;
          end;

          {$ENDREGION}

          {$REGION 'VALIDA LA LONGITUD DE LA HORA FINAL DE LA ACTIVIDAD'}

          sHPdaFin := excel.cells[ifilacab, 9].text;

          if length(trim(sHPdaFin)) = 0 then
          begin
            excel.cells[ifilacab, 9].interior.colorindex := 44;
            raise exception.Create('next f');
          end;

          {$ENDREGION}

          {$REGION 'VALIDA EL FORMATO DE LA HORA DE INICIO CONTRA FIN DE LA ACTIVIDAD'}

          if (sHPdaFin <> '24:00') and (sHPdaFin <> '24:00') then
          begin
            if StrToTime(sHPdaFin) > StrToTime(sHPdaFin) then
            begin
              excel.range[ifila, 7].addcomment('La hora de inicio es mayor a la final');
              excel.range[ifila, 7].comment.visivle := true;
              raise exception.Create('text f');
            end;
          end;              

          {$ENDREGION}


          while (Trim(Excel.cells[ifila,2].text) <> '') do
          begin
            try

              {$REGION 'LIMPIA LAS ETIQUETAS'}

                excel.cells[ifila, 2].ClearNotes;
                excel.cells[ifila, 3].ClearNotes;
                excel.cells[ifila, 4].ClearNotes ;
                excel.cells[ifila, 5].ClearNotes ;
                excel.cells[ifila, 6].ClearNotes ;
                excel.cells[ifila, 7].ClearNotes ;
                excel.cells[ifila, 8].ClearNotes ;
                excel.cells[ifila, 9].ClearNotes ;
                excel.cells[ifila, 10].ClearNotes;
                excel.cells[ifila, 11].ClearNotes;
                excel.cells[ifila, 12].ClearNotes;

              {$ENDREGION}

              {$REGION 'PREPARA CONSULTA DE iIdDiario DE LAS NOTAS'}

              qrNotas.Active := false;
              qrnotas.sql.Text := 'SELECT ' +
                                  ' MAX(iIdDiario) AS iIdDiario, ' +
                                  ' sWbs, ' +
                                  ' sNumeroActividad, ' +
                                  ' sNumeroOrden ' +
                                  'FROM bitacoradeactividades ' +
                                  'WHERE dIdFecha = :Fecha ' +
                                  'AND sNumeroOrden = :Orden ' +
                                  'AND sNumeroActividad = :Actividad ' +
                                  'AND sContrato = :Contrato AND sidTipoMovimiento = "ED" '+
                                  'AND sidclasificacion = :sidclasificacion ' ;

              {$ENDREGION}

              {$REGION 'CONSULTA LAS PARTIDAS REPORTADAS EN EL DIA Y EL FOLIO'}

              with qractividades do
              begin
                active := false;
                sql.text := 'SELECT b.sContrato, '+
                                   'b.sNumeroOrden, '+
                                   'b.dIdFecha, '+
                                   'b.sIdTurno, '+
                                   'b.sWbs, '+
                                   'b.sNumeroActividad, '+
                                   'b.sIdTipoMovimiento, '+
                                   'b.sHoraInicio, '+
                                   'b.sHoraFinal, '+
                                   'a.sMedida, '+
                                   'concat(b.sHoraInicio, "-", b.sHoraFinal) as hGerencial, '+
                                   'a.sTipoAnexo '+ #10 +
                            'from bitacoradeactividades b '+
                            'inner join tiposdemovimiento t '+
                              'on (b.sContrato = t.sContrato '+
                                  'And b.sIdTipoMovimiento = t.sIdTipoMovimiento '+
                                  'And t.sClasificacion <> "Tiempo Muerto") '+ #10 +

                            'left join actividadesxorden a '+
                              'on (a.sContrato = b.sContrato '+
                                  'and a.sNumeroOrden = b.sNumeroOrden '+
                                  'and a.sWbs = b.sWbs '+
                                  'and a.sNumeroActividad = b.sNumeroActividad '+
                                  'and a.sIdConvenio = :convenio) '+ #10 +

                            'Where b.sContrato = :contrato '+
                              'and b.dIdFecha = :fecha '+
                              'and b.sIdTipoMovimiento <> :Alcance '+
                              'and b.sNumeroOrden = :folio '+
                              'and b.sIdClasificacion="" '+
                              'and sIdTurno = :Turno '+

                            'order by b.iIdDiario ';
                  parambyname('contrato').asstring := param_contrato;
                  parambyname('convenio').asstring := global_convenio;
                  parambyname('fecha').AsDateTime := param_fecha;
                  parambyname('alcance').asstring := frm_connection.connection.configuracion.fieldbyname('sTipoAlcance').asstring;
                  parambyname('folio').asstring := sFolio;
                  parambyname('turno').asstring := global_turno_reporte;
                  open;
              end;

              {$ENDREGION}

              {$REGION 'ANTES DE LEER EL VALOR DE HH, VALIDAMOS SI ES TIEMPO EXTRA O PERSONAL'}

              {$ENDREGION}

              {$REGION 'ASIGNA VALORES A VALIDAR'}

              sActividad    := trim(excel.cells[ifila, 2].text);
              sHInicio      := trim(excel.cells[ifila, 3].text);
              sHTermino     := trim(excel.cells[ifila, 4].text);
              sRecurso      := trim(excel.cells[ifila, 5].text);
              sObra         := trim(excel.cells[ifila, 9].text);
              sPernocta     := trim(excel.cells[ifila, 10].text);
              sPlataforma   := trim(excel.cells[ifila, 11].text);
              sTipoPernocta := trim(excel.cells[ifila, 12].text);
              iTarea        := strtoint( trim( excel.cells[ifila, 13].text ) );

              {$ENDREGION}

              {$REGION 'VALIDA LA ACTIVIDAD'}

              if not qrActividades.locate('snumeroactividad', sActividad, []) then
              begin
                excel.range['B'+inttostr(ifila)].interior.colorindex := 46;
                excel.range['B'+inttostr(ifila)].addcomment('Actividad no registrada');
                excel.range['B'+inttostr(ifila)].comment.visible := true;

                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA ACTIVIDAD DENTRO DEL HORARIO'}

              with connection.QryBusca do
              begin
                active := false;
                sql.text := 'SELECT max(b.iIdDiario) as iIdDiario '+ #10 +
                            'from bitacoradeactividades b '+
                            'inner join tiposdemovimiento t '+
                              'on (b.sContrato = t.sContrato '+
                                  'And b.sIdTipoMovimiento = t.sIdTipoMovimiento '+
                                  'And t.sClasificacion <> "Tiempo Muerto") '+ #10 +

                            'left join actividadesxorden a '+
                              'on (a.sContrato = b.sContrato '+
                                  'and a.sNumeroOrden = b.sNumeroOrden '+
                                  'and a.sWbs = b.sWbs '+
                                  'and a.sNumeroActividad = b.sNumeroActividad '+
                                  'and a.sIdConvenio = :convenio) '+ #10 +

                            'Where b.sContrato = :contrato '+
                              'and b.dIdFecha = :fecha '+
                              'and b.sIdTipoMovimiento <> :Alcance '+
                              'and b.sNumeroOrden = :folio '+
                              'and b.sIdClasificacion <> "" '+
                              'and sIdTurno = :Turno '+
                              'and b.sNumeroActividad = :actividad '+
                              'and b.iIdTarea = :tarea '+
                              'and b.sHoraInicio = :inicio '+
                              'and b.sHoraFinal = :final';

                  parambyname('contrato').asstring := param_contrato;
                  parambyname('convenio').asstring := global_convenio;
                  parambyname('fecha').AsDateTime := param_fecha;
                  parambyname('alcance').asstring := frm_connection.connection.configuracion.fieldbyname('sTipoAlcance').asstring;
                  parambyname('folio').asstring := sFolio;
                  parambyname('turno').asstring := global_turno_reporte;
                  parambyname('actividad').asstring := sActividad;
                  parambyname('tarea').AsInteger := iTarea;
                  parambyname('inicio').asstring := sHPdaInicio;
                  parambyname('final').asstring  := sHPdaFin;
                  open;
              end;

              if connection.QryBusca.RecordCount = 0 then
              begin
                excel.range[ifilacab, 2].addcomment('Actividad ' + sActividad +', Tarea: '+ inttostr(iTarea) +' no registrada en la bitacora');
                excel.range[ifilacab, 2].comment.visible := false;
                raise exception.Create('next');
              end
              else
                iIdDiario := connection.qrybusca.FieldByName('iiddiario').asinteger;

              {$ENDREGION}

              {$REGION 'VALIDA QUE LLEVE ":" LA HORA DE INICIO DEL RECURSO'}
              if pos(':',sHInicio) = 0 then
              begin
                 excel.cells[ifila, 3].addcomment('Formato Erroneo');
                 excel.cells[ifila, 3].comment.visible     := true;
                 excel.cells[ifila, 3].interior.colorindex := 44;

                raise exception.Create('next');
              end;
              {$ENDREGION}

              {$REGION 'VALIDA LA LONGITUD DE LA HORA DE INICIO DEL RECURSO'}

              if length(trim(sHInicio)) = 0 then
              begin
                excel.range['C'+inttostr(ifila)].text.interior.colorindex := 44;

                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA EL FORMATO DE LA HORA DE INICIO DEL RECURSO'}

              try
                if sHInicio <> '24:00' then
                  StrToTime(sHInicio);
              except
                on e:exception do
                begin
                  excel.range['C'+inttostr(ifila)].interior.coloindex := 38;

                  raise exception.Create('next');
                end;
              end;

              {$ENDREGION}

              {$REGION 'VALIDA QUE LLEVE ":" LA HORA DE TERMINO DEL RECURSO'}
              if pos(':',sHTermino) = 0 then
              begin
                excel.cells[ifila, 4].addcomment('Formato Erroneo');
                excel.cells[ifila, 4].comment.visible     := true;
                excel.cells[ifila, 4].interior.colorindex := 44;

                raise exception.Create('next');
              end;
              {$ENDREGION}

              {$REGION 'VALIDA LA LONGITUD DE LA HORA DE TERMINO DEL RECURSO'}

              if length(trim(sHTermino)) = 0 then
              begin
                excel.range['D'+inttostr(ifila)].interior.colorindex := 44;

                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA EL FORMATO DE LA HORA DE TERMINO DEL RECURSO'}

              try
                if sHTermino <> '24:00' then
                  StrToTime(sHTermino);
              except
                on e:exception do
                begin
                  excel.range['D'+inttostr(ifila)].interior.colorindex := 38;
                  raise exception.Create('next');
                end;
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA HORA DE INICIO CONTRA LA HORA DE TERMINO'}

              if (sHInicio <> '24:00') and (sHTermino <> '24:00') then
              begin
                if StrToTime(sHInicio) > StrToTime(sHTermino) then
                begin
                  excel.cells[ifila, 3].addcomment('La hora de inicio no puede ser mayor a la hora de termino');
                  excel.cells[ifila, 3].comment.visible := true;
                  raise exception.Create('next');
                end;
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA HORA DE INICIO NO SEA MAYO QUE LA DE TERMINO'}
              
              if sHInicio = '24:00' then
              begin
                if StrToTime(sHInicio) > ( (StrToTime('23:59')+StrToTime('00:01')) ) then
                begin
                  excel.cells[ifila, 3].AddComment('La hora de inicio no puede ser mayor a la hora de termino');
                  excel.cells[ifila, 3].comment.visible := true;
                  raise exception.Create('next');
                end;
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA LONGITUD DEL ID DEL RECURSO'}

              if Length(Trim(sRecurso)) = 0  then
              begin
                excel.cells[ifila, 4].adcomment('Es necesario un Id de '+lowercase(stipo));
                excel.cells[ifila, 4].comment.visible := true;
                excel.cells[ifila, 4].interior.colorindex := 44; 
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA EL ID DEL RECURSO'}

              if qrTiempoExtra.RecordCount > 0 then
              begin
                  if qrTiempoExtra.Locate('sIdPersonal', sRecurso, []) then
                  begin
                      if qrFolios.Locate('sNumeroorden', sFolio,  []) then
                         iJornada := qrFolios.FieldByName('iJornada').AsInteger;

                      Excel.Range[ColumnaNombre(8)+IntToStr(iFila)+':'+ColumnaNombre(8)+IntToStr(iFila)].Select;
                      Excel.Selection.Formula := '='+ColumnaNombre(6)+IntToStr(ifila)+'*'+ColumnaNombre(7)+IntToStr(iFila) +'*'+ IntToStr(iJornada);
                  end;
              end;

              if not qrRecurso.Locate('sidrecurso', sRecurso, []) then
              begin
                excel.cells[ifila, 5].addcomment('No esta dado de alta en el MOE vigente');
                excel.cells[ifila, 5].comment.visible := true;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA EL ID DEL RECURSO EXISTA EN EL CATALGO MAESTRO'}

              with connection.qrybusca do
              begin
                active := false;
                SQL.text := 'select * from '+lowercase(stipo)+' where scontrato = :contrato and '+sId+' = :id';
                ParamByName('contrato').asstring := global_contrato_barco;
                ParamByName('id').asstring := sRecurso;
                Open;
              end;

              {$ENDREGION}

              {$REGION 'VALIDA SI NO SE REEMPLAZA TODO VALIDAR DUPLICIDAD EN BASE DE DATOS'}

              if not chkReemplazar.Checked then
              begin
                with connection.QryBusca2 do
                begin
                   active := false;
                   sql.text := 'select scontrato from bitacorade'+stipo+ ' '+
                               'where scontrato = :contrato '+
                               'and didfecha = :fecha and '+sid+' = :id '+
                               'and snumeroactividad = :actividad '+
                               'and shorainicio = :hora '+
                               'and snumeroorden = :folio '+
                               'and iiddiario = :diario ';
                   parambyname('contrato').asstring := param_contrato;
                   parambyname('fecha').asdatetime := param_fecha;
                   parambyname('id').asstring := sRecurso;
                   parambyname('actividad').asstring := sactividad;
                   parambyname('hora').asstring := sHInicio;
                   parambyname('folio').asstring := sFolio;
                   parambyname('diario').asinteger := iiddiario;
                   open;
                end;
                if connection.qrybusca2.RecordCount > 0 then
                begin
                  excel.cells[ifila, 5].interior.colorindex := 42;
                  raise exception.Create('next');
                end;
              end;

              {$ENDREGION}

              {$REGION 'VALIDA RESULTADO DE BUSQUEDA EN CATALOGO MAESTRO'}

              if connection.QryBusca.RecordCount = 0 then
              begin
                excel.cells[ifila, 4].addcomment('No dado de alta en el catalogo de '+lowercase(stipo));
                excel.cells[ifila, 4].comment.visible := true;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA LONGITUD DE LA CANTIDAD'}

              if length(trim(excel.cells[ifila, 6].text)) = 0 then
              begin
                excel.cells[ifila, 6].addcomment('Se ignora por tener valor nulo');
                excel.cells[ifila, 6].comment.visible := true;
                excel.cells[ifila, 6].interior.colorindex := 44;
                raise exception.Create('next');
              end;
              dCantidad   := excel.range[columnanombre(6)+inttostr(ifila)].value;

              {$ENDREGION}

              {$REGION 'VALIDA LA LONGITUD DE HORAS HOMBRE'}

              if Length(trim(excel.cells[ifila, 8].text)) = 0 then
              begin
                excel.cells[ifila, 8].addcomment('Se ignora por tener valor nulo');
                excel.cells[ifila, 8].comment.visible := true;
                excel.cells[ifila, 8].interior.colorindex := 44;
                raise exception.Create('next');
              end;
              dCantidadHH := excel.cells[ifila, 8].text;
              
              {$ENDREGION}

              {$REGION 'VALIDA LA CANTIDAD EN 0'}

              if dCantidad = 0 then
              begin
                excel.cells[ifila, 6].addcomment('Se ignora por tener valor 0');
                excel.cells[ifila, 6].comment.visible := true;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA CANTIDAD HORAS HOMBRE EN 0'}

              if dCantidadHH = 0 then
              begin
                excel.cells[ifila, 8].addcomment('Se ignora por tener valor nulo');
                excel.cells[ifila, 8].comment.visible := true;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA LONGITUD EL TIPO DE OBRA'}

              if length(Trim(sObra)) = 0 then
              begin
                excel.cells[ifila, 9].addcomment('Tipo de obra vacia');
                excel.cells[ifila, 9].comment.visible := true;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA EXISTA EL TIPO DE OBRA'}

              if not qrTObra.Locate('stipoobra', sObra,[]) then
              begin
                excel.cells[ifila, 9].addcomment('Tipo de obra no registrado');
                excel.cells[ifila, 9].comment.visible := true;
                raise exception.Create('next');
              end;

              {$ENDREGION}                           

              {$REGION 'VALIDA LA LONGITUD DE LA PERNOCTA EN LA QUE RESIDEN'}

              if length(trim(sPernocta)) = 0 then
              begin
                excel.cells[ifila, 10].addcomment('Pernocta no valida');
                excel.cells[ifila, 10].comment.visible := true;
                excel.cells[ifila, 10].interior.colorindex := 44;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA LOGITUD DE LA PLATAFORMA'}

              if length(trim(splataforma)) = 0 then
              begin
                excel.cells[ifila, 11].addcomment('Plataforma no valida');
                excel.cells[ifila, 11].comment.visible := true;
                excel.cells[ifila, 11].interior.colorindex := 44;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA LONGITUD DE LA PERNOCTA DE SERVICIOS'}

              if length(trim(sTipoPernocta)) = 0 then
              begin
                  excel.cells[ifila, 12].addcomment('Pernocta de servicios no valida');
                excel.cells[ifila, 12].comment.visible := true;
                excel.cells[ifila, 12].interior.colorindex := 44;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA EXISTENCIA DE LA PERNOCTA DE RESIDENCIA'}

              if not qrPernoctan.Locate('sIdPernocta', sPernocta, []) then
              begin
                excel.range[columnanombre(10)+inttostr(ifila)].addcomment('Pernocta no valida');
                excel.range[columnanombre(10)+inttostr(ifila)].comment.visible := true;
                excel.range[columnanombre(10)+inttostr(ifila)].interior.colorindex := 46;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'VALIDA LA EXISTENCIA DE LA PLATAFORMA'}

              if not qrPlataformas.Locate('sidplataforma', sPlataforma, []) then
              begin
                excel.cells[ifila, 11].addcomment('Plataforma no valida');
                excel.cells[ifila, 11].comment.visible := true;
                excel.cells[ifila, 11].interior.colorindex := 46;
                raise exception.Create('next');
              end;

              {$ENDREGION}

              {$REGION 'PREPARA DATOS PARA INSERTAR E INSERCCION DEL REGISTRO VALIDADO'}

              with connection.zCommand do
              begin
                active := false;
                sql.text := 'select * from bitacorade'+lowercase(stipo)+' where sContrato = :contrato and didfecha = :fecha';
                parambyname('contrato').asstring := param_contrato;
                parambyname('fecha').asdatetime := param_fecha;
                open;

                qrNotas.Active := false;
                qrNotas.ParamByName('fecha').AsDateTime          := param_fecha;
                qrNotas.ParamByName('orden').AsString            := sfolio;
                qrNotas.ParamByName('actividad').AsString        := sactividad;
                qrNotas.ParamByName('contrato').AsString         := param_contrato;
                qrNotas.ParamByName('sidclasificacion').AsString := 'TE';
                qrNotas.open;

                {$REGION 'VALIDA EL iIdDiario CON EL QUE SE LIGARÁ EL RECURSO'}

                if ( qrNotas.RecordCount = 0 ) or ( qrNotas.FieldByName('iiddiario').asinteger < 0 ) then
                begin
                  excel.cells[ifila, 12].addcomment('No se encontro nota para asignar el recurso.');
                  excel.cells[ifila, 12].comment.visible := true;

                  raise exception.Create('next');
                end;

                {$ENDREGION}

                try
                  Append;
                  fieldbyname('scontrato').asstring    := param_contrato;
                  fieldbyname('didfecha').asdatetime   := param_fecha;
                  fieldbyname('iiddiario').asinteger   := iiddiario;
                  fieldbyname('iitemorden').asinteger  := 0;
                  fieldbyname(sId).asstring            := sRecurso;
                  fieldbyname('sTipoObra').asstring    := sObra;
                  fieldbyname('sdescripcion').asstring := qrRecurso.FieldByName('sDescripcion').asstring;
                  fieldbyname('sidpernocta').asstring  := sPernocta;
                  fieldbyname('shorainicio').asstring  := sHInicio;
                  fieldbyname('shorafinal').asstring   := sHTermino;
                  fieldbyname('dcantidad').asfloat     := dcantidad;
                  fieldbyname('swbs').asstring         := qrActividades.FieldByName('swbs').asstring;
                  fieldbyname('snumeroactividad').asstring := sActividad;
                  fieldbyname('dCantHH').asfloat       := dCantidadHH;
                  FieldByName('dsolicitado').asfloat   := 0;
                  FieldByName('dsolicitado').asfloat   := 0;
                  FieldByName('sfactor').AsString      := '0';
                  FieldByName('dcostomn').AsInteger    := 0;
                  FieldByName('dcostodll').AsInteger   := 0;
                  FieldByName('dajuste').AsInteger     := 0;
                  fieldbyname('snumeroorden').asstring := sFolio;
                  FieldByName('dcanthhgenerador').AsFloat := 0;
                  FieldByName('iIdTarea').AsInteger    := iTarea;
                  if stipo = 'PERSONAL' then
                  begin
                    FieldByName('sagrupapersonal').AsString := '*';
                    FieldByName('laplicapernocta').AsString := 'Si';
                    fieldbyname('sidplataforma').asstring  := sPlataforma;
                    fieldbyname('sTipoPernocta').asstring := sTipoPernocta;
                  end;

                  Post;
                  Inc(iRegistros);
                except
                  on e:exception do
                    showmessage(e.message);
                end;

                excel.range['B'+inttostr(ifila)+':M'+inttostr(ifila)].interior.colorindex := 43;
              end;

              {$ENDREGION}

            except
              on e:exception do
                if e.Message = 'Next' then
                  ;
            end;

            inc(ifila);

          end;

          inc(ifila, 3);

        end;
      except
        on e:exception do
        begin
          if e.Message = 'next f' then
          begin
            while ( excel.cells[ifilasp, 2].text ) <> '' do
              inc(ifilasp);

            ifila := ifilasp + 3; 
          end;
        end;
      end;

      {$REGION 'CREAR LEYENDA'}

      excel.range['O4'].value := 'ACTIVIDAD NO VALIDA';
      excel.range['P4'].interior.colorindex := 46;
      excel.range['O5'].value := 'VACIO';
      excel.range['P5'].interior.colorindex := 44;
      excel.range['O6'].value := 'HORA MAL';
      excel.range['P6'].interior.colorindex := 38;
      excel.range['O7'].value := 'YA EXISTE';
      excel.range['P7'].interior.colorindex := 42;


      {$ENDREGION}

    end;
  except
    on e:exception do
    begin
      MessageDlg(e.Message, mtInformation, [mbOk], 0);
    end;
  end;
  MessageDlg('Se importaron : ' + inttostr(iregistros)+' registros', mtInformation, [mbOk], 0);
  excel.visible := True;
end;

procedure TFrmImportaCuadre.GeneraCuadreAnterior(Fecha: TDateTime; Personal: Boolean; Equipo: Boolean; Folios: TcxCheckListBox);
var
  Excel,
  Libro,
  Hoja,
  rango : Variant;

  qrMoe,
  qrBitacoraP,
  qrBitacoraE,
  qrActividades : TZReadOnlyQuery;
begin


end;

end.
