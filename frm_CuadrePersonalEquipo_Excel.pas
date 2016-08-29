unit frm_CuadrePersonalEquipo_Excel;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComObj, ComCtrls, ExtCtrls, OleCtrls, UnitExcel,
  DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Buttons,StrUtils,
  AdvSmoothPanel, AdvCombo, Mask, JvExMask, JvToolEdit, JvCombobox,
  AdvCircularProgress, frm_connection, global;

type
  TfrmCuadrePersonalEquipo_Excel = class(TForm)
    OpenXLS: TOpenDialog;
    ZQuery1: TZQuery;
    pBarraAvance: TProgressBar;
    cmbIgnorarColumnas: TCheckBox;
    cmbIgnorarRecursos: TCheckBox;
    cmbCombinarPartidas: TCheckBox;
    AdvSmoothPanel1: TAdvSmoothPanel;
    AdvSmoothPanel2: TAdvSmoothPanel;
    cmbSoloPersonal: TCheckBox;
    cmbSoloEquipo: TCheckBox;
    cmbReemplazar: TCheckBox;
    Cbx2: TCheckBox;
    ChkHorario: TCheckBox;
    btnCmdImportar: TSpeedButton;
    btn1: TSpeedButton;
    ChkEInfo: TCheckBox;
    ChkEPersonal: TCheckBox;
    ChkEEquipo: TCheckBox;
    CmbFolios: TJvCheckedComboBox;
    Progreso: TAdvCircularProgress;
    chkActualizaHE: TCheckBox;
    chkPartidas: TCheckBox;
    Label2: TLabel;
    chkHorasExtras: TCheckBox;
    procedure FormShow(Sender: TObject);
    Procedure LeerCuadre;
    function IsIn(Valor: string; Lista: TStringList): Boolean;
    procedure btnCmdImportarClick(Sender: TObject);
    procedure btn1Click(Sender: TObject);
    procedure chkHorasExtrasClick(Sender: TObject);
  private
    procedure ImportarCuadre;

    procedure ImportaCuadreHE;

    procedure ImportaCuadre(Pers, Equi, Sustituir: Boolean);
    procedure ExportarPlantilla(EPersonal, EEquipo, Info: Boolean);
    procedure CargarCheckCombo(ComboFolios: TJvCustomCheckedComboBox);
    procedure ExportarPlantillaActividades(EPersonal,EEquipo,Info:Boolean);
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmCuadrePersonalEquipo_Excel: TfrmCuadrePersonalEquipo_Excel;
  Cuadre_FechaReporte: TDateTime;
  ContratoDiario : String;

implementation

{$R *.dfm}

function TfrmCuadrePersonalEquipo_Excel.IsIn(Valor: string; Lista: TStringList): Boolean;
var
   nIdx: Integer;
begin
   Result := False;
   for nIdx := 0 to Lista.Count - 1 do begin
      if UpperCase(Lista[nIdx]) = UpperCase(Valor) then begin
         Result := True;
         Break;
      end;
   end;
end;

Procedure TfrmCuadrePersonalEquipo_Excel.LeerCuadre;
Var
  Excel,
  Libro,
  Hoja: Variant;

  i,
  iHojas,
  iHojasTotales,
  iFila,
  iColumna,
  iColumnaItem,
  iColumnaHorarioInicio,
  iColumnaClasificacion,
  iFilaItem,
  iFilaPartida,
  iColumnaPartida,
  Vacio,
  IdMoeActivo,
  IdDiario,IdDiarioAux,
  BarraAvance,
  TotalPartidas,
  TotalColumnas,
  iLeerPaginas: Integer;

  dCantidad, dHH: Real;

  sTipoCuadre,
  Valor,
  Valor2,
  ListaFolios,
  Folio,
  Partida,
  Item,
  ItemDescripcion, 
  Cantidad,
  HorasHombre,
  ArchivoExcel,
  sHoraInicio,
  sHoraFinal,
  sIdClasificacion,
  Tabla,
  sIdPlataforma,
  sIdPernocta,
  sWbs, sNumeroActividad: String;

  SQLExtra,
  IgnorarRecurso : TStringList;

  Founded, Error : Boolean;
  Query, Query2, zQCuentas  : TZQuery;

  HuboErrores,Colorear:Boolean;


  Function ExcelCloseWorkBooks(Excel : Variant; SaveAll: Boolean): Boolean;
  var
    loop: byte;
  Begin
    Result := True;
    Try
      For loop := 1 to Excel.Workbooks.Count Do
        Excel.Workbooks[1].Close[SaveAll];
    Except
      Result := False;
    End;
  End;
  Function ExcelClose(Excel : Variant; SaveAll: Boolean): Boolean;
  Begin
    Result := True;
    Try
      ExcelCloseWorkBooks(Excel, SaveAll);
      Excel.Quit;
    Except
      MessageDlg('No se puede cerrar la aplicación excel.', mtError, [mbOK], 0);
      Result := False;
    End;
  End;
begin
  Try
    //leer moe
    HuboErrores:=False;
    Query := TZQuery.Create(Self);
    Query.Connection := Connection.zConnection;
    Query2 := TZQuery.Create(Self);
    Query2.Connection := Connection.zConnection;

    zQCuentas := TZQuery.Create(Self);
    zQCuentas.Connection := Connection.zConnection;

    SQLExtra := TStringList.Create;
    IgnorarRecurso := TStringList.Create;

    zQCuentas.Active := False;
    zQCuentas.SQL.Clear;
    zQCuentas.SQL.Add('select sIdPernocta from cuentas limit 1');
    zQCuentas.Open;

    Query.Active := False;
    Query.SQL.Clear;
    Query.SQL.Add('' +
                  'SELECT ' +
                  '	* ' +
                  'FROM moe AS m ' +
                  '	INNER JOIN moerecursos AS mr ' +
                  '		ON (mr.iIdMoe = m.iIdMoe) ' +
                  'WHERE m.dIdFecha = ( SELECT MAX(dIdFecha) FROM moe WHERE dIdFecha <= :Fecha AND sContrato = :Contrato) ' +
                  ' AND sContrato = :Contrato ' +
                  'LIMIT 1 ');
    Query.Params.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
    Query.Params.ParamByName('Contrato').AsString := ContratoDiario;
    Query.Open;

    if Query.RecordCount > 0 then
    begin
      IdMoeActivo := Query.FieldByName('iIdMoe').AsInteger;
    end
    else
    begin
      ShowMessage('No existen movimientos de personal y equipo en el sistema, crée uno para continuar.');
      Exit;
    end;
    Try
      Excel := CreateOleObject('Excel.Application');
    Except
      On E: Exception do
      begin
        FreeAndNil(Excel);
        ShowMessage(E.Message);
        Exit;
      end;
    End;

    if OpenXLS.Execute then
    begin
      ArchivoExcel := OpenXLS.FileName;
      Excel.WorkBooks.Open(ArchivoExcel);
      Libro := Excel.WorkBooks[excel.workbooks.count];
    end
    else
    begin
      exit;
    end;

    Excel.Visible := False;
    Excel.DisplayAlerts:= False;

    pBarraAvance.Min := 0;
    pBarraAvance.Position := 0;
    iHojasTotales := Libro.Sheets.Count;

    if cmbSoloPersonal.Checked then
    begin
      iLeerPaginas := 1;
      Query2.Active := False;
      Query2.SQL.Clear;
      Query2.SQL.Text := 'DELETE FROM bitacoradepersonal WHERE dIdFecha = :Fecha AND sContrato = :Contrato AND sTipoObra = "PU"';
      Query2.Params.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
      Query2.Params.ParamByName('Contrato').AsString := ContratoDiario;
      Query2.ExecSQL;
    end;

    if cmbSoloEquipo.Checked then
    begin
      Query2.Active := False;
      Query2.SQL.Clear;
      Query2.SQL.Text := 'DELETE FROM bitacoradeequipos WHERE dIdFecha = :Fecha AND sContrato = :Contrato';
      Query2.Params.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
      Query2.Params.ParamByName('Contrato').AsString := ContratoDiario;
      Query2.ExecSQL;
    end;     

    if (cmbSoloEquipo.Checked) AND (Not cmbSoloPersonal.Checked) then
    begin
      iLeerPaginas := 2;
    end;

    if (Not cmbSoloEquipo.Checked) AND (Not cmbSoloPersonal.Checked) then
    begin
      ShowMessage('No se importará información.');
      Exit;
    end;

    for iHojas := iLeerPaginas to Libro.Sheets.Count do
    begin
      IgnorarRecurso.Clear;

      if (cmbSoloPersonal.Checked) AND (Not cmbSoloEquipo.Checked) then
      begin
        if iHojas = 2 then
        begin
          Continue;
        end;
      end;

      if (cmbSoloEquipo.Checked) AND (Not cmbSoloPersonal.Checked) then
      begin
        if iHojas = 1 then
        begin
          Continue;
        end;
      end;

      if iHojas = 1 then
      begin
        sTipoCuadre := 'Personal';
      end
      else
      if iHojas = 2 then
      begin
        sTipoCuadre := 'Equipo';
      end
      else
      begin
        Exit;
      end;

      Libro.WorkSheets[iHojas].Activate;

      Query.Active := False;
      Query.SQL.Clear;
      Query.SQL.Add('' +
                    'SELECT ' +
                    '	* ' +
                    'FROM moerecursos AS mre ' +
                    'WHERE mre.iIdMoe = :IdMoe AND eTipoRecurso = :TipoRecurso');
      Query.ParamByName('IdMoe').AsInteger := IdMoeActivo;
      Query.ParamByName('TipoRecurso').AsString := sTipoCuadre;
      Query.Open;

      Connection.QryBusca.Active := False;
      Connection.QryBusca.SQL.Text := 'SELECT iIdDiario FROM bitacoradeactividades WHERE sContrato = :Contrato AND dIdFecha = :Fecha AND sIdTipoMovimiento = "ED" ';

      Connection.QryBusca.ParamByName('Contrato').AsString := ContratoDiario;
      Connection.QryBusca.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
      Connection.QryBusca.Open;

      pBarraAvance.Max := Connection.QryBusca.RecordCount;
      TotalPartidas := Connection.QryBusca.RecordCount;
      TotalColumnas := 0;

      iColumnaPartida := 2;
      iColumnaItem := 7;

      iColumna := 7;
      iFila := 1;
      iFilaPartida := 1;
      iFilaItem := 0;

      IdDiario := Connection.QryBusca.FieldByName('iIdDiario').AsInteger;

      Founded := False;
      Vacio := 0;
      ListaFolios := '';
      Error := False;

      while Vacio < 30 do
      begin
        Excel.activeworkbook.activesheet.Range[ColumnaNombre(iColumnaPartida)+IntToStr(iFila)+':'+ColumnaNombre(iColumnaPartida)+IntToStr(iFila)].Select;
        Valor := Excel.Selection.value;
        if Valor = '' then
        begin
          Inc(Vacio);
        end
        else
        begin
          Vacio := 0;
          if (Valor = 'PDA') OR (Valor = 'PDA.') then
          begin
            if iFilaItem = 0 then begin
              iFilaItem := iFila - 2;
            end;
            iFilaPartida := iFila;
            Partida := excel.activeworkbook.activesheet.Cells[(iFila - 1), (iColumnaPartida + 1)].Text;
            Folio := Trim(Partida);
            Connection.QryBusca.Active := False;
            Connection.QryBusca.SQL.Text := 'SELECT * FROM ordenesdetrabajo WHERE sNumeroOrden = ' + QuotedStr(Folio) + ' AND sContrato = ' + QuotedStr(ContratoDiario);
            Connection.QryBusca.Open;
            Inc(iFilaPartida);
            if Connection.QryBusca.RecordCount > 0 then begin
              sIdPlataforma := Connection.QryBusca.FieldByName('sIdPlataforma').AsString;
              sIdPernocta := Connection.QryBusca.FieldByName('sIdPernocta').AsString;
              while Excel.Cells[(iFilaPartida), (iColumnaPartida)].Text <> '' do
              begin
                Partida := Excel.Cells[(iFilaPartida), (iColumnaPartida)].Text;
                sIdClasificacion := Excel.Cells[(iFilaPartida), (iColumnaPartida + 1)].Text;
                sHoraInicio := Excel.Cells[(iFilaPartida), (iColumnaPartida + 2)].Text;
                sHoraFinal := Excel.Cells[(iFilaPartida), (iColumnaPartida + 3)].Text;

                connection.QryBusca.Active := False;
                connection.QryBusca.SQL.Clear;


                Connection.QryBusca.SQL.Add('SELECT ' +
                                            ' MAX(iIdDiario) AS iIdDiario, ' +
                                            ' sWbs, ' +
                                            ' sNumeroActividad, ' +
                                            ' sNumeroOrden ' +
                                            'FROM bitacoradeactividades ' +
                                            'WHERE dIdFecha = :Fecha ' +
                                            'AND sNumeroOrden = :Orden ' +
                                            'AND sNumeroActividad = :Actividad ' +
                                            'AND sContrato = :Contrato AND sidTipoMovimiento = "ED" '+

                                            'AND sidclasificacion = :sidclasificacion ');

                if ChkHorario.Checked then
                begin
                  connection.QryBusca.SQL.Add(' AND sHorainicio = :sHoraInicio AND sHoraFinal = :sHoraFinal ');
                  connection.QryBusca.ParamByName('sHoraInicio').AsString    := sHoraInicio;
                  if sHorafinal <> '00:00' then
                    connection.QryBusca.ParamByName('sHoraFinal').AsString    := sHoraFinal
                  else
                    connection.QryBusca.ParamByName('sHoraFinal').AsString    := '24:00';
                end;
                connection.QryBusca.ParamByName('Contrato').AsString := ContratoDiario;
                connection.QryBusca.ParamByName('Orden').AsString    := Folio;
                connection.QryBusca.ParamByName('Fecha').AsDate      := Cuadre_FechaReporte;
                connection.QryBusca.ParamByName('Actividad').AsString    := Partida;
                connection.QryBusca.ParamByName('sidclasificacion').AsString    := sIdClasificacion;

                connection.QryBusca.Open;
                Colorear := False;
                if Connection.QryBusca.FieldByName('iIdDiario').AsInteger = 0 then
                begin
                  //rojo no se encontro el iddiario
                  //ShowMessage('No se puede importar la información de: '+Contratodiario+' con folio: '+folio+' partida: '+partida);
                  excel.Cells[(iFilaPartida), (iColumnaItem)].Interior.ColorIndex := 3;//no se encuantra el iddiario
                  Colorear := True;
                  HuboErrores:=True;
                  //Exit;
                end;

                if Connection.QryBusca.RecordCount > 0 then
                begin
                  IdDiario := Connection.QryBusca.FieldByName('iIdDiario').AsInteger;
                  sWbs := Connection.QryBusca.FieldByName('sWbs').AsString;
                  sNumeroActividad := Connection.QryBusca.FieldByName('sNumeroActividad').AsString;
                end
                else
                begin
                  ShowMessage('No se encontró reportada la actividad ' + Partida + ' en el folio ' + Folio);
                  Exit;
//                  IdDiario := 10;
                end;

                iColumnaItem := 7;
                while Excel.Cells[(iFilaItem), (iColumnaItem)].Text <> '' do
                begin

                  if cmbIgnorarColumnas.Checked then
                  begin
                    if Excel.Columns[ColumnaNombre(iColumnaItem)+':'+ColumnaNombre(iColumnaItem)].Hidden then begin
                      pBarraAvance.Position := pBarraAvance.Position + 1;
                      Inc(iColumnaItem, 2);
                      Continue;
                    end;
                  end;

                  if (iColumnaItem - 6) > TotalColumnas then
                  begin
                    pBarraAvance.Max := TotalPartidas * TotalColumnas;
                    Inc(TotalColumnas);
                  end;
                  Error := False;
                  Item := Trim(Excel.Cells[(iFilaItem), (iColumnaItem)].Value);
                  ItemDescripcion := Trim(Excel.Cells[(iFilaItem - 1), (iColumnaItem)].Value);

                  Cantidad := Trim(Excel.Cells[(iFilaPartida), (iColumnaItem)].Value);
                  if Colorear then                 //amarillo
                    Excel.Cells[(iFilaPartida), (iColumnaItem)].Interior.ColorIndex := 6;
                  if Cantidad = '' then
                  begin
                    Cantidad := '0';
                  end;

                  HorasHombre := Excel.Cells[(iFilaPartida), (iColumnaItem + 1)].Value;
                  if HorasHombre = '' then
                  begin
                    HorasHombre := '0';
                  end;

                  if Query.Locate('sIdRecurso', Item, [loCaseInsensitive]) then
                  begin
                    Item := Query.FieldByName('sIdRecurso').AsString;
                    Try
                      dCantidad := StrToFloat(Cantidad);
                      dHH := StrToFloat(HorasHombre);
                    Except
                        dCantidad := 0;
                        dHH := 0;
                    End;

                    if (dCantidad > 0) OR (dHH > 0) then
                    begin
                      Error := False;
                    end
                    else
                    begin
                      Error := True;
                    end;
                  end
                  else
                  begin
                    if cmbIgnorarRecursos.Checked then
                    begin
                      IgnorarRecurso.Add(Item);
                      Error := True;
                    end
                    else
                    begin
                      if Not IsIn(Item, IgnorarRecurso) then begin
                        if MessageDlg('No se encontró el recurso ' + Item + ': ' + #10#13 + ItemDescripcion + #10#13 + ' en la embarcación ¿desea que el sistema lo de en alta automáticamente?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then begin
                          Query2.Active := False;
                          Query2.SQL.Clear;
                          Query2.SQL.Add('INSERT IGNORE INTO moerecursos (iIdMoe, sIdRecurso, eTipoRecurso, sDescripcion, dCantidad) VALUES ('+IntToStr(IdMoeActivo)+', '+QuotedStr(Item)+', '+QuotedStr(sTipoCuadre)+', '+QuotedStr(ItemDescripcion)+', 0); ');
                          Try
                            Query2.ExecSQL;
                            Error := False;
                            Query.Refresh;
                          Except
                            Error := True;
                          End;
                        end
                        else
                        begin
                          IgnorarRecurso.Add(Item);
                          Error := True;
                        end;
                      end
                      else
                      begin
                        Error := True;
                      end;
                    end;
                  end;

                  if (Not Error) and (not huboerrores) then
                  begin

                    SQLExtra.Clear;
                    if sTipoCuadre = 'Personal' then
                    begin
                      Tabla := 'bitacoradepersonal';
                      SQLExtra.Add('sIdPlataforma, ');//0
                      SQLExtra.Add('"' + sIdPlataforma + '", ');//1
                      SQLExtra.Add('sTipoPernocta, ');
                      SQLExtra.Add(' '+zQCuentas.FieldByName('sIdPernocta').AsString+' , ');
                    end
                    else
                    if sTipoCuadre = 'Equipo' then
                    begin
                      Tabla := 'bitacoradeequipos';
                      SQLExtra.Add('');
                      SQLExtra.Add('');
                      SQLExtra.Add('');
                      SQLExtra.Add('');
                    end;

                    Query2.Active := False;
                    Query2.SQL.Clear;
                    Query2.SQL.Add( '' +
                                    'INSERT IGNORE INTO ' + Tabla + ' ' +
                                    ' (sContrato, sNumeroOrden, dIdFecha, iIdDiario, sId'+sTipoCuadre+', sTipoObra, sDescripcion, sIdPernocta, '+SQLExtra[0]+' sHoraInicio, sHoraFinal, dCantidad, '+SQLExtra[2]+' dCantHH, sNumeroActividad, sWbs)' +
                                    ' VALUES ' +
                                    ' (:Contrato, :Orden, :Fecha, :IdDiario, :IdItem, "PU", :Descripcion, :IdPernocta, '+SQLExtra[1]+' :HoraInicio, :HoraFinal, :Cantidad, '+SQLExtra[3]+' :CantidadHH, :Partida, :Wbs); ' +
                                    '');
                    Query2.ParamByName('Contrato').AsString := ContratoDiario;
                    Query2.ParamByName('Orden').AsString := Folio;
                    Query2.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
                    Query2.ParamByName('IdDiario').AsInteger := iddiario;
                    Query2.ParamByName('IdItem').AsString := Item;
                    Query2.ParamByName('Descripcion').AsString := ItemDescripcion;
                    Query2.ParamByName('IdPernocta').AsString := sIdPernocta;
                    Query2.ParamByName('HoraInicio').AsString := sHoraInicio;
                    Query2.ParamByName('HoraFinal').AsString := sHoraFinal;
                    Query2.ParamByName('Cantidad').AsFloat := dCantidad;
                    Query2.ParamByName('CantidadHH').AsFloat := dHH;
                    Query2.ParamByName('Partida').AsString := sNumeroActividad;
                    Query2.ParamByName('Wbs').AsString := sWbs;
                    Query2.ExecSQL;

                  end;
                  pBarraAvance.Position := pBarraAvance.Position + 1;
                  Inc(iColumnaItem, 2);
                end;

                Inc(iFilaPartida);
              end;
            end
            else
            begin
              ShowMessage('El folio ' + Folio + ' no fue encontrado en los registros, favor de verificar.');
              HuboErrores := True;
            end;
          end;
        end;
        Inc(iFila);
      end;
    end;
    if not HuboErrores then
      ShowMessage('Se importó el cuadre para el día ' + FormatDateTime('dd-mm-yyyy', Cuadre_FechaReporte))
    else
      ShowMessage('Se importó el cuadre para el día ' + FormatDateTime('dd-mm-yyyy', Cuadre_FechaReporte)+#10+' Pero hubo errores, porfavor verifique los registros marcados en excel'+#10+' en el apartado cantidad e intente de nuevo.')
  Finally
    pBarraAvance.Position := 0;
    Query2.Free;
    Query.Free;
    zQCuentas.Free;
    if HuboErrores then
      Excel.Visible := True
    else
      if not ExcelClose(Excel, False) then
        ShowMessage('No se puede cerrar la aplicacion excel.');
  End;
end;

procedure TfrmCuadrePersonalEquipo_Excel.ExportarPlantilla(EPersonal,EEquipo,Info:Boolean);
var
  EExcel,
  ELibro,
  EHojaPErsonal,EHojaEquipo: Variant;

  IELibro, IEHojaPersonal,IEhojaEquipo,
  CVeces:Integer;
  I: Integer;

  zEMoe,ZEFolios,ZEActividades,ZEMovimientos:TZReadOnlyQuery;
  IMActivo:Integer;
  procedure CreaEstructura(Tipo:String);
  const
    IClPlataforma = 1;
    IclPda = 2;
    IcltipoMov = 3;
    IClInicia = 4;
    IClFinaliza = 5;
    IClDuracion = 6;
  var
    ZEMoeR:TZReadOnlyQuery;
    IEFila, x, Ucolumna,
    IClRecurso, IIndex, Ifolio
    :Integer;
    VHoja, VRango : Variant;
  begin
    try
      if info then
      begin

        ZEMoeR := TZReadOnlyQuery.Create(nil);
        zEMoeR.Connection := connection.zConnection;
        zEMoeR.Active := False;
        if Tipo = 'Personal' then
          zEMoeR.SQL.Text :=  'SELECT mre.sIdRecurso ,p.* '+
                              'FROM moerecursos AS mre '+
                              'inner join personal p on (mre.sidrecurso = p.sidpersonal and p.scontrato = :Contrato) '+
                              'WHERE mre.iIdMoe = :IdMoe  order by p.iItemOrden';
        if Tipo = 'Equipo' then
          zEMoeR.SQL.Text :=  'SELECT mre.sIdRecurso ,e.* '+
                              'FROM moerecursos AS mre '+
                              'inner join equipos e on (mre.sidrecurso = e.sidequipo and e.scontrato = :Contrato) '+
                              'WHERE mre.iIdMoe = :IdMoe  order by e.iItemOrden';
        ZEMoeR.ParamByName('IdMoe').AsInteger := IMActivo;
        ZEMoeR.ParamByName('Contrato').AsString := ContratoDiario;
        ZEMoeR.Open;

      end;

      IEFila := 1;

      if Tipo = 'Personal' then
      begin
        Vhoja := EHojaPErsonal;
        iindex := iEHojaPErsonal;
      end;

      if Tipo = 'Equipo' then
      begin
        Vhoja := EHojaEquipo;
        iindex := iEHojaEquipo;
      end;
      // creando paquetes
      ZEFolios.first;
      while not ZEFolios.Eof do
      begin
        Ifolio := CmbFolios.Items.IndexOf(ZEFolios.FieldByName('snumeroorden').AsString);
        if Ifolio > -1 then
        if CmbFolios.IsChecked(Ifolio) then
        begin

          IClRecurso := 7;

          VHoja.Rows[inttostr(IEFila)+':'+inttostr(IEFila)].RowHeight := 66;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IcltipoMov)+IntToStr(IEFila)+':'+ColumnaNombre(IcltipoMov+1)+IntToStr(IEFila+1)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'TOTAL';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza)+IntToStr(IEFila)+':'+ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila+1)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := '0';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IcltipoMov)+IntToStr(IEFila+2)+':'+ColumnaNombre(IcltipoMov+2)+IntToStr(IEFila+2)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := ZEFolios.FieldByName('snumeroorden').AsString;
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClPlataforma)+IntToStr(IEFila+2)+':'+ColumnaNombre(IClPlataforma)+IntToStr(IEFila+2)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := ZEFolios.FieldByName('sidplataforma').AsString;
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;


          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClPlataforma)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClPlataforma)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'Referencia';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IclPda)+IntToStr(IEFila+3)+':'+ColumnaNombre(IclPda)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'PDA';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IcltipoMov)+IntToStr(IEFila+3)+':'+ColumnaNombre(IcltipoMov)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'N.E.C.';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClInicia)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClInicia)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'Inicia';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClFinaliza)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'Finaliza';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'Duración';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          ZEMoeR.First;
          while not ZEMoeR.Eof do
          begin
            Ucolumna := iclrecurso;
            //Recursos
            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso)+IntToStr(IEFila)+':'+ColumnaNombre(IClRecurso+1)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := ZEMoeR.FieldByName('sdescripcion').AsString;
            VRango.NumberFormat := '@';
            VRango.Interior.Color := $00BBBBBB;
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso)+IntToStr(IEFila+1)+':'+ColumnaNombre(IClRecurso+1)+IntToStr(IEFila+1)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 11;
            VRango.Font.Bold := true;
            VRango.WrapText := True;
            VRango.NumberFormat := '@';
            VRango.value := ZEMoeR.FieldByName('sidrecurso').AsString;
            VRango.Interior.Color := $00BBBBBB;
            VRango.NumberFormat := '@';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClRecurso)+IntToStr(IEFila+3)];
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := true;
            VRango.WrapText := True;
            VRango.NumberFormat := '@';
            VRango.value := 'Cant.';
            VRango.Interior.Color := $00BBBBBB;
            Vrango.Borders.LineStyle := xlContinuous;
            VHoja.Cells[IEFila, IClRecurso].ColumnWidth := 4.3;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso+1)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClRecurso+1)+IntToStr(IEFila+3)];
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := false;
            VRango.WrapText := True;
            VRango.NumberFormat := '@';//'0.00000000';
            VRango.value := 'h.h.';
            VRango.Interior.Color := $00BBBBBB;
            Vrango.Borders.LineStyle := xlContinuous;
            VHoja.Cells[IEFila, IClRecurso+1].ColumnWidth := 12.3;

            IClRecurso := IClRecurso+2;
            ZEMoeR.Next;
          end;
          IEFila := IEFila + 4;

          for x := 0 to 4 do
          begin
            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClPlataforma)+IntToStr(IEFila)+':'+ColumnaNombre(IClPlataforma)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := '';
            VRango.NumberFormat := '@';
            //VRango.Interior.Color := $00BBBBBB;
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IclPda)+IntToStr(IEFila)+':'+ColumnaNombre(IclPda)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            //VRango.value := 'PDA';
            VRango.NumberFormat := '@';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IcltipoMov)+IntToStr(IEFila)+':'+ColumnaNombre(IcltipoMov)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            //VRango.value := 'N.E.C.';
            VRango.NumberFormat := '@';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClInicia)+IntToStr(IEFila)+':'+ColumnaNombre(IClInicia)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := '00:00';
            VRango.NumberFormat := '[hh]:mm';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza)+IntToStr(IEFila)+':'+ColumnaNombre(IClFinaliza)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := ('1/1/1900 0:00');
            VRango.NumberFormat := '[hh]:mm';
            //VRango.Interior.Color := $00BBBBBB;
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila)+':'+ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := '=+'+ColumnaNombre(IClFinaliza)+inttostr(IEFila)+'-'+ColumnaNombre(IClInicia)+inttostr(IEFila);
            VRango.NumberFormat := '0.00000000';
            Vrango.Borders.LineStyle := xlContinuous;





            IClRecurso := 7;
            while IClRecurso <= Ucolumna do
            begin
              VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso)+IntToStr(IEFila)+':'+ColumnaNombre(IClRecurso)+IntToStr(IEFila)];
              VRango.HorizontalAlignment :=xlCenter;
              VRango.VerticalAlignment := xlCenter;
              VRango.Font.Size := 8;
              VRango.Font.Bold := true;
              VRango.WrapText := True;
              VRango.NumberFormat := '0';
              VRango.value := '0';
              Vrango.Borders.LineStyle := xlContinuous;


              VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso+1)+IntToStr(IEFila)+':'+ColumnaNombre(IClRecurso+1)+IntToStr(IEFila)];
              VRango.HorizontalAlignment :=xlCenter;
              VRango.VerticalAlignment := xlCenter;
              VRango.Font.Size := 8;
              VRango.Font.Bold := false;
              VRango.WrapText := True;
              VRango.NumberFormat :='0.00000000';
              if tipo = 'Equipo' then
                VRango.value := '=+('+ColumnaNombre(IClRecurso)+inttostr(IEFila)+'*$F'+inttostr(IEFila)+')';
              if Tipo = 'Personal' then
                VRango.value := '=+('+ColumnaNombre(IClRecurso)+inttostr(IEFila)+'*$F'+inttostr(IEFila)+')*2';

              Vrango.Borders.LineStyle := xlContinuous;
              IClRecurso := IClRecurso + 2;
            end;
            Inc(IEFila);
          end;  
          IEFila := IEfila+4;

        end;
        ZEFolios.next;
      end;
    finally
      if info then
      begin
        try
          ZEMoeR.free;
        except
          ;
        end;
      end;
    end;
  end;
begin
//
  try
    Try
      EExcel := CreateOleObject('Excel.Application');
    Except
      raise Exception.Create('Error al enlazar con la aplicación de terceros (excel).');
    End;

    EExcel.Visible := True;
    EExcel.DisplayAlerts:= True;
    ELibro := EExcel.Workbooks.Add;
    IElibro := EExcel.Workbooks.count;

    for  IEHojaPersonal := Eexcel.Workbooks[ielibro].Sheets.Count  downto 2 do
      EExcel.workbooks[iELibro].workSheets[iEHojaPErsonal].Delete;

    if Epersonal then
    begin
      IEHojaPersonal := 1;
      EHojaPersonal := Eexcel.Workbooks[ielibro].Sheets[IEHojapersonal];
      EHojaPersonal.Name := 'CUADRE PERSONAL';

    end;

    if EEquipo then
    begin
      if not EPersonal then
      begin
        IEhojaEquipo := 1;
        EHojaEquipo := Eexcel.Workbooks[ielibro].Sheets[IEhojaEquipo];
        EHojaEquipo.Name := 'CUADRE EQUIPO';
      end
      else
      begin
        EHojaEquipo := EExcel.Workbooks[ielibro].Sheets.Add;
        IEhojaEquipo := 1;
        EHojaEquipo.Name := 'CUADRE EQUIPO';
        IEHojaPersonal := 2;
        EHojaPersonal := Eexcel.Workbooks[ielibro].Sheets[IEHojapersonal];
      end;
    end;
    try
      Progreso.Visible := True;
      if Info then
      begin
        zEMoe :=TZReadOnlyQuery.Create(nil);
        try
          zEMoe.Connection := connection.zConnection;
          zEMoe.Active := False;
          zEMoe.SQL.Text := ' SELECT iidmoe,dIdFecha FROM moe WHERE dIdFecha <= :Fecha AND sContrato = :Contrato order by didfecha desc limit 1 ';
          ZEMoe.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
          ZEMoe.ParamByName('Contrato').AsString := ContratoDiario;
          if info then
          begin
            zEMoe.Open;
            IMActivo := zEMoe.FieldByName('iidmoe').AsInteger;
          end;
        finally
          zEMoe.Free;
        end;


        ZEFolios := TZReadOnlyQuery.Create(nil);
        ZEFolios.Connection := connection.zConnection;
        ZEFolios.Active := False;
        ZEFolios.SQL.Text :=  'SELECT * FROM ordenesdetrabajo WHERE  sContrato = :Contrato';
        ZEFolios.ParamByName('Contrato').AsString := ContratoDiario;
        ZEFolios.Open;

        ZEActividades := TZReadOnlyQuery.Create(nil);
        ZEActividades.Connection := connection.zConnection;
        ZEActividades.Active := False;
        ZEActividades.SQL.Text :=  'SELECT * FROM actividadesxorden WHERE scontrato = :Contrato and snumeroorden = :Folio';
        ZEActividades.ParamByName('Contrato').asstring := ContratoDiario;
        //ZEActividades.ParamByName('Folio').AsString := tipo;
        //ZEActividades.Open;

        ZEMovimientos := TZReadOnlyQuery.Create(nil);
        ZEMovimientos.Connection := connection.zConnection;
        ZEMovimientos.Active := False;
        ZEMovimientos.SQL.Text :=  'select * from tiposdemovimiento where sClasificacion <> "movimiento de barco" and sContrato = :Contrato order by iorden ';
        ZEMovimientos.ParamByName('Contrato').asstring := ContratoDiario;
        ZEMovimientos.Open;
      end;

      if EPersonal then
      begin
        EExcel.Workbooks[iElibro].Worksheets[IEHojaPersonal].select;
        CreaEstructura('Personal');
      end;

      if EEquipo then
      begin
        EExcel.Workbooks[iElibro].Worksheets[IEhojaEquipo].select;
        CreaEstructura('Equipo');
      end;


    finally
      Progreso.Visible := False;
      EExcel.Visible := True;
      try
        ZEFolios.free;
      except
        ;
      end;
      try
        ZEActividades.free;
      except
        ;
      end;
      try
        ZEMovimientos.free;
      except
        ;
      end;
    end;
    //EHoja := EExcel.Workbooks[ielibro].Sheets.Add;
    //IEHoja := Eexcel.Workbooks[ielibro].Sheets.Count;
   // EHoja.Name := 'CUADRE EQUIPO';

   // ActiveWindow.SelectedSheets.Delete
  except
    on e:Exception do
      ShowMessage('No se pudo completar la generación de la plantilla por el siguiente motivo: '+#10+e.Message);

  end;

end;

procedure TfrmCuadrePersonalEquipo_Excel.ImportaCuadre(Pers,Equi,Sustituir:Boolean);
var
  ZMoe,ZmoeR,ZFolio,ZActiv,ZMovtos:TZReadOnlyQuery;
  CVeces:Integer;
  I,IdMoe,
  Ilibro, //Libro actual
  IPagina
  : Integer;

  PathArchivo:string;

  AExcel,ALibro,AHoja:Variant;

  EPagina:String;

  ErrorEncontrado:Boolean;

 {$Region'Direccion de archivo'}
  function DireccionPlantilla :Boolean;
  var ResDP:Boolean;
    DlgPat:TOpenDialog;
    Cancelar :Boolean;
  begin
    ResDP := False;
    Cancelar := False;
    try
      DlgPat := TOpenDialog.Create(nil);
      try
        DlgPat.Filter :=  'Archivo Excel  (*.xls,*xlsx)|*.XLS;*.XLSX';
        DlgPat.FilterIndex := 0;
        repeat
          if DlgPat.Execute then
          begin
            if (AnsiEndsText('.xls',lowercase(DlgPat.FileName))) or (AnsiEndsText('.xlsx',lowercase(DlgPat.FileName))) then
            begin
              PathArchivo := DlgPat.FileName;
              ResDP := True
            end
            else
              ShowMessage('El archivo seleccionado no corresponde al formato excel requerido por el sistema.'+#10+'Intente de nuevo o cancele el proceso porfavor.')
          end
          else
            Cancelar := True;
        until ResDP or Cancelar;
      finally
        DlgPat.Free;
      end;
    finally
      Result := ResDP;
    end;
  end;
  {$ENDREGION}

  {$REGION 'Buscar hoja y eliminaacentos'}
  Function BuscaHoja(ExcelImp:Variant;Nomb:String):Integer;
  var ResBuscaHoja,x:Integer;
  begin
    ResBuscaHoja := -1;
    try
      x := 1;
      while (x <= AExcel.WorkBooks[Ilibro].Sheets.Count) and (ResBuscaHoja = -1) do
      begin
        if AExcel.workbooks[iLibro].workSheets[x].visible then
          if LowerCase(trim(AExcel.workbooks[iLibro].workSheets[x].Name)) = LowerCase(Nomb) then
            ResBuscaHoja := x;
        Inc(x);
      end;
    finally
      Result := ResBuscaHoja;
    end;
  end;

  Function EliminaAcentos(Texto:string):string;
  const Acentos = 'áéíóúÁÉÍÓÚ'; NoAcentos = 'aeiouAEIOU';
  var i: integer;
  begin
    for i:= 1 to length(Texto) do
      begin
        if pos(Texto[i], Acentos) <> 0 then
        Texto:= StringReplace(Texto, Acentos[pos(Texto[i], Acentos)], NoAcentos[pos(Texto[i], Acentos)], [rfReplaceAll, rfIgnoreCase]);
      end;
    Result:= Texto;
  end;
  {$ENDREGION}

  procedure Lectura(Tipo:String);
  const
    ClPlataforma = 1;
    ClFolio = 3;
    ClPartida = 2;
    Clclasificacion = 3;
    ClHinicio = 4;
    ClHFin = 5;

    Vacio = 46;//Naranja
    NoExiste = 3;//Rojo
    NoImportar= 6;//amarillo
    ImportarL = 35; //verde
    NoEnCatalogo = 22;//  Folio no existe en el catalogo
    NoEnActividades = 40;//  Folio no existe en actividades

  var
    //Variables para guardar las cadenas
    sVPlataforma,
    sVPernocta,
    sVFolio,
    sVPartida,
    sVClasificacion,
    sVHinicio,
    sVHfin,
    sVCant,
    sVHh,
    sVIdRecurso,
    sVDescripcion
    :string;

    ClCant,Clhh,ClRecurso,
    iFila,
    iFilaRecurso,
    i,maxBlanca,lineab ,CActual
    :Integer;

    LineaLeyenda: Integer;

    ZRecursos:tzreadonlyquery;
    ZUptCuadre, zQCuentas: TZQuery;
    ZNotas:TZReadOnlyQuery;


    procedure crealeyenda(Doc:Variant);
    begin
      Doc.Workbooks[ilibro].Worksheets[ipagina].select;
      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Interior.Colorindex := Vacio;
      Doc.Selection.Value := 'LEYENDA';
      Doc.Selection.VerticalAlignment := xlCenter;
      Doc.Selection.Font.Size := 11;
      Doc.Selection.Font.Bold := True;
      Doc.Selection.WrapText := True;
      Doc.Selection.Interior.Color := $00BBBBBB;
      Doc.Selection.Borders.LineStyle := xlContinuous;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda+5)].Select;
      Doc.Selection.Interior.Colorindex := Vacio;
      Doc.Selection.VerticalAlignment := xlCenter;
      Doc.Selection.Font.Size := 11;
      Doc.Selection.Font.Bold := False;
      Doc.Selection.WrapText := True;
      Doc.Selection.Borders.LineStyle := xlContinuous;

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'No importar';
      Doc.Selection.Interior.Colorindex := NoImportar;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'Vacíos';
      Doc.Selection.Interior.Colorindex := Vacio;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'No existe';
      Doc.Selection.Interior.Colorindex := NoExiste ;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'Importable';
      Doc.Selection.Interior.Colorindex := ImportarL ;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'No en catálogo';
      Doc.Selection.Interior.Colorindex := NoEnCatalogo ;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'No en actividades';
      Doc.Selection.Interior.Colorindex := NoEnActividades ;
      Inc(linealeyenda);

    end;
  begin

    zQCuentas := TZQuery.Create(Self);
    zQCuentas.Connection := Connection.zConnection;

    zQCuentas.Active := False;
    zQCuentas.SQL.Clear;
    zQCuentas.SQL.Add('select sIdPernocta from cuentas limit 1');
    zQCuentas.Open;

    ZRecursos := TZReadOnlyQuery.Create(nil);
    ZRecursos.Connection := connection.zConnection;
    ZRecursos.Active := False;
    ZRecursos.SQL.Clear;
    if lowercase(Tipo) = 'equipo' then
      ZRecursos.SQL.Text := 'Select * from equipos where scontrato = :contrato and sidequipo = :idrecurso';
    if lowercase(Tipo) = 'personal' then
      ZRecursos.SQL.Text := 'Select * from personal where scontrato = :contrato and sidpersonal = :idrecurso';
    ZRecursos.ParamByName('contrato').AsString := ContratoDiario;
    {ZRecursos.Open;}
    try
      ZUptCuadre := TZQuery.Create(nil);
      ZUptCuadre.Connection := connection.zConnection;
      ZUptCuadre.Active := False;
      ZUptCuadre.SQL.Clear;
      if Sustituir then
      begin
        if lowercase(Tipo) = 'equipo' then
          ZUptCuadre.SQL.Text := 'delete from bitacoradeequipos where scontrato = :Contrato  and didfecha = :fecha'
        else
          ZUptCuadre.SQL.Text := 'delete from bitacoradepersonal where scontrato = :Contrato  and didfecha = :fecha';

        ZUptCuadre.ParamByName('contrato').AsString := ContratoDiario;
        ZUptCuadre.ParamByName('fecha').AsDateTime := cuadre_fechareporte;
        ZUptCuadre.ExecSQL;
      end;

      ZNotas := TZReadOnlyQuery.Create(nil);
      try
        ZNotas.Connection := connection.zConnection;
        ZNotas.Active := False;
        ZNotas.SQL.Clear;
        ZNotas.SQL.Text := ''+
                            'SELECT ' +
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


        try
          //Localizar el primer folio
          sVPartida := '';
          ifila := 1;
          ClCant := ClHFin+1;
          Clhh := Clcant+1;
          maxBlanca := 0;
          //Recorrer en busca de paquetes
          while (maxBlanca < 30)  do
          begin
            //modulo de datos
            sVPartida := AHoja.Cells[ifila,ClPartida].Text;
            sVClasificacion := AHoja.cells[ifila,ClClasificacion].Text;
            sVHinicio := AHoja.cells[ifila,ClHinicio].Text;

            if length(Trim(sVPartida)) + length(Trim(sVClasificacion)) + length(Trim(sVHinicio)) = 0 then
              Inc(maxBlanca);

            try
              if (lowercase(sVPartida) = 'pda') or (LowerCase(sVPartida) = 'pda.') or (LowerCase(sVClasificacion)= 'clasif') or (LowerCase(sVHinicio)= 'inicia') then
              begin
                maxBlanca := 0;
                iFilaRecurso := ifila-2;
                sVFolio := AHoja.Cells[ifila-1,ClFolio].Text;

                ZFolio.close;
                ZFolio.ParamByName('folio').AsString := sVFolio;
                ZFolio.open;
                if ZFolio.recordcount <> 1 then
                begin
                  AHoja.Cells[ifila-1,ClFolio].interior.colorindex := NoExiste;
                  AHoja.Cells[ifila-1,ClFolio].AddComment('No se encontró el fólio en la BD, paso al siguiente fólio.');
                  AHoja.Cells[ifila-1,ClFolio].Comment.Visible := True;
                  ErrorEncontrado := True;
                  Raise exception.create('siguiente paquete');
                end;

                zactiv.Close;
                ZActiv.ParamByName('folio').AsString := sVFolio;
                ZActiv.ParamByName('contrato').AsString := ZFolio.FieldByName('scontrato').AsString;
                ZActiv.Open;

                sVPlataforma := ZFolio.FieldByName('sidplataforma').AsString;
                sVPernocta := ZFolio.FieldByName('sidpernocta').AsString;


                Inc(ifila);
                sVPartida := AHoja.Cells[ifila,ClPartida].Text;
                sVClasificacion := AHoja.cells[ifila,ClClasificacion].Text;
                sVHinicio := AHoja.cells[ifila,ClHinicio].Text;
                sVHfin := AHoja.cells[ifila,ClHfin].Text;
                lineab := 0;
                //Recorrer contenido de paquete
                while (lineab < 2) and (svpartida <> 'pda') and (sVPartida <> 'pda.') do
                begin
                  try
                    sVPartida := AHoja.Cells[ifila,ClPartida].Text;
                    sVClasificacion := AHoja.cells[ifila,ClClasificacion].Text;
                    sVHinicio := AHoja.cells[ifila,ClHinicio].Text;
                    sVHfin := AHoja.cells[ifila,ClHfin].Text;

                    if (length(trim(sVPartida)) = 0) then
                    begin
                      AHoja.Cells[ifila,ClPartida].interior.colorindex :=  Vacio;
                      Inc(lineab);
                      raise Exception.Create('siguiente fila');
                    end;

                    if not ZActiv.Locate('snumeroactividad',sVPartida,[]) then
                    begin
                      AHoja.Cells[ifila,ClPartida].interior.colorindex :=  NoExiste;
                      AHoja.Cells[ifila,ClPartida].AddComment('Partida no asignada al folio.');
                      AHoja.Cells[ifila,ClPartida].Comment.Visible := True;
                    end;

                    if (length(trim(sVClasificacion)) = 0) then
                    begin
                      AHoja.Cells[ifila,ClClasificacion].interior.colorindex :=  Vacio;
                      Inc(lineab);
                      raise Exception.Create('siguiente fila');
                    end;

                    if not ZMovtos.Locate('sidtipomovimiento',sVClasificacion,[]) then
                    begin
                      AHoja.Cells[ifila,ClClasificacion].interior.colorindex :=  NoExiste;
                      AHoja.Cells[ifila,ClClasificacion].AddComment('El tipo de movimiento no esta dado de alta en el catálogo de su contrato.');
                      AHoja.Cells[ifila,ClClasificacion].Comment.Visible := True;
                      ErrorEncontrado := True;
                      raise Exception.Create('siguiente fila')
                    end;

                    if (length(trim(sVHinicio)) = 0) then
                    begin
                      AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                      Inc(lineab);
                      raise Exception.Create('siguiente fila');
                    end;

                    try
                      if sVHinicio <> '24:00' then
                        StrToTime(sVHinicio);
                    Except
                      on e:Exception do
                      begin
                        AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                        AHoja.Cells[ifila,ClHinicio].AddComment('El rango de valor aceptado es 00:00-24:00.');
                        AHoja.Cells[ifila,ClHinicio].Comment.Visible := True;
                        ErrorEncontrado := True;
                        raise Exception.Create('siguiente fila');
                      end;
                    end;

                    if (length(trim(sVHfin)) = 0) then
                    begin
                      AHoja.Cells[ifila,ClHFin].interior.colorindex :=  Vacio;
                      Inc(lineab);
                      raise Exception.Create('siguiente fila');
                    end;


                    try
                      if sVHfin <> '24:00' then
                        StrToTime(sVHfin);
                    Except
                      on e:Exception do
                      begin
                        AHoja.Cells[ifila,ClHFin].interior.colorindex :=  Vacio;
                        AHoja.Cells[ifila,ClHFin].AddComment('El rango de valor aceptado es 00:00-24:00.');
                        AHoja.Cells[ifila,ClHFin].Comment.Visible := True;
                        ErrorEncontrado := True;
                        raise Exception.Create('siguiente fila');
                      end;
                    end;

                    if (sVHinicio <>  '24:00') and (sVHfin <> '24:00') then
                    begin
                      if StrToTime(sVHinicio) > StrToTime(sVHfin) then
                      begin
                        AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                        AHoja.Cells[ifila,ClHinicio].AddComment('La hora de inicio no debe ser mayor a la hora de fin.');
                        AHoja.Cells[ifila,ClHinicio].Comment.Visible := True;
                        ErrorEncontrado := True;
                        raise Exception.Create('siguiente fila');
                      end;
                    end;

                    if sVHfin = '24:00' then
                    begin
                      if StrToTime(sVHinicio) > (StrToTime('23:59')+strtotime('00:01')) then
                      begin
                        AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                        AHoja.Cells[ifila,ClHinicio].AddComment('La hora de inicio no debe ser mayor a la hora de fin.');
                        AHoja.Cells[ifila,ClHinicio].Comment.Visible := True;
                        ErrorEncontrado := True;
                        raise Exception.Create('siguiente fila');
                      end;
                    end;

                    if (sVHinicio = '24:00') and (sVHfin = '24:00') then
                    begin
                      AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                      AHoja.Cells[ifila,ClHinicio].AddComment('La hora de inicio y fin no deben ser iguales a 24:00');
                      AHoja.Cells[ifila,ClHinicio].Comment.Visible := True;
                      ErrorEncontrado := True;
                      raise Exception.Create('siguiente fila');
                    end;

                    //Recorrer columnas
                    ClRecurso := 7;
                    sVIdRecurso :=  AHoja.cells[iFilaRecurso,Clrecurso].Text;
                    if length(trim(sVIdRecurso)) = 0 then
                    begin
                      AHoja.Cells[iFilaRecurso,ClRecurso].interior.colorindex := Vacio;
                      AHoja.Cells[iFilaRecurso,Clrecurso].AddComment('Es necesario un id de recurso y descripción favor de poner el encabezado al paquete de datos.');
                      ErrorEncontrado := True;
                      AHoja.Cells[iFilaRecurso,Clrecurso].Comment.Visible := True;
                      raise Exception.Create('siguiente paquete');
                    end;

                    while length(trim(sVIdRecurso)) > 0 do
                    begin
                      sVDescripcion := '';
                      try
                        sVIdRecurso :=  AHoja.cells[iFilaRecurso,Clrecurso].Text;
                        if length(trim(sVIdRecurso)) = 0 then
                        begin
                          if Clrecurso = 7 then
                          begin
                            AHoja.Cells[iFilaRecurso,ClRecurso].interior.colorindex := Vacio;
                            AHoja.Cells[iFilaRecurso,Clrecurso].AddComment('Es necesario un id de recurso y descripción favor de poner el encabezado al paquete de datos.');
                            AHoja.Cells[iFilaRecurso,Clrecurso].Comment.Visible := True;
                          end;
                          raise Exception.Create('siguiente linea');
                        end;

                        if not ZmoeR.Locate('sidrecurso',sVIdRecurso,[]) then
                        begin
                          AHoja.Cells[iFilaRecurso,ClRecurso].interior.colorindex :=  NoExiste;
                          AHoja.Cells[iFilaRecurso,Clrecurso].AddComment('No esta dado de alta en el moe vigente.');
                          AHoja.Cells[iFilaRecurso,Clrecurso].Comment.Visible := True;
                          ErrorEncontrado := True;
                          raise exception.Create('siguiente columna');
                        end;

                        sVDescripcion :=  AHoja.cells[iFilaRecurso-1,Clrecurso].Text;

                        ZRecursos.Close;
                        ZRecursos.ParamByName('idrecurso').AsString := sVIdRecurso;
                        ZRecursos.Open;

                        if ZRecursos.RecordCount = 0 then
                        begin
                          AHoja.Cells[iFilaRecurso,ClRecurso].interior.colorindex :=  NoExiste;
                          AHoja.Cells[iFilaRecurso,Clrecurso].AddComment('No esta dado de alta en el catalogo de '+tipo+' por contrato.');
                          AHoja.Cells[iFilaRecurso,Clrecurso].Comment.Visible := True;
                          ErrorEncontrado := True;
                          raise exception.Create('siguiente columna');
                        end;

                        if length(trim(sVDescripcion)) = 0 then
                        begin
                          AHoja.Cells[iFilaRecurso-1,ClRecurso].interior.colorindex :=  Vacio;
                          AHoja.Cells[iFilaRecurso-1,ClRecurso].AddComment('Es obligatorio una descripción.');
                          AHoja.Cells[iFilaRecurso-1,ClRecurso].Comment.Visible := True;
                          ErrorEncontrado := True;
                          raise exception.Create('siguiente columna');
                        end;

                        lineab := 0;

                        sVCant := VarToStr(trim(AHoja.cells[iFila,Clrecurso].value));
                        sVHh := vartostr(trim(AHoja.cells[iFila,Clrecurso+1].value));

                        if (Length(trim(sVCant)) = 0) or (Length(trim(sVHh)) = 0) then
                        begin
                          AHoja.Cells[iFila,ClRecurso].interior.colorindex :=  Vacio;
                          AHoja.Cells[iFila,ClRecurso].AddComment('Se ignora por tener valor nulo en cantidad o hh.');
                          AHoja.Cells[iFila,ClRecurso].Comment.Visible := False;
                          raise Exception.Create('siguiente columna');
                        end;

                        if StrToFloat(sVHh) = 0 then
                        begin
                          AHoja.Cells[iFila,ClRecurso+1].interior.colorindex :=  Vacio;
                          AHoja.Cells[iFila,ClRecurso+1].AddComment('Se ignora por tener valor = 0');
                          AHoja.Cells[iFila,ClRecurso+1].Comment.Visible := False;
                          raise Exception.Create('siguiente columna');
                        end;

                        ZUptCuadre.Active := False;
                        ZUptCuadre.SQL.Clear;
                        if LowerCase(tipo) = 'equipo' then
                          ZUptCuadre.SQL.Text := 'select * from bitacoradeequipos where scontrato = :contrato and didfecha = :fecha'
                        else
                          ZUptCuadre.SQL.Text := 'select * from bitacoradepersonal where scontrato = :contrato and didfecha = :fecha';
                        ZUptCuadre.ParamByName('contrato').AsString := ContratoDiario;
                        ZUptCuadre.ParamByName('fecha').AsDateTime := Cuadre_FechaReporte;
                        ZUptCuadre.Open;

                        ZNotas.Active := False;
                        ZNotas.ParamByName('fecha').AsDateTime := Cuadre_FechaReporte;
                        ZNotas.ParamByName('orden').AsString := sVFolio;
                        ZNotas.ParamByName('actividad').AsString := sVPartida;
                        ZNotas.ParamByName('contrato').AsString := ContratoDiario;
                        ZNotas.ParamByName('sidclasificacion').AsString := sVClasificacion;
                        ZNotas.Open;

                        if (ZNotas.RecordCount = 0) or (ZNotas.FieldByName('iiddiario').AsInteger < 0) then
                        begin
                          AHoja.Cells[iFila,ClRecurso+1].AddComment('No se encontro nota para asignar el recurso.');
                          AHoja.Cells[iFila,ClRecurso+1].interior.colorindex :=  NoEnActividades;
                          ErrorEncontrado := True;
                          raise Exception.Create('siguiente columna');
                        end;

                        try
                          ZUptCuadre.Append;
                          ZUptCuadre.FieldByName('scontrato').AsString:= ContratoDiario;
                          ZUptCuadre.FieldByName('didfecha').AsDateTime := Cuadre_FechaReporte;
                          ZUptCuadre.fieldbyname('snumeroorden').AsString := sVFolio;
                          ZUptCuadre.FieldByName('iiddiario').AsInteger := ZNotas.FieldByName('iiddiario').AsInteger;
                          ZUptCuadre.FieldByName('swbs').AsString := ZNotas.FieldByName('swbs').AsString;
                          if LowerCase(Tipo) = 'equipo' then
                          begin
                            ZUptCuadre.FieldByName('sidequipo').AsString := sVIdRecurso;
                          end
                          else
                          begin
                            ZUptCuadre.FieldByName('sidpersonal').AsString := sVIdRecurso;
                            ZUptCuadre.FieldByName('sidplataforma').AsString := sVPlataforma;
                            ZUptCuadre.FieldByName('stipopernocta').AsString := zQCuentas.FieldByName('sIdPernocta').AsString;
                            ZUptCuadre.FieldByName('sagrupapersonal').AsString := '*';
                            ZUptCuadre.FieldByName('laplicapernocta').AsString := 'Si';
                          end;
                          ZUptCuadre.FieldByName('dcanthhgenerador').AsFloat := 0;
                          ZUptCuadre.FieldByName('iitemorden').AsInteger := ZRecursos.FieldByName('iitemorden').AsInteger;
                          ZUptCuadre.FieldByName('sdescripcion').AsString := sVDescripcion;
                          ZUptCuadre.FieldByName('sidpernocta').AsString := sVPernocta;
                          ZUptCuadre.FieldByName('stipoobra').AsString := 'PU';
                          ZUptCuadre.FieldByName('shorainicio').AsString := sVHinicio;
                          ZUptCuadre.FieldByName('shorafinal').AsString := sVHfin;
                          ZUptCuadre.FieldByName('dcantidad').AsFloat := StrToFloat(svcant);
                          ZUptCuadre.FieldByName('snumeroactividad').AsString := sVPartida;
                          ZUptCuadre.FieldByName('dsolicitado').asfloat := 0;
                          ZUptCuadre.FieldByName('sfactor').AsString := '0';
                          ZUptCuadre.FieldByName('dcostomn').AsInteger :=0;
                          ZUptCuadre.FieldByName('dcostodll').AsInteger :=0;
                          ZUptCuadre.FieldByName('dcanthh').AsFloat := StrToFloat(sVHh);
                          ZUptCuadre.FieldByName('dajuste').AsInteger := 0;
                          ZUptCuadre.Post;

                          AHoja.Cells[iFila,ClRecurso+1].interior.colorindex :=  ImportarL;

                        except
                          on e:exception do
                          begin
                            AHoja.Cells[iFila,ClRecurso+1].AddComment('Error encontrado: '+e.message);
                            AHoja.Cells[iFila,ClRecurso+1].Comment.Visible := false;
                            ErrorEncontrado := True;
                          end;

                        end;



                      except
                        on e:Exception do
                          if e.Message = 'siguiente columna' then
                            ;
                      end;
                      ClRecurso := ClRecurso+2;
                    end;


                  except
                    on e:Exception do
                    if e.message = 'siguiente fila' then
                      ;
                  end;
                  Inc(ifila);
                end;

              end;
            except
              on e:exception do
                if e.message <> 'siguiente paquete' then
                   raise;

            end;

            Inc(iFila);

            if maxBlanca = 29 then
            begin
              LineaLeyenda := iFila;
              CreaLeyenda(AExcel);
            end;
          end;



        finally
          ZUptCuadre.free;
        end;

      finally
        ZNotas.Free;
      end;

    finally
      ZRecursos.Free;
      zQCuentas.Free;
    end;
  end;


begin
  try
    ZMoe := TZReadOnlyQuery.Create(nil);
    try
      if not Equi and not Pers then
        raise Exception.Create('No seleccionó opciones de importacion.');

      //Abrir excel
      if not DireccionPlantilla then
        raise exception.create('proceso cancelado por el usuario');

      ZMoe.Connection := connection.zConnection;
      ZMoe.Active := False;
      ZMoe.SQL.Clear;
      ZMoe.SQL.Text :=  ' SELECT iidmoe,dIdFecha FROM moe WHERE dIdFecha <= :Fecha AND sContrato = :Contrato  order by dIdFecha desc limit 1 ';
      ZMoe.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
      ZMoe.ParamByName('Contrato').AsString := ContratoDiario;
      ZMoe.Open;

      IdMoe := ZMoe.FieldByName('iidmoe').AsInteger;

    finally
      ZMoe.Free;
    end;


    CVeces:=0;
    if Pers then
      Inc(Cveces);
    if Equi then
      inc(CVeces);

    try
      AExcel:=CreateOleObject('Excel.Application');
      AExcel.Visible := False;
    except
      raise Exception.Create('No tiene instalado microsoft excel o bién ocurre un problema con el mismo.');
    end;

    AExcel.Workbooks.Open(PathArchivo);
    ALibro := AExcel.Workbooks[AExcel.Workbooks.count];
    ILibro := AExcel.Workbooks.count;

    ZmoeR := TZReadOnlyQuery.Create(nil);
    try
      ZMoeR.Connection := connection.zConnection;
      ZMoeR.Active := False;
      ZMoeR.SQL.Clear;
      ZMoeR.SQL.Text :=  'SELECT ' +
                    '	* ' +
                    'FROM moerecursos AS mre ' +
                    'WHERE mre.iIdMoe = :IdMoe AND eTipoRecurso = :TipoRecurso ';
      ZMoeR.ParamByName('IdMoe').AsInteger := IdMoe;

      ZFolio := TZReadOnlyQuery.Create(nil);
      try
        ZFolio.Connection := connection.zConnection;
        ZFolio.Active := False;
        ZFolio.SQL.Clear;
        ZFolio.SQL.Text := 'SELECT * FROM ordenesdetrabajo WHERE sNumeroOrden = :Folio AND sContrato = :Contrato';
        ZFolio.ParamByName('Contrato').AsString := ContratoDiario;

        ZActiv := TZReadOnlyQuery.Create(nil);
        try
          ZActiv.Connection := connection.zConnection;
          ZActiv.Active := False;
          ZActiv.SQL.Clear;
          ZActiv.SQL.Text := 'SELECT * FROM actividadesxorden WHERE scontrato = :Contrato and snumeroorden = :Folio';

          ZMovtos := Tzreadonlyquery.create(nil);
          try
            ZMovtos.Connection := connection.zConnection;
            ZMovtos.Active := False;
            ZMovtos.SQL.Clear;
            ZMovtos.SQL.Text := 'select * from tiposdemovimiento where sClasificacion <> "movimiento de barco" and sContrato = :Contrato order by iorden';
            ZMovtos.parambyname('contrato').asstring := ContratoDiario;
            ZMovtos.Open;



            EPagina := '';
            //Error de pagina
            for I := 0 to CVeces-1 do
            begin
              ErrorEncontrado := False;
              Ipagina := -1;
              try
                if Pers and (i = 0) then
                begin
                  //importar personal
                  //buscar hoja cuadre p
                  IPagina:= BuscaHoja(AExcel,'cuadre personal');

                  if Ipagina = -1 then
                    raise exception.create('No se encontró la pestaña con el texto cuadre personal');

                  AHoja := ALibro.worksheets[IPagina];
                  Zmoer.close;
                  ZMoeR.ParamByName('TipoRecurso').AsString := 'Personal';
                  ZMoeR.Open;

                  if ZmoeR.RecordCount = 0 then
                    raise Exception.Create('No existe personal registrado en el oficio de movimiento de personal y equipo.');

                  if Zmoer.Connection.InTransaction then
                    ZmoeR.Connection.Rollback;
                  ZmoeR.Connection.StartTransaction;

                  Lectura('personal');



                  if ErrorEncontrado then
                  begin
                    ZmoeR.Connection.Rollback;
                    ShowMessage('Se encontraron errores en la pestaña de cuadre personal, favor de corregir.');
                  end
                  else
                  begin
                    ZmoeR.Connection.Commit;
                    ShowMessage('Se importaron registros para el cuadre de personal.');
                  end;
                end;

                if (Equi and (i = 1)) or (Equi and (i = 0) and (not pers)) then
                begin
                  //importar equipo  cuadre e
                  //buscar hoja cuadre e
                  IPagina:= BuscaHoja(AExcel,'cuadre equipo');
                  if Ipagina = -1 then
                    raise exception.create('No se encontró la pestaña con el texto cuadre equipo');

                  AHoja := ALibro.worksheets[IPagina];
                  zmoer.close;
                  ZMoeR.ParamByName('TipoRecurso').AsString := 'Equipo';
                  ZMoeR.Open;

                  if ZmoeR.RecordCount = 0 then
                    raise Exception.Create('No existe equipo registrado en el oficio de movimiento de personal y equipo.');

                  if Zmoer.Connection.InTransaction then
                    ZmoeR.Connection.Rollback;

                  ZmoeR.Connection.StartTransaction;

                  Lectura('equipo');


                  if ErrorEncontrado then
                  begin
                    ZmoeR.Connection.Rollback;
                    ShowMessage('Se encontraron errores en la pestaña de cuadre equipo, favor de corregir.');
                  end
                  else
                  begin
                    ZmoeR.Connection.Commit;
                    ShowMessage('Se importaron registros para el cuadre de equipos');
                  end;
                  //Lectura('Equipo',ZmoeR,ZFolio,ZActiv,ALibro,AExcel,Ahoja)
                end;

              except
                on e:Exception do
                  ShowMessage('Ocurrió el siguiente error al tratar de importar el cuadre '+e.Message);
              end;
            end;
          finally
            zmovtos.free;
          end;
        finally
          Zactiv.free;
        end;
      finally
        Zfolio.free;
      end;

    finally
      ZmoeR.Free;
      AExcel.Visible := True;
      AExcel := Unassigned;
    end;

  finally

  end;
end;

procedure TfrmCuadrePersonalEquipo_Excel.btn1Click(Sender: TObject);
begin
  if chkpartidas.Checked then
  begin
    ExportarPlantillaActividades(ChkEPersonal.Checked, ChkEEquipo.Checked, ChkEInfo.Checked);
    exit;
  end
  else
    ExportarPlantilla(ChkEPersonal.Checked, ChkEEquipo.Checked, ChkEInfo.Checked);
end;

procedure TfrmCuadrePersonalEquipo_Excel.btnCmdImportarClick(Sender: TObject);
begin
  {if ChkVer2.Checked then
    ImportarCuadre()
  else   }
  //LeerCuadre;
  if cbx2.Checked then
    ImportaCuadre(cmbSoloPersonal.Checked,cmbSoloEquipo.Checked,cmbReemplazar.Checked)
  else
  if chkHorasExtras.Checked then
    ImportaCuadreHE
  else
    LeerCuadre;
end;

procedure TfrmCuadrePersonalEquipo_Excel.FormShow(Sender: TObject);
begin
  Label2.Caption := FormatDateTime('dd-mm-yyyy', Cuadre_FechaReporte);
  if Connection.Contrato.FieldByName('sIdResidencia').AsString = '03' then
    ChkHorario.Checked := True
  else
    ChkHorario.Checked := False;
  CargarCheckCombo(CmbFolios);  
end;

procedure TfrmCuadrePersonalEquipo_Excel.CargarCheckCombo(ComboFolios: TJvCustomCheckedComboBox);
var
  QrFolios:TZReadOnlyQuery;
begin
  ComboFolios.Clear;
  QrFolios:=TZReadOnlyQuery.Create(nil);
  QrFolios.Connection:=connection.zConnection;

  try
    with QrFolios,CmbFolios do
    begin
      Connection:=frm_connection.connection.zConnection;
      SQL.Text:='select ot.sNumeroorden from ordenesdetrabajo ot  ' +
                  'where ot.sContrato=:ContratoNormal group by ot.sContrato,ot.sNumeroorden ';
      ParamByName('ContratoNormal').AsString :=ContratoDiario;
      Open;

      while not Eof do
      begin
        ComboFolios.Items.Add(FieldByName('snumeroorden').AsString);
        Next;
      end;
    end;
  finally
    FreeAndNil(QrFolios);
  end;
end;

procedure TfrmCuadrePersonalEquipo_Excel.chkHorasExtrasClick(Sender: TObject);
begin
  if chkHorasExtras.Checked then
  begin
    Cbx2.Checked := False;
    cmbSoloPersonal.Checked := False;
    cmbSoloEquipo.Checked :=  False;
    cmbReemplazar.Checked := False;

    chkActualizaHE.Enabled := True;
  end;
end;

procedure TfrmCuadrePersonalEquipo_Excel.ImportarCuadre();
const
  iColumnaPartida = 2;
Var
  //referencias de excel
  Excel,
  Libro,
  Hoja: Variant;

  Moe,Actividades,ZActividades:TZReadOnlyQuery;
  ErrorEnPlantilla:Boolean;

  Vacio, //lineas vacias
  iFila,  //Contador de filas
  iLeerPaginas, //iniciar a leer desde pagina
  iFilaItem,
  IdDiario     //almacena el id del moe actual
  :Integer;

  Valor,  //Valor actual
  Folio, //Almacena folio leido
  Partida //almacena partida leida
  :string;
  i,
  iHojas,
  iHojasTotales,
  iColumna,
  iColumnaItem,
  iColumnaHorarioInicio,
  iColumnaClasificacion,

  iFilaPartida,


  IdMoeActivo,
  IdDiarioAux,
  BarraAvance,
  TotalPartidas,
  TotalColumnas: Integer;

  dCantidad, dHH: Real;

  sTipoCuadre,

  Valor2,
  ListaFolios,

  Item,
  ItemDescripcion, 
  Cantidad,
  HorasHombre,
  ArchivoExcel,
  sHoraInicio,
  sHoraFinal,
  sIdClasificacion,
  Tabla,
  sIdPlataforma,
  sIdPernocta,
  sWbs, sNumeroActividad: String;

  SQLExtra,
  IgnorarRecurso: TStringList;

  Founded, Error: Boolean;
  Query, Query2: TZQuery;

  HuboErrores,Colorear:Boolean;

  Function ExcelCloseWorkBooks(Excel : Variant; SaveAll: Boolean): Boolean;
  var
    loop: byte;
  Begin
    Result := True;
    Try
      For loop := 1 to Excel.Workbooks.Count Do
        Excel.Workbooks[1].Close[SaveAll];
    Except
      Result := False;
    End;
  End;

  Function ExcelClose(Excel : Variant; SaveAll: Boolean): Boolean;
  Begin
    Result := True;
    Try
      ExcelCloseWorkBooks(Excel, SaveAll);
      Excel.Quit;
    Except
      MessageDlg('No se puede cerrar la aplicación excel.', mtError, [mbOK], 0);
      Result := False;
    End;
  End;
begin

  try
    Try
      Excel := CreateOleObject('Excel.Application');
    Except
      raise Exception.Create('Error al enlazar con la aplicación de terceros (excel).');
    End;

    Excel.Visible := False;
    Excel.DisplayAlerts:= False;

    if OpenXLS.Execute then
    begin
      ArchivoExcel := OpenXLS.FileName;
      Libro := Excel.WorkBooks.Open(ArchivoExcel);
    end
    else
      raise Exception.Create('Proceso cancelado por el usuario.');

    ErrorEnPlantilla := False; //Determina si se visualiza el excel a lo ultimo

    HuboErrores:=False;
    Moe := TZReadOnlyQuery.Create(nil);
    try
      Moe.Connection := Connection.zConnection;
      Query2 := TZQuery.Create(nil);
      try
        Query2.Connection := Connection.zConnection;
        SQLExtra := TStringList.Create;
        try
          IgnorarRecurso := TStringList.Create;
          ZActividades := TZReadOnlyQuery.Create(nil);
          try
            ZActividades.Connection := connection.zConnection;
            //Leer moe obteniendo personal y equipo
            Moe.Active := False;
            Moe.SQL.Clear;
            Moe.SQL.Add('' +
                          'SELECT ' +
                          '	* ' +
                          'FROM moe AS m ' +
                          '	INNER JOIN moerecursos AS mr ' +
                          '		ON (mr.iIdMoe = m.iIdMoe) ' +
                          'WHERE m.dIdFecha = ( SELECT MAX(dIdFecha) FROM moe WHERE dIdFecha <= :Fecha AND sContrato = :Contrato) ' +
                          ' AND sContrato = :Contrato ' +
                          'LIMIT 1 ');
            Moe.Params.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
            Moe.Params.ParamByName('Contrato').AsString := ContratoDiario;
            Moe.Open;

            if Moe.RecordCount = 0 then
              raise Exception.Create('No existen movimientos de personal y equipo en el sistema, crée uno para continuar.');

            IdMoeActivo := Moe.FieldByName('iIdMoe').AsInteger;

            ZActividades.Active := False;
            ZActividades.SQL.Clear;
            ZActividades.SQL.Text := 'select * from bitacoradeactividades where dIdFecha = :didfecha and '+
                                     'sContrato = :scontrato and sidturno = "A" and sIdTipoMovimiento = "ED"';
            ZActividades.Params.ParamByName('sContrato').AsString := ContratoDiario;
            ZActividades.Params.ParamByName('dIdFecha').AsDateTime := Cuadre_FechaReporte;
            ZActividades.Open;

            if ZActividades.RecordCount = 0 then
              raise Exception.Create('No hay actividades registradas para esa fecha.');

            iHojasTotales := Libro.Sheets.Count;

            if (Not cmbSoloEquipo.Checked) AND (Not cmbSoloPersonal.Checked) then
              raise Exception.Create('Es necesario que identifique que quiere importar, personal, equipo o ambos.');
          

            if (cmbSoloPersonal.Checked) and (cmbreemplazar.checked)then
            begin
              iLeerPaginas := 1;
              Query2.Active := False;
              Query2.SQL.Clear;
              Query2.SQL.Text := 'DELETE FROM bitacoradepersonal WHERE dIdFecha = :Fecha AND sContrato = :Contrato AND sTipoObra = "PU"';
              Query2.Params.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
              Query2.Params.ParamByName('Contrato').AsString := ContratoDiario;
              Query2.ExecSQL;
            end;

            if (cmbSoloEquipo.Checked) and (cmbreemplazar.checked) then
            begin
              Query2.Active := False;
              Query2.SQL.Clear;
              Query2.SQL.Text := 'DELETE FROM bitacoradeequipos WHERE dIdFecha = :Fecha AND sContrato = :Contrato';
              Query2.Params.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
              Query2.Params.ParamByName('Contrato').AsString := ContratoDiario;
              Query2.ExecSQL;
            end;



            if cmbsolopersonal.checked then
              iLeerPaginas := 1;

            if (cmbSoloEquipo.Checked) AND (Not cmbSoloPersonal.Checked) then
              iLeerPaginas := 2;

            //Recorrer todas las hojas desde la inicial
            for iHojas := iLeerPaginas to Libro.Sheets.Count do
            begin
              IgnorarRecurso.Clear;

              if (cmbSoloPersonal.Checked) AND (Not cmbSoloEquipo.Checked) then
                if iHojas = 2 then
                  Continue;

              if (cmbSoloEquipo.Checked) AND (Not cmbSoloPersonal.Checked) then
                if iHojas = 1 then
                  Continue;

              if iHojas = 1 then
              begin
                sTipoCuadre := 'Personal'
              end
              else
              if iHojas = 2 then
              begin
                sTipoCuadre := 'Equipo';
              end
              else
              begin
                Exit;
              end;

              Libro.WorkSheets[iHojas].Activate;

              Moe.Active := False;
              Moe.SQL.Clear;
              Moe.SQL.Add('' +
                            'SELECT ' +
                            '	* ' +
                            'FROM moerecursos AS mre ' +
                            'WHERE mre.iIdMoe = :IdMoe AND eTipoRecurso = :TipoRecurso');
              Moe.ParamByName('IdMoe').AsInteger := IdMoeActivo;
              Moe.ParamByName('TipoRecurso').AsString := sTipoCuadre;
              Moe.Open;




              //Leer excel hasta que el limite de lineas vacias sean 30
              Vacio := 0;
              iFila := 1; //iniciar desde la fila1
              iFilaItem := 0;
              //validaciones
              while Vacio < 30 do
              begin
                Excel.Range[ColumnaNombre(iColumnaPartida)+IntToStr(iFila)+':'+ColumnaNombre(iColumnaPartida)+IntToStr(iFila)].Select;
                Valor := Excel.Selection.Value;
                if Valor = '' then
                begin
                  Inc(Vacio); //si el valor es nulo entonces es una linea vacia
                end
                else
                begin
                  Vacio := 0; //si hay algun valor entonces reestablecer los vacias
                  if (Valor = 'PDA') OR (Valor = 'PDA.') then
                  begin
                    if iFilaItem = 0 then
                    begin
                      iFilaItem := iFila - 2;
                    end;
                    iFilaPartida := iFila;
                    Partida := Excel.Cells[(iFila - 1), (iColumnaPartida + 1)].Text;
                    Folio := Trim(Partida);



                  end;
                end;




                Inc(iFila);
              end;
            end;

          finally
            ZActividades.Free;
          end;




        finally
          SQLExtra.free;
        end;
      finally
        Query2.Free;
      end;
    finally
      Query.free;
    end;

    ExcelClose(Excel,False);

  except
    on e:exception do
    begin
      ShowMessage('No se pudo importar el formato por el siguiente motivo: '+e.message);
      try  //Si ocurren errores mostrarlo
        Excel.visible := True;
      except
        ExcelClose(Excel,False);
      end;
    end;

  end;

end;

procedure TfrmCuadrePersonalEquipo_Excel.ImportaCuadreHE;
var
  ZGPersonal, ZFolio,ZActiv,ZMovtos:TZReadOnlyQuery;
  CVeces:Integer;
  I,IdMoe,
  Ilibro, //Libro actual
  IPagina
  : Integer;

  Pers, sustituir : Boolean;

  PathArchivo:string;

  AExcel,ALibro,AHoja:Variant;

  EPagina:String;

  ErrorEncontrado:Boolean;

  {$Region'Direccion de archivo'}
  function DireccionPlantilla :Boolean;
  var ResDP:Boolean;
    DlgPat:TOpenDialog;
    Cancelar :Boolean;
  begin
    ResDP := False;
    Cancelar := False;
    try
      DlgPat := TOpenDialog.Create(nil);
      try
        DlgPat.Filter :=  'Archivo Excel  (*.xls,*xlsx)|*.XLS;*.XLSX';
        DlgPat.FilterIndex := 0;
        repeat
          if DlgPat.Execute then
          begin
            if (AnsiEndsText('.xls',lowercase(DlgPat.FileName))) or (AnsiEndsText('.xlsx',lowercase(DlgPat.FileName))) then
            begin
              PathArchivo := DlgPat.FileName;
              ResDP := True
            end
            else
              ShowMessage('El archivo seleccionado no corresponde al formato excel requerido por el sistema.'+#10+'Intente de nuevo o cancele el proceso porfavor.')
          end
          else
            Cancelar := True;
        until ResDP or Cancelar;
      finally
        DlgPat.Free;
      end;
    finally
      Result := ResDP;
    end;
  end;
  {$ENDREGION}

  {$REGION 'Buscar hoja y eliminaacentos'}
  Function BuscaHoja(ExcelImp:Variant;Nomb:String):Integer;
  var ResBuscaHoja,x:Integer;
  begin
    ResBuscaHoja := -1;
    try
      x := 1;
      while (x <= AExcel.WorkBooks[Ilibro].Sheets.Count) and (ResBuscaHoja = -1) do
      begin
        if AExcel.workbooks[iLibro].workSheets[x].visible then
          if LowerCase(trim(AExcel.workbooks[iLibro].workSheets[x].Name)) = LowerCase(Nomb) then
            ResBuscaHoja := x;
        Inc(x);
      end;
    finally
      Result := ResBuscaHoja;
    end;
  end;

  Function EliminaAcentos(Texto:string):string;
  const Acentos = 'áéíóúÁÉÍÓÚ'; NoAcentos = 'aeiouAEIOU';
  var i: integer;
  begin
    for i:= 1 to length(Texto) do
      begin
        if pos(Texto[i], Acentos) <> 0 then
        Texto:= StringReplace(Texto, Acentos[pos(Texto[i], Acentos)], NoAcentos[pos(Texto[i], Acentos)], [rfReplaceAll, rfIgnoreCase]);
      end;
    Result:= Texto;
  end;
  {$ENDREGION}

  procedure Lectura(Tipo :String);
  const
    ClPlataforma = 1;
    ClFolio = 3;
    ClPartida = 2;
    Clclasificacion = 3;
    ClHinicio = 4;
    ClHFin = 5;

    Vacio = 46;//Naranja
    NoExiste = 3;//Rojo
    NoImportar= 6;//amarillo
    ImportarL = 43; //verde
    NoEnCatalogo = 22;//  Folio no existe en el catalogo
    NoEnActividades = 40;//  Folio no existe en actividades

  var
    //Variables para guardar las cadenas
    sVPlataforma,
    sVPernocta,
    sVFolio,
    sVPartida,
    sVClasificacion,
    sVHinicio,
    sVHfin,
    sVCant,
    sVHh,
    sVIdRecurso,
    sVDescripcion
    :string;

    ClCant,Clhh,ClRecurso,
    iFila,
    iFilaRecurso,
    i,maxBlanca,lineab ,CActual
    :Integer;

    LineaLeyenda: Integer;

    ZRecursos:tzreadonlyquery;
    ZUptCuadre, zQCuentas : TZQuery;
    ZNotas:TZReadOnlyQuery;


    procedure crealeyenda(Doc:Variant);
    begin
      Doc.Workbooks[ilibro].Worksheets[ipagina].select;
      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Interior.Colorindex := Vacio;
      Doc.Selection.Value := 'LEYENDA';
      Doc.Selection.VerticalAlignment := xlCenter;
      Doc.Selection.Font.Size := 11;
      Doc.Selection.Font.Bold := True;
      Doc.Selection.WrapText := True;
      Doc.Selection.Interior.Color := $00BBBBBB;
      Doc.Selection.Borders.LineStyle := xlContinuous;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda+5)].Select;
      Doc.Selection.Interior.Colorindex := Vacio;
      Doc.Selection.VerticalAlignment := xlCenter;
      Doc.Selection.Font.Size := 11;
      Doc.Selection.Font.Bold := False;
      Doc.Selection.WrapText := True;
      Doc.Selection.Borders.LineStyle := xlContinuous;

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'No importar';
      Doc.Selection.Interior.Colorindex := NoImportar;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'Vacíos';
      Doc.Selection.Interior.Colorindex := Vacio;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'No existe';
      Doc.Selection.Interior.Colorindex := NoExiste ;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'Importable';
      Doc.Selection.Interior.Colorindex := ImportarL ;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'No en catálogo';
      Doc.Selection.Interior.Colorindex := NoEnCatalogo ;
      Inc(linealeyenda);

      Doc.Range['J'+inttostr(LineaLeyenda)+':J'+inttostr(LineaLeyenda)].Select;
      Doc.Selection.Value := 'No en actividades';
      Doc.Selection.Interior.Colorindex := NoEnActividades ;
      Inc(linealeyenda);

    end;
  begin

    zQCuentas := TZQuery.Create(Self);
    zQCuentas.Connection := Connection.zConnection;

    zQCuentas.Active := False;
    zQCuentas.SQL.Clear;
    zQCuentas.SQL.Add('select sIdPernocta from cuentas limit 1');
    zQCuentas.Open;

    ZRecursos := TZReadOnlyQuery.Create(nil);
    ZRecursos.Connection := connection.zConnection;
    ZRecursos.Active := False;
    ZRecursos.SQL.Clear;
    ZRecursos.SQL.Text := 'Select * from personal where scontrato = :contrato and sidpersonal = :idrecurso';
    ZRecursos.ParamByName('contrato').AsString := ContratoDiario;
    try
      ZUptCuadre := TZQuery.Create(nil);
      ZUptCuadre.Connection := connection.zConnection;
      ZUptCuadre.Active := False;
      ZUptCuadre.SQL.Clear;
      if chkActualizaHE.Checked then
      begin
        ZUptCuadre.Active := False;
        ZUptCuadre.SQL.Text := 'delete from horasextras where scontrato = :Contrato  and didfecha = :fecha';
        ZUptCuadre.ParamByName('contrato').AsString := ContratoDiario;
        ZUptCuadre.ParamByName('fecha').AsDateTime := cuadre_fechareporte;
        ZUptCuadre.ExecSQL;
      end;

      ZNotas := TZReadOnlyQuery.Create(nil);
      try
        ZNotas.Connection := connection.zConnection;
        ZNotas.Active := False;
        ZNotas.SQL.Clear;
        ZNotas.SQL.Text := ''+
                            'SELECT ' +
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


        try
          //Localizar el primer folio
          sVPartida := '';
          ifila := 1;
          ClCant := ClHFin+1;
          Clhh := Clcant+1;
          maxBlanca := 0;
          //Recorrer en busca de paquetes
          while (maxBlanca < 30)  do
          begin
            //modulo de datos
            sVPartida := AHoja.Cells[ifila,ClPartida].Text;
            sVClasificacion := AHoja.cells[ifila,ClClasificacion].Text;
            sVHinicio := AHoja.cells[ifila,ClHinicio].Text;

            if length(Trim(sVPartida)) + length(Trim(sVClasificacion)) + length(Trim(sVHinicio)) = 0 then
              Inc(maxBlanca);

            try
              if (lowercase(sVPartida) = 'pda') or (LowerCase(sVPartida) = 'pda.') or (LowerCase(sVClasificacion)= 'clasif') or (LowerCase(sVHinicio)= 'inicia') then
              begin
                maxBlanca := 0;
                iFilaRecurso := ifila-2;
                sVFolio := AHoja.Cells[ifila-1,ClFolio].Text;

                ZFolio.close;
                ZFolio.ParamByName('folio').AsString := sVFolio;
                ZFolio.open;
                if ZFolio.recordcount <> 1 then
                begin
                  AHoja.Cells[ifila-1,ClFolio].interior.colorindex := NoExiste;
                  AHoja.Cells[ifila-1,ClFolio].AddComment('No se encontró el fólio en la BD, paso al siguiente fólio.');
                  AHoja.Cells[ifila-1,ClFolio].Comment.Visible := True;
                  ErrorEncontrado := True;
                  Raise exception.create('siguiente paquete');
                end;

                zactiv.Close;
                ZActiv.ParamByName('folio').AsString := sVFolio;
                ZActiv.ParamByName('contrato').AsString := ZFolio.FieldByName('scontrato').AsString;
                ZActiv.Open;

                sVPlataforma := ZFolio.FieldByName('sidplataforma').AsString;
                sVPernocta := ZFolio.FieldByName('sidpernocta').AsString;


                Inc(ifila);
                sVPartida := AHoja.Cells[ifila,ClPartida].Text;
                sVClasificacion := AHoja.cells[ifila,ClClasificacion].Text;
                sVHinicio := AHoja.cells[ifila,ClHinicio].Text;
                sVHfin := AHoja.cells[ifila,ClHfin].Text;
                lineab := 0;
                //Recorrer contenido de paquete
                while (lineab < 2) and (svpartida <> 'pda') and (sVPartida <> 'pda.') do
                begin
                  try
                    sVPartida := AHoja.Cells[ifila,ClPartida].Text;
                    sVClasificacion := AHoja.cells[ifila,ClClasificacion].Text;
                    sVHinicio := AHoja.cells[ifila,ClHinicio].Text;
                    sVHfin := AHoja.cells[ifila,ClHfin].Text;

                    //Si la longitud de la cadena de la actividad no es cero
                    if (length(trim(sVPartida)) = 0) then
                    begin
                      AHoja.Cells[ifila,ClPartida].interior.colorindex :=  Vacio;
                      Inc(lineab);
                      raise Exception.Create('siguiente fila');
                    end;

                    //Si no encuentra la actividad en la consulara que obtiene todas las partidas del contrato en el folio
                    if not ZActiv.Locate('snumeroactividad',sVPartida,[]) then
                    begin
                      AHoja.Cells[ifila,ClPartida].interior.colorindex :=  NoExiste;
                      AHoja.Cells[ifila,ClPartida].AddComment('Partida no asignada al folio.');
                      AHoja.Cells[ifila,ClPartida].Comment.Visible := True;
                    end;

                    //Si la longitud de la cadena de la clasificacion de la actividad es cero
                    if (length(trim(sVClasificacion)) = 0) then
                    begin
                      AHoja.Cells[ifila,ClClasificacion].interior.colorindex :=  Vacio;
                      Inc(lineab);
                      raise Exception.Create('siguiente fila');
                    end;

                    //Si encuentra el tipo de movimiento en la consulta que ibtiene los tipos de movimiento la encuentra
                    if not ZMovtos.Locate('sidtipomovimiento',sVClasificacion,[]) then
                    begin
                      AHoja.Cells[ifila,ClClasificacion].interior.colorindex :=  NoExiste;
                      AHoja.Cells[ifila,ClClasificacion].AddComment('El tipo de movimiento no esta dado de alta en el catálogo de su contrato.');
                      AHoja.Cells[ifila,ClClasificacion].Comment.Visible := True;
                      ErrorEncontrado := True;
                      raise Exception.Create('siguiente fila')
                    end;

                    //Si la longitud de la cadena de la hora de inicio de la actividad es cero
                    if (length(trim(sVHinicio)) = 0) then
                    begin
                      AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                      Inc(lineab);
                      raise Exception.Create('siguiente fila');
                    end;

                    //Valida si el formato de la cadena de la hora de inicio de la actividad es valida
                    try
                      if sVHinicio <> '24:00' then
                        StrToTime(sVHinicio);
                    Except
                      on e:Exception do
                      begin
                        AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                        AHoja.Cells[ifila,ClHinicio].AddComment('El rango de valor aceptado es 00:00-24:00.');
                        AHoja.Cells[ifila,ClHinicio].Comment.Visible := True;
                        ErrorEncontrado := True;
                        raise Exception.Create('siguiente fila');
                      end;
                    end;

                    //Si la longitud de la hora de termino de la actividad es cero
                    if (length(trim(sVHfin)) = 0) then
                    begin
                      AHoja.Cells[ifila,ClHFin].interior.colorindex :=  Vacio;
                      Inc(lineab);
                      raise Exception.Create('siguiente fila');
                    end;

                    //Si el formato de la hoa de termino de la partida es correcto
                    try
                      if sVHfin <> '24:00' then
                        StrToTime(sVHfin);
                    Except
                      on e:Exception do
                      begin
                        AHoja.Cells[ifila,ClHFin].interior.colorindex :=  Vacio;
                        AHoja.Cells[ifila,ClHFin].AddComment('El rango de valor aceptado es 00:00-24:00.');
                        AHoja.Cells[ifila,ClHFin].Comment.Visible := True;
                        ErrorEncontrado := True;
                        raise Exception.Create('siguiente fila');
                      end;
                    end;

                    //Si las horas tiene valores distintos a 24:00 se comprueba si la hora de inicio es menor a la hora de termino
                    if (sVHinicio <>  '24:00') and (sVHfin <> '24:00') then
                    begin
                      if StrToTime(sVHinicio) > StrToTime(sVHfin) then
                      begin
                        AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                        AHoja.Cells[ifila,ClHinicio].AddComment('La hora de inicio no debe ser mayor a la hora de fin.');
                        AHoja.Cells[ifila,ClHinicio].Comment.Visible := True;
                        ErrorEncontrado := True;
                        raise Exception.Create('siguiente fila');
                      end;
                    end;

                    //Si la hora de termino es igual a 24:00 y si la hora de inicio mas un minuto mas es menor a la hora de termino
                    if sVHfin = '24:00' then
                    begin
                      if StrToTime(sVHinicio) > (StrToTime('23:59')+strtotime('00:01')) then
                      begin
                        AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                        AHoja.Cells[ifila,ClHinicio].AddComment('La hora de inicio no debe ser mayor a la hora de fin.');
                        AHoja.Cells[ifila,ClHinicio].Comment.Visible := True;
                        ErrorEncontrado := True;
                        raise Exception.Create('siguiente fila');
                      end;
                    end;

                    if (sVHinicio = '24:00') and (sVHfin = '24:00') then
                    begin
                      AHoja.Cells[ifila,ClHinicio].interior.colorindex :=  Vacio;
                      AHoja.Cells[ifila,ClHinicio].AddComment('La hora de inicio y fin no deben ser iguales a 24:00');
                      AHoja.Cells[ifila,ClHinicio].Comment.Visible := True;
                      ErrorEncontrado := True;
                      raise Exception.Create('siguiente fila');
                    end;

                    //Recorrer columnas
                    ClRecurso := 7;
                    sVIdRecurso :=  AHoja.cells[iFilaRecurso,Clrecurso].Text;
                    if length(trim(sVIdRecurso)) = 0 then
                    begin
                      AHoja.Cells[iFilaRecurso,ClRecurso].interior.colorindex := Vacio;
                      AHoja.Cells[iFilaRecurso,Clrecurso].AddComment('Es necesario un id de recurso y descripción favor de poner el encabezado al paquete de datos.');
                      ErrorEncontrado := True;
                      AHoja.Cells[iFilaRecurso,Clrecurso].Comment.Visible := True;
                      raise Exception.Create('siguiente paquete');
                    end;

                    while length(trim(sVIdRecurso)) > 0 do
                    begin
                      sVDescripcion := '';
                      try
                        sVIdRecurso :=  AHoja.cells[iFilaRecurso,Clrecurso].Text;
                        if length(trim(sVIdRecurso)) = 0 then
                        begin
                          if Clrecurso = 7 then
                          begin
                            AHoja.Cells[iFilaRecurso,ClRecurso].interior.colorindex := Vacio;
                            AHoja.Cells[iFilaRecurso,Clrecurso].AddComment('Es necesario un id de recurso y descripción favor de poner el encabezado al paquete de datos.');
                            AHoja.Cells[iFilaRecurso,Clrecurso].Comment.Visible := True;
                          end;
                          raise Exception.Create('siguiente linea');
                        end;

                        if ZGPersonal.Locate('sIdGrupo',sVIdRecurso,[]) then
                        begin
                          if not ((Trim(ZGPersonal.FieldByName('sDescripcion').AsString) = 'TIEMPO EXTRA.') or (Trim(ZGPersonal.FieldByName('sDescripcion').AsString) = 'TIEMPO EXTRA')) then
                          begin
//                            AHoja.Cells[iFilaRecurso,ClRecurso].interior.colorindex :=  NoExiste;
//                            AHoja.Cells[iFilaRecurso,Clrecurso].AddComment('No esta dado de alta en el catalogo de personal como tiempo extra.');
//                            AHoja.Cells[iFilaRecurso,Clrecurso].Comment.Visible := True;
                            ErrorEncontrado := True;
                            raise exception.Create('siguiente columna');
                          end
                        end
                        else
                        begin
                          AHoja.Cells[iFilaRecurso,ClRecurso].interior.colorindex :=  NoExiste;
                          AHoja.Cells[iFilaRecurso,Clrecurso].AddComment('No esta dado de alta en el catalogo de personal como tiempo extra.');
                          AHoja.Cells[iFilaRecurso,Clrecurso].Comment.Visible := True;
                          ErrorEncontrado := True;
                          raise exception.Create('siguiente columna');
                        end;

                        sVDescripcion :=  AHoja.cells[iFilaRecurso-1,Clrecurso].Text;


                        ZRecursos.Close;
                        ZRecursos.ParamByName('idrecurso').AsString := sVIdRecurso;
                        ZRecursos.Open;

                        if ZRecursos.RecordCount = 0 then
                        begin
                          AHoja.Cells[iFilaRecurso,ClRecurso].interior.colorindex :=  NoExiste;
                          AHoja.Cells[iFilaRecurso,Clrecurso].AddComment('No esta dado de alta en el catalogo de '+tipo+' por contrato.');
                          AHoja.Cells[iFilaRecurso,Clrecurso].Comment.Visible := True;
                          ErrorEncontrado := True;
                          raise exception.Create('siguiente columna');
                        end;
                        

                        if length(trim(sVDescripcion)) = 0 then
                        begin
                          AHoja.Cells[iFilaRecurso-1,ClRecurso].interior.colorindex :=  Vacio;
                          AHoja.Cells[iFilaRecurso-1,ClRecurso].AddComment('Es obligatorio una descripción.');
                          AHoja.Cells[iFilaRecurso-1,ClRecurso].Comment.Visible := True;
                          ErrorEncontrado := True;
                          raise exception.Create('siguiente columna');
                        end;

                        lineab := 0;

                        sVCant := VarToStr(trim(AHoja.cells[iFila,Clrecurso].value));
                        sVHh := vartostr(trim(AHoja.cells[iFila,Clrecurso+1].Text));

                        if (Length(trim(sVCant)) = 0) or (Length(trim(sVHh)) = 0) then
                        begin
                          AHoja.Cells[iFila,ClRecurso].interior.colorindex :=  Vacio;
                          AHoja.Cells[iFila,ClRecurso].AddComment('Se ignora por tener valor nulo en cantidad o hh.');
                          AHoja.Cells[iFila,ClRecurso].Comment.Visible := False;
                          raise Exception.Create('siguiente columna');
                        end;
                        
                        {
                        if StrToFloat(sVHh) = 0 then
                        begin
                          AHoja.Cells[iFila,ClRecurso+1].interior.colorindex :=  Vacio;
                          AHoja.Cells[iFila,ClRecurso+1].AddComment('Se ignora por tener valor = 0');
                          AHoja.Cells[iFila,ClRecurso+1].Comment.Visible := False;
                          raise Exception.Create('siguiente columna');
                        end;
                        }

                        ZUptCuadre.Active := False;
                        ZUptCuadre.SQL.Clear;
                        ZUptCuadre.SQL.Text := 'select * from horasextras where scontrato = :contrato and didfecha = :fecha';
                        ZUptCuadre.ParamByName('contrato').AsString := ContratoDiario;
                        ZUptCuadre.ParamByName('fecha').AsDateTime := Cuadre_FechaReporte;
                        ZUptCuadre.Open;

                        ZNotas.Active := False;
                        ZNotas.ParamByName('fecha').AsDateTime := Cuadre_FechaReporte;
                        ZNotas.ParamByName('orden').AsString := sVFolio;
                        ZNotas.ParamByName('actividad').AsString := sVPartida;
                        ZNotas.ParamByName('contrato').AsString := ContratoDiario;
                        ZNotas.ParamByName('sidclasificacion').AsString := sVClasificacion;
                        ZNotas.Open;

                        if (ZNotas.RecordCount = 0) or (ZNotas.FieldByName('iiddiario').AsInteger <= 0) then
                        begin
                          AHoja.Cells[iFila,ClRecurso+1].AddComment('No se encontro nota para asignar el recurso.');
                          AHoja.Cells[iFila,ClRecurso+1].interior.colorindex :=  NoEnActividades;
                          ErrorEncontrado := True;
                          raise Exception.Create('siguiente columna');
                        end;

                        try
                          ZUptCuadre.Append;
                          ZUptCuadre.FieldByName('scontrato').AsString:= ContratoDiario;
                          ZUptCuadre.FieldByName('didfecha').AsDateTime := Cuadre_FechaReporte;
                          ZUptCuadre.fieldbyname('snumeroorden').AsString := sVFolio;
                          ZUptCuadre.FieldByName('iiddiario').AsInteger := ZNotas.FieldByName('iiddiario').AsInteger;
                          ZUptCuadre.FieldByName('swbs').AsString := ZNotas.FieldByName('swbs').AsString;
                          ZUptCuadre.FieldByName('sidpersonal').AsString := sVIdRecurso;
                          ZUptCuadre.FieldByName('sidplataforma').AsString := sVPlataforma;
                          ZUptCuadre.FieldByName('stipopernocta').AsString := zQCuentas.FieldByName('sIdPernocta').AsString;
                          ZUptCuadre.FieldByName('sagrupapersonal').AsString := '*';
                          ZUptCuadre.FieldByName('laplicapernocta').AsString := 'Si';
                          ZUptCuadre.FieldByName('dcanthhgenerador').AsFloat := 0;
                          ZUptCuadre.FieldByName('iitemorden').AsInteger := ZRecursos.FieldByName('iitemorden').AsInteger;
                          ZUptCuadre.FieldByName('sdescripcion').AsString := sVDescripcion;
                          ZUptCuadre.FieldByName('sidpernocta').AsString := sVPernocta;
                          ZUptCuadre.FieldByName('stipoobra').AsString := 'PU';
                          ZUptCuadre.FieldByName('shorainicio').AsString := sVHinicio;
                          ZUptCuadre.FieldByName('shorafinal').AsString := sVHfin;
                          ZUptCuadre.FieldByName('dcantidad').AsFloat := StrToFloat(svcant);
                          ZUptCuadre.FieldByName('snumeroactividad').AsString := sVPartida;
                          ZUptCuadre.FieldByName('dsolicitado').asfloat := 0;
                          ZUptCuadre.FieldByName('sfactor').AsString := '0';
                          ZUptCuadre.FieldByName('dcostomn').AsInteger :=0;
                          ZUptCuadre.FieldByName('dcostodll').AsInteger :=0;
                          ZUptCuadre.FieldByName('scanthh').AsString := sVHh;
                          ZUptCuadre.FieldByName('dajuste').AsInteger := 0;
                          ZUptCuadre.Post;

                          AHoja.Cells[iFila,ClRecurso+1].interior.colorindex :=  ImportarL;

                        except
                          on e:exception do
                          begin
                            AHoja.Cells[iFila,ClRecurso+1].AddComment('Error encontrado: '+e.message);
                            AHoja.Cells[iFila,ClRecurso+1].Comment.Visible := false;
                            ErrorEncontrado := True;
                          end;
                        end;
                      except
                        on e:Exception do
                          if e.Message = 'siguiente columna' then
                            ;
                      end;
                      ClRecurso := ClRecurso+2;
                    end;
                  except
                    on e:Exception do
                    if e.message = 'siguiente fila' then
                      ;
                  end;
                  Inc(ifila);
                end;     
              end;
            except
              on e:exception do
                if e.message <> 'siguiente paquete' then
                   raise;
            end;

            Inc(iFila);

            if maxBlanca = 29 then
            begin
              LineaLeyenda := iFila;
              CreaLeyenda(AExcel);
            end;
          end;



        finally
          ZUptCuadre.free;
        end;

      finally
        ZNotas.Free;
      end;

    finally
      ZRecursos.Free;
    end;
  end;
begin

  try
    try
      if not Pers then
        raise Exception.Create('No seleccionó opciones de importacion.');

      //Abrir excel
      if not DireccionPlantilla then
        raise exception.create('proceso cancelado por el usuario');
    finally
    end;

    try
      AExcel:=CreateOleObject('Excel.Application');
      AExcel.Visible := False;
    except
      raise Exception.Create('No tiene instalado microsoft excel o bién ocurre un problema con el mismo.');
    end;

    AExcel.Workbooks.Open(PathArchivo);
    ALibro := AExcel.Workbooks[AExcel.Workbooks.count];
    ILibro := AExcel.Workbooks.count;
    AExcel.Visible := True;

    ZFolio := TZReadOnlyQuery.Create(nil);
    try
      ZFolio.Connection := connection.zConnection;
      ZFolio.Active := False;
      ZFolio.SQL.Clear;
      ZFolio.SQL.Text := 'SELECT * FROM ordenesdetrabajo WHERE sNumeroOrden = :Folio AND sContrato = :Contrato';
      ZFolio.ParamByName('Contrato').AsString := ContratoDiario;

      ZGPersonal := TZReadOnlyQuery.Create(nil);
      try
        ZGPersonal.Connection := connection.zConnection;
        ZGPersonal.Active := False;
        ZGPersonal.SQL.Clear;
        ZGPersonal.SQL.Add(' select * from grupospersonal');
        ZGPersonal.Open;


        ZActiv := TZReadOnlyQuery.Create(nil);
        try
          ZActiv.Connection := connection.zConnection;
          ZActiv.Active := False;
          ZActiv.SQL.Clear;
          ZActiv.SQL.Text := 'SELECT * FROM actividadesxorden WHERE scontrato = :Contrato and snumeroorden = :Folio';

          ZMovtos := Tzreadonlyquery.create(nil);
          try
            ZMovtos.Connection := connection.zConnection;
            ZMovtos.Active := False;
            ZMovtos.SQL.Clear;
            ZMovtos.SQL.Text := 'select * from tiposdemovimiento where sClasificacion <> "movimiento de barco" and sContrato = :Contrato order by iorden';
            ZMovtos.parambyname('contrato').asstring := ContratoDiario;
            ZMovtos.Open;

            EPagina := '';
            //Error de pagina
            for I := 0 to 0 do
            begin
              ErrorEncontrado := False;
              Ipagina := -1;
              try

                  //importar personal horas extras
                  //buscar hoja cuadre p
                  IPagina:= BuscaHoja(AExcel,'cuadre personal');

                  if Ipagina = -1 then
                    raise exception.create('No se encontró la pestaña con el texto cuadre personal');

                  AHoja := ALibro.worksheets[IPagina];

                  if connection.zConnection.InTransaction then
                    connection.zConnection.Rollback;
                  connection.zConnection.StartTransaction;

                  Lectura('personal');

                  if ErrorEncontrado then
                  begin
                    connection.zConnection.Rollback;
                    ShowMessage('Se encontraron errores en la pestaña de cuadre personal, favor de corregir.');
                  end
                  else
                  begin
                    connection.zConnection.Commit;
                    ShowMessage('Se importaron registros para el cuadre de personal.');
                  end;

              except
                on e:Exception do
                  ShowMessage('Ocurrió el siguiente error al tratar de importar el cuadre '+e.Message);
              end;
            end;
          finally
            zmovtos.free;
          end;
        finally
          Zactiv.free;
        end;
      finally
        Zfolio.free;
      end;

    finally
      ZGPersonal.Free;
    end;
  finally

  end;
end;

procedure TfrmCuadrePersonalEquipo_Excel.ExportarPlantillaActividades(EPersonal,EEquipo,Info:Boolean);
var
  EExcel,
  ELibro,
  EHojaPErsonal,EHojaEquipo: Variant;

  IELibro, IEHojaPersonal,IEhojaEquipo,
  CVeces:Integer;
  I: Integer;

  zEMoe,ZEFolios,ZEActividades,ZEMovimientos:TZReadOnlyQuery;
  IMActivo:Integer;
  procedure CreaEstructura(Tipo:String);
  const
    IClPlataforma = 1;
    IclPda = 2;
    IcltipoMov = 3;
    IClInicia = 4;
    IClFinaliza = 5;
    IClDuracion = 6;
  var
    ZEMoeR,
    qrPartidas : TZReadOnlyQuery;
    IEFila, x, Ucolumna,
    IClRecurso, IIndex, Ifolio
    :Integer;
    VHoja, VRango : Variant;
  begin
    try
      qrPartidas := TZReadOnlyQuery.Create(nil);
      qrPartidas.Connection := connection.zconnection;
      if info then
      begin
        ZEMoeR := TZReadOnlyQuery.Create(nil);
        zEMoeR.Connection := connection.zConnection;
        zEMoeR.Active := False;
        if Tipo = 'Personal' then
          zEMoeR.SQL.Text :=  'SELECT mre.sIdRecurso, mre.dCantidad ,p.* '+
                              'FROM moerecursos AS mre '+
                              'inner join personal p on (mre.sidrecurso = p.sidpersonal and p.scontrato = :Contrato) '+
                              'WHERE mre.iIdMoe = :IdMoe  order by p.iItemOrden';
        if Tipo = 'Equipo' then
          zEMoeR.SQL.Text :=  'SELECT mre.sIdRecurso, mre.dCantidad ,e.* '+
                              'FROM moerecursos AS mre '+
                              'inner join equipos e on (mre.sidrecurso = e.sidequipo and e.scontrato = :Contrato) '+
                              'WHERE mre.iIdMoe = :IdMoe  order by e.iItemOrden';
        ZEMoeR.ParamByName('IdMoe').AsInteger := IMActivo;
        ZEMoeR.ParamByName('Contrato').AsString := ContratoDiario;
        ZEMoeR.Open;
      end;

      IEFila := 1;

      if Tipo = 'Personal' then
      begin
        Vhoja := EHojaPErsonal;
        iindex := iEHojaPErsonal;
      end;

      if Tipo = 'Equipo' then
      begin
        Vhoja := EHojaEquipo;
        iindex := iEHojaEquipo;
      end;
      ZEFolios.first;
      while not ZEFolios.Eof do
      begin
        Ifolio := CmbFolios.Items.IndexOf(ZEFolios.FieldByName('snumeroorden').AsString);
        if Ifolio > -1 then
        if CmbFolios.IsChecked(Ifolio) then
        begin
          IClRecurso := 7;

          VHoja.Rows[inttostr(IEFila)+':'+inttostr(IEFila)].RowHeight := 66;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IcltipoMov)+IntToStr(IEFila)+':'+ColumnaNombre(IcltipoMov+1)+IntToStr(IEFila+1)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'TOTAL';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza)+IntToStr(IEFila)+':'+ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila+1)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := '0';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IcltipoMov)+IntToStr(IEFila+2)+':'+ColumnaNombre(IcltipoMov+2)+IntToStr(IEFila+2)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := ZEFolios.FieldByName('snumeroorden').AsString;
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClPlataforma)+IntToStr(IEFila+2)+':'+ColumnaNombre(IClPlataforma)+IntToStr(IEFila+2)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := ZEFolios.FieldByName('sidplataforma').AsString;
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClPlataforma)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClPlataforma)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'Referencia';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IclPda)+IntToStr(IEFila+3)+':'+ColumnaNombre(IclPda)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'PDA';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IcltipoMov)+IntToStr(IEFila+3)+':'+ColumnaNombre(IcltipoMov)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'N.E.C.';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClInicia)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClInicia)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'Inicia';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClFinaliza)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'Finaliza';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila+3)];
          VRango.MergeCells := True;
          VRango.HorizontalAlignment :=xlCenter;
          VRango.VerticalAlignment := xlCenter;
          VRango.Font.Size := 8;
          VRango.Font.Bold := False;
          VRango.WrapText := True;
          VRango.value := 'Duración';
          VRango.NumberFormat := '@';
          VRango.Interior.Color := $00BBBBBB;
          Vrango.Borders.LineStyle := xlContinuous;

          ZEMoeR.First;
          while not ZEMoeR.Eof do
          begin
            Ucolumna := iclrecurso;
            //Recursos
            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso)+IntToStr(IEFila)+':'+ColumnaNombre(IClRecurso+1)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := ZEMoeR.FieldByName('sdescripcion').AsString;
            VRango.NumberFormat := '@';
            VRango.Interior.Color := $00BBBBBB;
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso)+IntToStr(IEFila+1)+':'+ColumnaNombre(IClRecurso+1)+IntToStr(IEFila+1)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 11;
            VRango.Font.Bold := true;
            VRango.WrapText := True;
            VRango.NumberFormat := '@';
            VRango.value := ZEMoeR.FieldByName('sidrecurso').AsString;
            VRango.Interior.Color := $00BBBBBB;
            VRango.NumberFormat := '@';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClRecurso)+IntToStr(IEFila+3)];
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := true;
            VRango.WrapText := True;
            VRango.NumberFormat := '@';
            VRango.value := 'Cant.';
            VRango.Interior.Color := $00BBBBBB;
            Vrango.Borders.LineStyle := xlContinuous;
            VHoja.Cells[IEFila, IClRecurso].ColumnWidth := 4.3;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso+1)+IntToStr(IEFila+3)+':'+ColumnaNombre(IClRecurso+1)+IntToStr(IEFila+3)];
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := false;
            VRango.WrapText := True;
            VRango.NumberFormat := '@';//'0.00000000';
            VRango.value := 'h.h.';
            VRango.Interior.Color := $00BBBBBB;
            Vrango.Borders.LineStyle := xlContinuous;
            VHoja.Cells[IEFila, IClRecurso+1].ColumnWidth := 12.3;

            IClRecurso := IClRecurso+2;
            ZEMoeR.Next;
          end;
          IEFila := IEFila + 4;

          {$REGION 'CONSULTA PARTIDAS'}
          with qrpartidas do
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
            parambyname('contrato').asstring := ContratoDiario;
            parambyname('convenio').asstring := global_convenio;
            parambyname('fecha').AsDateTime := Cuadre_FechaReporte;
            parambyname('alcance').asstring := frm_connection.connection.configuracion.fieldbyname('sTipoAlcance').asstring;
            parambyname('folio').asstring := ZEFolios.FieldByName('snumeroorden').AsString;
            parambyname('turno').asstring := global_turno_reporte;
            open;

            first;
          end;
          {$ENDREGION}

          while not qrpartidas.Eof do
          begin
            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClPlataforma)+IntToStr(IEFila)+':'+ColumnaNombre(IClPlataforma)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := '';
            VRango.NumberFormat := '@';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IclPda)+IntToStr(IEFila)+':'+ColumnaNombre(IclPda)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.Value := qrPartidas.FieldByName('sNumeroActividad').asstring;
            VRango.NumberFormat := '@';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IcltipoMov)+IntToStr(IEFila)+':'+ColumnaNombre(IcltipoMov)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.NumberFormat := '@';
            VRango.Value := 'TE';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClInicia)+IntToStr(IEFila)+':'+ColumnaNombre(IClInicia)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := qrPartidas.FieldByName('sHoraInicio').asstring;
            VRango.NumberFormat := '[hh]:mm';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza)+IntToStr(IEFila)+':'+ColumnaNombre(IClFinaliza)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := qrPartidas.FieldByName('sHoraFinal').asstring;
            VRango.NumberFormat := '[hh]:mm';
            Vrango.Borders.LineStyle := xlContinuous;

            VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila)+':'+ColumnaNombre(IClFinaliza+1)+IntToStr(IEFila)];
            VRango.MergeCells := True;
            VRango.HorizontalAlignment :=xlCenter;
            VRango.VerticalAlignment := xlCenter;
            VRango.Font.Size := 8;
            VRango.Font.Bold := False;
            VRango.WrapText := True;
            VRango.value := '=+'+ColumnaNombre(IClFinaliza)+inttostr(IEFila)+'-'+ColumnaNombre(IClInicia)+inttostr(IEFila);
            VRango.NumberFormat := '0.00000000';
            Vrango.Borders.LineStyle := xlContinuous;

            IClRecurso := 7;
            while IClRecurso <= Ucolumna do
            begin
              VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso)+IntToStr(IEFila)+':'+ColumnaNombre(IClRecurso)+IntToStr(IEFila)];
              VRango.HorizontalAlignment :=xlCenter;
              VRango.VerticalAlignment := xlCenter;
              VRango.Font.Size := 8;
              VRango.Font.Bold := true;
              VRango.WrapText := True;
              VRango.NumberFormat := '0';
              VRango.value := '0';
              Vrango.Borders.LineStyle := xlContinuous;
              
              VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso+1)+IntToStr(IEFila)+':'+ColumnaNombre(IClRecurso+1)+IntToStr(IEFila)];
              VRango.HorizontalAlignment :=xlCenter;
              VRango.VerticalAlignment := xlCenter;
              VRango.Font.Size := 8;
              VRango.Font.Bold := false;
              VRango.WrapText := True;
              VRango.NumberFormat :='0.00000000';
              if tipo = 'Equipo' then
                VRango.value := '=+('+ColumnaNombre(IClRecurso)+inttostr(IEFila)+'*$F'+inttostr(IEFila)+')';
              if Tipo = 'Personal' then
                VRango.value := '=+('+ColumnaNombre(IClRecurso)+inttostr(IEFila)+'*$F'+inttostr(IEFila)+')*2';

              Vrango.Borders.LineStyle := xlContinuous;
              IClRecurso := IClRecurso + 2;
            end;
            Inc(IEFila);
            qrPartidas.Next;
          end;
          IEFila := IEfila+4;

        end;
        ZEFolios.next;
      end;
      IClRecurso := 7;
      IEFila := IEFila - 3;
      ZEMoeR.First;
      ELibro.Sheets[iindex].rows[IEFila].select;
      EExcel.selection.interior.color := $00BBBBBB;

      VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso-4)+IntToStr(IEFila)+':'+ColumnaNombre(IClRecurso-4)+IntToStr(IEFila)];
      VRango.Value := 'Total solicitado';
      while not ZEMoeR.eof do
      begin
        VRango := ELibro.Sheets[iindex].Range[ColumnaNombre(IClRecurso+1)+IntToStr(IEFila)+':'+ColumnaNombre(IClRecurso+1)+IntToStr(IEFila)];
        VRango.HorizontalAlignment :=xlCenter;
        VRango.VerticalAlignment := xlCenter;
        VRango.Font.Size := 8;
        VRango.Font.Bold := false;
        VRango.WrapText := True;
        VRango.Value := ZEMoeR.Fieldbyname('dCantidad').asfloat;
        IClRecurso := IClRecurso+2;
        ZEMoeR.next;
      end;

    finally
      if info then
      begin
        try
          ZEMoeR.free;
        except
          ;
        end;
      end;
    end;
  end;
begin
  try
    Try
      EExcel := CreateOleObject('Excel.Application');
    Except
      raise Exception.Create('Error al enlazar con la aplicación de terceros (excel).');
    End;

    EExcel.Visible := True;
    EExcel.DisplayAlerts:= True;
    ELibro := EExcel.Workbooks.Add;
    IElibro := EExcel.Workbooks.count;

    for  IEHojaPersonal := Eexcel.Workbooks[ielibro].Sheets.Count  downto 2 do
      EExcel.workbooks[iELibro].workSheets[iEHojaPErsonal].Delete;

    if Epersonal then
    begin
      IEHojaPersonal := 1;
      EHojaPersonal := Eexcel.Workbooks[ielibro].Sheets[IEHojapersonal];
      EHojaPersonal.Name := 'CUADRE PERSONAL';

    end;

    if EEquipo then
    begin
      if not EPersonal then
      begin
        IEhojaEquipo := 1;
        EHojaEquipo := Eexcel.Workbooks[ielibro].Sheets[IEhojaEquipo];
        EHojaEquipo.Name := 'CUADRE EQUIPO';
      end
      else
      begin
        EHojaEquipo := EExcel.Workbooks[ielibro].Sheets.Add;
        IEhojaEquipo := 1;
        EHojaEquipo.Name := 'CUADRE EQUIPO';
        IEHojaPersonal := 2;
        EHojaPersonal := Eexcel.Workbooks[ielibro].Sheets[IEHojapersonal];
      end;
    end;
    try
      Progreso.Visible := True;
      if Info then
      begin
        zEMoe :=TZReadOnlyQuery.Create(nil);
        try
          zEMoe.Connection := connection.zConnection;
          zEMoe.Active := False;
          zEMoe.SQL.Text := ' SELECT iidmoe,dIdFecha FROM moe WHERE dIdFecha <= :Fecha AND sContrato = :Contrato order by didfecha desc limit 1 ';
          ZEMoe.ParamByName('Fecha').AsDateTime := Cuadre_FechaReporte;
          ZEMoe.ParamByName('Contrato').AsString := ContratoDiario;
          if info then
          begin
            zEMoe.Open;
            IMActivo := zEMoe.FieldByName('iidmoe').AsInteger;
          end;
        finally
          zEMoe.Free;
        end;

        ZEFolios := TZReadOnlyQuery.Create(nil);
        ZEFolios.Connection := connection.zConnection;
        ZEFolios.Active := False;
        ZEFolios.SQL.Text :=  'SELECT * FROM ordenesdetrabajo WHERE  sContrato = :Contrato';
        ZEFolios.ParamByName('Contrato').AsString := ContratoDiario;
        ZEFolios.Open;

        ZEActividades := TZReadOnlyQuery.Create(nil);
        ZEActividades.Connection := connection.zConnection;
        ZEActividades.Active := False;
        ZEActividades.SQL.Text :=  'SELECT * FROM actividadesxorden WHERE scontrato = :Contrato and snumeroorden = :Folio';
        ZEActividades.ParamByName('Contrato').asstring := ContratoDiario;

        ZEMovimientos := TZReadOnlyQuery.Create(nil);
        ZEMovimientos.Connection := connection.zConnection;
        ZEMovimientos.Active := False;
        ZEMovimientos.SQL.Text :=  'select * from tiposdemovimiento where sClasificacion <> "movimiento de barco" and sContrato = :Contrato order by iorden ';
        ZEMovimientos.ParamByName('Contrato').asstring := ContratoDiario;
        ZEMovimientos.Open;
      end;

      if EPersonal then
      begin
        EExcel.Workbooks[iElibro].Worksheets[IEHojaPersonal].select;
        CreaEstructura('Personal');
      end;

      if EEquipo then
      begin
        EExcel.Workbooks[iElibro].Worksheets[IEhojaEquipo].select;
        CreaEstructura('Equipo');
      end;

    finally
      Progreso.Visible := False;
      EExcel.Visible := True;
      try
        ZEFolios.free;
      except
        ;
      end;
      try
        ZEActividades.free;
      except
        ;
      end;
      try
        ZEMovimientos.free;
      except
        ;
      end;
    end;
  except
    on e:Exception do
      ShowMessage('No se pudo completar la generación de la plantilla por el siguiente motivo: '+#10+e.Message);
  end;
end;


end.
