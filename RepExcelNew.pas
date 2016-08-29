unit RepExcelNew;

interface

uses
  Windows, Messages, Forms, DB, Dialogs, ZAbstractRODataset,
  ZDataset, Grids, DBGrids, ComCtrls, Controls, StdCtrls, Classes, ExtCtrls;

type
  TActividad = class
    ListaPartida: TStringList;
    constructor Create;
    destructor Destroy;
  end;

  TfrmRepExcelNew = class(TForm)
    Panel1: TPanel;
    btnReporte: TButton;
    btnCerrar: TButton;
    roqReporte: TZReadOnlyQuery;
    SaveExcel: TSaveDialog;
    roqContratos: TZReadOnlyQuery;
    roqOrdenes: TZReadOnlyQuery;
    dsContratos: TDataSource;
    dsOrdenes: TDataSource;
    pnlGenerar: TPanel;
    cbActividad: TComboBox;
    Label5: TLabel;
    Label6: TLabel;
    cbPartida: TComboBox;
    btnProcesar: TButton;
    Panel3: TPanel;
    Panel5: TPanel;
    Panel2: TPanel;
    Label3: TLabel;
    Label4: TLabel;
    FechaInicio: TDateTimePicker;
    FechaFinal: TDateTimePicker;
    Panel4: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    DBGrid1: TDBGrid;
    Panel8: TPanel;
    Label1: TLabel;
    DBGrid2: TDBGrid;
    Panel9: TPanel;
    Label2: TLabel;
    btnBuscar: TButton;
    procedure btnCerrarClick(Sender: TObject);
    procedure btnReporteClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnProcesarClick(Sender: TObject);
    procedure roqReporteAfterOpen(DataSet: TDataSet);
    procedure roqReporteAfterClose(DataSet: TDataSet);
    procedure roqOrdenesAfterScroll(DataSet: TDataSet);
    procedure roqContratosAfterScroll(DataSet: TDataSet);
    procedure cbActividadChange(Sender: TObject);
    procedure btnBuscarClick(Sender: TObject);
    procedure roqOrdenesAfterClose(DataSet: TDataSet);
    procedure roqOrdenesAfterOpen(DataSet: TDataSet);
    procedure roqOrdenesAfterRefresh(DataSet: TDataSet);
  private
    Excel: OleVariant;
    ListaActividad: TStringList;
    procedure GenerarReporte;
    procedure EstablecerFormato(pInicial: Integer; Ren: Integer);
  public
    { Public declarations }
  end;

var
  frmRepExcelNew: TfrmRepExcelNew;

implementation

uses frm_connection;

{$R *.dfm}

constructor TActividad.Create;
begin
  ListaPartida := TStringList.Create;
  ListaPartida.Clear;
end;

destructor TActividad.Destroy;
begin
  ListaPartida.Clear;
  ListaPartida.Destroy;
end;

procedure TfrmRepExcelNew.btnProcesarClick(Sender: TObject);
var
  LocCursor: TCursor;
begin
  LocCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;

    roqReporte.ParamByName('sContrato').AsString := roqContratos.FieldByName('sContrato').AsString;       // 'OTm-001-2016';
    roqReporte.ParamByName('sNumeroOrden').AsString := roqOrdenes.FieldByName('sNumeroOrden').AsString;   // '0020-2016';
    roqReporte.ParamByName('FechaInicio').AsDate := FechaInicio.Date;   // StrToDate('03/03/2016');
    roqReporte.ParamByName('FechaFinal').AsDate := FechaFinal.Date;     // StrToDate('06/03/2016');

    roqReporte.Open;

    if roqReporte.RecordCount > 0 then
    begin
      btnReporte.SetFocus;
    end
    else
    begin
      roqReporte.Close;
      ShowMessage('No existen datos que reportar con el criterio especificado.');
    end;
  finally
    Screen.Cursor := LocCursor;
  end;
end;

procedure TfrmRepExcelNew.btnReporteClick(Sender: TObject);
begin
  GenerarReporte;
end;

procedure TfrmRepExcelNew.cbActividadChange(Sender: TObject);
var
  i, Indice: Integer;
begin
  cbPartida.Items.Clear;
  cbPartida.Items.Add('< TODAS >');

  if cbActividad.ItemIndex > 0 then
  begin
    // Cargar los datos de las partidas de anexo
    Indice := ListaActividad.IndexOf(cbActividad.Text);
    for i := 0 to TActividad(ListaActividad.Objects[Indice]).ListaPartida.Count -1 do
      cbPartida.Items.Add(TActividad(ListaActividad.Objects[Indice]).ListaPartida[i]);

    cbPartida.Enabled := True;
  end
  else
    cbPartida.Enabled := False;

  cbPartida.ItemIndex := 0;
end;

procedure TfrmRepExcelNew.btnBuscarClick(Sender: TObject);
begin
  roqContratos.ParamByName('FechaInicio').AsDate := FechaInicio.Date;
  roqContratos.ParamByName('FechaFinal').AsDate := FechaFinal.Date;
  if roqContratos.Active then
    roqContratos.Refresh
  else
    roqContratos.Open;
end;

procedure TfrmRepExcelNew.btnCerrarClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmRepExcelNew.GenerarReporte;
var
  LocCursor: TCursor;
  Tipo: TVarType;
  eRen: Integer;
  Libro: OleVariant;
  CadFecha: String;
  Inicial, nHoja, myCol: Word;
  CadFiltrado: String;
begin
  try
    LocCursor := Screen.Cursor;
    try
      Screen.Cursor := crHourGlass;

      if cbActividad.ItemIndex > 0 then
      begin
        roqReporte.Filtered := False;

        CadFiltrado := 'sNumeroActividad = ' + QuotedStr(cbActividad.Text);

        if cbPartida.ItemIndex > 0 then
          CadFiltrado := CadFiltrado + ' and sIdPartidaAnexo = ' + QuotedStr(cbPartida.Text);

        // Filtrar los datos
        roqReporte.Filter := CadFiltrado;

        roqReporte.Filtered := True;
      end;

      Excel := CreateOleObject('Excel.Application');
      Excel.Workbooks.Add;
      Excel.Visible := True;
      Excel.DisplayAlerts := False;

      // Crear las hojas necesarias
      while Excel.Sheets.Count < 3 do
      begin
        Excel.Sheets[1].Select;
        Excel.Sheets.Add(Null, Excel.ActiveSheet, Null, Null);
      end;

      while Excel.Sheets.Count > 3 do
        Excel.Sheets[1].Delete;

      Excel.Sheets[1].Name := 'Personal';
      Excel.Sheets[2].Name := 'Equipo';
      Excel.Sheets[3].Name := 'PU';

      Excel.Sheets[1].Select;   // Seleccionar la primer hoja (Mano de Obra)

      roqReporte.First;
      eRen := 6;
      Libro := Excel.ActiveWorkBook.ActiveSheet;
      Inicial := roqReporte.FieldByName('iIdOrdenTipoAnexo').AsInteger;
      case Inicial of
        0: Excel.Sheets[1].Select;
        1: Excel.Sheets[2].Select;
        2: Excel.Sheets[3].Select;
      end;
      nHoja := 0;
      while not roqReporte.Eof do
      begin
        if Inicial <> roqReporte.FieldByName('iIdOrdenTipoAnexo').AsInteger then
        begin
          EstablecerFormato(Inicial, eRen);    // DarFormato a la hoja de Mano de Obra
          eRen := 6;

          case Inicial of
            0: Excel.Sheets[2].Select;  // Seleccionar la siguiente hoja (Equipo)
            1: Excel.Sheets[3].Select;  // Seleccionar la siguiente hoja (PU)
          end;

          Inicial := roqReporte.FieldByName('iIdOrdenTipoAnexo').AsInteger;
        end;
        {case roqReporte.FieldByName('iIdOrdenTipoAnexo').AsInteger of
          1: begin
                  EstablecerFormato(eRen);    // DarFormato a la hoja de Mano de Obra

                  eRen := 6;
                  Excel.Sheets[2].Select;     // Seleccionar la siguiente hoja (Equipo)
                end;
          2: begin
                  EstablecerFormato(eRen);    // Dar Formato a la hoja de Equipo de Trabajo

                  eRen := 6;
                  Excel.Sheets[3].Select;     // Seleccionar la siguiente hoja (PU)
                end;
        end;}

        {if roqReporte.RecNo = 50 then
        begin
          EstablecerFormato(eRen);    // Dar Formato a la hoja de Equipo de Trabajo

          eRen := 6;
          Excel.Sheets[2].Select;     // Seleccionar la siguiente hoja (PU)
        end;}

        Excel.ActiveSheet.Cells[eRen, 1].Value := FormatDateTime('mm-dd-yyyy', roqReporte.FieldByName('dIdFecha').AsDateTime);
        Excel.ActiveSheet.Cells[eRen, 2].Value := roqReporte.FieldByName('sIdFolio').AsString;
        Excel.ActiveSheet.Cells[eRen, 3].Value := roqReporte.FieldByName('sNumeroActividad').AsString;
        Excel.ActiveSheet.Cells[eRen, 4].Value := roqReporte.FieldByName('mDescripcionOrden').AsString;
        Excel.ActiveSheet.Cells[eRen, 5].Value := roqReporte.FieldByName('mDescripcionBitacoraActividades').AsString;
        Excel.ActiveSheet.Cells[eRen, 6].Value := roqReporte.FieldByName('Frente').AsString;
        Excel.ActiveSheet.Cells[eRen, 7].Value := roqReporte.FieldByName('sIdClasificacion').AsString;
        Excel.ActiveSheet.Cells[eRen, 8].Value := #39 + roqReporte.FieldByName('sHoraInicio').AsString;
        Excel.ActiveSheet.Cells[eRen, 9].Value := #39 + roqReporte.FieldByName('sHoraFinal').AsString;
        Excel.ActiveSheet.Cells[eRen, 10].Value := #39 + roqReporte.FieldByName('sIdPartidaAnexo').AsString;
        Excel.ActiveSheet.Cells[eRen, 11].Value := roqReporte.FieldByName('sTituloPartidaAnexo').AsString;
        Excel.ActiveSheet.Cells[eRen, 12].Value := roqReporte.FieldByName('dCantidad').AsString;
        Excel.ActiveSheet.Cells[eRen, 13].Value := roqReporte.FieldByName('dJornada').AsString;

        if Inicial = 2 then
        begin
          Excel.ActiveSheet.Cells[eRen, 14].Value := roqReporte.FieldByName('sTituloPartidaAnexo').AsString;
          myCol := 14;
        end
        else
          myCol := 13;

        Excel.ActiveSheet.Cells[eRen, myCol +1].Value := roqReporte.FieldByName('dVentaMN').AsString;
        Excel.ActiveSheet.Cells[eRen, myCol +2].Value := roqReporte.FieldByName('dVentaDLL').AsString;
        Excel.ActiveSheet.Cells[eRen, myCol +3].Value := roqReporte.FieldByName('dCostoMN').AsString;
        Excel.ActiveSheet.Cells[eRen, myCol +4].Value := roqReporte.FieldByName('dCostoDLL').AsString;

        Inc(eRen);
        roqReporte.Next;
      end;

      // Dar formato a la hoja del PU
      EstablecerFormato(Inicial, eRen);

      // Solicitar el grabado del archivo
      if SaveExcel.Execute(Self.Handle) then
      begin
        Excel.DisplayAlerts := True;
        Excel.Visible := True;

        Excel.ActiveWorkbook.SaveAs(SaveExcel.FileName, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null, Null);
      end;
    finally
      roqReporte.Filtered := False;

      Tipo := TVarData(Excel).VType;
      if (Tipo > 0) then
      begin
        //Excel.Quit;
        Excel.DisplayAlerts := True;
        Excel := Null;
      end;

      Screen.Cursor := LocCursor;
    end;
  except
    on e:Exception do
      ShowMessage(e.Message);
  end;
end;

procedure TfrmRepExcelNew.roqContratosAfterScroll(DataSet: TDataSet);
begin
  roqReporte.Close;

  roqOrdenes.ParamByName('FechaInicio').AsDate := FechaInicio.Date;
  roqOrdenes.ParamByName('FechaFinal').AsDate := FechaFinal.Date;
  if roqOrdenes.Active then
    roqOrdenes.Refresh
  else
    roqOrdenes.Open;
end;

procedure TfrmRepExcelNew.roqOrdenesAfterClose(DataSet: TDataSet);
begin
  btnProcesar.Enabled := False;
end;

procedure TfrmRepExcelNew.roqOrdenesAfterOpen(DataSet: TDataSet);
begin
  btnProcesar.Enabled := roqOrdenes.RecordCount > 0;
end;

procedure TfrmRepExcelNew.roqOrdenesAfterRefresh(DataSet: TDataSet);
begin
  btnProcesar.Enabled := roqOrdenes.RecordCount > 0;
end;

procedure TfrmRepExcelNew.roqOrdenesAfterScroll(DataSet: TDataSet);
begin
  roqReporte.Close;
end;

procedure TfrmRepExcelNew.roqReporteAfterClose(DataSet: TDataSet);
begin
  btnReporte.Enabled := False;
  btnProcesar.Enabled := True;
  pnlGenerar.Visible := False;
end;

procedure TfrmRepExcelNew.roqReporteAfterOpen(DataSet: TDataSet);
var
  Actividad: TActividad;
  i, Indice: Integer;
begin
  btnReporte.Enabled := DataSet.RecordCount > 0;
  btnProcesar.Enabled := Not btnReporte.Enabled;
  pnlGenerar.Visible := btnReporte.Enabled;
  if pnlGenerar.Visible then
  begin
    // Obtener los datos de filtrado
    try
      roqReporte.DisableControls;
      roqReporte.First;
      ListaActividad := TStringList.Create;
      ListaActividad.Clear;

      while Not roqReporte.Eof do
      begin
        Indice := ListaActividad.IndexOf(roqReporte.FieldByName('sNumeroActividad').AsString);
        if Indice = -1 then
        begin
          Actividad := TActividad.Create;
          Actividad.ListaPartida.Add(roqReporte.FieldByName('sIdPartidaAnexo').AsString);

          ListaActividad.AddObject(roqReporte.FieldByName('sNumeroActividad').AsString, Actividad);
        end
        else
          if TActividad(ListaActividad.Objects[Indice]).ListaPartida.IndexOf(roqReporte.FieldByName('sIdPartidaAnexo').AsString) = -1 then
            TActividad(ListaActividad.Objects[Indice]).ListaPartida.Add(roqReporte.FieldByName('sIdPartidaAnexo').AsString);

        roqReporte.Next;
      end;

      // Cargar los datos a los combo
      cbActividad.Items.Clear;
      cbActividad.Items.Add('< TODAS >');
      cbPartida.Items.Clear;
      cbPartida.Enabled;
      for i := 0 to ListaActividad.Count -1 do
        cbActividad.Items.Add(ListaActividad[i]);
      cbActividad.ItemIndex := 0;
    finally
      roqReporte.First;
      roqReporte.EnableControls;
    end;
    cbActividad.SetFocus;
  end;
end;

procedure TfrmRepExcelNew.EstablecerFormato(pInicial: Integer; Ren: Integer);
const
  xlAutomatic = -4105;
  xlCenter = -4108;
  xlJustify = -4130;
  xlNone = -4142;
  xlContinuous = 1;
  xlThemeColorDark1 = 1;
  xlSolid = 1;
  xlThin = 2;
  xlDiagonalDown = 5;
  xlDiagonalUp = 6;
  xlEdgeLeft = 7;
  xlEdgeTop = 8;
  xlEdgeBottom = 9;
  xlEdgeRight = 10;
  xlInsideVertical = 11;
  xlInsideHorizontal = 12;

var
  myCol: Integer;

begin
  Excel.ActiveSheet.Cells[5, 1].Value := 'FECHA';
  Excel.ActiveSheet.Cells[5, 2].Value := 'FOLIO';
  Excel.ActiveSheet.Cells[5, 3].Value := 'ACTIVIDAD';
  Excel.ActiveSheet.Cells[5, 4].Value := 'DESCRIPCION DEL PROGRAMA';
  Excel.ActiveSheet.Cells[5, 5].Value := 'DESCRIPCION DE LA BITACORA';
  Excel.ActiveSheet.Cells[5, 6].Value := 'FRENTE';
  Excel.ActiveSheet.Cells[5, 7].Value := 'CLASIFICACION';
  Excel.ActiveSheet.Cells[5, 8].Value := 'INICIO';
  Excel.ActiveSheet.Cells[5, 9].Value := 'TERMINO';
  Excel.ActiveSheet.Cells[5, 10].Value := 'PARTIDA';
  Excel.ActiveSheet.Cells[5, 11].Value := 'DESCRIPCION';
  Excel.ActiveSheet.Cells[5, 12].Value := 'CANTIDAD';
  if pInicial = 2 then
  begin
    Excel.ActiveSheet.Cells[5, 13].Value := 'U.M.';
    Excel.ActiveSheet.Cells[5, 14].Value := 'Descripción Material';
    myCol := 14
  end
  else
  begin
    Excel.ActiveSheet.Cells[5, 13].Value := 'JORNADA';
    myCol := 13;
  end;

  Excel.ActiveSheet.Cells[5, myCol +1].Value := 'Precio Unitario (PESOS)';
  Excel.ActiveSheet.Cells[5, myCol +2].Value := 'Precio Unitario (USD)';
  Excel.ActiveSheet.Cells[5, myCol +3].Value := 'TOTAL PESOS';
  Excel.ActiveSheet.Cells[5, myCol +4].Value := 'TOTAL USD';

  // Ajustar las columnas, formato y alineación
  Excel.ActiveSheet.Columns[1].Select;
  Excel.Selection.NumberFormat := 'dd-mmm-aa;@';
  Excel.Selection.ColumnWidth := 9;

  Excel.ActiveSheet.Columns[2].Select;
  Excel.Selection.WrapText := True;
  Excel.Selection.ColumnWidth := 19.17;

  Excel.ActiveSheet.Columns[3].Select;
  Excel.Selection.WrapText := True;
  Excel.Selection.ColumnWidth := 13.17;

  Excel.ActiveSheet.Columns[4].Select;
  Excel.Selection.WrapText := True;
  Excel.Selection.ColumnWidth := 48.83;

  Excel.ActiveSheet.Columns[5].Select;
  Excel.Selection.HorizontalAlignment := xlJustify;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.WrapText := True;
  Excel.Selection.ColumnWidth := 60;

  Excel.ActiveSheet.Columns[6].Select;
  Excel.Selection.HorizontalAlignment := xlJustify;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.WrapText := True;
  Excel.Selection.ColumnWidth := 9;

  Excel.ActiveSheet.Columns[7].Select;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.WrapText := True;
  Excel.Selection.ColumnWidth := 11.17;

  Excel.ActiveSheet.Columns['H:I'].Select;
  Excel.Selection.ColumnWidth := 11;

  Excel.ActiveSheet.Columns[10].Select;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.WrapText := True;
  Excel.Selection.ColumnWidth := 11.17;

  Excel.ActiveSheet.Columns[11].Select;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.WrapText := True;
  Excel.Selection.ColumnWidth := 30;

  Excel.ActiveSheet.Columns[12].Select;
  Excel.Selection.ColumnWidth := 11.17;

  Excel.ActiveSheet.Columns[13].Select;
  Excel.Selection.NumberFormat := '#,##0.00000000_);(#,##0.00000000)';
  Excel.Selection.ColumnWidth := 11.17;

  if pInicial = 2 then
  begin
    Excel.ActiveSheet.Columns[14].Select;
    Excel.Selection.HorizontalAlignment := xlCenter;
    Excel.Selection.VerticalAlignment := xlCenter;
    Excel.Selection.WrapText := True;
    Excel.Selection.ColumnWidth := 30;

    Excel.ActiveSheet.Columns['O:R'].Select;
  end
  else
    Excel.ActiveSheet.Columns['N:Q'].Select;

  Excel.Selection.NumberFormat := '$#,##0.00;-$#,##0.00';
  Excel.Selection.ColumnWidth := 11.17;

  if pInicial = 2 then
    Excel.Range['A5:R5'].Select
  else
    Excel.Range['A5:Q5'].Select;
  Excel.Selection.Interior.Pattern := xlSolid;
  Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
  Excel.Selection.Interior.ThemeColor := xlThemeColorDark1;
  Excel.Selection.Interior.TintAndShade := -0.249946592608417;
  Excel.Selection.Interior.PatternTintAndShade := 0;
  Excel.Selection.Font.Bold := True;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;
  Excel.Selection.WrapText := True;

  if pInicial = 2 then
    Excel.Range['A5:R' + IntToStr(Ren -1)].Select
  else
    Excel.Range['A5:Q' + IntToStr(Ren -1)].Select;

  Excel.Selection.Borders[xlDiagonalDown].LineStyle := xlNone;
  Excel.Selection.Borders[xlDiagonalUp].LineStyle := xlNone;

  Excel.Selection.Borders[xlEdgeLeft].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlEdgeLeft].ColorIndex := 0;
  Excel.Selection.Borders[xlEdgeLeft].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeLeft].Weight := xlThin;

  Excel.Selection.Borders[xlEdgeTop].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlEdgeTop].ColorIndex := 0;
  Excel.Selection.Borders[xlEdgeTop].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeTop].Weight := xlThin;

  Excel.Selection.Borders[xlEdgeBottom].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlEdgeBottom].ColorIndex := 0;
  Excel.Selection.Borders[xlEdgeBottom].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeBottom].Weight := xlThin;

  Excel.Selection.Borders[xlEdgeRight].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlEdgeRight].ColorIndex := 0;
  Excel.Selection.Borders[xlEdgeRight].TintAndShade := 0;
  Excel.Selection.Borders[xlEdgeRight].Weight := xlThin;

  Excel.Selection.Borders[xlInsideVertical].LineStyle := xlContinuous;
  Excel.Selection.Borders[xlInsideVertical].ColorIndex := 0;
  Excel.Selection.Borders[xlInsideVertical].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideVertical].Weight := xlThin;

  Excel.Selection.Borders[xlInsideHorizontal].LineStyle := xlNone;
  Excel.Selection.Borders[xlInsideHorizontal].ColorIndex := 0;
  Excel.Selection.Borders[xlInsideHorizontal].TintAndShade := 0;
  Excel.Selection.Borders[xlInsideHorizontal].Weight := xlThin;

  Excel.Range['A:D'].Select;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;

  Excel.Range['F:R'].Select;
  Excel.Selection.HorizontalAlignment := xlCenter;
  Excel.Selection.VerticalAlignment := xlCenter;

  Excel.Selection.Font.Name := 'Arial';
  Excel.Selection.Font.Size := 8;

  Excel.Columns['J:K'].Select;
  Excel.Selection.Interior.PatternColorIndex := xlAutomatic;
  Excel.Selection.Interior.Color := 65535;
  Excel.Selection.Interior.TintAndShade := 0;
  Excel.Selection.Interior.PatternTintAndShade := 0;

  if pInicial = 2 then
    Excel.ActiveSheet.Columns['O:R'].Select
  else
    Excel.ActiveSheet.Columns['N:Q'].Select;
  Excel.Selection.NumberFormat := '$#,##0.00;-$#,##0.00';
  Excel.Selection.ColumnWidth := 11;

  Excel.ActiveSheet.Cells[1,1].Select;

  Excel.ActiveWindow.Zoom := 85;
end;

procedure TfrmRepExcelNew.FormShow(Sender: TObject);
begin
  FechaInicio.Date := Now() -30;
  FechaFinal.Date := Now();
end;

end.
