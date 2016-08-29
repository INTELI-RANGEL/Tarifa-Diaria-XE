unit frm_AjustesDiarios;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ZAbstractRODataset, ZDataset, ExtCtrls, Grids, DBGrids, DBClient,
  StdCtrls, DBCtrls, Mask, JvExMask, JvToolEdit, JvBaseEdits, frm_Connection,
  JvMemoryDataset, StrUtils;

type
  THorarios = class
    sHoraInicio,
    sHoraFinal: String;
  end;

  TfrmAjustesDiarios = class(TForm)
    roqOrdenes: TZReadOnlyQuery;
    roqReporte: TZReadOnlyQuery;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    DBGrid1: TDBGrid;
    dsOrdenes: TDataSource;
    Splitter1: TSplitter;
    dsDatos: TDataSource;
    Panel4: TPanel;
    Panel5: TPanel;
    gridActividades: TDBGrid;
    rgAnexo: TRadioGroup;
    cbActividad: TComboBox;
    cbHorarios: TComboBox;
    Label1: TLabel;
    Label2: TLabel;
    DBMemo1: TDBMemo;
    pnlEditar: TPanel;
    dJornadaAjustadaEdit: TJvCalcEdit;
    Label4: TLabel;
    btnAceptar: TButton;
    btnCancelar: TButton;
    memDatos: TJvMemoryData;
    btnGrabar: TButton;
    btnCerrar: TButton;
    roqExec: TZReadOnlyQuery;
    Panel6: TPanel;
    dJornadaEdit: TJvCalcEdit;
    Label3: TLabel;
    procedure FormShow(Sender: TObject);
    procedure roqOrdenesAfterScroll(DataSet: TDataSet);
    procedure cbActividadChange(Sender: TObject);
    procedure cbHorariosChange(Sender: TObject);
    procedure rgAnexoClick(Sender: TObject);
    procedure gridActividadesDblClick(Sender: TObject);
    procedure gridActividadesDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure btnGrabarClick(Sender: TObject);
    procedure btnCerrarClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    CargandoMem: Boolean;
    FormaEdit: TForm;
  public
    sContrato: String;
    sNumeroOrden: String;
    dIdFecha: TDate;
  end;

var
  frmAjustesDiarios: TfrmAjustesDiarios;

implementation

{$R *.dfm}

procedure TfrmAjustesDiarios.btnCerrarClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmAjustesDiarios.btnGrabarClick(Sender: TObject);
var
  Cadena,
  FechaSQL: String;
  Marca: LongInt;
  LocCursor: TCursor;
  wAnyo, wMes, wDia: Word;
  Dia, Mes, Año: String;
  OldFiltered: Boolean;
begin
  LocCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;

    DecodeDate(dIdFecha, wAnyo, wMes, wDia);
    Dia := RightStr('0' + vartostr(wDia), 2);
    Mes := RightStr('0' + vartostr(wMes), 2);
    Año := vartostr(wAnyo);

    FechaSQL := QuotedStr(Año + '/' + Mes + '/' + Dia);

    Marca := memDatos.RecNo;
    OldFiltered := memDatos.Filtered;
    try
      memDatos.DisableControls;
      OldFiltered := False;
      memDatos.First;
      while not memDatos.Eof do
      begin
        if StrToFloat(memDatos.FieldByName('dAjuste').AsString) = 0 then
        begin
          Cadena := 'DELETE FROM ' +
                    '  bitacoradeajustes ' +
                    'WHERE ' +
                    '  sContrato = ' + QuotedStr(memDatos.FieldByName('sContrato').AsString) + ' AND ' +
                    '  sNumeroOrden = ' + QuotedStr(memDatos.FieldByName('sNumeroOrden').AsString) + ' AND ' +
                    '  sNumeroActividad = ' + QuotedStr(memDatos.FieldByName('sNumeroActividad').AsString) + ' AND ' +
                    '  sIdPartidaAnexo = ' + QuotedStr(memDatos.FieldByName('sIdPartidaAnexo').AsString) + ' AND ' +
                    '  dFecha = ' + FechaSQL;
          if Not memDatos.FieldByName('sIdCategoria').IsNull then
            Cadena := Cadena + ' AND sIdCategoria = ' + QuotedStr(memDatos.FieldByName('sIdCategoria').AsString);
        end
        else
        begin
          Cadena := 'INSERT INTO bitacoradeajustes ' +
                    '(sContrato, sNumeroOrden, sNumeroActividad, sIdPartidaAnexo, dFecha, dAjuste';
          if Not memDatos.FieldByName('sIdCategoria').IsNull then
            Cadena := Cadena + ', sIdCategoria';
          Cadena := Cadena + ') VALUES (' + QuotedStr(memDatos.FieldByName('sContrato').AsString) + ', ' +
                              QuotedStr(memDatos.FieldByName('sNumeroOrden').AsString) + ', ' +
                              QuotedStr(memDatos.FieldByName('sNumeroActividad').AsString) + ', ' +
                              QuotedStr(memDatos.FieldByName('sIdPartidaAnexo').AsString) + ', ' +
                              FechaSQL + ', ' +
                              memDatos.FieldByName('dAjuste').AsString;
          if Not memDatos.FieldByName('sIdCategoria').IsNull then
            Cadena := Cadena + ', ' + QuotedStr(memDatos.FieldByName('sIdCategoria').AsString);
          Cadena := Cadena + ') ON DUPLICATE KEY UPDATE dAjuste = ' + memDatos.FieldByName('dAjuste').AsString;
        end;

        roqExec.SQL.Text := Cadena;
        roqExec.ExecSQL;

        memDatos.Next;
      end;
    finally
      memDatos.Filtered := OldFiltered;
      if Marca <= memDatos.RecordCount then
        memDatos.RecNo := Marca;

      memDatos.EnableControls;
    end;
  finally
    Screen.Cursor := LocCursor;
    btnGrabar.Enabled := False;
  end;
end;

procedure TfrmAjustesDiarios.cbActividadChange(Sender: TObject);
var
  i, Indice: Integer;
  Horarios: THorarios;
begin
  roqReporte.Filtered := False;
  try
    if cbActividad.ItemIndex = 0 then
      roqReporte.Filter := 'sNumeroOrden = ' + QuotedStr(roqOrdenes.FieldByName('sNumeroOrden').AsString) + ' AND iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex) + ' AND iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex)
    else
      roqReporte.Filter := 'sNumeroOrden = ' + QuotedStr(roqOrdenes.FieldByName('sNumeroOrden').AsString) + ' AND iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex) + ' AND iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex) + ' AND sNumeroActividad = ' + cbActividad.Text;
    roqReporte.Filtered := True;

    roqReporte.First;
    cbHorarios.Items.Clear;
    cbHorarios.Items.AddObject('< TODOS >', TObject.Create);
    while not roqReporte.Eof do
    begin
      Horarios := THorarios.Create;
      Horarios.sHoraInicio := roqReporte.FieldByName('sHoraInicio').AsString;
      Horarios.sHoraFinal := roqReporte.FieldByName('sHoraFinal').AsString;

      if cbHorarios.Items.IndexOf(roqReporte.FieldByName('sHoraInicio').AsString + ' - ' + roqReporte.FieldByName('sHoraFinal').AsString) < 0 then
        cbHorarios.Items.AddObject(roqReporte.FieldByName('sHoraInicio').AsString + ' - ' + roqReporte.FieldByName('sHoraFinal').AsString, Horarios);

      roqReporte.Next;
    end;
  finally
    roqReporte.Filtered := False;
  end;

  cbHorarios.ItemIndex := 0;
  cbHorarios.OnChange(cbHorarios);
end;

procedure TfrmAjustesDiarios.cbHorariosChange(Sender: TObject);
begin
  memDatos.Filtered := False;
  if cbHorarios.ItemIndex = 0 then
  begin
    if cbActividad.ItemIndex = 0 then
      memDatos.Filter := '(sNumeroOrden = ' + QuotedStr(roqOrdenes.FieldByName('sNumeroOrden').AsString) + ') AND (iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex) + ')'
    else
      memDatos.Filter := '(sNumeroOrden = ' + QuotedStr(roqOrdenes.FieldByName('sNumeroOrden').AsString) + ') AND (iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex) + ') AND (sNumeroActividad = ' + cbActividad.Text + ')';
  end
  else
    if cbActividad.ItemIndex = 0 then
      memDatos.Filter := '(sNumeroOrden = ' + QuotedStr(roqOrdenes.FieldByName('sNumeroOrden').AsString) + ') AND (iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex) + ') AND (sHoraInicio = ' + QuotedStr(THorarios(cbHorarios.Items.Objects[cbHorarios.ItemIndex]).sHoraInicio) + ') AND (sHoraFinal = ' + QuotedStr(THorarios(cbHorarios.Items.Objects[cbHorarios.ItemIndex]).sHoraInicio) + ')'
    else
      memDatos.Filter := '(sNumeroOrden = ' + QuotedStr(roqOrdenes.FieldByName('sNumeroOrden').AsString) + ') AND (iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex) + ') AND (sNumeroActividad = ' + cbActividad.Text + ') AND (sHoraInicio = ' + QuotedStr(THorarios(cbHorarios.Items.Objects[cbHorarios.ItemIndex]).sHoraInicio) + ') AND (sHoraFinal = ' + QuotedStr(THorarios(cbHorarios.Items.Objects[cbHorarios.ItemIndex]).sHoraFinal) + ')';
  memDatos.Filtered := True;
end;

procedure TfrmAjustesDiarios.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  CanClose := True;

  if btnGrabar.Enabled then
    case MessageDlg('¿Desea grabar los cambios antes de cerrar la ventana?', mtConfirmation, [mbYes, mbNo, mbCancel], 0) of
      mrYes: btnGrabar.Click;
      mrNo: ;
      mrCancel: CanClose := False;
    end;
end;

procedure TfrmAjustesDiarios.FormShow(Sender: TObject);
var
  LocCursor: TCursor;
begin
  LocCursor := Screen.Cursor;
  try
    Screen.Cursor := crHourGlass;

    roqOrdenes.ParamByName('sContrato').AsString := sNumeroOrden;
    roqOrdenes.ParamByName('Fecha').AsDate := dIdFecha;
    if roqOrdenes.Active then
      roqOrdenes.Refresh
    else
      roqOrdenes.Open;
  finally
    Screen.Cursor := LocCursor;
  end;
end;

procedure TfrmAjustesDiarios.gridActividadesDblClick(Sender: TObject);
var
  Ajuste: Real;
begin
  FormaEdit := TForm.Create(Self);
  try
    FormaEdit.Height := pnlEditar.Height +40;
    FormaEdit.Width := pnlEditar.Width +20;
    pnlEditar.Parent := FormaEdit;
    pnlEditar.Visible := True;
    pnlEditar.Align := alClient;
    dJornadaEdit.Value := memDatos.FieldByName('dJornadaRed').AsFloat;
    dJornadaAjustadaEdit.Value := memDatos.FieldByName('dJornadaAjustada').AsFloat;
    if FormaEdit.ShowModal = mrOk then
    begin
      // Actualizar los datos
      Ajuste := (dJornadaAjustadaEdit.Value - dJornadaEdit.Value);
      memDatos.Edit;
      memDatos.FieldByName('dAjuste').AsFloat := Ajuste;
      memDatos.FieldByName('dJornadaAjustada').AsFloat := dJornadaAjustadaEdit.Value;
      memDatos.Post;

      btnGrabar.Enabled := True;
    end;
  finally
    pnlEditar.Align := alClient;
    pnlEditar.Visible := False;
    pnlEditar.Parent := Self;
    FormaEdit.Destroy;
  end;
end;

procedure TfrmAjustesDiarios.gridActividadesDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
var
  i: Integer;
  UnColor: TColor;
begin
  if memDatos.FieldByName('dJornadaRed').AsString <> memDatos.FieldByName('dJornadaAjustada').AsString then
    UnColor := clBlue
  else
    UnColor := clBlack;

  case Column.Alignment of
    taCenter : i := (Rect.Right - Rect.Left) div 2 - (sender as TDBGrid).canvas.TextWidth(Column.Field.asstring) div 2;
    taLeftJustify : i := 0 +3;
    taRightJustify : i := (Rect.Right - Rect.Left) - (sender as TDBGrid).canvas.TextWidth(Column.Field.asstring) -3;
  end;

  (sender as TDBGrid).canvas.FillRect(Rect);
  with (Sender As TDBGrid).Canvas do
  begin
    Font.Color := UnColor;
    FillRect(Rect);
    TextOut(Rect.Left +i, Rect.Top +2, Column.Field.AsString);
  end;
end;

procedure TfrmAjustesDiarios.rgAnexoClick(Sender: TObject);
begin
  roqReporte.Filtered := False;
  roqReporte.Filter := '(sNumeroOrden = ' + QuotedStr(roqOrdenes.FieldByName('sNumeroOrden').AsString) + ') AND iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex) + ' AND (iIdOrdenTipoAnexo = ' + IntToStr(rgAnexo.ItemIndex) + ')';
  roqReporte.Filtered := True;

  roqReporte.First;
  cbActividad.Items.Clear;
  cbActividad.Items.AddObject('< TODOS >', TObject(-1));
  while not roqReporte.Eof do
  begin
    if cbActividad.Items.IndexOf(roqReporte.FieldByName('sNumeroActividad').AsString) < 0 then
      cbActividad.Items.AddObject(roqReporte.FieldByName('sNumeroActividad').AsString, TObject(roqReporte.RecordCount -1));

    roqReporte.Next;
  end;

  if cbActividad.Items.Count > 0 then
  begin
    cbActividad.ItemIndex := 0;
    cbActividad.OnChange(cbActividad);
  end;

  cbHorarios.Enabled := (rgAnexo.ItemIndex = 0) or (rgAnexo.ItemIndex = 1);
end;

procedure TfrmAjustesDiarios.roqOrdenesAfterScroll(DataSet: TDataSet);
var
  i: Integer;
  LocCursor: TCursor;
begin
  LocCursor := Screen.Cursor;
  CargandoMem := True;
  try
    Screen.Cursor := crHourGlass;

    roqReporte.Close;
    roqReporte.Filtered := False;

    roqReporte.ParamByName('sContrato').AsString := sContrato;
    roqReporte.ParamByName('sNumeroOrden').AsString := roqOrdenes.FieldByName('sNumeroOrden').AsString;
    roqReporte.ParamByName('Fecha').AsDate := dIdFecha;   // StrToDate('03/03/2016');

    roqReporte.Open;

    // Verificar si se han cargado los campos de edición de datos
    memDatos.Close;
    if memDatos.FieldDefs.Count <> roqReporte.FieldDefs.Count +1 then
    begin
      memDatos.FieldDefs.Clear;
      memDatos.FieldDefs.Add('Id', ftInteger, 0, True);
      for i := 0 to roqReporte.FieldDefs.Count -1 do
        if roqReporte.FieldDefs.Items[i].DataType = ftMemo then
          memDatos.FieldDefs.Add(roqReporte.FieldDefs.Items[i].Name, ftWideString, 800, False)    //roqReporte.FieldDefs.Items[i].Required)
        else
          memDatos.FieldDefs.Add(roqReporte.FieldDefs.Items[i].Name, roqReporte.FieldDefs.Items[i].DataType, roqReporte.FieldDefs.Items[i].Size, False);  //roqReporte.FieldDefs.Items[i].Required);

      //memDatos.CreateDataSet;
    end;

    if Not memDatos.Active then
      memDatos.Open;

    //memDatos.EmptyDataSet;
    memDatos.EmptyTable;
    roqReporte.First;
    cbActividad.Items.Clear;
    while Not roqReporte.Eof do
    begin
      memDatos.Append;
      memDatos.FieldByName('Id').AsInteger := memDatos.RecordCount;
      for i := 0 to roqReporte.FieldDefs.Count -1 do
        if roqReporte.FieldDefs.Items[i].DataType = ftMemo then
          memDatos.FieldByName(roqReporte.FieldDefs.Items[i].Name).AsString := roqReporte.FieldByName(roqReporte.FieldDefs.Items[i].Name).AsString
        else
          memDatos.FieldByName(roqReporte.FieldDefs.Items[i].Name).AsVariant := roqReporte.FieldByName(roqReporte.FieldDefs.Items[i].Name).AsVariant;
      memDatos.Post;

      roqReporte.Next;
    end;

    rgAnexo.ItemIndex := 0;
    rgAnexoClick(rgAnexo);
  finally
    CargandoMem := False;
    Screen.Cursor := LocCursor;
  end;
end;

end.
