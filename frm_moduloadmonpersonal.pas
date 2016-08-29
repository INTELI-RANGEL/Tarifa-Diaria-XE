unit frm_moduloadmonpersonal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, OleCtrls, 
  DB, ZAbstractRODataset, ZAbstractDataset, ZDataset, Global, 
  NxColumnClasses, NxColumns, NxScrollControl, NxCustomGridControl,
  NxCustomGrid, NxGrid, NxClasses, NxEdit, DateUtils;

type
  TfrmModuloAdmonPersonal = class(TForm)
    pBarraAvance: TProgressBar;
    QryFolios: TZQuery;
    dsFolios: TDataSource;
    QryFoliossContrato: TStringField;
    QryFoliossIdFolio: TStringField;
    QryFoliossNumeroOrden: TStringField;
    QryFoliossDescripcionCorta: TStringField;
    QryFoliossOficioAutorizacion: TStringField;
    QryFoliosmDescripcion: TMemoField;
    QryFoliossIdTipoOrden: TStringField;
    QryFoliossApoyo: TStringField;
    QryFoliossIdPlataforma: TStringField;
    QryFoliossIdPernocta: TStringField;
    QryFoliosdFiProgramado: TDateField;
    QryFoliosdFfProgramado: TDateField;
    QryFolioscIdStatus: TStringField;
    QryFoliosmComentarios: TMemoField;
    QryFoliossFormato: TStringField;
    QryFoliosiConsecutivo: TIntegerField;
    QryFoliosiConsecutivoTierra: TIntegerField;
    QryFoliosiJornada: TIntegerField;
    QryFolioslGeneraAnexo: TStringField;
    QryFolioslGeneraConsumibles: TStringField;
    QryFolioslGeneraPersonal: TStringField;
    QryFolioslGeneraEquipo: TStringField;
    QryFoliossDepsolicitante: TStringField;
    QryFoliosdFechaInicioT: TDateField;
    QryFoliosdFechaSitioM: TDateField;
    QryFoliossEquipo: TStringField;
    QryFoliossPozo: TStringField;
    QryFoliosdFechaElaboracion: TDateField;
    QryFoliossPuestoPep: TStringField;
    QryFoliossFirmantePep: TStringField;
    QryFoliossPuestocia: TStringField;
    QryFoliossFirmantecia: TStringField;
    QryFolioslMostrarAvanceProgramado: TStringField;
    QryFoliossTipoOrden: TStringField;
    QryFoliosbAvanceFrente: TStringField;
    QryFoliosbAvanceContrato: TStringField;
    QryFoliosbComentarios: TStringField;
    QryFoliosbPermisos: TStringField;
    QryFoliosbTipoAdmon: TStringField;
    QryFoliosbCostaFuera: TStringField;
    QryFoliossTipoPrograma: TStringField;
    QryFoliossTipoImpresionActividad: TStringField;
    QryFoliossTipoAvanceAdmon: TStringField;
    QryFoliosiDecimales: TIntegerField;
    QryFoliosiNiveles: TIntegerField;
    QryFolioslImprimeProgramado: TStringField;
    QryFolioslImprimeFisico: TStringField;
    QryFolioslImprimePlaticas: TStringField;
    QryFolioslImprimePersonalTM: TStringField;
    QryFolioslPersonalxPartida: TStringField;
    QryFolioslImprimeFases: TStringField;
    QryFolioslMostrarPartidasReportes: TStringField;
    QryFolioslMostrarPartidasGeneradores: TStringField;
    QryFoliosdFechaIniPReportes: TDateField;
    QryFoliosdFechaFinPReportes: TDateField;
    QryFoliosdFechaIniPGeneradores: TDateField;
    QryFoliosdFechaFinPGeneradores: TDateField;
    QryFolioslEstado: TStringField;
    Panel1: TPanel;
    Label1: TLabel;
    dIdFecha: TDateTimePicker;
    Label2: TLabel;
    ComboBox: TComboBox;
    Panel2: TPanel;
    NextGrid: TNextGrid;
    NxCheckBoxColumn1: TNxCheckBoxColumn;
    NxTextColumn1: TNxTextColumn;
    NxTextColumn2: TNxTextColumn;
    NxNumberColumn1: TNxNumberColumn;
    cmbCategoria: TComboBox;
    Label3: TLabel;
    Panel3: TPanel;
    Button1: TButton;
    NxComboBoxColumn1: TNxComboBoxColumn;
    NxTextColumn3: TNxTextColumn;
    Button2: TButton;
    Label4: TLabel;
    NxNumberEdit1: TNxNumberEdit;
    Label5: TLabel;
    LabelPernoctas: TLabel;
    function IsIn(Valor: string; Lista: TStringList): Boolean;
    procedure FormShow(Sender: TObject);
    procedure dIdFechaChange(Sender: TObject);
    procedure ComboBoxChange(Sender: TObject);
    procedure cmbCategoriaChange(Sender: TObject);
    Procedure CargarPersonalDB;
    procedure Button1Click(Sender: TObject);
    procedure NextGridAfterEdit(Sender: TObject; ACol, ARow: Integer;
      Value: WideString);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure NxComboBoxColumn1Select(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmModuloAdmonPersonal: TfrmModuloAdmonPersonal;
  PrimerFolio: String;

procedure Split(Delimiter: Char; Str: string; ListOfStrings: TStrings);

implementation

uses frm_connection;

{$R *.dfm}

procedure Split(Delimiter: Char; Str: string; ListOfStrings: TStrings);
begin
   ListOfStrings.Clear;
   ListOfStrings.Delimiter     := Delimiter;
   ListOfStrings.DelimitedText := Str;
end;

procedure TfrmModuloAdmonPersonal.Button1Click(Sender: TObject);
Var
  Query: TZQuery;
  i: Integer;
  Lista: TStringList;
  IdPersonal,
  sIdPlataforma,
  sIdPernocta,
  Folio: String;
  IdDiario: Integer;
begin
  Try
    if cmbCategoria.Text <> '' then begin
      Query := TZQuery.Create(Self);
      Query.Connection := Connection.zConnection;

      Lista := TStringList.Create;
      Split('|', cmbCategoria.Text, Lista);
      IdPersonal := Trim(Lista[0]);

      Query.Active := False;
      Query.SQL.Clear;
      Query.SQL.Text := 'SELECT MAX(iIdDiario) AS MaxDiario FROM bitacoradepersonal WHERE dIdFecha = :Fecha AND sContrato = :Contrato ';
      Query.ParamByName('Contrato').AsString := Global_Contrato;
      Query.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
      Query.Open;

      if Trim(ComboBox.Text) <> '' then begin
        Folio := Trim(ComboBox.Text);
        Connection.QryBusca.Active := False;
        Connection.QryBusca.SQL.Text := 'SELECT * FROM ordenesdetrabajo WHERE sNumeroOrden = ' + QuotedStr(Trim(ComboBox.Text)) + ' AND sContrato = ' + QuotedStr(Global_Contrato);
        Connection.QryBusca.Open;
        if Connection.QryBusca.RecordCount > 0 then begin
          sIdPlataforma := Connection.QryBusca.FieldByName('sIdPlataforma').AsString;
          sIdPernocta := Connection.QryBusca.FieldByName('sIdPernocta').AsString;
        end else begin
          sIdPlataforma := '@';
          sIdPernocta := '@';
        end;
      end else begin
        Folio := '@';
        sIdPlataforma := '@';
        sIdPernocta := '@';
      end;

      IdDiario := Query.FieldByName('MaxDiario').AsInteger;

      pBarraAvance.Max := NextGrid.RowCount;
      for i := 0 to NextGrid.RowCount - 1 do begin
        pBarraAvance.Position := pBarraAvance.Position + 1;
        
        Query.Active := False;
        Query.SQL.Text := 'DELETE FROM bitacoradepersonal WHERE sContrato = :Contrato AND sNumeroOrden = :Folio AND dIdFecha = :Fecha AND sIdPersonal = :Personal LIMIT 1';
        Query.ParamByName('Folio').AsString := Folio;
        Query.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
        Query.ParamByName('Contrato').AsString := Global_Contrato;
        Query.ParamByName('Personal').AsString := NextGrid.Cells[1, i];
        Query.ExecSQL;

        if NextGrid.Cells[0, i] = 'True' then begin
          Query.Active := False;
          Query.SQL.Clear;
          Query.SQL.Add( '' +
                          'INSERT IGNORE INTO bitacoradepersonal ' +
                          ' (sContrato, sNumeroOrden, dIdFecha, iIdDiario, sIdPersonal, sTipoObra, sDescripcion, sIdPernocta, sIdPlataforma, sHoraInicio, sHoraFinal, dCantidad, sTipoPernocta, dCantHH)' +
                          ' VALUES ' +
                          ' (:Contrato, :Orden, :Fecha, :IdDiario, :IdItem, :TipoPersonal, :Descripcion, :IdPernocta, :Plataforma, :HoraInicio, :HoraFinal, :Cantidad, :TipoPernocta, :CantidadHH); ' +
                          '');
          Query.ParamByName('Contrato').AsString := Global_Contrato;
          Query.ParamByName('Orden').AsString := Folio;
          Query.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
          Query.ParamByName('IdDiario').AsInteger := IdDiario;
          Query.ParamByName('IdItem').AsString := NextGrid.Cells[1, i];
          Query.ParamByName('TipoPersonal').AsString := IdPersonal;
          Query.ParamByName('Descripcion').AsString := NextGrid.Cells[2, i];
          Query.ParamByName('IdPernocta').AsString := sIdPernocta;
          Query.ParamByName('Plataforma').AsString := sIdPlataforma;
          Query.ParamByName('HoraInicio').AsString := '00:00';
          Query.ParamByName('HoraFinal').AsString := '00:00';
          Query.ParamByName('Cantidad').AsFloat := NextGrid.Cell[3, i].AsFloat;
          Query.ParamByName('CantidadHH').AsFloat := NextGrid.Cell[3, i].AsFloat;
          Query.ParamByName('TipoPernocta').AsString := NextGrid.Cells[5, NextGrid.SelectedRow];
          Query.ExecSQL;
          
          Inc(IdDiario);
        end;
      end;
      pBarraAvance.Position := 0;
      pBarraAvance.Max := 0;

      Query.SQL.Text := 'DELETE FROM bitacoradepernocta WHERE dIdFecha = :Fecha AND sContrato = :Contrato AND sNumeroOrden = :Folio ';
      Query.Params.ParamByName('Contrato').AsString := Global_Contrato;
      Query.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
      Query.Params.ParamByName('Folio').AsString := ComboBox.Text;
      Query.ExecSQL;

      Query.SQL.Text := 'INSERT INTO bitacoradepernocta (sContrato, dIdFecha, sNumeroOrden, dCantidad) ' +
                        'VALUES (:Contrato, :Fecha, :Folio, :Cantidad) ';
      Query.Params.ParamByName('Contrato').AsString := Global_Contrato;
      Query.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
      Query.Params.ParamByName('Folio').AsString := ComboBox.Text;
      Query.Params.ParamByName('Cantidad').AsFloat := NxNumberEdit1.Value;
      Query.ExecSQL;
      
      CargarPersonalDB;
    end;
  Finally
    Lista.Free;
    Query.Free;
  End;
end;

procedure TfrmModuloAdmonPersonal.Button2Click(Sender: TObject);
Var
  RegistrarFolio, RegistrarFolioAux: String;
  Query: TZQuery;
begin
  Try
    Query := TZQuery.Create(Self);
    Query.Connection := Connection.zConnection;

    Query.SQL.Text := 'START TRANSACTION;';
    Query.ExecSQL;

    if MessageDlg('Se borrará todo el personal del día de hoy y se importará el personal del día anterior. '+#10#13+'¿Desea Continuar?', mtConfirmation, [mbOk, mbCancel], 0) = mrOk then begin
      Connection.zCommand.Active := False;
      Connection.zCommand.SQL.Text := 'DELETE FROM bitacoradepersonal WHERE sContrato = :Contrato AND sTipoObra <> "PU" AND dIdFecha = :Fecha';
      Connection.zCommand.ParamByName('Contrato').AsString := Global_Contrato;
      Connection.zCommand.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
      Connection.zCommand.ExecSQL;

      Connection.QryBusca.SQL.Text := 'SELECT * FROM bitacoradepersonal WHERE sContrato = :Contrato AND sTipoObra <> "PU" AND dIdFecha = :Fecha';
      Connection.QryBusca.ParamByName('Contrato').AsString := Global_Contrato;
      Connection.QryBusca.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', IncDay(dIdFecha.Date, -1));
      Connection.QryBusca.Open;
      RegistrarFolioAux := Trim(PrimerFolio); //ComboBox.Items.ValueFromIndex[0];
      if Connection.QryBusca.RecordCount > 0 then begin
        while Not Connection.QryBusca.Eof do begin
            RegistrarFolio := Trim(Connection.QryBusca.FieldByName('sNumeroOrden').AsString);
            if RegistrarFolio <> '@' then begin
              if ComboBox.Items.IndexOf(RegistrarFolio) < 0 then begin
                if MessageDlg('No se encontró reportado el folio "' + RegistrarFolio + '" el día ' + FormatDateTime('yyyy-mm-dd', dIdFecha.Date) + ' ¿desea asignarlo al folio ' + RegistrarFolioAux + '?' +#10#13+ 'Si selecciona "No" ningún cambio se guardará.', mtConfirmation, [mbOk, mbNo], 0) = mrOk then begin
                  RegistrarFolio := RegistrarFolioAux;
                end else begin
                  Query.SQL.Text := 'ROLLBACK;';
                  Query.ExecSQL;
                  Break;
                  Exit;
                end;              
              end;
            end;

            Connection.zCommand.Active := False;
            Connection.zCommand.SQL.Text := 'INSERT IGNORE INTO bitacoradepersonal (sContrato, sNumeroOrden, dIdFecha, iIdDiario, sIdPersonal, iItemOrden, sTipoObra, sDescripcion, sIdPernocta, sIdPlataforma, sHoraInicio, sHoraFinal, dCantidad, ' +
                                            'dSolicitado, sFactor, dCostoMN, dCostoDLL, mMotivos, sAgrupaPersonal, sTipoPernocta, lAplicaPernocta, dCantHH, dAjuste) VALUES ' +
                                            '(:sContrato, :sNumeroOrden, :dIdFecha, :iIdDiario, :sIdPersonal, :iItemOrden, :sTipoObra, :sDescripcion, :sIdPernocta, :sIdPlataforma, ' +
                                            ':sHoraInicio, :sHoraFinal, :dCantidad, :dSolicitado, :sFactor, :dCostoMN, :dCostoDLL, :mMotivos, :sAgrupaPersonal, :sTipoPernocta, :lAplicaPernocta, :dCantHH, :dAjuste)';
            Connection.zCommand.ParamByName('sContrato').AsString := Connection.QryBusca.FieldByName('sContrato').AsString;
            Connection.zCommand.ParamByName('sNumeroOrden').AsString := RegistrarFolio;
            Connection.zCommand.ParamByName('dIdFecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
            Connection.zCommand.ParamByName('iIdDiario').AsString := Connection.QryBusca.FieldByName('iIdDiario').AsString;
            Connection.zCommand.ParamByName('sIdPersonal').AsString := Connection.QryBusca.FieldByName('sIdPersonal').AsString;
            Connection.zCommand.ParamByName('iItemOrden').AsString := Connection.QryBusca.FieldByName('iItemOrden').AsString;
            Connection.zCommand.ParamByName('sTipoObra').AsString := Connection.QryBusca.FieldByName('sTipoObra').AsString;
            Connection.zCommand.ParamByName('sDescripcion').AsString := Connection.QryBusca.FieldByName('sDescripcion').AsString;
            Connection.zCommand.ParamByName('sIdPernocta').AsString := Connection.QryBusca.FieldByName('sIdPernocta').AsString;
            Connection.zCommand.ParamByName('sIdPlataforma').AsString := Connection.QryBusca.FieldByName('sIdPlataforma').AsString;
            Connection.zCommand.ParamByName('sHoraInicio').AsString := Connection.QryBusca.FieldByName('sHoraInicio').AsString;
            Connection.zCommand.ParamByName('sHoraFinal').AsString := Connection.QryBusca.FieldByName('sHoraFinal').AsString;
            Connection.zCommand.ParamByName('dCantidad').AsString := Connection.QryBusca.FieldByName('dCantidad').AsString;
            Connection.zCommand.ParamByName('dSolicitado').AsString := Connection.QryBusca.FieldByName('dSolicitado').AsString;
            Connection.zCommand.ParamByName('sFactor').AsString := Connection.QryBusca.FieldByName('sFactor').AsString;
            Connection.zCommand.ParamByName('dCostoMN').AsString := Connection.QryBusca.FieldByName('dCostoMN').AsString;
            Connection.zCommand.ParamByName('dCostoDLL').AsString := Connection.QryBusca.FieldByName('dCostoDLL').AsString;
            Connection.zCommand.ParamByName('mMotivos').AsString := Connection.QryBusca.FieldByName('mMotivos').AsString;
            Connection.zCommand.ParamByName('sAgrupaPersonal').AsString := Connection.QryBusca.FieldByName('sAgrupaPersonal').AsString;
            Connection.zCommand.ParamByName('sTipoPernocta').AsString := Connection.QryBusca.FieldByName('sTipoPernocta').AsString;
            Connection.zCommand.ParamByName('lAplicaPernocta').AsString := Connection.QryBusca.FieldByName('lAplicaPernocta').AsString;
            Connection.zCommand.ParamByName('dCantHH').AsString := Connection.QryBusca.FieldByName('dCantHH').AsString;
            Connection.zCommand.ParamByName('dAjuste').AsString := Connection.QryBusca.FieldByName('dAjuste').AsString;
            Connection.zCommand.ExecSQL;

          Connection.QryBusca.Next;
        end;
      end else begin
        ShowMessage('No se encontraron registros el día de ayer');
      end;
    end;
    Query.SQL.Text := 'COMMIT;';
    Query.ExecSQL;
  Finally
    Query.Free;
  End;

end;

Procedure TfrmModuloAdmonPersonal.CargarPersonalDB;
Var
  Query: TZQuery;
  i: Integer;
  Lista: TStringList;
  IdPersonal: String;
begin
  Try
    if cmbCategoria.Text <> '' then begin
      Query := TZQuery.Create(Self);
      Query.Connection := Connection.zConnection;

      
      Lista := TStringList.Create;
      Split('|', cmbCategoria.Text, Lista);
      IdPersonal := Trim(Lista[0]);

      Query.SQL.Text := 'SELECT * FROM bitacoradepersonal WHERE dIdFecha = :Fecha AND sTipoObra = :IdPersonal AND sContrato = :Contrato ';
      if ComboBox.ItemIndex > -1 then begin
        Query.SQL.Text := Query.SQL.Text + 'AND sNumeroOrden = ' + QuotedStr(ComboBox.Text);
      end;
      Query.Params.ParamByName('Contrato').AsString := Global_Contrato;
      Query.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
      Query.Params.ParamByName('IdPersonal').AsString := IdPersonal;
      Query.Open;

      for i := 0 to NextGrid.RowCount - 1 do begin
        if Query.Locate('sIdPersonal', NextGrid.Cells[1, i], []) then begin
          NextGrid.Cells[0, i] := 'True';
          NextGrid.Cell[3, i].AsFloat := Query.FieldByName('dCantHH').AsFloat;
        end else begin
          NextGrid.Cells[0, i] := 'False';
          NextGrid.Cell[3, i].AsFloat := 0;
        end;
      end;
    end;
    NextGrid.CalculateFooter();

    if ComboBox.ItemIndex > -1 then begin
      NxNumberEdit1.Enabled := True;
      Query.SQL.Text := 'SELECT * FROM bitacoradepernocta WHERE dIdFecha = :Fecha AND sContrato = :Contrato AND sNumeroOrden = :Folio ';
      Query.Params.ParamByName('Contrato').AsString := Global_Contrato;
      Query.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
      Query.Params.ParamByName('Folio').AsString := ComboBox.Text;
      Query.Open;
      NxNumberEdit1.Value := Query.FieldByName('dCantidad').AsFloat;
    end else begin
      NxNumberEdit1.Value := 0;
      NxNumberEdit1.Enabled := False;
    end;
  Finally
    Lista.Free;
    Query.Free;
  End;

end;

procedure TfrmModuloAdmonPersonal.cmbCategoriaChange(Sender: TObject);
Var
  Lista: TStringList;
  IdPersonal, PrimerValor, PrimerId: String;
  QryPersonal: TZQuery;
begin
  Try
    Lista := TStringList.Create;
    Split('|', cmbCategoria.Text, Lista);
    IdPersonal := Trim(Lista[0]);
    QryPersonal := TZQuery.Create(Self);
    QryPersonal.Connection := Connection.zConnection;
    if IdPersonal = 'PU' then begin
      QryPersonal.Active := False;
      QryPersonal.SQL.Clear;
      QryPersonal.SQL.Add('' +
                    'SELECT ' +
                    ' p.sDescripcion AS sCategoria, ' +
                    ' p.sIdPersonal, ' +
                    ' p.sDescripcion ' +
                    'FROM moe AS m ' +
                    '	INNER JOIN moerecursos AS mr ' +
                    '		ON (mr.iidMoe=m.iidMoe) ' +
                    '	INNER JOIN personal AS p ' +
                    '		ON (p.scontrato=:Contrato AND p.sidpersonal=mr.sidRecurso) ' +
                    'WHERE ' +
                    '	m.didfecha = (SELECT max(didfecha) FROM moe WHERE didfecha <= :Fecha AND sContrato = :ContratoNormal) ' +
                    ' AND m.sContrato = :ContratoNormal ' +
                    '	AND mr.eTipoRecurso = "Personal" ' +
                    ' ORDER BY p.iItemOrden');
      QryPersonal.ParamByName('Contrato').AsString := Global_Contrato_Barco;
      QryPersonal.ParamByName('ContratoNormal').AsString := Global_Contrato;
      QryPersonal.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
    end else begin
      QryPersonal.Active := False;
      QryPersonal.SQL.Clear;
      QryPersonal.SQL.Text := 'SELECT tp.sDescripcion AS sCategoria, p.sIdPersonal, p.sDescripcion FROM personal AS p ' +
                              'INNER JOIN tiposdepersonal AS tp ON (tp.sIdTipoPersonal = p.sIdTipoPersonal) ' +
                              'WHERE p.sContrato = :ContratoBarco AND p.sIdTipoPersonal = :IdPersonal ORDER BY p.sIdTipoPersonal, p.iItemOrden;';
      QryPersonal.Params.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
      QryPersonal.Params.ParamByName('IdPersonal').AsString := IdPersonal;
    end;

    QryPersonal.Open;

    Connection.QryBusca.Active := False;
    Connection.QryBusca.SQL.Text := 'SELECT * FROM cuentas;';
    Connection.QryBusca.Open;

    NxComboBoxColumn1.DataMode := dmValueList;
    NxComboBoxColumn1.Items.Clear;
    PrimerValor := Connection.QryBusca.FieldByName('sDescripcion').AsString;
    PrimerId := Connection.QryBusca.FieldByName('sIdCuenta').AsString;
    while Not Connection.QryBusca.Eof do begin
      NxComboBoxColumn1.Items.Add(Connection.QryBusca.FieldByName('sIdCuenta').AsString + '=' + Connection.QryBusca.FieldByName('sDescripcion').AsString);
      Connection.QryBusca.Next;
    end;

//    ShowMessage(IntToStr(NxComboBoxColumn1.Items.IndexOfName('4.1')) + ' ' + IntToStr(NxComboBox1.ItemIndex));
//    NxComboBoxColumn1.Index := NxComboBoxColumn1.Items.IndexOfName('4.1');

//    NxListColumn1.Items.Add('4.1');
//    NxListColumn1.Items.Add('4.2');
//    NxListColumn1.DisplayMode := dmValueList;

    NextGrid.ClearRows;
    while Not QryPersonal.Eof do begin
      NextGrid.AddRow;
      NextGrid.Cells[0, NextGrid.LastAddedRow] := 'False';
      NextGrid.Cells[1, NextGrid.LastAddedRow] := QryPersonal.FieldByName('sIdPersonal').AsString;
      NextGrid.Cells[2, NextGrid.LastAddedRow] := QryPersonal.FieldByName('sDescripcion').AsString;
      NextGrid.Cells[3, NextGrid.LastAddedRow] := '0';
      NextGrid.Cells[4, NextGrid.LastAddedRow] := PrimerValor;
      NextGrid.Cells[5, NextGrid.LastAddedRow] := PrimerId;
//      TNxComboBox(NextGrid.Columns[4]).ItemIndex := 0;
//      NextGrid.Cells[4, NextGrid.LastAddedRow] := '4.1';
//      TNxComboBox(NextGrid.Columns.Item[4]).ItemIndex:= NxComboBoxColumn1.Items.IndexOfName('4.1');

//      TNxComboBox(NxComboBoxColumn1).ItemIndex:= NxComboBoxColumn1.Items.IndexOfName('4.1');
//      NextGrid.Cells[4, NextGrid.LastAddedRow] := '4.1=PERNOCTA COMPLETA CON SERVICIO DE CAFÉ';
      QryPersonal.Next;
    end;

    dIdFechaChange(dIdFecha);

//    if (IdPersonal = 'PEP') OR (IdPersonal = 'VPEP') then begin
//      ComboBox.Enabled := True;
//      ComboBox.ItemIndex := 0;
//      ComboBoxChange(ComboBox);
//    end else begin
//      ComboBox.Enabled := False;
//    end;
//    
//    CargarPersonalDB;
  Finally
    Lista.Free;
    QryPersonal.Free;
  End;
end;

procedure TfrmModuloAdmonPersonal.ComboBoxChange(Sender: TObject);
Var
  QryPersonal: TZQuery;
begin
  Try
    QryPersonal := TZQuery.Create(Self);
    QryPersonal.Connection := Connection.zConnection;
    QryPersonal.SQL.Text := '' +
                            'SELECT ' +
                            ' SUM(bp.dCantHH) AS dPersonalDia ' +
                            'FROM bitacoradepersonal AS bp ' +
                            '  INNER JOIN personal AS pe ' +
                            '    ON(pe.sContrato = :ContratoBarco AND bp.sIdPersonal = pe.sIdPersonal) ' +
                            'WHERE ' +
                            '  bp.sContrato = :Contrato ' +
                            '  AND bp.sNumeroOrden = :Folio ' +
                            '  AND bp.dIdFecha = :Fecha ' +
                            '  AND bp.sTipoPernocta = "4.1" ' +
                            '  AND (bp.sTipoObra = "PU" OR bp.sTipoObra = "PEP" OR bp.sTipoObra = "VPEP") ' +
                            '  AND lAplicaPernocta = "Si" ';
    QryPersonal.Params.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
    QryPersonal.Params.ParamByName('Contrato').AsString := Global_Contrato;
    QryPersonal.Params.ParamByName('Folio').AsString := ComboBox.Text;
    QryPersonal.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
    QryPersonal.Open;

    LabelPernoctas.Caption := QryPersonal.FieldByName('dPersonalDia').AsString;

    CargarPersonalDB;
  Finally
    QryPersonal.Free;
  End;
end;

procedure TfrmModuloAdmonPersonal.dIdFechaChange(Sender: TObject);
Var
  QueryX: TZQuery;
  Lista: TStringList;
  IdPersonal: String;
begin
  Try
    Lista := TStringList.Create;
    Split('|', cmbCategoria.Text, Lista);
    IdPersonal := Trim(Lista[0]);


    QueryX := TZQuery.Create(Self);
    QueryX.Connection := Connection.zConnection;

    QueryX.Active := False;
    QueryX.SQL.Text :=  '' +
                        'SELECT ' +
                        '	sNumeroOrden ' +
                        'FROM bitacoradeactividades AS ba ' +
                        'WHERE ' +
                        '	ba.sContrato = :Contrato ' +
                        '	AND ba.dIdFecha = :Fecha ' +
//                        ' AND (ba.sIdClasificacion = "TE" OR ba.sIdClasificacion = "SI" OR ba.sIdClasificacion = "FP" OR ba.sIdClasificacion = "NOTA")' +
                        'GROUP BY sNumeroOrden;';
    QueryX.Params.ParamByName('Contrato').AsString := Global_Contrato;
    QueryX.Params.ParamByName('Fecha').AsString := FormatDateTime('yyyy-mm-dd', dIdFecha.Date);
    QueryX.Open;
    QueryX.First;

    PrimerFolio := QueryX.FieldByName('sNumeroOrden').AsString;

    ComboBox.Items.Clear;
    while Not QueryX.Eof do begin
      ComboBox.Items.Add(QueryX.FieldByName('sNumeroOrden').AsString);
      QueryX.Next;
    end;

    if (IdPersonal = 'PEP') OR (IdPersonal = 'VPEP') then begin
      ComboBox.Enabled := True;
      ComboBox.ItemIndex := 0;
      ComboBoxChange(ComboBox);
    end else begin
      ComboBox.Enabled := False;
    end;

    CargarPersonalDB;
  Finally
    QueryX.Free;
  End;
end;

procedure TfrmModuloAdmonPersonal.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := CaFree;
end;

procedure TfrmModuloAdmonPersonal.FormShow(Sender: TObject);
Var
  Query: TZQuery;
begin
  Try
    dIdFecha.DateTime := IncDay(Now, -1);
    Query := TZQuery.Create(Self);
    Query.Connection := Connection.zConnection;

    Query.SQL.Text := 'SELECT * FROM tiposdepersonal;';
    Query.Open;
    cmbCategoria.Items.Clear;
    while Not Query.Eof do begin
//      if Query.FieldByName('sIdTipoPersonal').AsString <> 'PU' then begin
        cmbCategoria.Items.Add(Query.FieldByName('sIdTipoPersonal').AsString + ' | ' + Query.FieldByName('sDescripcion').AsString);
//      end;
      Query.Next;
    end;

    cmbCategoria.ItemIndex := 0;
    cmbCategoriaChange(cmbCategoria);
  Finally
    Query.Free;
  End;
end;

function TfrmModuloAdmonPersonal.IsIn(Valor: string; Lista: TStringList): Boolean;
var
   nIdx: Integer;
begin
   Result := False;
   for nIdx := 0 to Lista.Count - 1 do begin
      if Lista[nIdx] = Valor then begin
         Result := True;
         Break;
      end;
   end;
end;

procedure TfrmModuloAdmonPersonal.NextGridAfterEdit(Sender: TObject; ACol,
  ARow: Integer; Value: WideString);
begin
  if ACol = 3 then begin
    if NextGrid.Cell[ACol, ARow].AsFloat > 0 then begin
      NextGrid.Cell[0, ARow].AsBoolean := True;
    end else begin
      NextGrid.Cell[0, ARow].AsBoolean := False;
    end;
  end;
  NextGrid.CalculateFooter();
end;

procedure TfrmModuloAdmonPersonal.NxComboBoxColumn1Select(Sender: TObject);
begin
  with Sender as TNxComboBoxColumn do begin
    NextGrid.Cells[5, NextGrid.SelectedRow] := NxComboBoxColumn1.Items.Names[TNxComboBox(Editor).ItemIndex];
  end;
end;

end.
