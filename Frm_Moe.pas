unit Frm_Moe;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, ComCtrls, ImgList, Grids, DBGrids,
  JvExDBGrids, JvDBGrid, JvDBUltimGrid, NxPageControl, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, StdCtrls, AdvDateTimePicker,
  AdvDBDateTimePicker, DBCtrls, RXDBCtrl, Newpanel, Menus, Mask;

type
  TFrmMoe = class(TForm)
    pnl2: TPanel;
    pnl3: TPanel;
    pnl1: TPanel;
    pnl4: TPanel;
    pnl5: TPanel;
    pnl6: TPanel;
    JDbUG1: TJvDBUltimGrid;
    ImgBtns: TImageList;
    pgDatos: TNxPageControl;
    NxTshDatos: TNxTabSheet;
    pgPersonal: TNxTabSheet;
    QMoe: TZQuery;
    dsMoe: TDataSource;
    AvDbDtpFecha: TAdvDBDateTimePicker;
    dbmmoComentario: TDBMemo;
    lbl1: TLabel;
    pgEquipo: TNxTabSheet;
    JDbUGPersonal: TJvDBUltimGrid;
    QMoePersonal: TZQuery;
    QMoeEquipo: TZQuery;
    dsMoePersonal: TDataSource;
    dsMoeEquipo: TDataSource;
    JDbUGEquipo: TJvDBUltimGrid;
    Panel: tNewGroupBox;
    ListaObjeto: TRxDBGrid;
    QMoeEquipoiIdMoe: TIntegerField;
    QMoeEquiposIdRecurso: TStringField;
    QMoeEquipoeTipoRecurso: TStringField;
    QMoeEquiposDescripcion: TStringField;
    QMoeEquipodCantidad: TFloatField;
    BuscaObjeto: TZReadOnlyQuery;
    ds_buscaobjeto: TDataSource;
    mnuVigencia: TPopupMenu;
    mnuCarga: TMenuItem;
    Contratos: TZQuery;
    DsContratos: TDataSource;
    QMoePersonaliIdMoe: TIntegerField;
    QMoePersonalsIdRecurso: TStringField;
    QMoePersonaleTipoRecurso: TStringField;
    QMoePersonalsDescripcion: TStringField;
    QMoePersonaldCantidad: TFloatField;
    Label2: TLabel;
    tsLicitacion: TDBEdit;
    procedure FormShow(Sender: TObject);
    procedure QMoeAfterInsert(DataSet: TDataSet);
    procedure QMoePersonalAfterInsert(DataSet: TDataSet);
    procedure QMoeEquipoAfterInsert(DataSet: TDataSet);
    procedure QMoePersonalsIdRecursoChange(Sender: TField);
    procedure QMoeEquiposIdRecursoChange(Sender: TField);
    procedure ListaObjetoDblClick(Sender: TObject);
    procedure ListaObjetoKeyPress(Sender: TObject; var Key: Char);
    procedure ListaObjetoExit(Sender: TObject);
    procedure QMoeAfterScroll(DataSet: TDataSet);
    procedure mnuVigenciaPopup(Sender: TObject);
    procedure NewMenu(var mi : TMenuItem);
    procedure TestClick(Sender: TObject);
    procedure TestClick2(Sender: TObject);
    procedure QMoePersonalAfterPost(DataSet: TDataSet);
    procedure QMoeEquipoAfterPost(DataSet: TDataSet);
    procedure QMoePersonalBeforeDelete(DataSet: TDataSet);
    procedure QMoeEquipoBeforeDelete(DataSet: TDataSet);
    procedure QMoeBeforePost(DataSet: TDataSet);
    procedure AvDbDtpFechaKeyPress(Sender: TObject; var Key: Char);
    procedure tsLicitacionKeyPress(Sender: TObject; var Key: Char);
  private

    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmMoe: TFrmMoe;
  local_contrato : string;

implementation

uses frm_connection, global;

{$R *.dfm}


procedure TFrmMoe.AvDbDtpFechaKeyPress(Sender: TObject; var Key: Char);
begin
   if key=#13 then
      tsLicitacion.SetFocus;
end;

procedure TFrmMoe.FormShow(Sender: TObject);
begin
  QMoe.Active:=False;
  QMoe.ParamByName('Contrato').AsString := Global_Contrato;
  QMoe.Open;
  Self.Caption := 'M O V I M I E N T O S  D E  P E R S O N A L  Y  E Q U I P O   P A R A ('+global_contrato+')';
  //Contratos.Open;
end;

procedure TFrmMoe.ListaObjetoDblClick(Sender: TObject);
begin
  if pgDatos.ActivePageIndex = 1 then
    JDbUGPersonal.SetFocus
  else
    if pgDatos.ActivePageIndex = 2 then
      JDbUGEquipo.SetFocus;
end;

procedure TFrmMoe.ListaObjetoExit(Sender: TObject);
begin
  if Panel.Visible = True then
  begin
    if BuscaObjeto.RecordCount > 0 then
      if pgDatos.ActivePageIndex = 1 then
      begin
        QMoePersonal.FieldValues['sIdRecurso']  := BuscaObjeto.FieldValues['sNumeroActividad'];
        //BitacoradeEquipos.FieldValues['iItemOrden'] := BuscaObjeto.FieldValues['iItemOrden'];
      end
      else
        if pgDatos.ActivePageIndex = 2 then
        begin
          QMoeEquipo.FieldValues['sIdRecurso'] := BuscaObjeto.FieldValues['sNumeroActividad'];
          //GridMaterialesxPartida.SetFocus
        end;

    Panel.Visible := False;
  end
end;

procedure TFrmMoe.ListaObjetoKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    if pgDatos.ActivePageIndex = 1 then
      JDbUGPersonal.SetFocus
    else
      if pgDatos.ActivePageIndex = 2 then
        JDbUGEquipo.SetFocus;
end;

procedure TFrmMoe.mnuVigenciaPopup(Sender: TObject);
var
    zqOrdenes, zqFechas: TZQuery;
    sOT : string;
    Popup: TPopupMenu;

    i, x, n : Integer;
    misOrdenes  : array[0..50] of TMenuItem;
    misFechas   : array[0..50] of TMenuItem;
begin
    zqOrdenes := TZQuery.Create(Self);
    zqOrdenes.Connection := Connection.zConnection;

    zqFechas := TZQuery.Create(Self);
    zqFechas.Connection := Connection.zConnection;

    //Consultamos las ordenes y fechas
    zqOrdenes.Active := False;
    zqOrdenes.SQL.Clear;
    zqOrdenes.SQL.Add('select * from moe group by sContrato, dIdFecha order by sContrato, dIdFecha ');
    zqOrdenes.Open;

    sOT := '';
    mnuVigencia.Items[0].Clear;
    i := 0;
    //Aqui cramos los submenus de las ordenes..
    zqOrdenes.First;
    while not zqOrdenes.Eof do
    begin
        if sOt <> zqOrdenes.FieldValues['sContrato'] then
        begin
            misOrdenes[i] := tMenuItem.Create(mnuCarga);
            misOrdenes[i].Caption := zqOrdenes.FieldValues['sContrato'];
            misOrdenes[i].OnClick := TestClick2;
            mnuVigencia.Items[0].Add(misOrdenes[i]);
            sOT := zqOrdenes.FieldValues['sContrato'];
            inc(i);
        end;
        zqOrdenes.Next;
    end;

    x := 0;
    n := 0;
    //Aqui cramos los submenus de las fechas de vigencia de las ordenes..
    while x < i do
    begin
        zqOrdenes.First;
        while not zqOrdenes.Eof do
        begin
            if mnuVigencia.Items[0].Items[x].Caption = zqOrdenes.FieldValues['sContrato'] then
            begin
                misFechas[n] := tMenuItem.Create(mnuCarga);
                misFechas[n].Caption := DateToStr(zqOrdenes.FieldValues['dIdFecha']);
                mnuVigencia.Items[0].Items[x].Add(misFechas[n]);
                NewMenu(misFechas[n]);
                inc(n);
            end;
            zqOrdenes.Next;
        end;
        inc(x);
    end;

    if pgDatos.ActivePageIndex = 1 then
       mnuCarga.Caption := 'Cargar personal del Oficio :';

    if pgDatos.ActivePageIndex = 2 then
       mnuCarga.Caption := 'Cargar Equipo del Oficio :';
end;

procedure TFrmMoe.NewMenu(var mi: TMenuItem);
begin
    mi.OnClick := TestClick;
end;


procedure TFrmMoe.TestClick(Sender: TObject);
var
  Cursor: TCursor;
  zMoePersonal: TZQuery;
  zMoeEquipo: TZQuery;
begin
 (Sender as tMenuItem).Checked := not (Sender as tMenuItem).Checked;
  try
    Cursor := Screen.Cursor;
    Screen.Cursor := crAppStart;

    zMoePersonal := TZQuery.Create(Self);
    zMoePersonal.Active := False;
    zMoePersonal.Connection := connection.zConnection;
    zMoePersonal.SQL.Clear;
    zMoePersonal.SQL.Text := 'Select MoeB.* From MoeRecursos_aBordo as MoeB where MoeB.iIdMoe = -9'; // -9 para que solo regrese la estructura
    zMoePersonal.Open;

    zMoeEquipo := TZQuery.Create(Self);
    zMoeEquipo.Active := False;
    zMoeEquipo.Connection := connection.zConnection;
    zMoeEquipo.SQL.Clear;
    zMoeEquipo.SQL.Text := 'Select MoeE.* From MoeRecursos_aBordo as MoeE Where MoeE.iIdMoe = -9';
    zMoeEquipo.Open;
    try
      if pgDatos.ActivePage = pgPersonal then
      begin
        //Primero consultamos si ya existe personal cargado a la vigencia..
        connection.QryBusca2.Active := False;
        connection.QryBusca2.SQL.Clear;
        connection.QryBusca2.SQL.Add('select * from moerecursos where iIdMoe =:Id and eTipoRecurso="Personal" ');
        connection.QryBusca2.ParamByName('Id').AsString := QMoe.FieldValues['iIdMoe'];
        connection.QryBusca2.Open;

        if connection.QryBusca2.RecordCount > 0  then
        begin
          if MessageDlg('¿Desea eliminar la Vigencia de Personal del dia '+ DateToStr(QMoe.FieldValues['dIdFecha'])+'?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            connection.QryBusca2.Active := False;
            connection.QryBusca2.SQL.Clear;
            connection.QryBusca2.SQL.Add('delete from moerecursos where iIdMoe =:Id and eTipoRecurso="Personal" ');
            connection.QryBusca2.ParamByName('Id').AsString := QMoe.FieldValues['iIdMoe'];
            connection.QryBusca2.ExecSQL;
            QMoePersonal.Refresh;
            JDbUGPersonal.Refresh;
          end
          else
            exit;
        end;

        Connection.QryBusca2.Active := False;
        Connection.QryBusca2.Connection := Connection.zConnection;
        Connection.QryBusca2.SQL.Clear;
        Connection.QryBusca2.SQL.Add('' +
                      'select ' +
                      'mr.*, mr.dCantidad, ' +
                      'p.sMedida ' +
                      'from moerecursos mr ' +
                      '	inner join moe as m ' +
                      '		on (m.iidMoe = mr.iidMoe) ' +
                      '	INNER JOIN personal AS p ' +
                      '		ON (p.scontrato=:ContratoBarco AND p.sidpersonal=mr.sidRecurso) ' +
                      'where ' +
                      '	m.sContrato = :Contrato ' +
                      '	and didfecha = (select max(didfecha) from moe where didfecha <= :Fecha and sContrato = :Contrato) ' +
                      '	and mr.eTipoRecurso="Personal" ' +
                      ' order by p.iItemOrden');
        Connection.QryBusca2.ParamByName('Contrato').AsString      := StringReplace(local_Contrato, '&','',[]);
        Connection.QryBusca2.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
        Connection.QryBusca2.ParamByName('Fecha').AsDate           := StrToDate(StringReplace((Sender as tMenuItem).Caption, '&','',[]));
        Connection.QryBusca2.Open;

        while Not Connection.QryBusca2.Eof do
        begin
          try
            QMoePersonal.Append;
            QMoePersonal.FieldByName('sIdRecurso').AsString   := Connection.QryBusca2.FieldByName('sIdRecurso').AsString;
            QMoePersonal.FieldByName('sDescripcion').AsString := Connection.QryBusca2.FieldByName('sDescripcion').AsString;
            QMoePersonal.FieldByName('dCantidad').AsFloat     := Connection.QryBusca2.FieldByName('dCantidad').AsFloat;

            zMoePersonal.Append;
            zMoePersonal.FieldByName('iIdMoe').AsInteger:=QMoe.FieldByName('iIdMoe').AsInteger;
            zMoePersonal.FieldByName('eTipoRecurso').AsString:='Personal';
            zMoePersonal.FieldByName('sIdRecurso').AsString   := Connection.QryBusca2.FieldByName('sIdRecurso').AsString;
            zMoePersonal.FieldByName('sDescripcion').AsString := Connection.QryBusca2.FieldByName('sDescripcion').AsString;
            zMoePersonal.FieldByName('dCantidad').AsFloat     := Connection.QryBusca2.FieldByName('dCantidad').AsFloat;
            QMoePersonal.Post;
            zMoePersonal.Post;
          Except
            QMoePersonal.Cancel;
            zMoePersonal.Cancel;
          end;
          Connection.QryBusca2.Next;
        end;
        QMoe.Refresh;
        JDbUG1.Refresh;
        QMoePersonal.Refresh;
        zMoePersonal.Refresh;
        JDbUGPersonal.Refresh;
      end;

      if pgDatos.ActivePage = pgEquipo then
      begin
        //Primero consultamos si ya existe equipo cargado a la vigencia..
        connection.QryBusca2.Active := False;
        connection.QryBusca2.SQL.Clear;
        connection.QryBusca2.SQL.Add('select * from moerecursos where iIdMoe =:Id and eTipoRecurso="Equipo" ');
        connection.QryBusca2.ParamByName('Id').AsString := QMoe.FieldValues['iIdMoe'];
        connection.QryBusca2.Open;

        if connection.QryBusca2.RecordCount > 0  then
        begin
          if MessageDlg('¿Desea eliminar la Vigencia de Equipo del dia '+ DateToStr(QMoe.FieldValues['dIdFecha'])+'?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
          begin
            connection.QryBusca2.Active := False;
            connection.QryBusca2.SQL.Clear;
            connection.QryBusca2.SQL.Add('delete from moerecursos where iIdMoe =:Id and eTipoRecurso="Equipo" ');
            connection.QryBusca2.ParamByName('Id').AsString := QMoe.FieldValues['iIdMoe'];
            connection.QryBusca2.ExecSQL;
            QMoeEquipo.Refresh;
            JDbUGEquipo.Refresh;
          end
          else
             exit;
        end;

        Connection.QryBusca2.Active := False;
        Connection.QryBusca2.Connection := Connection.zConnection;
        Connection.QryBusca2.SQL.Clear;
        Connection.QryBusca2.SQL.Add('' +
                      'SELECT ' +
                      '	mr.*, ' +
                      '	p.sMedida ' +
                      'FROM moerecursos AS mr ' +
                      '	INNER JOIN moe AS m ' +
                      '		ON (m.iidMoe=mr.iidMoe) ' +
                      '	INNER JOIN equipos AS p ' +
                      '		ON (p.scontrato= :ContratoBarco AND p.sIdEquipo=mr.sidRecurso) ' +
                      'WHERE ' +
                      '	m.scontrato= :Contrato ' +
                      '	AND m.didfecha = (SELECT max(didfecha) FROM moe WHERE didfecha <= :Fecha AND sContrato = :Contrato) ' +
                      '	AND mr.eTipoRecurso="Equipo" ' +
                      ' ORDER BY p.iItemOrden');
        Connection.QryBusca2.ParamByName('Contrato').AsString      := StringReplace(local_Contrato, '&','',[]);
        Connection.QryBusca2.ParamByName('ContratoBarco').AsString := Global_Contrato_Barco;
        Connection.QryBusca2.ParamByName('Fecha').AsDate           := StrToDate(StringReplace((Sender as tMenuItem).Caption, '&','',[]));
        Connection.QryBusca2.Open;

        while Not Connection.QryBusca2.Eof do
        begin
          try
            QMoeEquipo.Append;
            QMoeEquipo.FieldByName('sIdRecurso').AsString   := Connection.QryBusca2.FieldByName('sIdRecurso').AsString;
            QMoeEquipo.FieldByName('sDescripcion').AsString := Connection.QryBusca2.FieldByName('sDescripcion').AsString;
            QMoeEquipo.FieldByName('dCantidad').AsFloat     := Connection.QryBusca2.FieldByName('dCantidad').AsFloat;

            zMoeEquipo.Append;
            zMoeEquipo.FieldByName('iIdMoe').AsInteger:=QMoe.FieldByName('iIdMoe').AsInteger;
            zMoeEquipo.FieldByName('eTipoRecurso').AsString:='Equipo';
            zMoeEquipo.FieldByName('sIdRecurso').AsString   := Connection.QryBusca2.FieldByName('sIdRecurso').AsString;
            zMoeEquipo.FieldByName('sDescripcion').AsString := Connection.QryBusca2.FieldByName('sDescripcion').AsString;
            zMoeEquipo.FieldByName('dCantidad').AsFloat     := Connection.QryBusca2.FieldByName('dCantidad').AsFloat;
            QMoeEquipo.Post;
            zMoeEquipo.Post;
          except
            QMoeEquipo.Cancel;
            zMoeEquipo.Cancel;
          end;
          Connection.QryBusca2.Next;
        end;

        QMoe.Refresh;
        JDbUG1.Refresh;
        QMoeEquipo.Refresh;
        zMoeEquipo.Refresh;
        JDbUGEquipo.Refresh;
      end;
    finally
      Screen.Cursor := Cursor;
      if Assigned(zMoeEquipo) then
        zMoeEquipo.Destroy;
        
      if Assigned(zMoePersonal) then
        zMoePersonal.Destroy;
    end;
  except
    on e: Exception do
      MessageDlg('Ha Ocurrido un error inesperado, informar al administrador del sistema del siguiente error :' + e.Message, mtError, [mbOK], 0);
  end;
end;


procedure TFrmMoe.TestClick2(Sender: TObject);
begin
   (Sender as tMenuItem).Checked := not (Sender as tMenuItem).Checked;
    local_contrato := (Sender as tMenuItem).Caption;
end;

procedure TFrmMoe.tsLicitacionKeyPress(Sender: TObject; var Key: Char);
begin
     if key=#13 then
      dbmmoComentario.SetFocus;
end;

procedure TFrmMoe.QMoeAfterInsert(DataSet: TDataSet);
begin
  QMoe.FieldByName('sContrato').AsString:= Global_Contrato;
  QMoe.FieldByName('sfolio').AsString:= '*';
  QMoe.FieldByName('dIdFecha').AsDateTime:= Now;
end;

procedure TFrmMoe.QMoeAfterScroll(DataSet: TDataSet);
begin
  QMoePersonal.Active:=False;
  QMoePersonal.ParamByName('Id').AsInteger:=QMoe.FieldByName('iIdMoe').AsInteger;
  QMoePersonal.Open;

  QMoeEquipo.Active:=False;
  QMoeEquipo.ParamByName('Id').AsInteger:=QMoe.FieldByName('iIdMoe').AsInteger;
  QMoeEquipo.Open;
end;

procedure TFrmMoe.QMoeBeforePost(DataSet: TDataSet);
begin
    QMoe.FieldByName('sOficio').AsString:= tsLicitacion.Text;
end;

procedure TFrmMoe.QMoeEquipoAfterInsert(DataSet: TDataSet);
begin
  QMoeEquipo.FieldByName('iIdMoe').AsInteger:=QMoe.FieldByName('iIdMoe').AsInteger;
  QMoeEquipo.FieldByName('eTipoRecurso').AsString:='Equipo';
end;

procedure TFrmMoe.QMoeEquipoAfterPost(DataSet: TDataSet);
begin
   try
      connection.zCommand.Active := False;
      connection.zCommand.SQL.Text := ' insert into moerecursos_abordo (iIdMoe, sIdRecurso, eTipoRecurso, sDescripcion, dCantidad) '+
                                      ' values (:idmoe, :idrecurso, "Equipo", :descripcion, :cantidad) ';
      connection.zCommand.ParamByName('idmoe').AsInteger := QMoe.FieldByName('iIdMoe').AsInteger;
      connection.zCommand.ParamByName('idrecurso').AsString := QMoeEquipo.FieldByName('sidrecurso').AsString;
      connection.zCommand.ParamByName('descripcion').AsString := QMoeEquipo.FieldByName('sdescripcion').AsString;
      connection.zCommand.ParamByName('cantidad').AsFloat := QMoeEquipo.FieldByName('dCantidad').AsFloat;
      connection.zCommand.ExecSQL;
   Except
   end;
end;

procedure TFrmMoe.QMoeEquipoBeforeDelete(DataSet: TDataSet);
begin
  connection.zCommand.Active := False;
  connection.zCommand.SQL.Text := 'delete from moerecursos_abordo where iIdMoe = :moe and sIdRecurso = :recurso and eTipoRecurso = "Equipo"';
  connection.zCommand.ParamByName('moe').AsInteger := QMoe.FieldByName('iidmoe').AsInteger;
  connection.zCommand.ParamByName('recurso').AsString := QMoeEquipo.FieldByName('sidrecurso').AsString;
  connection.zCommand.ExecSQL;
end;

procedure TFrmMoe.QMoeEquiposIdRecursoChange(Sender: TField);
var
  sDescripcion:string;
begin
  Connection.qryBusca.Active := False;
  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('select iItemOrden, sDescripcion, dCostoDLL, dCostoMN from equipos where sContrato = :Contrato and sIdEquipo = :Equipo');
  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value := Global_Contrato_Barco;
  Connection.qryBusca.Params.ParamByName('Equipo').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Equipo').Value := QMoeEquiposIdRecurso.Text;
  Connection.qryBusca.Open;
  if Connection.qryBusca.RecordCount > 0 then
  begin
      QMoeEquipo.FieldValues['sDescripcion'] := Connection.qryBusca.FieldValues['sDescripcion'];
  end
  else
    if not QMoeEquipo.FieldByName('sIdRecurso').IsNull then
      if Trim(QMoeEquipo.FieldValues['sIdRecurso']) <> '' then
      begin
        sDescripcion := '%' + Trim(QMoeEquipo.FieldValues['sIdRecurso']) + '%';
        BuscaObjeto.Active := False;
        ListaObjeto.Columns.Clear;
        ListaObjeto.Columns.Add;
        ListaObjeto.Columns[0].FieldName := 'sDescripcion';

        BuscaObjeto.SQL.Clear;
        BuscaObjeto.SQL.Add('Select iItemOrden, sIdEquipo as sNumeroActividad, sDescripcion, dCostoDLL, dCostoMN from equipos Where ' +
          'sContrato = :Contrato And sDescripcion Like :Descripcion Order by sDescripcion');
        BuscaObjeto.Params.ParamByName('Contrato').DataType := ftString;
        BuscaObjeto.Params.ParamByName('Contrato').Value := Global_Contrato_Barco;
        BuscaObjeto.Params.ParamByName('Descripcion').DataType := ftString;
        BuscaObjeto.Params.ParamByName('Descripcion').Value := sDescripcion;
        BuscaObjeto.Open;
        Panel.Visible := True;
        Panel.Height  := 358;
        Panel.Width   := 590;
        ListaObjeto.Columns[0].Width := 680;
        ListaObjeto.SetFocus
      end
end;

procedure TFrmMoe.QMoePersonalAfterInsert(DataSet: TDataSet);
begin
    QMoePersonal.FieldByName('iIdMoe').AsInteger:=QMoe.FieldByName('iIdMoe').AsInteger;
    QMoePersonal.FieldByName('eTipoRecurso').AsString:='Personal';
end;

procedure TFrmMoe.QMoePersonalAfterPost(DataSet: TDataSet);
begin
  try
    connection.zCommand.Active := False;
    connection.zCommand.SQL.Text := ' insert into moerecursos_abordo (iIdMoe, sIdRecurso, eTipoRecurso, sDescripcion, dCantidad) '+
                                    ' values (:idmoe, :idrecurso, "Personal", :descripcion, :cantidad) ';
    connection.zCommand.ParamByName('idmoe').AsInteger := QMoe.FieldByName('iIdMoe').AsInteger;
    connection.zCommand.ParamByName('idrecurso').AsString := QMoePersonal.FieldByName('sidrecurso').AsString;
    connection.zCommand.ParamByName('descripcion').AsString := QMoePersonal.FieldByName('sdescripcion').AsString;
    connection.zCommand.ParamByName('cantidad').AsFloat     := QMoePersonal.FieldByName('dCantidad').AsFloat;
    connection.zCommand.ExecSQL;
  Except
  end;
end;

procedure TFrmMoe.QMoePersonalBeforeDelete(DataSet: TDataSet);
begin
  connection.zCommand.Active := False;
  connection.zCommand.SQL.Text := 'delete from moerecursos_abordo where iIdMoe = :moe and sIdRecurso = :recurso and eTipoRecurso = "Personal"';
  connection.zCommand.ParamByName('moe').AsInteger := QMoe.FieldByName('iidmoe').AsInteger;
  connection.zCommand.ParamByName('recurso').AsString := QMoePersonal.FieldByName('sidrecurso').AsString;
  connection.zCommand.ExecSQL;
end;

procedure TFrmMoe.QMoePersonalsIdRecursoChange(Sender: TField);
var
  sDescripcion:string;
begin
  Connection.qryBusca.Active := False;
  Connection.qryBusca.SQL.Clear;
  Connection.qryBusca.SQL.Add('select iItemOrden, sDescripcion, dCostoDLL, dCostoMN, sAgrupaPersonal from personal where sContrato = :Contrato And sIdPersonal = :Personal');
  Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Contrato').Value := Global_Contrato_Barco;
  Connection.qryBusca.Params.ParamByName('Personal').DataType := ftString;
  Connection.qryBusca.Params.ParamByName('Personal').Value := QMoePersonalsIdRecurso.Text;
  Connection.qryBusca.Open;
  if Connection.qryBusca.RecordCount > 0 then
  begin
      QMoePersonal.FieldByName('sDescripcion').AsString := Connection.qryBusca.FieldByName('sDescripcion').AsString;
  end
  else
    if not QMoePersonal.FieldByName('sIdRecurso').IsNull then
      if Trim(QMoePersonal.FieldValues['sIdRecurso']) <> '' then
      begin
        sDescripcion := '%' + Trim(QMoePersonal.FieldValues['sIdRecurso']) + '%';
        BuscaObjeto.Active := False;
        ListaObjeto.Columns.Clear;
        ListaObjeto.Columns.Add;
        ListaObjeto.Columns[0].FieldName := 'sNumeroActividad';
        ListaObjeto.Columns.Add;
        ListaObjeto.Columns[1].FieldName := 'sDescripcion';
        BuscaObjeto.SQL.Clear;
        BuscaObjeto.SQL.Add('Select iItemOrden, sIdPersonal as sNumeroActividad, sDescripcion, dCostoDLL, dCostoMN  from personal Where ' +
          'sContrato = :Contrato And sDescripcion Like :Descripcion Order by sDescripcion');
        BuscaObjeto.Params.ParamByName('Contrato').DataType := ftString;
        BuscaObjeto.Params.ParamByName('Contrato').Value := Global_Contrato_Barco;
        BuscaObjeto.Params.ParamByName('Descripcion').DataType := ftString;
        BuscaObjeto.Params.ParamByName('Descripcion').Value := sDescripcion;
        BuscaObjeto.Open;
        Panel.Visible := True;
        Panel.Height  := 358;
        Panel.Width   := 590;
        ListaObjeto.Columns[0].Width := 50;
        ListaObjeto.Columns[1].Width := 680;
        ListaObjeto.SetFocus;
      end;
end;

end.
