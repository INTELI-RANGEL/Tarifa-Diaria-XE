unit frm_anexosCotemar;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, global, ComCtrls, StdCtrls, ExtCtrls,
  DBCtrls, db, Menus, OleCtrls, Buttons, ZAbstractRODataset, ZDataset,
  ZAbstractDataset, rxToolEdit, rxCurrEdit, RxMemDS, RXDBCtrl, UnitExcepciones,
  frm_seleccionarpaquetecotemar;

type
  TfrmAnexosCotemar = class(TForm)
    Progress: TProgressBar;
    btnOk: TBitBtn;
    btnExit: TBitBtn;
    pgCatalogos: TPageControl;
    pgPersonal: TTabSheet;
    pgEquipo: TTabSheet;
    pgMaterial: TTabSheet;
    pgAnexo: TTabSheet;
    Grid_Personal: TRxDBGrid;
    rxPersonal: TRxMemoryData;
    ds_Personal_2: TDataSource;
    btnAsigna: TBitBtn;
    btnBorra: TBitBtn;
    bntTodo: TBitBtn;
    rxPersonal_2: TRxMemoryData;
    Grid_Personal_2: TRxDBGrid;
    Grid_Equipo: TRxDBGrid;
    btnAsigna2: TBitBtn;
    btnBorra2: TBitBtn;
    btnTodo2: TBitBtn;
    Grid_Equipo_2: TRxDBGrid;
    ds_Personal: TDataSource;
    Q_Insertar: TZQuery;
    ds_Insertar: TDataSource;
    rxEquipo: TRxMemoryData;
    ds_Equipo: TDataSource;
    rxEquipo_2: TRxMemoryData;
    ds_Equipo_2: TDataSource;
    Grid_Material: TRxDBGrid;
    btnAsigna3: TBitBtn;
    btnBorra3: TBitBtn;
    btnTodo3: TBitBtn;
    Grid_Material_2: TRxDBGrid;
    Grid_Anexo: TRxDBGrid;
    btnAsigna4: TBitBtn;
    btnBorra4: TBitBtn;
    btnTodo4: TBitBtn;
    rxMaterial: TRxMemoryData;
    ds_Material: TDataSource;
    rxMaterial_2: TRxMemoryData;
    ds_Material_2: TDataSource;
    rxAnexo: TRxMemoryData;
    ds_Anexo: TDataSource;
    rxAnexo_2: TRxMemoryData;
    ds_Anexo_2: TDataSource;
    Label1: TLabel;
    Grid_Anexo_2: TRxDBGrid;
    ComboAnexos: TComboBox;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    txtBuscarPersonal: TEdit;
    cmdBuscaPersonal: TButton;
    cmdBuscarEquipo: TButton;
    txtBuscarEquipo: TEdit;
    Label5: TLabel;
    Label6: TLabel;
    txtBuscarMateriales: TEdit;
    cmdBuscarMateriales: TButton;
    Label7: TLabel;
    txtBuscarAnexo: TEdit;
    cmdBuscarAnexo: TButton;
    pgHerramienta: TTabSheet;
    pgBasicos: TTabSheet;
    Label8: TLabel;
    txtBuscaHerramienta: TEdit;
    cmdBuscaHerramienta: TButton;
    Grid_Herramienta: TRxDBGrid;
    btnAsigna5: TBitBtn;
    btnBorra5: TBitBtn;
    btnTodo5: TBitBtn;
    Grid_Herramienta2: TRxDBGrid;
    rxHerramienta: TRxMemoryData;
    ds_Herramienta: TDataSource;
    rxHerramienta2: TRxMemoryData;
    ds_Herramienta2: TDataSource;
    Label9: TLabel;
    txtBuscaBasico: TEdit;
    cmdBuscaBasico: TButton;
    Grid_Basico: TRxDBGrid;
    btnAsigna6: TBitBtn;
    btnBorrar6: TBitBtn;
    btnTodos6: TBitBtn;
    Grid_Basico2: TRxDBGrid;
    rxBasicos: TRxMemoryData;
    ds_basicos: TDataSource;
    rxBasicos2: TRxMemoryData;
    ds_basicos2: TDataSource;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnExitClick(Sender: TObject);
    procedure equipoypersonal(Anexo: string);
    procedure FormShow(Sender: TObject);
    procedure btnAsignaClick(Sender: TObject);
    procedure ds_Personal_2DataChange(Sender: TObject; Field: TField);
    procedure btnBorraClick(Sender: TObject);
    procedure bntTodoClick(Sender: TObject);
    procedure btnOkClick(Sender: TObject);
    procedure btnAsigna2Click(Sender: TObject);
    procedure btnBorra2Click(Sender: TObject);
    procedure btnTodo2Click(Sender: TObject);
    procedure pgCatalogosChange(Sender: TObject);
    procedure btnAsigna3Click(Sender: TObject);
    procedure btnBorra3Click(Sender: TObject);
    procedure btnTodo3Click(Sender: TObject);
    procedure btnAsigna4Click(Sender: TObject);
    procedure btnBorra4Click(Sender: TObject);
    procedure btnTodo4Click(Sender: TObject);
    procedure ComboAnexosExit(Sender: TObject);
    procedure cmdBuscaPersonalClick(Sender: TObject);
    procedure cmdBuscarEquipoClick(Sender: TObject);
    procedure cmdBuscarMaterialesClick(Sender: TObject);
    procedure filtrarAnexo();
    procedure cargarMaterial();
    procedure cargarAnexo();
    procedure cargaHerramienta();
    procedure cargarBasicos();
    procedure cmdBuscarAnexoClick(Sender: TObject);
    procedure txtBuscarAnexoKeyPress(Sender: TObject; var Key: Char);
    procedure txtBuscarMaterialesKeyPress(Sender: TObject; var Key: Char);
    procedure txtBuscarEquipoKeyPress(Sender: TObject; var Key: Char);
    procedure txtBuscarPersonalKeyPress(Sender: TObject; var Key: Char);
    procedure Grid_AnexoDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure cmdBuscaHerramientaClick(Sender: TObject);
    procedure txtBuscaHerramientaKeyPress(Sender: TObject; var Key: Char);
    procedure btnAsigna5Click(Sender: TObject);
    procedure btnBorra5Click(Sender: TObject);
    procedure btnTodo5Click(Sender: TObject);
    procedure cmdBuscaBasicoClick(Sender: TObject);
    procedure btnAsigna6Click(Sender: TObject);
    procedure btnBorrar6Click(Sender: TObject);
    procedure btnTodos6Click(Sender: TObject);
    procedure txtBuscaBasicoKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmAnexosCotemar: TfrmAnexosCotemar;
  sGrupo: string;
  sOpcion: string;
  qryAnexo: TZReadOnlyQuery;
  qryAgregar: TzQuery;
  Cta: integer;
  Q_Recurso: TZReadOnlyQuery;
  Q_Fases: TZReadOnlyQuery;
  fase: string;
implementation

{$R *.dfm}

procedure TfrmAnexosCotemar.cargarAnexo();
begin
  Q_Fases.Active := False;
  Q_Fases.SQL.Clear;
  Q_Fases.SQL.Add('select * from anexos');
  Q_Fases.Open;

  if Q_Fases.RecordCount > 0 then
  begin
    comboAnexos.Clear;
    ComboAnexos.Items.Add('Todos');
    while not Q_Fases.Eof do
    begin
      ComboAnexos.Items.Add(Q_Fases.FieldValues['sAnexo']);
      Q_Fases.Next;
    end;
  end;
end;

procedure TfrmAnexosCotemar.cargarMaterial();
begin
        {AHOR EL MATERIAL...}
  rxMaterial.FieldDefs.Clear;
  rxMaterial_2.FieldDefs.Clear;

  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select *, SubStr(mDescripcion, 1,255) as sDescripcion from insumos ' +
    'where sContrato =:Contrato order by sIdInsumo ');
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

        {Material Anexo}
  try
    for Cta := 0 to Q_Recurso.FieldDefs.Count do
      rxMaterial.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
  except

  end;

        {Material Orden}
  try
    for Cta := 0 to Q_Recurso.FieldDefs.Count do
      rxMaterial_2.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
  except

  end;

  rxMaterial_2.Open;
  rxMaterial_2.EmptyTable;

  rxMaterial.Open;
  rxMaterial.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxMaterial.Append;
    for Cta := 1 to rxMaterial.FieldDefs.Count do
      rxMaterial.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxMaterial.Post;
    Q_Recurso.next;
  end;
end;

procedure TfrmAnexosCotemar.filtrarAnexo();
begin
  try
    {Buscamos la fase a la que corresponde..}
    if comboAnexos.Text = 'Todos' then
      fase := ''
    else
    begin
      Q_Fases.Active := False;
      Q_Fases.SQL.Clear;
      Q_Fases.SQL.Add('select sAnexo from anexos where sAnexo =:Anexo ');
      Q_Fases.ParamByName('Anexo').AsString := ComboAnexos.Text;
      Q_Fases.Open;

      if Q_Fases.RecordCount > 0 then
        fase := Q_Fases.FieldValues['sAnexo']
      else
        fase := '*';
    end;

    rxAnexo.FieldDefs.Clear;
    rxAnexo_2.FieldDefs.Clear;

    Q_Recurso.Active := False;
    Q_Recurso.SQL.Clear;
    Q_Recurso.SQL.Add('Select *, substr(mDescripcion, 1,255) as sDescripcion from actividadesxanexo where sContrato =:Contrato and sAnexo =:Anexo ');
    Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
    Q_Recurso.ParamByName('Anexo').AsString := fase;

    if txtBuscarAnexo.Text <> '' then
    begin
      Q_Recurso.SQL.Add(' and ( sNumeroActividad like :Texto  ');
      Q_Recurso.SQL.Add(' or mDescripcion like :Texto  ) ');
      Q_Recurso.ParamByName('Texto').AsString := '%' + txtBuscarAnexo.Text + '%';
    end;

    Q_Recurso.Open;

    {AnexoAnexo}
    try
      for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
        rxAnexo.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
    except

    end;

    {Anexo Orden}
    try
      for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
        rxAnexo_2.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
    except

    end;

    rxAnexo_2.Open;
    rxAnexo_2.EmptyTable;

    rxAnexo.Open;
    rxAnexo.EmptyTable;
    while not Q_Recurso.Eof do
    begin
      rxAnexo.Append;
      for Cta := 1 to rxAnexo.FieldDefs.Count do
        rxAnexo.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
      rxAnexo.Post;
      Q_Recurso.next;
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogos Maestros', 'Al elegir anexo', 0);
    end;
  end;
end;


procedure TfrmAnexosCotemar.btnBorra2Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Equipo_2.DataSource.DataSet.GetBookmark;
  try
    with Grid_Equipo_2.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Equipo_2.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Equipo_2.SelectedRows.Items[iGrid]));
        rxEquipo_2.Edit;
        rxEquipo_2.Delete;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnBorra3Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Material_2.DataSource.DataSet.GetBookmark;
  try
    with Grid_Material_2.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Material_2.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Material_2.SelectedRows.Items[iGrid]));
        rxMaterial_2.Edit;
        rxMaterial_2.Delete;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnBorra4Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Anexo_2.DataSource.DataSet.GetBookmark;
  try
    with Grid_Anexo_2.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Anexo_2.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Anexo_2.SelectedRows.Items[iGrid]));
        rxAnexo_2.Edit;
        rxAnexo_2.Delete;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnBorra5Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Herramienta2.DataSource.DataSet.GetBookmark;
  try
    with Grid_Herramienta2.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Herramienta2.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Herramienta2.SelectedRows.Items[iGrid]));
        rxHerramienta2.Edit;
        rxHerramienta2.Delete;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnBorraClick(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Personal_2.DataSource.DataSet.GetBookmark;
  try
    with Grid_Personal_2.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Personal_2.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Personal_2.SelectedRows.Items[iGrid]));
        rxPersonal_2.Edit;
        rxPersonal_2.Delete;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnBorrar6Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Basico2.DataSource.DataSet.GetBookmark;
  try
    with Grid_Basico2.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Basico2.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Basico2.SelectedRows.Items[iGrid]));
        rxBasicos2.Edit;
        rxBasicos2.Delete;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;

end;

procedure TfrmAnexosCotemar.bntTodoClick(Sender: TObject);
begin
  rxPersonal_2.EmptyTable;
  rxPersonal.First;
  while not rxPersonal.Eof do
  begin
    rxPersonal_2.Append;
    for Cta := 1 to rxPersonal.FieldDefs.Count do
      rxPersonal_2.Fields.FieldByNumber(Cta).AsVariant := rxPersonal.Fields.FieldByNumber(Cta).AsVariant;
    rxPersonal_2.Post;
    rxPersonal.Next;
  end;
end;

procedure TfrmAnexosCotemar.btnAsigna2Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Equipo.DataSource.DataSet.GetBookmark;
  try
    with Grid_Equipo.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Equipo.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Equipo.SelectedRows.Items[iGrid]));
        rxEquipo_2.Append;
        for Cta := 1 to rxEquipo.FieldDefs.Count do
          rxEquipo_2.Fields.FieldByNumber(Cta).AsVariant := rxEquipo.Fields.FieldByNumber(Cta).AsVariant;
        rxEquipo_2.Post;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnAsigna3Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Material.DataSource.DataSet.GetBookmark;
  try
    with Grid_Material.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Material.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Material.SelectedRows.Items[iGrid]));
        rxMaterial_2.Append;
        for Cta := 1 to rxMaterial.FieldDefs.Count do
          rxMaterial_2.Fields.FieldByNumber(Cta).AsVariant := rxMaterial.Fields.FieldByNumber(Cta).AsVariant;
        rxMaterial_2.Post;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnAsigna4Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Anexo.DataSource.DataSet.GetBookmark;
  try
    with Grid_Anexo.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Anexo.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Anexo.SelectedRows.Items[iGrid]));
        rxAnexo_2.Append;
        for Cta := 1 to rxAnexo.FieldDefs.Count do
          rxAnexo_2.Fields.FieldByNumber(Cta).AsVariant := rxAnexo.Fields.FieldByNumber(Cta).AsVariant;
        rxAnexo_2.Post;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnAsigna5Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Herramienta.DataSource.DataSet.GetBookmark;
  try
    with Grid_Herramienta.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Herramienta.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Herramienta.SelectedRows.Items[iGrid]));
        rxHerramienta2.Append;
        for Cta := 1 to rxHerramienta.FieldDefs.Count do
          rxHerramienta2.Fields.FieldByNumber(Cta).AsVariant := rxHerramienta.Fields.FieldByNumber(Cta).AsVariant;
        rxHerramienta2.Post;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;

end;

procedure TfrmAnexosCotemar.btnAsigna6Click(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_basico.DataSource.DataSet.GetBookmark;
  try
    with Grid_basico.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_basico.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_basico.SelectedRows.Items[iGrid]));
        rxBasicos2.Append;
        for Cta := 1 to rxBasicos.FieldDefs.Count do
          rxBasicos2.Fields.FieldByNumber(Cta).AsVariant := rxBasicos.Fields.FieldByNumber(Cta).AsVariant;
        rxBasicos2.Post;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnAsignaClick(Sender: TObject);
var
  SavePlace: TBookmark;
  iGrid: integer;
begin
  SavePlace := Grid_Personal.DataSource.DataSet.GetBookmark;
  try
    with Grid_Personal.DataSource.DataSet do
    begin
      for iGrid := 0 to Grid_Personal.SelectedRows.Count - 1 do
      begin
        GotoBookmark(pointer(Grid_Personal.SelectedRows.Items[iGrid]));
        rxPersonal_2.Append;
        for Cta := 1 to rxPersonal.FieldDefs.Count do
          rxPersonal_2.Fields.FieldByNumber(Cta).AsVariant := rxPersonal.Fields.FieldByNumber(Cta).AsVariant;
        rxPersonal_2.Post;
      end;
    end;
  except
    messageDLG('Haga clic sobre un Registro!', mtInformation, [mbOK], 0);
  end;
end;

procedure TfrmAnexosCotemar.btnExitClick(Sender: TObject);
begin
  close;
end;



procedure TfrmAnexosCotemar.btnOkClick(Sender: TObject);
var
  Q_Valida: TZReaDOnlyQuery;
  campos,
    valores: string;
  total, copiados, duplicados: integer;
  sWbs, sWbsAnterior, sNumeroActividad: string;
  lContinuar: boolean;
  sWbsBandera: string;
  iNivel : Integer;
begin
  try

    Q_Valida := TZReadOnlyQuery.Create(self);
    Q_Valida.Connection := connection.zConnection;

    {Primero el personal seleccionadoo...}
    if pgCatalogos.ActivePageIndex = 0 then
    begin
        total := 0;
        copiados := 0;
        duplicados := 0;
        if rxPersonal_2.RecordCount > 0 then
        begin
          Q_Valida.Active := False;
          Q_Valida.SQL.Clear;
          Q_Valida.SQL.Add('Select * from personal where sContrato =:Contrato ');
          Q_Valida.ParamByName('Contrato').AsString := global_contrato;
          Q_Valida.Open;

             {Validamos si existen registross}
          if Q_Valida.RecordCount > 0 then
          begin
            if MessageDlg('Ya existe Personal en el Catalgo, Desea Continuar ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin
              if MessageDlg('Desea Eliminar el Catalogo de Personal Existente ? ', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                          {Reemplazamos registros existentes}
                Q_Valida.Active := False;
                Q_Valida.SQL.Clear;
                Q_Valida.SQL.Add('Delete from personal where sContrato =:Contrato ');
                Q_Valida.ParamByName('Contrato').AsString := global_contrato;
                Q_Valida.ExecSQL;
              end;
            end
            else
            begin
              Q_Valida.Destroy;
              exit;
            end;
          end;
             {insetamos datos nuevos..}
          Q_Insertar.FieldDefs.Clear;
          for Cta := 1 to rxPersonal_2.FieldDefs.Count - 1 do
            Q_Insertar.FieldDefs.Add(rxPersonal_2.FieldDefs.Items[Cta].Name, rxPersonal_2.FieldDefs.Items[Cta].DataType, rxPersonal_2.FieldDefs.Items[Cta].Size, rxPersonal_2.FieldDefs.Items[Cta].Required);

             {Copiamos las Especialidades de personal }
          Q_Insertar.Active := False;
          Q_Insertar.SQL.Clear;
          Q_Insertar.SQL.Add('select * from personal where sContrato =:Contrato');
          Q_Insertar.ParamByName('Contrato').AsString := global_contrato;
          Q_Insertar.Open;

          rxPersonal_2.First;
          total := rxPersonal_2.RecordCount;
          while not rxPersonal_2.Eof do
          begin
            try
              Q_Insertar.Append;
              for Cta := 1 to rxPersonal_2.FieldDefs.Count - 1 do
              begin
                if Cta = 1 then
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_contrato
                else
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := rxPersonal_2.Fields.FieldByNumber(Cta).AsVariant;
              end;
              Q_Insertar.Post;
              Inc(copiados);
            except
              Inc(duplicados);
            end;
            rxPersonal_2.Next;
          end;
          MessageDlg('Proceso Terminado con Exito!' + #13
            + '  Total = ' + IntToStr(total) + '  Insertados = ' + IntToStr(copiados) +
            ' Duplicados = ' + IntToStr(duplicados), mtConfirmation, [mbOk], 0);
        end;
    end;


    {Ahora el equipo seleccionadoo...}
    if pgCatalogos.ActivePageIndex = 1 then
    begin
        total := 0;
        copiados := 0;
        duplicados := 0;
        if rxEquipo_2.RecordCount > 0 then
        begin
          Q_Valida.Active := False;
          Q_Valida.SQL.Clear;
          Q_Valida.SQL.Add('Select * from equipos where sContrato =:Contrato ');
          Q_Valida.ParamByName('Contrato').AsString := global_contrato;
          Q_Valida.Open;

             {Validamos si existen registross}
          if Q_Valida.RecordCount > 0 then
          begin
            if MessageDlg('Ya existe Equipo en el Catalgo, Desea Continuar ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin
              if MessageDlg('Desea Eliminar el Catalogo de Equipo Existente ? ', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                          {Reemplazamos registros existentes}
                Q_Valida.Active := False;
                Q_Valida.SQL.Clear;
                Q_Valida.SQL.Add('Delete from equipos where sContrato =:Contrato ');
                Q_Valida.ParamByName('Contrato').AsString := global_contrato;
                Q_Valida.ExecSQL;
              end;
            end
            else
            begin
              Q_Valida.Destroy;
              exit;
            end;
          end;
             {insetamos datos nuevos..}
          Q_Insertar.FieldDefs.Clear;
          for Cta := 1 to rxEquipo_2.FieldDefs.Count - 1 do
            Q_Insertar.FieldDefs.Add(rxEquipo_2.FieldDefs.Items[Cta].Name, rxEquipo_2.FieldDefs.Items[Cta].DataType, rxEquipo_2.FieldDefs.Items[Cta].Size, rxEquipo_2.FieldDefs.Items[Cta].Required);

          Q_Insertar.Active := False;
          Q_Insertar.SQL.Clear;
          Q_Insertar.SQL.Add('select * from equipos where sContrato =:Contrato');
          Q_Insertar.ParamByName('Contrato').AsString := global_contrato;
          Q_Insertar.Open;

          rxEquipo_2.First;
          total := rxEquipo_2.RecordCount;
          while not rxEquipo_2.Eof do
          begin
            try
              Q_Insertar.Append;
              for Cta := 1 to rxEquipo_2.FieldDefs.Count - 1 do
              begin
                if Cta = 1 then
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_contrato
                else
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := rxEquipo_2.Fields.FieldByNumber(Cta).AsVariant;
              end;
              Q_Insertar.Post;
              Inc(copiados);
            except
              Inc(duplicados);
            end;
            rxEquipo_2.Next;
          end;
          MessageDlg('Proceso Terminado con Exito!' + #13
            + '  Total = ' + IntToStr(total) + '  Insertados = ' + IntToStr(copiados) +
            ' Duplicados = ' + IntToStr(duplicados), mtConfirmation, [mbOk], 0);
        end;
    end;

    {Ahora la herramienta seleccionada...}
    if pgCatalogos.ActivePageIndex = 4 then
    begin
        total := 0;
        copiados := 0;
        duplicados := 0;
        if rxHerramienta2.RecordCount > 0 then
        begin
          Q_Valida.Active := False;
          Q_Valida.SQL.Clear;
          Q_Valida.SQL.Add('Select * from herramientas where sContrato =:Contrato ');
          Q_Valida.ParamByName('Contrato').AsString := global_contrato;
          Q_Valida.Open;

             {Validamos si existen registross}
          if Q_Valida.RecordCount > 0 then
          begin
            if MessageDlg('Ya existe Herramienta en el Catalgo, Desea Continuar ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin
              if MessageDlg('Desea Eliminar el Catalogo de Herramienta Existente ? ', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                          {Reemplazamos registros existentes}
                Q_Valida.Active := False;
                Q_Valida.SQL.Clear;
                Q_Valida.SQL.Add('Delete from herramientas where sContrato =:Contrato ');
                Q_Valida.ParamByName('Contrato').AsString := global_contrato;
                Q_Valida.ExecSQL;
              end;
            end
            else
            begin
              Q_Valida.Destroy;
              exit;
            end;
          end;
             {insetamos datos nuevos..}
          Q_Insertar.FieldDefs.Clear;
          for Cta := 1 to rxHerramienta2.FieldDefs.Count - 1 do
            Q_Insertar.FieldDefs.Add(rxHerramienta2.FieldDefs.Items[Cta].Name, rxHerramienta2.FieldDefs.Items[Cta].DataType, rxHerramienta2.FieldDefs.Items[Cta].Size, rxHerramienta2.FieldDefs.Items[Cta].Required);

          Q_Insertar.Active := False;
          Q_Insertar.SQL.Clear;
          Q_Insertar.SQL.Add('select * from herramientas where sContrato =:Contrato');
          Q_Insertar.ParamByName('Contrato').AsString := global_contrato;
          Q_Insertar.Open;

          rxHerramienta2.First;
          total := rxHerramienta2.RecordCount;
          while not rxHerramienta2.Eof do
          begin
            try
              Q_Insertar.Append;
              for Cta := 1 to rxHerramienta2.FieldDefs.Count - 1 do
              begin
                if Cta = 1 then
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_contrato
                else
                begin
                    Q_Insertar.Fields.FieldByName('fFecha').AsDateTime := date;
                    Q_Insertar.Fields.FieldByNumber(Cta).AsVariant     := rxHerramienta2.Fields.FieldByNumber(Cta).AsVariant;
                end;
              end;
              Q_Insertar.Post;
              Inc(copiados);
            except
              Inc(duplicados);
            end;
            rxHerramienta2.Next;
          end;
          MessageDlg('Proceso Terminado con Exito!' + #13
            + '  Total = ' + IntToStr(total) + '  Insertados = ' + IntToStr(copiados) +
            ' Duplicados = ' + IntToStr(duplicados), mtConfirmation, [mbOk], 0);
        end;
    end;


    {Ahora el basico seleccionada...}
    if pgCatalogos.ActivePageIndex = 5 then
    begin
        total := 0;
        copiados := 0;
        duplicados := 0;
        if rxBasicos2.RecordCount > 0 then
        begin
          Q_Valida.Active := False;
          Q_Valida.SQL.Clear;
          Q_Valida.SQL.Add('Select * from basicos where sContrato =:Contrato ');
          Q_Valida.ParamByName('Contrato').AsString := global_contrato;
          Q_Valida.Open;

             {Validamos si existen registross}
          if Q_Valida.RecordCount > 0 then
          begin
            if MessageDlg('Ya existe Basico en el Catalgo, Desea Continuar ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin
              if MessageDlg('Desea Eliminar el Catalogo de Basicos Existente ? ', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                          {Reemplazamos registros existentes}
                Q_Valida.Active := False;
                Q_Valida.SQL.Clear;
                Q_Valida.SQL.Add('Delete from basicos where sContrato =:Contrato ');
                Q_Valida.ParamByName('Contrato').AsString := global_contrato;
                Q_Valida.ExecSQL;
              end;
            end
            else
            begin
              Q_Valida.Destroy;
              exit;
            end;
          end;
             {insetamos datos nuevos..}
          Q_Insertar.FieldDefs.Clear;
          for Cta := 1 to rxBasicos2.FieldDefs.Count - 1 do
            Q_Insertar.FieldDefs.Add(rxBasicos2.FieldDefs.Items[Cta].Name, rxBasicos2.FieldDefs.Items[Cta].DataType, rxBasicos2.FieldDefs.Items[Cta].Size, rxBasicos2.FieldDefs.Items[Cta].Required);

          Q_Insertar.Active := False;
          Q_Insertar.SQL.Clear;
          Q_Insertar.SQL.Add('select * from basicos where sContrato =:Contrato');
          Q_Insertar.ParamByName('Contrato').AsString := global_contrato;
          Q_Insertar.Open;

          rxBasicos2.First;
          total := rxBasicos2.RecordCount;
          while not rxBasicos2.Eof do
          begin
            try
              Q_Insertar.Append;
              for Cta := 1 to rxBasicos2.FieldDefs.Count  do
              begin
                if Cta = 1 then
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_contrato
                else
                    Q_Insertar.Fields.FieldByNumber(Cta).AsVariant     := rxBasicos2.Fields.FieldByNumber(Cta).AsVariant;
              end;
              Q_Insertar.Post;
              Inc(copiados);
            except
              Inc(duplicados);
            end;
            rxBasicos2.Next;
          end;
          MessageDlg('Proceso Terminado con Exito!' + #13
            + '  Total = ' + IntToStr(total) + '  Insertados = ' + IntToStr(copiados) +
            ' Duplicados = ' + IntToStr(duplicados), mtConfirmation, [mbOk], 0);
        end;
    end;


    {Ahora el material seleccionadoo...}
    if pgCatalogos.ActivePageIndex = 2 then
    begin
        total := 0;
        copiados := 0;
        duplicados := 0;
        if rxMaterial_2.RecordCount > 0 then
        begin
          Q_Valida.Active := False;
          Q_Valida.SQL.Clear;
          Q_Valida.SQL.Add('Select * from insumos where sContrato =:Contrato ');
          Q_Valida.ParamByName('Contrato').AsString := global_contrato;
          Q_Valida.Open;

             {Validamos si existen registross}
          if Q_Valida.RecordCount > 0 then
          begin
            if MessageDlg('Ya existen Materiales en el Catalgo, Desea Continuar ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin
              if MessageDlg('Desea Eliminar el Catalogo de Materiales Existente ? ', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                          {Reemplazamos registros existentes}
                Q_Valida.Active := False;
                Q_Valida.SQL.Clear;
                Q_Valida.SQL.Add('Delete from insumos where sContrato =:Contrato ');
                Q_Valida.ParamByName('Contrato').AsString := global_contrato;
                Q_Valida.ExecSQL;
              end;
            end
            else
            begin
              Q_Valida.Destroy;
              exit;
            end;
          end;
             {insetamos datos nuevos..}
          Q_Insertar.FieldDefs.Clear;
          for Cta := 1 to rxMaterial_2.FieldDefs.Count - 1 do
            Q_Insertar.FieldDefs.Add(rxMaterial_2.FieldDefs.Items[Cta].Name, rxMaterial_2.FieldDefs.Items[Cta].DataType, rxMaterial_2.FieldDefs.Items[Cta].Size, rxMaterial_2.FieldDefs.Items[Cta].Required);

          Q_Insertar.Active := False;
          Q_Insertar.SQL.Clear;
          Q_Insertar.SQL.Add('select * from insumos where sContrato =:Contrato');
          Q_Insertar.ParamByName('Contrato').AsString := global_contrato;
          Q_Insertar.Open;

          rxMaterial_2.First;
          total := rxMaterial_2.RecordCount;
          while not rxMaterial_2.Eof do
          begin
            try
              Q_Insertar.Append;
              for Cta := 1 to rxMaterial_2.FieldDefs.Count - 1 do
              begin
                if Cta = 1 then
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_contrato
                else
                begin
                  if (rxMaterial_2.Fields.FieldByNumber(Cta).AsVariant = Null) and ((Cta <> 3) and (Cta <> 19)) then
                  begin
                     if rxMaterial_2.Fields.FieldByNumber(Cta).DataType = ftDate then
                        Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := date
                     else
                        if rxMaterial_2.Fields.FieldByNumber(Cta).DataType = ftFloat then
                           Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := 0
                        else
                           Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := '';
                  end
                  else
                    Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := rxMaterial_2.Fields.FieldByNumber(Cta).AsVariant;
                end;
              end;
              Q_Insertar.Post;
              Inc(copiados);
            except
              Inc(duplicados);
            end;
            rxMaterial_2.Next;
          end;

          MessageDlg('Proceso Terminado con Exito!' + #13
            + '  Total = ' + IntToStr(total) + '  Insertados = ' + IntToStr(copiados) +
            ' Duplicados = ' + IntToStr(duplicados), mtConfirmation, [mbOk], 0);
        end;
    end;


    {Ahora el Catalogo de Anexo seleccionado...}
    if pgCatalogos.ActivePageIndex = 3 then
    begin
        total := 0;
        copiados := 0;
        duplicados := 0;
        if rxAnexo_2.RecordCount > 0 then
        begin

          Q_Valida.Active := False;
          Q_Valida.SQL.Clear;
          Q_Valida.SQL.Add('Select * from actividadesxanexo where sContrato =:Contrato and sIdConvenio =:Convenio and sAnexo =:Anexo ');
          Q_Valida.ParamByName('Contrato').AsString := global_contrato;
          Q_Valida.ParamByName('Convenio').AsString := global_convenio;
          Q_Valida.ParamByName('Anexo').AsString := fase;
          Q_Valida.Open;

             {Validamos si existen registross}
          if Q_Valida.RecordCount > 0 then
          begin
            if MessageDlg('Ya existe el Catalgo de Anexo, Desea Continuar ?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
            begin
              if MessageDlg('Desea Eliminar el Catalogo de Anexo Existente ? ', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
              begin
                          {Reemplazamos registros existentes}
                Q_Valida.Active := False;
                Q_Valida.SQL.Clear;
                Q_Valida.SQL.Add('Delete from actividadesxanexo where sContrato =:Contrato and sIdConvenio =:Convenio and sAnexo =:Anexo ');
                Q_Valida.ParamByName('Anexo').AsString := fase;
                Q_Valida.ParamByName('Convenio').AsString := global_convenio;
                Q_Valida.ParamByName('Contrato').AsString := global_contrato;
                Q_Valida.ExecSQL;
              end;
            end
            else
            begin
              Q_Valida.Destroy;
              exit;
            end;
          end;
             {Solicitar a que paquete se va a destinar la actividad, si al Paquete principal o uno personalizado
             , el personalizado va asobre el paquete Principal}
          iNivel := 0;
         // if MessageDlg('Desea Adicionar las partidas al paquete Principal?', mtConfirmation, [mbYes, mbNo], 0) = mrNo then
         if true then
          begin
            Application.CreateForm(TfrmSeleccionarAnexoCotemar, frmSeleccionarAnexoCotemar);
            frmSeleccionarAnexoCotemar.Visible := false;
            frmSeleccionarAnexoCotemar.ShowModal;
            frmSeleccionarAnexoCotemar.Visible := true;

            if frmSeleccionarAnexoCotemar.chkSeleccionar.Checked then
            begin
              sWbs := frmSeleccionarAnexoCotemar.qryPaquetes.FieldValues['sWbs'];
              sNumeroActividad := frmSeleccionarAnexoCotemar.qryPaquetes.FieldValues['sNumeroActividad'];
              sWbsAnterior := frmSeleccionarAnexoCotemar.qryPaquetes.FieldValues['sWbsAnterior'];
              iNivel := frmSeleccionarAnexoCotemar.qryPaquetes.FieldValues['iNivel'] ;
              if iNivel = 0 then iNivel := 1;
            end
            else
            begin
              Q_Insertar.Active := False;
              Q_Insertar.SQL.Clear;
              Q_Insertar.SQL.Add('insert into actividadesxanexo set sContrato=:contrato, sIdConvenio=:Convenio,  ' +
                ' sTipoActividad="Paquete" , ' +
                ' sWbsAnterior = "A", ' +
                ' iNivel=1, ' +
                ' sWbs = concat("A.",:actividad) , ' +
                ' mDescripcion=:Descripcion, ' +
                ' lExtraordinario="No" ,  ' +
                ' sAnexo =:Anexo , ' +
                ' sNumeroActividad=:actividad , ' +
                ' sSimbolo ="-" , ' +
                ' sTipoAnexo="", ' +
                ' sEspecificacion="",' +
                ' sActividadAnterior="",' +
                ' iItemOrden="00000001",' +
                ' dFechaInicio=:dFechaInicio ,' +
                ' dDuracion=:duracion, ' +
                ' dFechaFinal=:dFechaFinal, ' +
                ' dPonderado=0, ' +
                ' dCostoMN=0,' +
                ' dCostoDll=0,' +
                ' dVentaMN=0,' +
                ' dVentaDLL=0,' +
                ' lCalculo="No",' +
                ' sMedida="", ' +
                ' dCantidadAnexo=0, ' +
                ' dCargado=0,' +
                ' dInstalado=0,' +
                ' dExcedente=0,' +
                ' iColor=0,' +
                ' sIdFase="", ' +
                ' sPred="",' +
                ' sSucesor="" ');

              Q_Insertar.ParamByName('Anexo').AsString := fase;
              Q_Insertar.ParamByName('Convenio').AsString := global_convenio;
              Q_Insertar.ParamByName('Contrato').AsString := global_contrato;
              Q_Insertar.ParamByName('Descripcion').AsString := frmSeleccionarAnexoCotemar.mDescripcion.Text;
              Q_Insertar.ParamByName('dFechaInicio').asDate := frmSeleccionarAnexoCotemar.dFechaInicio.Date;
              Q_Insertar.ParamByName('dFechaFinal').asDate := frmSeleccionarAnexoCotemar.dFechaFinal.Date;
              Q_Insertar.ParamByName('actividad').AsString := frmSeleccionarAnexoCotemar.sNumeroActividad.Text;
              Q_Insertar.ParamByName('duracion').AsString := frmSeleccionarAnexoCotemar.Duracion.Text;
              Q_Insertar.ExecSQL;

              sWbs := 'A.' + frmSeleccionarAnexoCotemar.sNumeroActividad.Text;
              sNumeroActividad := frmSeleccionarAnexoCotemar.sNumeroActividad.Text;
              sWbsAnterior := 'A.';
              iNivel := 1;

            end;
          end
          else
            sWbs := '<||>';
             {insetamos datos nuevos..}
          Q_Insertar.FieldDefs.Clear;
          for Cta := 1 to rxAnexo_2.FieldDefs.Count - 1 do
            Q_Insertar.FieldDefs.Add(rxAnexo_2.FieldDefs.Items[Cta].Name, rxAnexo_2.FieldDefs.Items[Cta].DataType, rxAnexo_2.FieldDefs.Items[Cta].Size, rxAnexo_2.FieldDefs.Items[Cta].Required);

          Q_Insertar.Active := False;
          Q_Insertar.SQL.Clear;
          Q_Insertar.SQL.Add('select * from actividadesxanexo where sContrato =:Contrato and sIdConvenio =:Convenio and sAnexo =:Anexo ');
          Q_Insertar.ParamByName('Anexo').AsString := fase;
          Q_Insertar.ParamByName('Convenio').AsString := global_convenio;
          Q_Insertar.ParamByName('Contrato').AsString := global_contrato_barco;
          Q_Insertar.Open;

          rxAnexo_2.First;
          total := rxAnexo_2.RecordCount;
          while not rxAnexo_2.Eof do
          begin
            try
              {Insertar los paquetes superiores de la partida O  actividad}
              lContinuar := true;
              sWbsBandera := rxAnexo_2.Fields.FieldByNumber(6).AsVariant;
              while lContinuar do
              begin
                Connection.QryBusca.Active := false;
                Connection.QryBusca.SQL.Clear;
                Connection.QryBusca.SQL.Add('select * from actividadesxanexo where ' +
                  ' sContrato=:contrato  and ' +
                  ' sWbs=:sWbsAnterior and sAnexo=:anexo and sTipoActividad ="Paquete" order by iNivel desc');
                Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato_barco;
                Connection.QryBusca.Params.ParamByName('sWbsAnterior').Value := sWbsBandera;
                Connection.QryBusca.Params.ParamByName('anexo').Value := fase;
                Connection.QryBusca.Open;

                if Connection.QryBusca.RecordCount > 0 then
                begin
                  Q_Insertar.Append;
                  for Cta := 1 to rxAnexo_2.FieldDefs.Count - 1 do
                  begin
                    if rxAnexo_2.Fields.FieldByNumber(Cta).AsVariant = Null then
                      Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := ''
                    else
                      Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := Connection.QryBusca.Fields.FieldByNumber(Cta).AsVariant;

                    if Cta = 1 then
                      Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_contrato;

                    if Cta = 2 then
                      Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_convenio;

                    if Cta = 3 then   //iNivel
                      Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := Q_Insertar.Fields.FieldByNumber(Cta).AsVariant + iNivel;

                    if sWbs <> '<||>' then
                    begin
                      if Cta = 5 then //sWbs
                      begin
                      if Connection.QryBusca.Fields.FieldByNumber(6).AsVariant = '' then
                          Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := sWbs + '.' +
                          Connection.QryBusca.Fields.FieldByNumber(7).AsVariant
                      else
                          Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := sWbs + '.' +
                          Connection.QryBusca.Fields.FieldByNumber(6).AsVariant + '.'  +
                          Connection.QryBusca.Fields.FieldByNumber(7).AsVariant

                      end;

                      if Cta = 6 then //sWbsAnterior
                      begin
                        Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := sWbs + '.' + Connection.QryBusca.Fields.FieldByNumber(Cta).AsVariant;
                      end;

                      if Cta = 31 then //sAnexo
                      begin
                        Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := fase;
                      end;
                    end;
                  end;
                  try
                    Q_Insertar.Post;
                  except
                    on E: Exception do
                    begin
                      //No hacer nada en caso de error
                    end;
                  end;
                  sWbsBandera := Connection.QryBusca.FieldValues['sWbsAnterior'];
                  if Connection.QryBusca.FieldValues['iNivel'] <= 0 then
                  begin
                    lContinuar := false;
                  end;
                end
                else
                begin
                  lContinuar := false;
                end;

              end;

              {insertar la actividad o paquete}
              Q_Insertar.Append;
              for Cta := 1 to rxAnexo_2.FieldDefs.Count - 1 do
              begin
                if rxAnexo_2.Fields.FieldByNumber(Cta).AsVariant = Null then
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := ''
                else
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := rxAnexo_2.Fields.FieldByNumber(Cta).AsVariant;

                if Cta = 1 then
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_contrato;

                if Cta = 2 then
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_convenio;

                if Cta = 3 then   //iNivel
                  Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := Q_Insertar.Fields.FieldByNumber(Cta).AsVariant + iNivel;

                if sWbs <> '<||>' then
                begin
                  if Cta = 5 then //sWbs
                  begin
                    sWbsBandera := rxAnexo_2.Fields.FieldByNumber(Cta).AsVariant;
                    if rxAnexo_2.Fields.FieldByNumber(6).AsVariant = ''  then
                        Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := sWbs + '.'
                       + '.' + rxAnexo_2.Fields.FieldByNumber(7).AsVariant
                    else
                         Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := sWbs + '.' + rxAnexo_2.Fields.FieldByNumber(6).AsVariant
                        + '.' + rxAnexo_2.Fields.FieldByNumber(7).AsVariant;
                  end;

                  if Cta = 6 then //sWbsAnterior
                  begin
                    Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := sWbs + '.' + rxAnexo_2.Fields.FieldByNumber(Cta).AsVariant;
                  end;

                end;
              end;
              Q_Insertar.Post;
              {Insertar los paquetes y activdades inferiores en caso de tenerlos contenido y de ser paquete}
              //if rxAnexo_2.Fields.FieldByNumber(8).AsString = 'Paquete' then
              //begin
                lContinuar := true;
            
                while lContinuar do
                begin
                  Connection.QryBusca.Active := false;
                  Connection.QryBusca.SQL.Clear;
                  Connection.QryBusca.SQL.Add(' select * from actividadesxanexo where ' +
                    ' sContrato=:contrato  and ' +
                    ' sWbs like :sWbs and sAnexo=:anexo order by iNivel desc');
                  Connection.QryBusca.Params.ParamByName('Contrato').Value := global_contrato_barco;
                  Connection.QryBusca.Params.ParamByName('sWbs').Value := sWbsBandera+'.%';
                  Connection.QryBusca.Params.ParamByName('anexo').Value := fase;
                  Connection.QryBusca.Open;
                  lContinuar := false;
                  while Not Connection.QryBusca.EOF do
                  begin
                    Q_Insertar.Append;
                    for Cta := 1 to rxAnexo_2.FieldDefs.Count - 1 do
                    begin
                      if rxAnexo_2.Fields.FieldByNumber(Cta).AsVariant = Null then
                        Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := ''
                      else
                        Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := Connection.QryBusca.Fields.FieldByNumber(Cta).AsVariant;

                      if Cta = 1 then
                        Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_contrato;

                      if Cta = 2 then
                        Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := global_convenio;

                    if Cta = 3 then   //iNivel
                      Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := Q_Insertar.Fields.FieldByNumber(Cta).AsVariant + iNivel;                    

                      if sWbs <> '<||>' then
                      begin
                        if Cta = 5 then //sWbs
                        begin
                        if Connection.QryBusca.Fields.FieldByNumber(6).AsVariant = '' then
                           Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := sWbs + '.' +
                          Connection.QryBusca.Fields.FieldByNumber(7).AsVariant
                        else
                          Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := sWbs + '.' +
                          Connection.QryBusca.Fields.FieldByNumber(6).AsVariant + '.'  +
                          Connection.QryBusca.Fields.FieldByNumber(7).AsVariant

                        end;

                        if Cta = 6 then //sWbsAnterior
                        begin
                          Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := sWbs + '.' + Connection.QryBusca.Fields.FieldByNumber(Cta).AsVariant;
                        end;

                        if Cta = 31 then //sAnexo
                        begin
                          Q_Insertar.Fields.FieldByNumber(Cta).AsVariant := fase;
                        end;
                      end;
                    end;
                    try
                      Q_Insertar.Post;
                    except
                      on E: Exception do
                      begin
                      //No hacer nada en caso de error
                      end;
                    end;
                    Connection.QryBusca.Next;
                  end;


                end;
              //end;
              Inc(copiados);
            except
              Inc(duplicados);
            end;
            rxAnexo_2.Next;
          end;
          MessageDlg('Proceso Terminado con Exito!' + #13
            + '  Total = ' + IntToStr(total) + '  Insertados = ' + IntToStr(copiados) +
            ' Duplicados = ' + IntToStr(duplicados), mtConfirmation, [mbOk], 0);
        end;
    end;

  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogos Maestros', 'Al ejecutar proceso', 0);
    end;
  end;

  {Borramos obejtoss}
  Q_Valida.Destroy;
end;

procedure TfrmAnexosCotemar.btnTodo2Click(Sender: TObject);
begin
  rxEquipo_2.EmptyTable;
  rxEquipo.First;
  while not rxEquipo.Eof do
  begin
    rxEquipo_2.Append;
    for Cta := 1 to rxEquipo.FieldDefs.Count do
      rxEquipo_2.Fields.FieldByNumber(Cta).AsVariant := rxEquipo.Fields.FieldByNumber(Cta).AsVariant;
    rxEquipo_2.Post;
    rxEquipo.Next;
  end;
end;

procedure TfrmAnexosCotemar.btnTodo3Click(Sender: TObject);
begin
  rxMaterial_2.EmptyTable;
  rxMaterial.First;
  while not rxMaterial.Eof do
  begin
    rxMaterial_2.Append;
    for Cta := 1 to rxMaterial.FieldDefs.Count do
      rxMaterial_2.Fields.FieldByNumber(Cta).AsVariant := rxMaterial.Fields.FieldByNumber(Cta).AsVariant;
    rxMaterial_2.Post;
    rxMaterial.Next;
  end;
end;

procedure TfrmAnexosCotemar.btnTodo4Click(Sender: TObject);
begin
  try
    rxAnexo_2.EmptyTable;
    rxAnexo.First;
    while not rxAnexo.Eof do
    begin
      rxAnexo_2.Append;
      for Cta := 1 to rxAnexo.FieldDefs.Count do
        rxAnexo_2.Fields.FieldByNumber(Cta).AsVariant := rxAnexo.Fields.FieldByNumber(Cta).AsVariant;
      rxAnexo_2.Post;
      rxAnexo.Next;
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogos Maestros', 'Al agregar todos', 0);
    end;
  end;
end;

procedure TfrmAnexosCotemar.btnTodo5Click(Sender: TObject);
begin
      rxHerramienta2.EmptyTable;
      rxHerramienta.First;
      while not rxHerramienta.Eof do
      begin
        rxHerramienta2.Append;
        for Cta := 1 to rxHerramienta.FieldDefs.Count do
          rxHerramienta2.Fields.FieldByNumber(Cta).AsVariant := rxHerramienta.Fields.FieldByNumber(Cta).AsVariant;
        rxHerramienta2.Post;
        rxHerramienta.Next;
      end;
end;

procedure TfrmAnexosCotemar.btnTodos6Click(Sender: TObject);
begin
    rxBasicos2.EmptyTable;
    rxBasicos.First;
    while not rxBasicos.Eof do
    begin
      rxBasicos2.Append;
      for Cta := 1 to rxBasicos.FieldDefs.Count do
        rxBasicos2.Fields.FieldByNumber(Cta).AsVariant := rxBasicos.Fields.FieldByNumber(Cta).AsVariant;
      rxBasicos2.Post;
      rxBasicos.Next;
    end;
end;

procedure TfrmAnexosCotemar.cmdBuscaBasicoClick(Sender: TObject);
begin
  {FILTRO BASICO...}
  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select * from basicos where sContrato =:Contrato ');
  if txtBuscaBasico.Text <> '' then
  begin
    Q_Recurso.SQL.Add(' and ( sIdBasico like :basico  ');
    Q_Recurso.SQL.Add(' or sDescripcion like :basico ) ');
    Q_Recurso.ParamByName('basico').AsString := '%' + txtBuscaBasico.Text + '%';
  end;
  Q_Recurso.SQL.Add(' order by sIdBasico ');
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

  rxBasicos.Open;
  rxBasicos.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxBasicos.Append;
    for Cta := 1 to rxBasicos.FieldDefs.Count do
      rxBasicos.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxBasicos.Post;
    Q_Recurso.next;
  end;
end;

procedure TfrmAnexosCotemar.cmdBuscaHerramientaClick(Sender: TObject);
begin
 {FILTRO HERRAMIENTA...}
  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select * from herramientas where sContrato =:Contrato ');
  if txtBuscaHerramienta.Text <> '' then
  begin
    Q_Recurso.SQL.Add(' and ( sIdHerramientas like :herramienta  ');
    Q_Recurso.SQL.Add(' or sDescripcion like :herramienta ) ');
    Q_Recurso.ParamByName('herramienta').AsString := '%' + txtBuscaHerramienta.Text + '%';
  end;
  Q_Recurso.SQL.Add(' order by sIdHerramientas ');
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

  rxHerramienta.Open;
  rxHerramienta.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxHerramienta.Append;
    for Cta := 1 to rxHerramienta.FieldDefs.Count do
      rxHerramienta.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxHerramienta.Post;
    Q_Recurso.next;
  end;
end;

procedure TfrmAnexosCotemar.cmdBuscaPersonalClick(Sender: TObject);
begin
    {FILTRO EL PERSONAL..}
  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select p.*, t.sDescripcion as Tipo from personal p ' +
    'left join tiposdepersonal t on (t.sIdTipoPersonal = p.sIdTipoPersonal) ' +
    'where p.sContrato =:Contrato ');
  if txtBuscarPersonal.Text <> '' then
  begin
    Q_Recurso.SQL.Add(' and ( p.sIdPersonal like :personal  ');
    Q_Recurso.SQL.Add(' or p.sDescripcion like :personal  ');
    Q_Recurso.SQL.Add(' or t.sDescripcion like :personal  ) ');
    Q_Recurso.ParamByName('personal').AsString := '%' + txtBuscarPersonal.Text + '%';
  end;
  Q_Recurso.SQL.Add(' order by p.iItemOrden ');
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

  rxPersonal.FieldDefs.Clear;
  rxPersonal_2.FieldDefs.Clear;

    {Personal Anexo}
  for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
    rxPersonal.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);

  rxPersonal.Open;
  rxPersonal.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxPersonal.Append;
    for Cta := 1 to rxPersonal.FieldDefs.Count do
      rxPersonal.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxPersonal.Post;
    Q_Recurso.next;
  end;
end;

procedure TfrmAnexosCotemar.cmdBuscarAnexoClick(Sender: TObject);
begin
  filtrarAnexo();
end;

procedure TfrmAnexosCotemar.cmdBuscarEquipoClick(Sender: TObject);
begin
    {FILTRO EL EQUIPO...}
  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select e.*, t.sDescripcion as Tipo from equipos e ' +
    'left join tiposdeequipo t on (t.sIdTipoEquipo = e.sIdTipoEquipo) ' +
    'where e.sContrato =:Contrato ');
  if txtBuscarEquipo.Text <> '' then
  begin
    Q_Recurso.SQL.Add(' and ( e.sIdEquipo like :equipo  ');
    Q_Recurso.SQL.Add(' or e.sDescripcion like :equipo  ');
    Q_Recurso.SQL.Add(' or t.sDescripcion like :equipo  ) ');
    Q_Recurso.ParamByName('equipo').AsString := '%' + txtBuscarEquipo.Text + '%';
  end;
  Q_Recurso.SQL.Add(' order by e.iItemOrden ');
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

  rxEquipo.Open;
  rxEquipo.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxEquipo.Append;
    for Cta := 1 to rxEquipo.FieldDefs.Count do
      rxEquipo.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxEquipo.Post;
    Q_Recurso.next;
  end;
end;

procedure TfrmAnexosCotemar.cmdBuscarMaterialesClick(Sender: TObject);
begin
        {FILTRAR EL MATERIAL...}
  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select *, SubStr(mDescripcion, 1,255) as sDescripcion from insumos ' +
    'where sContrato =:Contrato order by sIdInsumo ');
  if txtBuscarMateriales.Text <> '' then
  begin
    Q_Recurso.SQL.Add(' and ( sIdInsumo like :material  ');
    Q_Recurso.SQL.Add(' or sTipoActividad like :material  ');
    Q_Recurso.SQL.Add(' or mDescripcion like :material  ) ');
    Q_Recurso.ParamByName('material').AsString := '%' + txtBuscarMateriales.Text + '%';
  end;
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

  rxMaterial.Open;
  rxMaterial.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxMaterial.Append;
    for Cta := 1 to rxMaterial.FieldDefs.Count do
      rxMaterial.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxMaterial.Post;
    Q_Recurso.next;
  end;
end;

procedure TfrmAnexosCotemar.ComboAnexosExit(Sender: TObject);
begin
  filtrarAnexo();
end;

procedure TfrmAnexosCotemar.ds_Personal_2DataChange(Sender: TObject;
  Field: TField);
begin

end;

//Aqui van Materiales

procedure TfrmAnexosCotemar.pgCatalogosChange(Sender: TObject);
begin
  try
    if pgCatalogos.ActivePage.Name = 'pgMaterial' then
    begin
      cargarMaterial();
    end;

    if pgCatalogos.ActivePage.Name = 'pgAnexo' then
    begin
      cargarAnexo();
    end;

    if pgCatalogos.ActivePage.Name = 'pgHerramienta' then
    begin
      cargaHerramienta();
    end;

    if pgCatalogos.ActivePage.Name = 'pgBasicos' then
    begin
      cargarBasicos();
    end;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogos Maestros', 'Al cambiar de pestaña', 0);
    end;
  end;
end;

procedure TfrmAnexosCotemar.txtBuscaBasicoKeyPress(Sender: TObject;
  var Key: Char);
begin
    if txtBuscaBasico.Text <> '' then
       cmdBuscaBasico.Click;
end;

procedure TfrmAnexosCotemar.txtBuscaHerramientaKeyPress(Sender: TObject;
  var Key: Char);
begin
   if txtBuscaHerramienta.Text <> '' then
       cmdBuscaHerramienta.Click;
end;

procedure TfrmAnexosCotemar.txtBuscarAnexoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if txtBuscarAnexo.Text <> '' then
    cmdBuscarAnexo.Click;
end;

procedure TfrmAnexosCotemar.txtBuscarEquipoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if txtBuscarEquipo.Text <> '' then
    cmdBuscarEquipo.Click;
end;

procedure TfrmAnexosCotemar.txtBuscarMaterialesKeyPress(Sender: TObject;
  var Key: Char);
begin
  if txtBuscarMateriales.Text <> '' then
    cmdBuscarMateriales.Click;

end;

procedure TfrmAnexosCotemar.txtBuscarPersonalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if txtBuscarPersonal.Text <> '' then
    cmdBuscaPersonal.Click;
end;

procedure tFrmAnexosCotemar.equipoypersonal(Anexo: string);
var
  sContrato: string;
  iRecord: Byte;
begin

end;


procedure TfrmAnexosCotemar.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := cafree;
end;

procedure TfrmAnexosCotemar.FormShow(Sender: TObject);
begin
  fase := '*';
  Q_Fases := TZReadOnlyQuery.Create(self);
  Q_Fases.Connection := connection.zConnection;

  Q_Recurso := TZReadOnlyQuery.Create(self);
  Q_Recurso.Connection := connection.zConnection;

    {PRIMERO EL PERSONAL..}
  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select p.*, t.sDescripcion as Tipo from personal p ' +
    'left join tiposdepersonal t on (t.sIdTipoPersonal = p.sIdTipoPersonal) ' +
    'where p.sContrato =:Contrato order by p.iItemOrden');
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

  rxPersonal.FieldDefs.Clear;
  rxPersonal_2.FieldDefs.Clear;

    {Personal Anexo}
  for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
    rxPersonal.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);

    {Personal Orden}
  for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
    rxPersonal_2.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);

  rxPersonal_2.Open;
  rxPersonal_2.EmptyTable;

  rxPersonal.Open;
  rxPersonal.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxPersonal.Append;
    for Cta := 1 to rxPersonal.FieldDefs.Count do
      rxPersonal.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxPersonal.Post;
    Q_Recurso.next;
  end;

    {AHOR EL EQUIPO...}
  rxEquipo.FieldDefs.Clear;
  rxEquipo_2.FieldDefs.Clear;

  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select e.*, t.sDescripcion as Tipo from equipos e ' +
    'left join tiposdeequipo t on (t.sIdTipoEquipo = e.sIdTipoEquipo) ' +
    'where e.sContrato =:Contrato order by e.iItemOrden');
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

    {Equipo Anexo}
  try
    for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
      rxEquipo.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
  except

  end;

    {Equipo Orden}
  try
    for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
      rxEquipo_2.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
  except

  end;

  rxEquipo_2.Open;
  rxEquipo_2.EmptyTable;

  rxEquipo.Open;
  rxEquipo.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxEquipo.Append;
    for Cta := 1 to rxEquipo.FieldDefs.Count do
      rxEquipo.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxEquipo.Post;
    Q_Recurso.next;
  end;


end;

procedure TfrmAnexosCotemar.Grid_AnexoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if rxAnexo.FieldByName('sTipoActividad').AsString = 'Paquete' then
  begin
    Grid_Anexo.Canvas.Brush.Color := clGray;
    Grid_Anexo.Canvas.Font.Color := clWhite;
  end;

  Grid_Anexo.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TfrmAnexosCotemar.cargaHerramienta;
begin
  {AHOR LA HERRAMIENTA...}
  rxHerramienta.FieldDefs.Clear;
  rxHerramienta2.FieldDefs.Clear;

  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select * from herramientas where sContrato =:Contrato order by sIdHerramientas ');
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

    {Herramienta Anexo}
  try
    for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
      rxHerramienta.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
  except

  end;

    {Herramienta Orden}
  try
    for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
      rxHerramienta2.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
  except

  end;

  rxHerramienta2.Open;
  rxHerramienta2.EmptyTable;

  rxHerramienta.Open;
  rxHerramienta.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxHerramienta.Append;
    for Cta := 1 to rxHerramienta.FieldDefs.Count do
      rxHerramienta.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxHerramienta.Post;
    Q_Recurso.next;
  end;
end;

procedure TfrmAnexosCotemar.cargarBasicos;
begin
  {AHOR LOS BASICOS...}
  rxBasicos.FieldDefs.Clear;
  rxBasicos2.FieldDefs.Clear;

  Q_Recurso.Active := False;
  Q_Recurso.SQL.Clear;
  Q_Recurso.SQL.Add('select * from basicos where sContrato =:Contrato order by sIdBasico ');
  Q_Recurso.ParamByName('Contrato').AsString := global_contrato_barco;
  Q_Recurso.Open;

    {Basicos Anexo}
  try
    for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
      rxBasicos.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
  except

  end;

    {Basicos Orden}
  try
    for Cta := 0 to Q_Recurso.FieldDefs.Count - 1 do
      rxBasicos2.FieldDefs.Add(Q_Recurso.FieldDefs.Items[Cta].Name, Q_Recurso.FieldDefs.Items[Cta].DataType, Q_Recurso.FieldDefs.Items[Cta].Size, Q_Recurso.FieldDefs.Items[Cta].Required);
  except

  end;

  rxBasicos2.Open;
  rxBasicos2.EmptyTable;

  rxBasicos.Open;
  rxBasicos.EmptyTable;
  while not Q_Recurso.Eof do
  begin
    rxBasicos.Append;
    for Cta := 1 to rxBasicos.FieldDefs.Count do
      rxBasicos.Fields.FieldByNumber(Cta).AsVariant := Q_Recurso.Fields.FieldByNumber(Cta).AsVariant;
    rxBasicos.Post;
    Q_Recurso.next;
  end;
end;

end.

