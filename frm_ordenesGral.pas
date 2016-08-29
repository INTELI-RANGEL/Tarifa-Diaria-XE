unit frm_ordenesGral;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, global, db, frm_connection, frm_barra, ComCtrls,
  DBCtrls, StdCtrls, Mask, Menus, ADODB, ZDataset, DateUtils,
  ZAbstractRODataset, ZAbstractDataset, rxToolEdit, UnitTBotonesPermisos,
  UnitExcepciones, udbgrid, UnitValidaTexto, unitactivapop,CommCtrl, DBDateTimePicker,
  ExtCtrls;

type
  TfrmOrdenesGeneral = class(TForm)
    ds_ordenesdetrabajo: TDataSource;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    N3: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    N4: TMenuItem;
    Salir1: TMenuItem;
    OrdenesdeTrabajoGral: TZQuery;
    pnl6: TPanel;
    Panel1: TPanel;
    tvFrentes: TTreeView;
    pnl2: TPanel;
    pnl3: TPanel;
    pnl4: TPanel;
    frmBarra1: TfrmBarra;
    pnl5: TPanel;
    grid_ordenes: TDBGrid;
    pnl1: TPanel;
    Label7: TLabel;
    Label13: TLabel;
    Label16: TLabel;
    Label2: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Label5: TLabel;
    Label3: TLabel;
    Label10: TLabel;
    GroupBox1: TGroupBox;
    chkConsolidada: TDBCheckBox;
    chkDiaria: TDBCheckBox;
    tsIdFolio: TDBEdit;
    tsDescripcionCorta: TDBEdit;
    tsOficioAutorizacion: TDBEdit;
    tmDescripcion: TDBMemo;
    sDias: TDBEdit;
    tdFechaFinal: TDBDateTimePicker;
    tdFechaInicio: TDBDateTimePicker;
    tdFechaAlta: TDBDateTimePicker;
    tmComentarios: TDBMemo;
    QrFrentes: TZReadOnlyQuery;
    lbl1: TLabel;
    dbedtAutoriza: TDBEdit;
    procedure tdFechaInicioKeyPress(Sender: TObject; var Key: Char);
    procedure tsApoyoKeyPress(Sender: TObject; var Key: Char);
    procedure tcIdStatusKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure grid_ordenesEnter(Sender: TObject);
    procedure grid_ordenesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grid_ordenesKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grid_ordenesCellClick(Column: TColumn);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tmDescripcionEnter(Sender: TObject);
    procedure tmDescripcionExit(Sender: TObject);
    procedure tdFechaInicioEnter(Sender: TObject);
    procedure tdFechaInicioExit(Sender: TObject);
    procedure tdFechaFinalEnter(Sender: TObject);
    procedure tdFechaFinalExit(Sender: TObject);
    procedure tmComentariosEnter(Sender: TObject);
    procedure tmComentariosExit(Sender: TObject);
    procedure tsDescripcionCortaEnter(Sender: TObject);
    procedure tsDescripcionCortaExit(Sender: TObject);
    procedure tsDescripcionCortaKeyPress(Sender: TObject; var Key: Char);
    procedure tsIdFolioEnter(Sender: TObject);
    procedure tsIdFolioExit(Sender: TObject);
    procedure tsIdFolioKeyPress(Sender: TObject; var Key: Char);
    procedure tsOficioAutorizacionKeyPress(Sender: TObject; var Key: Char);
    procedure tsOficioAutorizacionEnter(Sender: TObject);
    procedure tsOficioAutorizacionExit(Sender: TObject);
    procedure tdFechaAltaExit(Sender: TObject);
    procedure tdFechaAltaEnter(Sender: TObject);
    procedure tdFechaAltaKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure tdFechaFinalKeyPress(Sender: TObject; var Key: Char);
    procedure sDiasEnter(Sender: TObject);
    procedure sDiasExit(Sender: TObject);
    procedure sDiasKeyPress(Sender: TObject; var Key: Char);
    procedure grid_ordenesMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_ordenesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_ordenesTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure tdFechaFinalChange(Sender: TObject);
    procedure tdFechaInicioChange(Sender: TObject);
    procedure chkConsolidadaClick(Sender: TObject);
    procedure ChkDiariaClick(Sender: TObject);
    procedure ActualizaFechas();
    function SeCruzaFecha(dFecha: TDateTime; Comparar: string): Boolean;
    procedure CargarArbol;
    procedure FormCreate(Sender: TObject);
    procedure tvFrentesChange(Sender: TObject; Node: TTreeNode);
    procedure tvFrentesCollapsing(Sender: TObject; Node: TTreeNode;
      var AllowCollapse: Boolean);
    procedure tvFrentesCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode;
      State: TCustomDrawState; var DefaultDraw: Boolean);

  private
    { Private declarations }
  public
    { Public declarations }
    Param_Frente:string;
  end;

var
  frmOrdenesGeneral: TfrmOrdenesGeneral;
  Opcion: string;
  sTipo, sPlataforma, sPernocta: string;
  dFechaOld, dFechaiOld, dFechafOld: tDate;
  sOrdenOld,
    sOficioAnterior: string;
  BotonPermiso: TBotonesPermisos;
  utgrid: ticdbgrid;
  ListaContratos:TStringList;
  nodoseleccionado:Integer;
implementation

uses UFunctionsGHH;

{$R *.dfm}


procedure TfrmOrdenesGeneral.CargarArbol;
var
    nodo, hijo,
    nodoprim    : TTreeNode;
    apNodo      :tbookmark;
    indice      :integer;
    padre       :integer;
    hItem       : HTREEITEM;
    vContrato:TContrato;
begin
  QrFrentes.Filtered:=false;
  QrFrentes.DisableControls;

  QrFrentes.First;
  ListaContratos.Clear;
  TvFrentes.Items.Clear;

  while not QrFrentes.Eof do
  begin
    apNodo:=QrFrentes.GetBookmark;
    with TvFrentes.Items,QrFrentes do
    begin
      if FieldByName('IndiceP').AsInteger=0 then
      begin
        nodo := add(nil,FieldByName('sNumeroOrden').AsString);
        nodo.ImageIndex:=0;
        nodo.SelectedIndex:=0;
        nodoprim:=nodo;

        vContrato:=TContrato.Create;
        vContrato.sContrato:=FieldByName('sContrato').AsString;
        vContrato.sNumeroOrden:=FieldByName('sNumeroOrden').AsString;
        vContrato.Id:=IntToStr(Integer(nodoprim.ItemId));
        ListaContratos.AddObject(vContrato.Id,vContrato);

        if(FieldByName('sContrato').AsString=param_global_contrato)
          and(FieldByName('sNumeroOrden').AsString=Param_Frente) then
          hijo:=nodo;

        nodo:=addchild(nodoprim,FieldByName('Frente').AsString);
        nodo.ImageIndex:=nodo.Level;
        nodo.SelectedIndex:=nodo.Level;
        vContrato:=TContrato.Create;
        vContrato.sContrato:=FieldByName('sContrato').AsString;
        vContrato.sNumeroOrden:=FieldByName('Frente').AsString;
        vContrato.Id:=IntToStr(Integer(nodo.ItemId));
        ListaContratos.AddObject(vContrato.Id,vContrato);

        if(FieldByName('sContrato').AsString=param_global_contrato)
          and(FieldByName('Frente').AsString=Param_Frente) then
          hijo:=nodo;
      end
      else
      begin
        if nodoprim<>nil then
        begin
          nodo:=addchild(nodoprim,FieldByName('sNumeroOrden').AsString);
          nodo.ImageIndex:=nodo.Level;
          nodo.SelectedIndex:=nodo.Level;
          vContrato:=TContrato.Create;
          vContrato.sContrato:=FieldByName('sContrato').AsString;
          vContrato.sNumeroOrden:=FieldByName('sNumeroOrden').AsString;
          vContrato.Id:=IntToStr(Integer(nodo.ItemId));
          ListaContratos.AddObject(vContrato.Id,vContrato);
          if(FieldByName('sContrato').AsString=param_global_contrato)
          and(FieldByName('sNumeroOrden').AsString=Param_Frente) then
            hijo:=nodo;
        end;
      end;
    end;

    QrFrentes.Next;
  end;

  QrFrentes.EnableControls;
  TvFrentes.FullExpand;
  if hijo<>nil then
    TvFrentes.Select(hijo);

End;



procedure TfrmOrdenesGeneral.tsOficioAutorizacionEnter(Sender: TObject);
begin
  tsOficioAutorizacion.Color := global_color_entrada
end;

procedure TfrmOrdenesGeneral.tsOficioAutorizacionExit(Sender: TObject);
begin
  tsOficioAutorizacion.Color := global_color_salida
end;

procedure TfrmOrdenesGeneral.tsOficioAutorizacionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tmDescripcion.SetFocus
end;


procedure TfrmOrdenesGeneral.tvFrentesChange(Sender: TObject; Node: TTreeNode);
begin
  nodoseleccionado:=Integer(node.ItemId);

  frmBarra1.btnCancel.Click;

  param_global_contrato:=TContrato(ListaContratos.Objects[ListaContratos.IndexOf(IntToStr(Integer(Node.ItemId)))]).sContrato;
  Param_Frente:=TContrato(ListaContratos.Objects[ListaContratos.IndexOf(IntToStr(Integer(Node.ItemId)))]).sNumeroOrden;

  OrdenesdeTrabajoGral.Active := False;
 // OrdenesdeTrabajoGral.SQL.Clear;
  //11OrdenesdeTrabajoGral.SQL.Add('select * from ordenesdetrabajogral where sContrato = :Contrato order by dFechaVigencia');
  OrdenesdeTrabajoGral.Params.ParamByName('Contrato').DataType := ftString;
  OrdenesdeTrabajoGral.Params.ParamByName('Contrato').Value := param_global_contrato;
  OrdenesdeTrabajoGral.ParamByName('Orden').AsString:=Param_Frente;
  OrdenesdeTrabajoGral.Open;




end;

procedure TfrmOrdenesGeneral.tvFrentesCollapsing(Sender: TObject;
  Node: TTreeNode; var AllowCollapse: Boolean);
begin
  AllowCollapse:=False;
end;

procedure TfrmOrdenesGeneral.tvFrentesCustomDrawItem(Sender: TCustomTreeView;
  Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
  Sender.Canvas.Font.Color := clblack;
  Sender.Canvas.Font.Style := [];
  //Sender.Canvas.Font.Size:=tam;
  if (cdsFocused in State) then
  begin
    Sender.Canvas.Font.Color := clwhite;
    sender.Canvas.Font.Style:=[fsbold];
    //sender.Canvas.Brush.Color:=clwindow;
  end
  else
    if Integer(node.ItemId)=nodoseleccionado then
    begin
      Sender.Canvas.Font.Color := clwhite;
      sender.Canvas.Font.Style:=[fsbold];
    end;
end;

procedure TfrmOrdenesGeneral.tdFechaInicioKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tdFechaFinal.SetFocus
end;

procedure TfrmOrdenesGeneral.tsApoyoKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdFechaInicio.SetFocus
end;

procedure TfrmOrdenesGeneral.tcIdStatusKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tmcomentarios.SetFocus
end;


procedure TfrmOrdenesGeneral.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  utgrid.Destroy;
  ordenesdetrabajogral.Cancel;
  action := cafree;
  BotonPermiso.Free;
end;

procedure TfrmOrdenesGeneral.FormCreate(Sender: TObject);
begin
  ListaContratos:=TStringList.Create;
end;

procedure TfrmOrdenesGeneral.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'oficmodif', PopupPrincipal);
  UtGrid := TicdbGrid.create(grid_ordenes);
  OpcButton := '';
  
  Param_Frente := param_global_contrato;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
  QrFrentes.Active:=False;
  QrFrentes.ParamByName('contrato').AsString:=global_contrato_Barco;
  QrFrentes.Open;
  CargarArbol;
end;

procedure TfrmOrdenesGeneral.grid_ordenesEnter(Sender: TObject);
begin
  if frmbarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;

  if OrdenesdeTrabajoGral.RecordCount > 0 then
  begin
    tdFechaInicio.Date := OrdenesdeTrabajoGral.FieldValues['dFIProgramado'];
    tdFechaFinal.Date := OrdenesdeTrabajoGral.FieldValues['dFFProgramado'];
  end
end;

procedure TfrmOrdenesGeneral.grid_ordenesKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if frmbarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;

  if OrdenesdeTrabajoGral.RecordCount > 0 then
  begin
    tdFechaInicio.Date := OrdenesdeTrabajoGral.FieldValues['dFIProgramado'];
    tdFechaFinal.Date := OrdenesdeTrabajoGral.FieldValues['dFFProgramado'];
  end
end;

procedure TfrmOrdenesGeneral.grid_ordenesKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if frmbarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;

  if OrdenesdeTrabajoGral.RecordCount > 0 then
  begin
    tdFechaInicio.Date := OrdenesdeTrabajoGral.FieldValues['dFIProgramado'];
    tdFechaFinal.Date := OrdenesdeTrabajoGral.FieldValues['dFFProgramado'];
  end
end;

procedure TfrmOrdenesGeneral.grid_ordenesMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x, y);
end;

procedure TfrmOrdenesGeneral.grid_ordenesMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender, Button, Shift, X, Y);
end;

procedure TfrmOrdenesGeneral.grid_ordenesTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmOrdenesGeneral.grid_ordenesCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
    frmBarra1.btnCancel.Click;

  if OrdenesdeTrabajoGral.RecordCount > 0 then
  begin
    tdFechaInicio.Date := OrdenesdeTrabajoGral.FieldValues['dFIProgramado'];
    tdFechaFinal.Date := OrdenesdeTrabajoGral.FieldValues['dFFProgramado'];
  end
end;


procedure TfrmOrdenesGeneral.frmBarra1btnAddClick(Sender: TObject);
var
  qry: TZReadOnlyQuery;
  FechaFinal: TDateTime;
begin
  frmBarra1.btnAddClick(Sender);
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;
  tsIdFolio.setFocus;
  OrdenesdeTrabajoGral.Append;
  tdFechaAlta.Date := Date;
  tdFechaInicio.Date := Date;
  tdFechaFinal.Date := Date;
  chkConsolidada.Checked := True;
  activapop(frmOrdenesGeneral, popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;

    {identificar el ultimo registro guardado
    Tomar la fecha de vigencia programada final}
  qry := TZReadOnlyQuery.Create(Self);
  qry.Connection := Connection.zConnection;
  qry.Active := false;
  qry.SQL.Clear;
  qry.SQL.Add('select max(dFfProgramado) as dFfProgramado from ordenesdetrabajogral ' +
    ' where sContrato=:Contrato ');
  qry.Params.ParamByName('Contrato').DataType := ftString;
  qry.Params.ParamByName('Contrato').Value := param_global_contrato;
  qry.Open;

  FechaFinal := now;
  if qry.RecordCount > 0 then
  begin
    if qry.FieldValues['dFfProgramado'] <> NULL then
      FechaFinal := qry.FieldValues['dFfProgramado'];
  end;

  tdFechaAlta.DateTime := FechaFinal;
  tdFechaInicio.DateTime := IncDay(FechaFinal);
  tdFechaFinal.DateTime := IncDay(FechaFinal);


end;

procedure TfrmOrdenesGeneral.frmBarra1btnEditClick(Sender: TObject);
begin

  MessageDLG('Si cambia las fechas de vigencias, el sistema ajustará de forma automatica ' +
    'el personal, equipo y pernocta registradas en este folio, eliminando o agregando la programacion diaria en ceros.' + Chr(13) + Chr(10) +
    'Si hay reportes diarios dentro de las folios de fechas de vigencias, estos no podran ser actualizados. ', mtInformation, [mbOk], 0);

  frmBarra1.btnEditClick(Sender);
  Insertar1.Enabled := False;
  Editar1.Enabled := False;
  Registrar1.Enabled := True;
  Can1.Enabled := True;
  Eliminar1.Enabled := False;
  Refresh1.Enabled := False;
  Salir1.Enabled := False;
  try
    dFechaOld := OrdenesdeTrabajoGral.FieldValues['dFechaVigencia'];
    dFechaiOld := OrdenesdeTrabajoGral.FieldValues['dFiProgramado'];
    dFechafOld := OrdenesdeTrabajoGral.FieldValues['dFfProgramado'];
    sOrdenOld := OrdenesdeTrabajoGral.FieldValues['sNumeroOrden'];
    sOficioAnterior := OrdenesdeTrabajoGral.FieldValues['sOficioAutorizacion'];
    OrdenesdeTrabajoGral.Edit;
  except
    on e: exception do
    begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Oficio Modificaciones Personal ', 'Al editar registro', 0);
      frmBarra1.btnCancel.Click;
    end;
  end;
  tsIdFolio.SetFocus;
  activapop(frmOrdenesGeneral, popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmOrdenesGeneral.frmBarra1btnPostClick(Sender: TObject);
var
  QryActualiza: TZReadOnlyQuery;
  lEdita: boolean;
  nombres, cadenas: TStringList;
  qry: TZReadOnlyQuery;
begin
  if OrdenesdeTrabajoGral.State in [dsInsert] then
  begin
     dFechaiOld := tdFechaInicio.Date;
     dFechafOld := tdFechaFinal.Date;
  end;

  if OrdenesdeTrabajoGral.State in [dsInsert, dsEdit] then
  begin
    qry := TZReadOnlyQuery.Create(Self);
    qry.Connection := Connection.zConnection;
    qry.Active := false;
    qry.SQL.Clear;
    qry.SQL.Add(' select * from bitacoradepersonal where sContrato=:contrato and dIdFecha between :fechai and :fechaf ');
    qry.Params.ParamByName('contrato').AsString := param_global_contrato;
    qry.Params.ParamByName('fechai').AsDate := dFechaiOld;
    qry.Params.ParamByName('fechaf').AsDate := dFechafOld;
    qry.Open;

    if qry.RecordCount > 0 then
    begin
      MessageDLG('Existe personal reportado en reportes diarios, no se permite la edicion. ', mtError, [mbOk], 0);
      exit;
    end;

    qry.Active := false;
    qry.SQL.Clear;
    qry.SQL.Add(' select * from bitacoradeequipos where sContrato=:contrato and dIdFecha between :fechai and :fechaf ');
    qry.Params.ParamByName('contrato').AsString := param_global_contrato;
    qry.Params.ParamByName('fechai').AsDate := dFechaiOld;
    qry.Params.ParamByName('fechaf').AsDate := dFechafOld;
    qry.Open;

    if qry.RecordCount > 0 then
    begin
      MessageDLG('Existe equipo reportado en reportes diarios, no se permite la edicion. ', mtError, [mbOk], 0);
      exit;
    end;
  end;


  OrdenesdetrabajoGral.FieldValues['sDias'] := DaysBetween(tdfechaInicio.date, tdFechaFinal.date) + 1;

    {Validacion de campos}
  nombres := TStringList.Create; cadenas := TStringList.Create;
  nombres.Add('Folio'); nombres.Add('Titulo Corto Obra');
  nombres.Add('Oficio de Autorizacion'); nombres.Add('Descripcion');
  nombres.Add('Dias');
  cadenas.Add(tsIdFolio.Text); cadenas.Add(tsDescripcionCorta.Text);
  cadenas.Add(tsOficioAutorizacion.Text); cadenas.Add(tmDescripcion.Text);
  cadenas.Add(sDias.Text);

  if not validaTexto(nombres, cadenas, '', '') then
  begin
    MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
    exit;
  end;
    {Continua insercion de datos}

  if tdFechaFinal.Date < tdFechaInicio.Date then
  begin
    showmessage('la fecha final es menor a la fecha inicial');
    tdFechaFinal.SetFocus;
    exit;
  end;

  if OrdenesdeTrabajoGral.State = dsEdit then
    lEdita := True
  else
    lEdita := False;

  QryActualiza := TZReadOnlyQuery.Create(self);
  QryActualiza.Connection := connection.zConnection;

  if lEdita = False then
  begin
    QryActualiza.Active := False;
    QryActualiza.SQL.Clear;
    QryActualiza.SQL.Add('Select sOficioAutorizacion from ordenesdetrabajogral where sContrato =:Contrato and sNumeroOrden =:Orden and dFechaVigencia =:Fecha and sOficioAutorizacion =:Oficio ');
    QryActualiza.ParamByName('Contrato').AsString := param_global_contrato;
    QryActualiza.ParamByName('Orden').AsString := Param_Frente;
    QryActualiza.ParamByName('Oficio').AsString := tsOficioAutorizacion.Text;
    QryActualiza.ParamByName('Fecha').AsDate := tdFechaAlta.Date;
    QryActualiza.Open;

    if QryActualiza.RecordCount > 0 then
    begin
      MessageDLG('El Oficio ya se encuentra dado de Alta!', mtInformation, [mbOk], 0);
      exit;
    end;
  end;

  try
    OrdenesdeTrabajoGral.FieldValues['sContrato'] := param_global_contrato;
    OrdenesdeTrabajoGral.FieldValues['sNumeroOrden'] := Param_Frente;
    OrdenesdeTrabajoGral.FieldValues['lGeneraAnexo'] := 'No';
    OrdenesdeTrabajoGral.FieldValues['lGeneraConsumibles'] := 'No';
    OrdenesdeTrabajoGral.FieldValues['lGeneraPersonal'] := 'No';
    OrdenesdeTrabajoGral.FieldValues['lGeneraEquipo'] := 'No';
    OrdenesdeTrabajoGral.FieldValues['dFechaVigencia'] := tdFechaAlta.date;
    OrdenesdeTrabajoGral.FieldValues['dFiProgramado'] := tdFechaInicio.date;
    OrdenesdeTrabajoGral.FieldValues['dFfProgramado'] := tdFechaFinal.date;
    OrdenesdeTrabajoGral.FieldValues['iFolioOficio'] :=0;
    OrdenesdeTrabajoGral.FieldValues['sDias'] := sDias.text;
    if tsIdFolio.text = '' then
      tsIdFolio.text := '*';
    if tmComentarios.Text = '' then
      OrdenesdeTrabajoGral.FieldValues['mComentarios'] := '*';
    if chkDiaria.Checked then
      OrdenesdeTrabajoGral.FieldValues['sTipoVigencia'] := 'DIARIA';
    if chkConsolidada.Checked then
      OrdenesdeTrabajoGral.FieldValues['sTipoVigencia'] := 'CONSOLIDADA';
    OrdenesdeTrabajoGral.Post;

    if lEdita then
    begin
       //Actulizamos la fecha de la vigencia..
      QryActualiza.Active := False;
      QryActualiza.SQL.Clear;
      QryActualiza.SQL.Add('Update movtorecursosxoficio set dFechaVigencia =:Fecha ' +
        'where iFolioOficio =:Folio and sContrato=:contrato');
      QryActualiza.ParamByName('Folio').AsString := OrdenesdetrabajoGral.FieldValues['iFolioOficio'];
      QryActualiza.ParamByName('contrato').AsString := param_global_contrato;
      QryActualiza.ParamByName('Fecha').AsDateTime := tdFechaAlta.date;

      QryActualiza.ExecSQL;

      ActualizaFechas;

      //Actulizamos los recurso de la vigencia..
      QryActualiza.Active := False;
      QryActualiza.SQL.Clear;
      QryActualiza.SQL.Add('Update movtorecursosxoficio set dCantidad =( ' +
        ' select if(sum(dCantidad) is null,0,sum(dCantidad)) as Cantidad from detallerecursosxoficio where sContrato =:contrato and iFolioOficio = :Folio ' +
        ' and detallerecursosxoficio.sNumeroActividad=movtorecursosxoficio.sNumeroActividad' +
        ' ) ' +
        'where iFolioOficio =:Folio and sContrato=:contrato ');
      QryActualiza.ParamByName('Folio').AsString := OrdenesdetrabajoGral.FieldValues['iFolioOficio'];
      QryActualiza.ParamByName('contrato').AsString := param_global_contrato;

      QryActualiza.ExecSQL;

    end;

    Insertar1.Enabled := True;
    Editar1.Enabled := True;
    Registrar1.Enabled := False;
    Can1.Enabled := False;
    Eliminar1.Enabled := True;
    Refresh1.Enabled := True;
    Salir1.Enabled := True;
    frmBarra1.btnPostClick(Sender);
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Oficio Modificaciones Personal ', 'Al actualizar registro', 0);
    end;
  end;
  QryActualiza.Destroy;
  desactivapop(popupprincipal);
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmOrdenesGeneral.frmBarra1btnCancelClick(Sender: TObject);
begin
  desactivapop(popupprincipal);
  frmBarra1.btnCancelClick(Sender);
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  OrdenesdeTrabajoGral.Cancel;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmOrdenesGeneral.frmBarra1btnDeleteClick(Sender: TObject);
var
  QryBuscaRecursos: TZReadOnlyQuery;
  qry: TZReadOnlyQuery;
begin

  qry := TZReadOnlyQuery.Create(Self);
  qry.Connection := Connection.zConnection;

  qry.Active := false;
  qry.SQL.Clear;
  qry.SQL.Add(' select * from bitacoradepersonal where sContrato=:contrato and dIdFecha between :fechai and :fechaf ');
  qry.Params.ParamByName('contrato').AsString := param_global_contrato;
  qry.Params.ParamByName('fechai').AsDate := OrdenesdeTrabajoGral.FieldValues['dFiProgramado'];
  qry.Params.ParamByName('fechaf').AsDate := OrdenesdeTrabajoGral.FieldValues['dFfProgramado'];
  qry.Open;

  if qry.RecordCount > 0 then
  begin
    MessageDLG('Existe personal reportado en reportes diarios, no se permite la eliminacion. ', mtError, [mbOk], 0);
    exit;
  end;

  qry.Active := false;
  qry.SQL.Clear;
  qry.SQL.Add(' select * from bitacoradeequipos where sContrato=:contrato and dIdFecha between :fechai and :fechaf ');
  qry.Params.ParamByName('contrato').AsString := param_global_contrato;
  qry.Params.ParamByName('fechai').AsDate := OrdenesdeTrabajoGral.FieldValues['dFiProgramado'];
  qry.Params.ParamByName('fechaf').AsDate := OrdenesdeTrabajoGral.FieldValues['dFfProgramado'];
  qry.Open;

  if qry.RecordCount > 0 then
  begin
    MessageDLG('Existe equipo reportado en reportes diarios, no se permite la eliminacion. ', mtError, [mbOk], 0);
    exit;
  end;


  QryBuscaRecursos := TZReadOnlyQuery.Create(self);
  QryBuscaRecursos.Connection := connection.zConnection;

    //Actulizamos los recurso de la vigencia..
  QryBuscaRecursos.Active := False;
  QryBuscaRecursos.SQL.Clear;
  QryBuscaRecursos.SQL.Add('Select iFolioOficio from movtorecursosxoficio where iFolioOficio =:Folio ');
  QryBuscaRecursos.ParamByName('Folio').AsString := Ordenesdetrabajogral.FieldValues['iFolioOficio'];
  QryBuscaRecursos.Open;

  if QryBuscaRecursos.RecordCount > 0 then
  begin
    messageDLG('No se puede eliminar el Oficio, contiene personal y equipo. Favor de Verificar!', mtInformation, [mbOk], 0);
    exit;
  end;

  if OrdenesdeTrabajoGral.RecordCount > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        OrdenesdeTrabajoGral.Delete;
      except
        on e: exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Oficio Modificaciones Personal ', 'Al eliminar registro', 0);
        end;
      end
    end
end;

procedure TfrmOrdenesGeneral.frmBarra1btnRefreshClick(Sender: TObject);
begin
  OrdenesdeTrabajoGral.Active;
  OrdenesdeTrabajoGral.Open;
end;

procedure TfrmOrdenesGeneral.frmBarra1btnExitClick(Sender: TObject);
begin
  frmBarra1.btnExitClick(Sender);
  Insertar1.Enabled := True;
  Editar1.Enabled := True;
  Registrar1.Enabled := False;
  Can1.Enabled := False;
  Eliminar1.Enabled := True;
  Refresh1.Enabled := True;
  Salir1.Enabled := True;
  Close;
end;


procedure TfrmOrdenesGeneral.Insertar1Click(Sender: TObject);
begin
  frmBarra1.btnAdd.click
end;

procedure TfrmOrdenesGeneral.Paste1Click(Sender: TObject);
begin
  try
    UtGrid.AddRowsFromClip;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Orden General de Trabajo', 'Al pegar registro', 0);
    end;
  end;
end;

procedure TfrmOrdenesGeneral.tdFechaAltaEnter(Sender: TObject);
begin
  tdFechaalta.Color := global_Color_Entrada;
end;

procedure TfrmOrdenesGeneral.tdFechaAltaExit(Sender: TObject);
begin
  tdFechaAlta.Color := Global_Color_Salida;
  if OrdenesdeTrabajoGral.State in [dsInsert,dsEdit] then
  if SeCruzaFecha(tdFechaAlta.DateTime, 'REGISTRO') then
  begin
    MessageDlg('La fecha de vigencia no debe ser mayor a la fecha inicial de la vigencia.', mtError, [mbOk], 0);
    tdFechaAlta.SetFocus;
  end;
end;

procedure TfrmOrdenesGeneral.tdFechaAltaKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdFechaInicio.SetFocus;
end;

procedure TfrmOrdenesGeneral.chkConsolidadaClick(Sender: TObject);
begin
  if chkConsolidada.Checked then
    chkDiaria.Checked := False;
end;

procedure TfrmOrdenesGeneral.ChkDiariaClick(Sender: TObject);
begin
  if chkDiaria.Checked then
    chkConsolidada.Checked := False;
end;

procedure TfrmOrdenesGeneral.Copy1Click(Sender: TObject);
begin
  try
    UtGrid.CopyRowsToClip;
  except
    on e: exception do begin
      UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Orden General de Trabajo', 'Al copiar registro', 0);
    end;
  end;
end;

procedure TfrmOrdenesGeneral.Editar1Click(Sender: TObject);
begin
  frmBarra1.btnEdit.Click
end;

procedure TfrmOrdenesGeneral.Registrar1Click(Sender: TObject);
begin
  frmBarra1.btnPost.Click
end;

procedure TfrmOrdenesGeneral.Can1Click(Sender: TObject);
begin
  frmBarra1.btnCancel.Click
end;

procedure TfrmOrdenesGeneral.Eliminar1Click(Sender: TObject);
begin
  frmBarra1.btnDelete.Click
end;

procedure TfrmOrdenesGeneral.Refresh1Click(Sender: TObject);
begin
  frmBarra1.btnRefresh.Click
end;

procedure TfrmOrdenesGeneral.Salir1Click(Sender: TObject);
begin
  frmBarra1.btnExit.Click
end;

procedure TfrmOrdenesGeneral.sDiasEnter(Sender: TObject);
begin
  sdias.Color := global_color_entrada
end;

procedure TfrmOrdenesGeneral.sDiasExit(Sender: TObject);
begin
  sdias.color := global_color_salida
end;

procedure TfrmOrdenesGeneral.sDiasKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tmcomentarios.SetFocus
end;

procedure TfrmOrdenesGeneral.tmDescripcionEnter(Sender: TObject);
begin
  tmDescripcion.Color := global_color_entrada
end;

procedure TfrmOrdenesGeneral.tmDescripcionExit(Sender: TObject);
begin
  tmDescripcion.Color := global_color_salida
end;


procedure TfrmOrdenesGeneral.tdFechaInicioChange(Sender: TObject);
begin

  if frmBarra1.btnPost.Enabled then
//      OrdenesdetrabajoGral.FieldValues['sDias'] := DaysBetween(OrdenesdetrabajoGral.FieldValues['dFiProgramado'], OrdenesdetrabajoGral.FieldValues['dFfProgramado']) + 1 ;
    OrdenesdetrabajoGral.FieldValues['sDias'] := DaysBetween(tdfechaInicio.date, tdFechaFinal.date) + 1;


end;

procedure TfrmOrdenesGeneral.tdFechaInicioEnter(Sender: TObject);
begin
  tdFechaInicio.Color := global_color_entrada
end;

procedure TfrmOrdenesGeneral.tdFechaInicioExit(Sender: TObject);
begin
  tdFechaInicio.Color := global_color_salida;
  if OrdenesdeTrabajoGral.State in [dsInsert,dsEdit] then
  if SeCruzaFecha(tdFechaInicio.DateTime, 'RANGO') then
  begin
    MessageDlg('La fecha seleccionada se encuentra dentro del rango de otro oficio.', mtError, [mbOk], 0);
    tdFechaInicio.SetFocus;
  end;
end;

procedure TfrmOrdenesGeneral.tdFechaFinalChange(Sender: TObject);
begin

  if frmBarra1.btnPost.Enabled then
//      OrdenesdetrabajoGral.FieldValues['sDias'] := DaysBetween(OrdenesdetrabajoGral.FieldValues['dFiProgramado'], OrdenesdetrabajoGral.FieldValues['dFfProgramado']) + 1 ;
    OrdenesdetrabajoGral.FieldValues['sDias'] := DaysBetween(tdfechaInicio.date, tdFechaFinal.date) + 1;



end;

procedure TfrmOrdenesGeneral.tdFechaFinalEnter(Sender: TObject);
begin
  tdFechaFinal.Color := global_color_entrada
end;

procedure TfrmOrdenesGeneral.tdFechaFinalExit(Sender: TObject);
begin
  tdFechaFinal.Color := global_color_salida;
  if frmBarra1.btnPost.Enabled then
    OrdenesdetrabajoGral.FieldValues['sDias'] := DaysBetween(tdfechaInicio.date, tdFechaFinal.date) + 1;
  if OrdenesdeTrabajoGral.State in [dsInsert,dsEdit] then
  if SeCruzaFecha(tdFechaFinal.DateTime, 'RANGO') then
  begin
    MessageDlg('La fecha seleccionada se encuentra dentro del rango de otro oficio.', mtError, [mbOk], 0);
    tdFechaFinal.SetFocus;
  end;
end;

procedure TfrmOrdenesGeneral.tdFechaFinalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
  begin
    tdFechaFinal.Color := global_color_salida;
    tmComentarios.SetFocus;
  end;
end;

procedure TfrmOrdenesGeneral.tmComentariosEnter(Sender: TObject);
begin
  tmComentarios.Color := global_color_entrada
end;

procedure TfrmOrdenesGeneral.tmComentariosExit(Sender: TObject);
begin
  tmComentarios.Color := global_color_salida
end;

procedure TfrmOrdenesGeneral.tsDescripcionCortaEnter(Sender: TObject);
begin
  tsDescripcionCorta.Color := global_color_entrada
end;

procedure TfrmOrdenesGeneral.tsDescripcionCortaExit(Sender: TObject);
begin
  tsDescripcionCorta.Color := global_color_salida
end;

procedure TfrmOrdenesGeneral.tsDescripcionCortaKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsOficioAutorizacion.SetFocus
end;


procedure TfrmOrdenesGeneral.tsIdFolioEnter(Sender: TObject);
begin
  tsIdFolio.Color := global_color_entrada
end;

procedure TfrmOrdenesGeneral.tsIdFolioExit(Sender: TObject);
begin
  tsIdFolio.Color := global_color_salida
end;

procedure TfrmOrdenesGeneral.tsIdFolioKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tsDescripcionCorta.SetFocus
end;

procedure TfrmOrdenesGeneral.ActualizaFechas;
var
  Mifecha, MifechaI, MiFechaF: tDate;
  i, indice: integer;
begin

  MiFecha := OrdenesdeTrabajoGral.FieldValues['dFiProgramado'];
  MiFechaI := OrdenesdeTrabajoGral.FieldValues['dFiProgramado'];
  MiFechaF := OrdenesdeTrabajoGral.FieldValues['dFfProgramado'];

     {Eliminamos los registros menores a la nueva fecha de incio}
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('Delete from detallerecursosxoficio where sContrato =:Contrato and iFolioOficio =:Oficio and dFechaDia < :Dia ');
  connection.QryBusca.ParamByName('Contrato').AsString := param_global_contrato;
  connection.QryBusca.ParamByName('Oficio').AsInteger := OrdenesdeTrabajoGral.FieldValues['iFolioOficio'];
  connection.QryBusca.ParamByName('Dia').AsDate := MiFecha;
  connection.QryBusca.ExecSQL;

     {Eliminamos los registros mayores a la nueva fecha final}
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.Clear;
  connection.QryBusca.SQL.Add('Delete from detallerecursosxoficio where sContrato =:Contrato and iFolioOficio =:Oficio and dFechaDia > :Dia ');
  connection.QryBusca.ParamByName('Contrato').AsString := param_global_contrato;
  connection.QryBusca.ParamByName('Oficio').AsInteger := OrdenesdeTrabajoGral.FieldValues['iFolioOficio'];
  connection.QryBusca.ParamByName('Dia').AsDate := MiFechaF;
  connection.QryBusca.ExecSQL;

  indice := 0;
  for i := 1 to (DaysBetween(MiFechaF, MiFechaI) + 1) do
  begin
         {Ahora buscamos las fechas que no esten registradas.}
    connection.QryBusca.Active := False;
    connection.QryBusca.SQL.Clear;
    connection.QryBusca.SQL.Add('Select dFechaDia from detallerecursosxoficio where sContrato =:Contrato and iFolioOficio =:Oficio and dFechaDia =:Dia ');
    connection.QryBusca.ParamByName('Contrato').AsString := param_global_contrato;
    connection.QryBusca.ParamByName('Oficio').AsInteger := OrdenesdeTrabajoGral.FieldValues['iFolioOficio'];
    connection.QryBusca.ParamByName('Dia').AsDate := MiFecha;
    connection.QryBusca.Open;

    if connection.QryBusca.RecordCount <= 0 then
    begin
             {Se consultan los anexos con sus partidas..}
      connection.QryBusca2.Active := False;
      connection.QryBusca2.SQL.Clear;
      connection.QryBusca2.SQL.Add('Select sNumeroActividad, sAnexo from detallerecursosxoficio where sContrato =:Contrato and iFolioOficio =:Oficio group by sAnexo, sNumeroActividad ');
      connection.QryBusca2.ParamByName('Contrato').AsString := param_global_contrato;
      connection.QryBusca2.ParamByName('Oficio').AsInteger := OrdenesdeTrabajoGral.FieldValues['iFolioOficio'];
      connection.QryBusca2.Open;

      while not connection.QryBusca2.Eof do
      begin
                 {Insertamos los registros..}
        connection.zCommand.Active := False;
        connection.zCommand.SQL.Clear;
        connection.zCommand.SQL.Add('insert into detallerecursosxoficio (sContrato, iFolioOficio, sAnexo, sNumeroActividad, dFechaDia, dCantidad) ' +
          ' values (:Contrato, :Folio, :Anexo, :Actividad, :Fecha, :Cantidad)');
        connection.zCommand.ParamByName('Contrato').AsString := param_global_contrato;
        connection.zCommand.ParamByName('Folio').AsInteger := OrdenesdeTrabajoGral.FieldValues['iFolioOficio'];
        connection.zCommand.ParamByName('Anexo').AsString := connection.QryBusca2.FieldValues['sAnexo'];
        connection.zCommand.ParamByName('Actividad').AsString := connection.QryBusca2.FieldValues['sNumeroActividad'];
        connection.zCommand.ParamByName('Fecha').AsDate := MiFecha;
        connection.zCommand.ParamByName('Cantidad').AsFloat := 0;
        connection.zCommand.ExecSQL;

        connection.QryBusca2.Next;
      end;
    end;
    MiFecha := IncDay(MiFecha);
  end;
end;

function TfrmOrdenesGeneral.SeCruzaFecha(dFecha: TDateTime; Comparar: string): Boolean;
var
  qry: TZReadOnlyQuery;
begin
  qry := TZReadOnlyQuery.Create(nil);
  qry.Connection := Connection.zConnection;
  qry.Active := false;
  qry.SQL.Clear;

  if Comparar = 'REGISTRO' then
  begin
    if tdFechaAlta.Date > tdFechaInicio.Date then
      SeCruzaFecha := True
    else
      SeCruzaFecha := false;
  end;

  if Comparar = 'RANGO' then
  begin
    qry.SQL.Add(' select sContrato from ordenesdetrabajogral  ' +
      ' where sContrato =:Contrato and :Fecha between dFiProgramado  ' +
      ' and dFfProgramado  ');
    qry.Params.ParamByName('Contrato').Value := param_global_contrato;
    qry.Params.ParamByName('Fecha').Value := dFecha;
    qry.Open;
    if qry.RecordCount > 0 then
      SeCruzaFecha := True
    else
      SeCruzaFecha := false;
  end;
end;

end.

