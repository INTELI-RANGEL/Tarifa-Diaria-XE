unit frm_Fases;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, global, frm_barra, Grids, DBGrids, StdCtrls,
  ExtCtrls, DBCtrls, Mask, DB, Menus, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, unitexcepciones, udbgrid, unittbotonespermisos,
  UnitValidaTexto, UnitTablasImpactadas, unitactivapop;

type
  TfrmFases = class(TForm)
    grid_fases: TDBGrid;
    Label1: TLabel;
    Label2: TLabel;
    tsIdFase: TDBEdit;
    tsDescripcion: TDBEdit;
    frmBarra1: TfrmBarra;
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
    ds_Fases: TDataSource;
    QryFases: TZQuery;
    procedure tsIdEmbarcacionKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure grid_fasesCellClick(Column: TColumn);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure tsDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tsIdFaseEnter(Sender: TObject);
    procedure tsIdFaseExit(Sender: TObject);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure grid_fasesMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_fasesMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_fasesTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    function tablasDependientes(idOrig: string): boolean;
    function posibleBorrar(idOrig: string): boolean;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmFases: TfrmFases;
  utgrid:ticdbgrid;
  botonpermiso:tbotonespermisos;
  sIdOrig : string;

implementation

{$R *.dfm}

procedure TfrmFases.tsIdEmbarcacionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsDescripcion.SetFocus
end;

procedure TfrmFases.tsIdFaseEnter(Sender: TObject);
begin
  tsidfase.color:= global_color_entrada
end;

procedure TfrmFases.tsIdFaseExit(Sender: TObject);
begin
  tsidfase.color:=global_color_salida
end;

procedure TfrmFases.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  action := cafree ;
  utgrid.destroy;
  botonpermiso.Free;
end;

procedure TfrmFases.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cFases', PopupPrincipal);
  UtGrid:=TicdbGrid.create(grid_fases);
  OpcButton := '' ;
  sIdOrig := '';
  frmbarra1.btnCancel.Click ;
  QryFases.Active := False ;
  QryFases.ParamByName('contrato').Value := global_contrato;
  QryFases.Open ;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmFases.grid_fasesCellClick(Column: TColumn);
begin
  if frmBarra1.btnCancel.Enabled = True then
      frmbarra1.btnCancel.Click ;
end;

procedure TfrmFases.grid_fasesMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
   UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmFases.grid_fasesMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmFases.grid_fasesTitleClick(Column: TColumn);
begin
   UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmFases.frmBarra1btnAddClick(Sender: TObject);
begin
  activapop(frmFases,popupprincipal);
  frmBarra1.btnAddClick(Sender);
  Insertar1.Enabled := False ;
  Editar1.Enabled := False ;
  Registrar1.Enabled := True ;
  Can1.Enabled := True ;
  Eliminar1.Enabled := False ;
  Refresh1.Enabled := False ;
  Salir1.Enabled := False ;
  tsIdFase.SetFocus ;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
  QryFases.Append ;
end;

procedure TfrmFases.frmBarra1btnEditClick(Sender: TObject);
begin
   activapop(frmFases,popupprincipal);
   frmBarra1.btnEditClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   sIdOrig := QryFases.FieldByName('sIdFase').AsString;
   try
      QryFases.Edit ;
   except
     on e : exception do begin
       UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Fases', 'Al editar registro', 0);
       frmbarra1.btnCancel.Click ;
     end;
   end ;
   tsIdFase.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmFases.frmBarra1btnPostClick(Sender: TObject);
var
   nombres, cadenas: TStringList;
   lEdita: boolean;
begin
   {Validacion de campos}
   nombres:=TStringList.Create;cadenas:=TStringList.Create;
   nombres.Add('Descripcion');cadenas.Add(tsDescripcion.Text);
   if not validaTexto(nombres, cadenas, 'Fase', tsIdFase.Text) then
   begin
     MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
     exit;
   end;

   lEdita := false;
   if QryFases.State = dsEdit then
     lEdita := true;

   {Continua insercion de datos}
   try
      desactivapop(popupprincipal);
      QryFases.FieldByName('sContrato').Value := global_contrato;
      QryFases.Post ;
      Insertar1.Enabled := True ;
      Editar1.Enabled := True ;
      Registrar1.Enabled := False ;
      Can1.Enabled := False ;
      Eliminar1.Enabled := True ;
      Refresh1.Enabled := True ;
      Salir1.Enabled := True ;
      frmBarra1.btnPostClick(Sender);
   except
       on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Fases', 'Al salvar registro', 0);
        frmbarra1.btnCancel.Click ;
        lEdita := false;//cancelar la actualizacion de tablas dependientes
       end;
   end;
   if (lEdita) and (QryFases.FieldByName('sIdFase').AsString <> sIdOrig) then
   begin
       tablasDependientes(sIdOrig);
   end;
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmFases.frmBarra1btnCancelClick(Sender: TObject);
begin
   desactivapop(popupprincipal);
   frmBarra1.btnCancelClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   QryFases.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmFases.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If QryFases.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if not posibleBorrar(QryFases.FieldByName('sIdFase').AsString) then
      begin
        MessageDlg('No es posible eliminar el registro, existen registros dependientes.', mtInformation, [mbOk], 0);
        exit;
      end;
      try
        QryFases.Delete ;
      except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Fases', 'Al eliminar registro', 0);
       end;
      end
    end
end;

procedure TfrmFases.frmBarra1btnRefreshClick(Sender: TObject);
begin
   QryFases.Refresh ;
end;

procedure TfrmFases.frmBarra1btnExitClick(Sender: TObject);
begin
   frmBarra1.btnExitClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   close
end;

function TfrmFases.tablasDependientes(idOrig: string): boolean;
var
  ParamNamesSET,ParamValuesSET,ParamNamesWHERE,ParamValuesWHERE: TStringList;
begin
  result := true;
  ParamNamesSET:=TStringList.Create;ParamValuesSET:=TStringList.Create;ParamNamesWHERE:=TStringList.Create;ParamValuesWHERE:=TStringList.Create;
  ParamNamesSET.Add('sIdFase');ParamValuesSET.Add(QryFases.FieldByName('sIdFase').AsString);
  ParamNamesWHERE.Add('sContrato');ParamValuesWHERE.Add(global_contrato);
  ParamNamesWHERE.Add('sIdFase');ParamValuesWHERE.Add(idOrig);
  if not UnitTablasImpactadas.impactar('fases',ParamNamesSET,ParamValuesSET,ParamNamesWHERE,ParamValuesWHERE) then
  begin
    result := false;
    showmessage('Ocurrio un error al actualizar las tablas dependientes: ' + UnitTablasImpactadas.xError);
  end;
end;

function TfrmFases.posibleBorrar(idOrig: string): boolean;
var
  ParamNamesWHERE,ParamValuesWHERE: TStringList;
begin
  result := true;
  ParamNamesWHERE:=TStringList.Create;ParamValuesWHERE:=TStringList.Create;
  ParamNamesWHERE.Add('sContrato');ParamValuesWHERE.Add(global_contrato);
  ParamNamesWHERE.Add('sIdFase');ParamValuesWHERE.Add(idOrig);
  result := not UnitTablasImpactadas.hayDependientes('fases',ParamNamesWHERE,ParamValuesWHERE);
end;

procedure TfrmFases.tsDescripcionEnter(Sender: TObject);
begin
  tsdescripcion.Color:=global_color_entrada
end;

procedure TfrmFases.tsDescripcionExit(Sender: TObject);
begin
  tsdescripcion.color:=global_color_salida
end;

procedure TfrmFases.tsDescripcionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsIdFase.SetFocus
end;

procedure TfrmFases.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure TfrmFases.Paste1Click(Sender: TObject);
begin
utgrid.AddRowsFromClip
end;

procedure TfrmFases.Copy1Click(Sender: TObject);
begin
utgrid.CopyRowsToClip
end;

procedure TfrmFases.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click
end;

procedure TfrmFases.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click
end;

procedure TfrmFases.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click
end;

procedure TfrmFases.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click
end;

procedure TfrmFases.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmFases.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

end.
