unit frm_trinomios;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, global, frm_barra, Grids, DBGrids, StdCtrls,
  ExtCtrls, DBCtrls, Mask, DB, Menus, ADODB, RxCombos, 
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UnitExcepciones, UnitTBotonesPermisos, udbgrid, UnitValidaTexto,
  UnitTablasImpactadas, unitactivapop, UnitValidacion;
type
  TfrmTrinomios = class(TForm)
    grid_trinomios: TDBGrid;
    frmBarra1: TfrmBarra;
    ds_trinomio: TDataSource;
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
    N4: TMenuItem;
    Salir1: TMenuItem;
    Label1: TLabel;
    Label2: TLabel;
    tsFondo: TDBEdit;
    Label3: TLabel;
    tsPosicionFinanciera: TDBEdit;
    Label4: TLabel;
    tsCentroGestorEjecutor: TDBEdit;
    Label5: TLabel;
    tsCuentaMayor: TDBEdit;
    Label6: TLabel;
    tsElementoPep: TDBEdit;
    tsVigente: TDBCheckBox;
    trinomio: TZQuery;
    dbUtilizacion: TDBEdit;
    Label7: TLabel;
    dbCosto: TDBEdit;
    dbBeneficio: TDBEdit;
    Label8: TLabel;
    Label9: TLabel;
    dbActividad: TDBEdit;
    Label10: TLabel;
    dbInstalacion: TDBLookupComboBox;
    qryOrdenes: TZReadOnlyQuery;
    ds_Ordenes: TDataSource;
    trinomiosContrato: TStringField;
    trinomiosInstalacion: TStringField;
    trinomiosFondo: TStringField;
    trinomiosPosicionFinanciera: TStringField;
    trinomiosCentroGestorEjecutor: TStringField;
    trinomiosCuentaMayor: TStringField;
    trinomiosElementoPep: TStringField;
    trinomiolVigente: TStringField;
    trinomiosUtilizacionFondo: TStringField;
    trinomiosCosto: TStringField;
    trinomiosBeneficio: TStringField;
    trinomiosActividad: TStringField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure grid_trinomiosCellClick(Column: TColumn);
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
    procedure tsInstalacionEnter(Sender: TObject);
    procedure tsInstalacionExit(Sender: TObject);
    procedure tsInstalacionKeyPress(Sender: TObject; var Key: Char);
    procedure tsFondoEnter(Sender: TObject);
    procedure tsFondoExit(Sender: TObject);
    procedure tsFondoKeyPress(Sender: TObject; var Key: Char);
    procedure tsPosicionFinancieraEnter(Sender: TObject);
    procedure tsPosicionFinancieraExit(Sender: TObject);
    procedure tsPosicionFinancieraKeyPress(Sender: TObject; var Key: Char);
    procedure tsCentroGestorEjecutorEnter(Sender: TObject);
    procedure tsCentroGestorEjecutorExit(Sender: TObject);
    procedure tsCentroGestorEjecutorKeyPress(Sender: TObject;
      var Key: Char);
    procedure tsCuentaMayorEnter(Sender: TObject);
    procedure tsCuentaMayorExit(Sender: TObject);
    procedure tsCuentaMayorKeyPress(Sender: TObject; var Key: Char);
    procedure tsElementoPepEnter(Sender: TObject);
    procedure tsElementoPepExit(Sender: TObject);
    procedure tsElementoPepKeyPress(Sender: TObject; var Key: Char);
    procedure dbUtilizacionEnter(Sender: TObject);
    procedure dbUtilizacionExit(Sender: TObject);
    procedure dbCostoEnter(Sender: TObject);
    procedure dbCostoExit(Sender: TObject);
    procedure dbBeneficioEnter(Sender: TObject);
    procedure dbBeneficioExit(Sender: TObject);
    procedure dbActividadEnter(Sender: TObject);
    procedure dbActividadExit(Sender: TObject);
    procedure dbCostoKeyPress(Sender: TObject; var Key: Char);
    procedure dbBeneficioKeyPress(Sender: TObject; var Key: Char);
    procedure dbUtilizacionKeyPress(Sender: TObject; var Key: Char);
    procedure dbInstalacionEnter(Sender: TObject);
    procedure dbInstalacionExit(Sender: TObject);
    procedure dbInstalacionKeyPress(Sender: TObject; var Key: Char);
    procedure grid_trinomiosMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_trinomiosMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_trinomiosTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    function tablasDependientes(idOrig: string): boolean;
    function posibleBorrar(idOrig: string): boolean;
    procedure dbActividadKeyPress(Sender: TObject; var Key: Char);
    procedure trinomiosCostoSetText(Sender: TField; const Text: string);
    procedure trinomiosCuentaMayorSetText(Sender: TField; const Text: string);
    procedure dbCostoChange(Sender: TObject);
    procedure tsCuentaMayorChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmTrinomios: TfrmTrinomios;
  BotonPermiso: TBotonesPermisos;
  utgrid:ticdbgrid;
  sIdOrig : string;
implementation
uses
   frm_estimainstalado;

{$R *.dfm}

procedure TfrmTrinomios.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  global_frmActivo := 'ninguno';
  Trinomio.Cancel ;
  action := cafree ;
  BotonPermiso.free;
  utgrid.destroy;
end;

procedure TfrmTrinomios.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cTrinomios', PopupPrincipal);
  UtGrid:=TicdbGrid.create(grid_trinomios);
  OpcButton := '' ;
  sIdOrig := '';
  frmBarra1.btnCancel.Click ;
  Trinomio.Active := False ;
  Trinomio.Params.ParamByName('Contrato').DataType := ftString ;
  Trinomio.Params.ParamByName('Contrato').Value := Global_Contrato ;
  Trinomio.Open ;

  qryOrdenes.Active := False ;
  qryOrdenes.Params.ParamByName('Contrato').DataType := ftString ;
  qryOrdenes.Params.ParamByName('Contrato').Value := Global_Contrato ;
  qryOrdenes.Open ;

  if (global_frmActivo = 'frm_estimainstalado') then
      frmBarra1.btnAdd.Click;

  if Global_OptGrafica = 'Frente' then     // Agregar directamente centro presupuestal..
  begin
      frmBarra1.btnAdd.Click;
      dbInstalacion.Field.Text := global_FrenteTrabajo;
      dbCosto.SetFocus;
      dbCosto.text                := '0' ;
      tsCentroGestorEjecutor.Text := '*';
      dbBeneficio.Text            := '*';
      tsCuentaMayor.Text          := '0'  ;
      tsPosicionFinanciera.Text   := '*';
      tsFondo.Text := '*';
      tsElementoPep.Text          := '*';
      dbUtilizacion.Text          := '*';
      dbActividad.Text            := '*';
  end;

  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmTrinomios.grid_trinomiosCellClick(Column: TColumn);
begin
  if frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;

end;

procedure TfrmTrinomios.grid_trinomiosMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmTrinomios.grid_trinomiosMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmTrinomios.grid_trinomiosTitleClick(Column: TColumn);
begin
 UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmTrinomios.frmBarra1btnAddClick(Sender: TObject);
begin
   activapop(frmTrinomios, popupprincipal);
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   Trinomio.Append ;
   Trinomio.FieldValues [ 'sContrato' ]  := Global_Contrato ;
   dbInstalacion.SetFocus ;
   tsVigente.Checked := True ;
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
   grid_trinomios.Enabled := False;
end;

procedure TfrmTrinomios.frmBarra1btnEditClick(Sender: TObject);
begin
   activapop(frmTrinomios, popupprincipal);
   frmBarra1.btnEditClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   sIdOrig := trinomio.FieldByName('sInstalacion').AsString;
   try
      Trinomio.Edit ;
   except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Trinomios del Contrato', 'Al editar registro', 0);
           frmbarra1.btnCancel.Click ;
       end;
   end ;
   dbInstalacion.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
   grid_trinomios.Enabled := False;
end;

procedure TfrmTrinomios.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
  lEdita: boolean;
  sTrinomio : string;
begin
   {Validacion de campos}
   nombres:=TStringList.Create;cadenas:=TStringList.Create;
   nombres.Add('Instalacion');nombres.Add('Costo');nombres.Add('Gestor');nombres.Add('Cuenta de Mayor');
   nombres.Add('Posicion Financiera');nombres.Add('Fondo');nombres.Add('Elemento Pep');
   cadenas.Add(dbInstalacion.Text);cadenas.Add(dbCosto.Text);cadenas.Add(tsCentroGestorEjecutor.Text);cadenas.Add(tsCuentaMayor.Text);
   cadenas.Add(tsPosicionFinanciera.Text);cadenas.Add(tsFondo.Text);cadenas.Add(tsElementoPep.Text);
   if not validaTexto(nombres, cadenas, '', '') then
   begin
     MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
     exit;
   end;
   {Continua insercion de datos}
   lEdita := false;
   if Trinomio.State = dsEdit then
      lEdita := true;
   try
      Trinomio.FieldValues['lVigente'] := 'Si' ;
      sTrinomio := Trinomio.FieldValues['sInstalacion'];
      Trinomio.Post ;

      if (global_frmActivo = 'frm_estimainstalado') then
      begin
           frmEstimaInstalado.tsInstalacion.Items.Add(sTrinomio) ;
           frmEstimaInstalado.tsInstalacion.ItemIndex := frmEstimaInstalado.tsInstalacion.Items.IndexOf(sTrinomio);
           frmEstimaInstalado.tsInstalacion.SetFocus;
      end;
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
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Trinomios del Contrato', 'Al salvar registro', 0);
           frmbarra1.btnCancel.Click ;
           lEdita := false;//cancelar la actualizacion de tablas dependientes
       end;
   end;
   if (lEdita) and (Trinomio.FieldByName('sInstalacion').AsString <> sIdOrig) then
   begin
       tablasDependientes(sIdOrig);
   end;

   if (global_frmActivo = 'frm_estimainstalado') then
   begin
        global_frmActivo := 'ninguno';
        frmbarra1.btnCancel.Click;
        frmbarra1.btnExit.Click;
   end;

   desactivapop(popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
   grid_trinomios.Enabled := True;
end;

procedure TfrmTrinomios.frmBarra1btnCancelClick(Sender: TObject);
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
   Trinomio.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
   grid_trinomios.Enabled := True;
end;

procedure TfrmTrinomios.frmBarra1btnDeleteClick(Sender: TObject);
var
   sTrinomio : string;
begin
  If Trinomio.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if not posibleBorrar(Trinomio.FieldByName('sInstalacion').AsString) then
      begin
        MessageDlg('No es posible eliminar el registro, existen registros dependientes.', mtInformation, [mbOk], 0);
        exit;
      end;
      try
        sTrinomio := trinomio.FieldValues['sInstalacion'];
        Trinomio.Delete ;

        if (global_frmActivo = 'frm_estimainstalado') then
        begin
            frmEstimaInstalado.tsInstalacion.Items.Delete(frmEstimaInstalado.tsInstalacion.Items.IndexOf(sTrinomio));
            frmEstimaInstalado.tsInstalacion.Refresh ;
            frmEstimaInstalado.tsInstalacion.SetFocus;
        end;
      except
        on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Trinomios del Contrato', 'Al eliminar registro', 0);
        end;
      end
    end ;
   if (global_frmActivo = 'frm_estimainstalado') then
   begin
       global_frmActivo := 'ninguno';
       frmbarra1.btnCancel.Click;
       frmbarra1.btnExit.Click;
   end;
end;

procedure TfrmTrinomios.frmBarra1btnRefreshClick(Sender: TObject);
begin
  Trinomio.Active := False ;
  Trinomio.Open
end;

procedure TfrmTrinomios.frmBarra1btnExitClick(Sender: TObject);
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

procedure TfrmTrinomios.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click 
end;

procedure TfrmTrinomios.Paste1Click(Sender: TObject);
begin
UtGrid.AddRowsFromClip;
end;

procedure TfrmTrinomios.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure TfrmTrinomios.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click 
end;

procedure TfrmTrinomios.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmTrinomios.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click
end;

procedure TfrmTrinomios.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmTrinomios.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

procedure TfrmTrinomios.tsInstalacionEnter(Sender: TObject);
begin
    dbInstalacion.Color := global_color_entrada
end;

procedure TfrmTrinomios.tsInstalacionExit(Sender: TObject);
begin
    dbInstalacion.Color := global_color_salida
end;

procedure TfrmTrinomios.tsInstalacionKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsFondo.SetFocus
end;

procedure TfrmTrinomios.tsFondoEnter(Sender: TObject);
begin
    tsFondo.Color := global_Color_entrada
end;

procedure TfrmTrinomios.tsFondoExit(Sender: TObject);
begin
    tsFondo.Color := global_Color_salida
end;

procedure TfrmTrinomios.tsFondoKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
       tsElementoPep.SetFocus
end;

procedure TfrmTrinomios.tsPosicionFinancieraEnter(Sender: TObject);
begin
    tsPosicionFinanciera.Color := global_color_entrada
end;

procedure TfrmTrinomios.tsPosicionFinancieraExit(Sender: TObject);
begin
    tsPosicionFinanciera.Color := global_color_salida
end;

procedure TfrmTrinomios.tsPosicionFinancieraKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsFondo.SetFocus ;
        
end;

function TfrmTrinomios.tablasDependientes(idOrig: string): boolean;
var
  ParamNamesSET,ParamValuesSET,ParamNamesWHERE,ParamValuesWHERE: TStringList;
begin
  result := true;
  ParamNamesSET:=TStringList.Create;ParamValuesSET:=TStringList.Create;ParamNamesWHERE:=TStringList.Create;ParamValuesWHERE:=TStringList.Create;
  ParamNamesSET.Add('sInstalacion');ParamValuesSET.Add(Trinomio.FieldByName('sInstalacion').AsString);
  ParamNamesWHERE.Add('sContrato');ParamValuesWHERE.Add(global_contrato);
  ParamNamesWHERE.Add('sInstalacion');ParamValuesWHERE.Add(idOrig);
  if not UnitTablasImpactadas.impactar('contrato_trinomio',ParamNamesSET,ParamValuesSET,ParamNamesWHERE,ParamValuesWHERE) then
  begin
    result := false;
    showmessage('Ocurrio un error al actualizar las tablas dependientes: ' + UnitTablasImpactadas.xError);
  end;
end;
procedure TfrmTrinomios.trinomiosCostoSetText(Sender: TField;
  const Text: string);
begin
  Sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmTrinomios.trinomiosCuentaMayorSetText(Sender: TField;
  const Text: string);
begin
  Sender.Value:=abs(StrToFloatDef(text,0));
end;

function TfrmTrinomios.posibleBorrar(idOrig: string): boolean;
var
  ParamNamesWHERE,ParamValuesWHERE: TStringList;
begin
  result := true;
  ParamNamesWHERE:=TStringList.Create;ParamValuesWHERE:=TStringList.Create;
  ParamNamesWHERE.Add('sContrato');ParamValuesWHERE.Add(global_contrato);
  ParamNamesWHERE.Add('sInstalacion');ParamValuesWHERE.Add(idOrig);
  result := not UnitTablasImpactadas.hayDependientes('contrato_trinomio',ParamNamesWHERE,ParamValuesWHERE);
end;

procedure TfrmTrinomios.tsCentroGestorEjecutorEnter(Sender: TObject);
begin
    tsCentroGestorEjecutor.Color := global_color_entrada
end;

procedure TfrmTrinomios.tsCentroGestorEjecutorExit(Sender: TObject);
begin
    tsCentroGestorEjecutor.Color := global_color_salida
end;

procedure TfrmTrinomios.tsCentroGestorEjecutorKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        dbBeneficio.SetFocus ;
end;

procedure TfrmTrinomios.tsCuentaMayorChange(Sender: TObject);
begin
tdbeditchangef(tsCuentaMayor,'Cuenta de Mayor');
end;

procedure TfrmTrinomios.tsCuentaMayorEnter(Sender: TObject);
begin
    tsCuentaMayor.Color := global_color_entrada
end;

procedure TfrmTrinomios.tsCuentaMayorExit(Sender: TObject);
begin
    tsCuentaMayor.Color := global_color_salida
end;

procedure TfrmTrinomios.tsCuentaMayorKeyPress(Sender: TObject;
  var Key: Char);
begin
    if not keyFiltroTdbedit(tsCuentaMayor,key) then
      key:=#0;
    If Key = #13 Then
        tsPosicionFinanciera.SetFocus ;
end;

procedure TfrmTrinomios.tsElementoPepEnter(Sender: TObject);
begin
    tselementoPEp.Color := global_color_entrada
end;

procedure TfrmTrinomios.tsElementoPepExit(Sender: TObject);
begin
    tselementoPEp.Color := global_color_salida
end;

procedure TfrmTrinomios.tsElementoPepKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        dbUtilizacion.SetFocus ;
end;

procedure TfrmTrinomios.dbUtilizacionEnter(Sender: TObject);
begin
   dbUtilizacion.Color  := Global_Color_Entrada;
end;

procedure TfrmTrinomios.dbUtilizacionExit(Sender: TObject);
begin
   dbUtilizacion.Color  := Global_Color_Salida ;
end;

procedure TfrmTrinomios.dbCostoChange(Sender: TObject);
begin
tdbeditchangef(dbCosto,'Costo');
end;

procedure TfrmTrinomios.dbCostoEnter(Sender: TObject);
begin
  dbCosto.Color := Global_Color_Entrada ;
end;

procedure TfrmTrinomios.dbCostoExit(Sender: TObject);
begin
  dbCosto.Color := Global_Color_Salida ;
end;

procedure TfrmTrinomios.dbBeneficioEnter(Sender: TObject);
begin
  dbBeneficio.Color := Global_Color_Entrada ;
end;

procedure TfrmTrinomios.dbBeneficioExit(Sender: TObject);
begin
   dbBeneficio.Color := Global_Color_Salida  ;
end;

procedure TfrmTrinomios.Copy1Click(Sender: TObject);
begin
UtGrid.CopyRowsToClip;
end;

procedure TfrmTrinomios.dbActividadEnter(Sender: TObject);
begin
  dbActividad.Color := Global_Color_Entrada ;
end;

procedure TfrmTrinomios.dbActividadExit(Sender: TObject);
begin
   dbActividad.Color  := Global_Color_Salida ;
end;

procedure TfrmTrinomios.dbActividadKeyPress(Sender: TObject; var Key: Char);
begin
//  if not keyFiltroTdbedit(dbActividad,key) then
//      key:=#0;
  If key=#13 Then
    dbCosto.SetFocus ;
end;
procedure TfrmTrinomios.dbCostoKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTdbedit(dbcosto,key) then
      key:=#0;
  If Key=#13 Then
    tsCentroGestorEjecutor.SetFocus ;
end;

procedure TfrmTrinomios.dbBeneficioKeyPress(Sender: TObject;
  var Key: Char);
begin
  If key=#13 Then
    tsCuentaMayor.SetFocus
end;

procedure TfrmTrinomios.dbUtilizacionKeyPress(Sender: TObject;
  var Key: Char);
begin
  If key=#13 Then
    dbActividad.SetFocus ;
end;

procedure TfrmTrinomios.dbInstalacionEnter(Sender: TObject);
begin
  dbInstalacion.Color  := Global_Color_Entrada ;
end;

procedure TfrmTrinomios.dbInstalacionExit(Sender: TObject);
begin
  dbInstalacion.Color  := Global_Color_Salida ;
end;

procedure TfrmTrinomios.dbInstalacionKeyPress(Sender: TObject; var Key: Char);
begin
  If key=#13 Then
    dbCosto.SetFocus ;
End;

end.
