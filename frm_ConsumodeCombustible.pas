unit frm_ConsumodeCombustible;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, frm_barra, StdCtrls, DBCtrls,
  Mask, ExtCtrls, DB, Global, Menus, frxClass, frxDBSet,
  ZAbstractRODataset, ZAbstractDataset, ZDataset, UdbGrid,
  UnitExcepciones, unittbotonespermisos, unitactivapop, 
  ComCtrls, JvExComCtrls, JvDateTimePicker, DateUtils, RXDBCtrl, Newpanel,
  NxCollection, AdvGlowButton, Buttons;

type
  TfrmConsumodeCombustible = class(TForm)
    grid_plataformas: TDBGrid;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    frmBarra1: TfrmBarra;
    DBPlataformas: TfrxDBDataset;
    frxPlataformas: TfrxReport;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    Imprimir1: TMenuItem;
    N3: TMenuItem;
    Copy1: TMenuItem;
    Paste1: TMenuItem;
    N4: TMenuItem;
    Salir1: TMenuItem;
    ds_Combustibles: TDataSource;
    zq_consumosporequipo: TZQuery;
    tSeleccionarFecha: TJvDateTimePicker;
    zq_equipos: TZQuery;
    zq_recursos: TZQuery;
    zq_recursosiIdRecursoExistencia: TLargeintField;
    zq_recursossMedida: TStringField;
    zq_recursossDescripcion: TStringField;
    zq_recursoslCombustible: TStringField;
    zq_equipossContrato: TStringField;
    zq_equipossIdEquipo: TStringField;
    zq_equiposiItemOrden: TIntegerField;
    zq_equipossDescripcion: TStringField;
    zq_equipossIdTipoEquipo: TStringField;
    zq_equipossMedida: TStringField;
    zq_equiposdCantidad: TFloatField;
    zq_equiposdCostoMN: TFloatField;
    zq_equiposdCostoDLL: TFloatField;
    zq_equiposdVentaMN: TFloatField;
    zq_equiposdVentaDLL: TFloatField;
    zq_equiposdFechaInicio: TDateField;
    zq_equiposdFechaFinal: TDateField;
    zq_equiposlProrrateo: TStringField;
    zq_equiposlCobro: TStringField;
    zq_equiposlImprime: TStringField;
    zq_equiposiJornada: TIntegerField;
    zq_equiposlDistribuye: TStringField;
    zq_equiposlCuadraEquipo: TStringField;
    Label4: TLabel;
    ds_Equipos: TDataSource;
    ds_Recursos: TDataSource;
    tsdCantidad: TDBEdit;
    cmbConsumo: TDBLookupComboBox;
    ZQuery1: TZQuery;
    StringField1: TStringField;
    StringField2: TStringField;
    IntegerField1: TIntegerField;
    StringField3: TStringField;
    StringField4: TStringField;
    StringField5: TStringField;
    FloatField1: TFloatField;
    FloatField2: TFloatField;
    FloatField3: TFloatField;
    FloatField4: TFloatField;
    FloatField5: TFloatField;
    DateField1: TDateField;
    DateField2: TDateField;
    StringField6: TStringField;
    StringField7: TStringField;
    StringField8: TStringField;
    IntegerField2: TIntegerField;
    StringField9: TStringField;
    StringField10: TStringField;
    DataSource1: TDataSource;
    ZQuery2: TZQuery;
    LargeintField1: TLargeintField;
    StringField11: TStringField;
    StringField12: TStringField;
    StringField13: TStringField;
    DataSource2: TDataSource;
    zq_consumosporequipoiId: TIntegerField;
    zq_consumosporequiposIdEquipo: TStringField;
    zq_consumosporequiposNumerosDeSerie: TStringField;
    zq_consumosporequipoiIdTipoConsumo: TIntegerField;
    zq_consumosporequiposMedida: TStringField;
    zq_consumosporequipodCantidad: TFloatField;
    zq_consumosporequipodIdFecha: TDateField;
    zq_consumosporequiposTurno: TStringField;
    zq_consumosporequiposContrato: TStringField;
    zq_consumosporequiposNumeroOrden: TStringField;
    Label6: TLabel;
    lbl1: TLabel;
    tsMedida: TDBEdit;
    zq_equiposlAplicaDiesel: TStringField;
    zq_equipossDescripcionDiesel: TStringField;
    ZQuery1lAplicaDiesel: TStringField;
    ZQuery1sDescripcionDiesel: TStringField;
    raerconsumosdeldaanterior1: TMenuItem;
    zq_consumosporequiposEquipo: TStringField;
    lbl2: TLabel;
    tsNumeroSerie: TDBEdit;
    zqryPlatafomas: TZQuery;
    ds_plataformas: TDataSource;
    Panel: tNewGroupBox;
    gridListaObjeto: TRxDBGrid;
    lbl3: TLabel;
    tsDescripcion: TDBEdit;
    cmbPlataforma: TDBLookupComboBox;
    zqryPlatafomassIdPlataforma: TStringField;
    zqryPlatafomassDescripcion: TStringField;
    zq_consumosporequiposIdPlataforma: TStringField;
    zq_consumosporequiposDescripcion: TStringField;
    zq_consumosporequiposPlataforma: TStringField;
    GridCombustibles: TDBGrid;
    PnlFiltro: TPanel;
    PnlBarra: TPanel;
    PnlPosterior: TPanel;
    pnlSuperior: TPanel;
    Splitter1: TSplitter;
    EdtBusca: TEdit;
    PanelMateriales: tNewGroupBox;
    PanelDatos: TNxFlipPanel;
    Label5: TLabel;
    Label7: TLabel;
    lbl4: TLabel;
    Label9: TLabel;
    pmMaterial: TPopupMenu;
    NuevoMaterial: TMenuItem;
    Tm1: TDBMemo;
    EdtUM: TDBEdit;
    EdtIDM: TDBEdit;
    DsBuscaobjeto: TDataSource;
    BtnCancelSel: TAdvGlowButton;
    BtnSeleccionarMat: TAdvGlowButton;
    BtnAgregarMAt: TAdvGlowButton;
    BtnCancelMat: TAdvGlowButton;
    BtnGuardarMat: TAdvGlowButton;
    EdtSuma: TEdit;
    Label8: TLabel;
    lbl5: TLabel;
    cmbFolios: TDBLookupComboBox;
    zqFolios: TZQuery;
    ds_folios: TDataSource;
    zq_consumosporequiposIdProveedor: TStringField;
    Label10: TLabel;
    tsProveedor: TDBLookupComboBox;
    zq_proveedores: TZQuery;
    ds_proveedores: TDataSource;
    zq_proveedoressIdProveedor: TStringField;
    zq_proveedoressRazon: TStringField;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure grid_plataformasCellClick(Column: TColumn);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure grid_plataformasMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure grid_plataformasMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_plataformasTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure Imprimir1Click(Sender: TObject);
    procedure tSeleccionarFechaExit(Sender: TObject);
    procedure dblkcbbsIdEquipoEnter(Sender: TObject);
    procedure DBMemo1Enter(Sender: TObject);
    procedure cmbConsumoEnter(Sender: TObject);
    procedure tsdCantidadEnter(Sender: TObject);
    procedure dblkcbbsIdEquipoExit(Sender: TObject);
    procedure DBMemo1Exit(Sender: TObject);
    procedure cmbConsumoExit(Sender: TObject);
    procedure tsdCantidadExit(Sender: TObject);
    procedure tsdCantidadKeyPress(Sender: TObject; var Key: Char);
    procedure raerconsumosdeldaanterior1Click(Sender: TObject);
    procedure gridListaObjetoKeyPress(Sender: TObject; var Key: Char);
    procedure gridListaObjetoCellClick(Column: TColumn);
    procedure zq_consumosporequipoCalcFields(DataSet: TDataSet);
    procedure tsIdEquipoKeyPress(Sender: TObject; var Key: Char);
    procedure gridListaObjetoDblClick(Sender: TObject);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure cmbPlataformaEnter(Sender: TObject);
    procedure cmbPlataformaExit(Sender: TObject);
    procedure tsNumeroSerieEnter(Sender: TObject);
    procedure tsNumeroSerieExit(Sender: TObject);
    procedure tsMedidaEnter(Sender: TObject);
    procedure tsMedidaExit(Sender: TObject);
    procedure cmbConsumoKeyPress(Sender: TObject; var Key: Char);
    procedure cmbPlataformaKeyPress(Sender: TObject; var Key: Char);
    procedure ds_CombustiblesStateChange(Sender: TObject);
    procedure zq_consumosporequipoAfterScroll(DataSet: TDataSet);
    procedure EdtBuscaExit(Sender: TObject);
    procedure EdtBuscaEnter(Sender: TObject);
    procedure NuevoMaterialClick(Sender: TObject);
    procedure BtnCancelMatClick(Sender: TObject);
    procedure BtnGuardarMatClick(Sender: TObject);
    procedure BtnAgregarMAtClick(Sender: TObject);
    procedure BtnSeleccionarMatClick(Sender: TObject);
    procedure BtnCancelSelClick(Sender: TObject);
    procedure cmbFoliosKeyPress(Sender: TObject; var Key: Char);
    procedure cmbFoliosEnter(Sender: TObject);
    procedure cmbFoliosExit(Sender: TObject);
  private
  sMenuP: String;
  VCancelar:Boolean;
    procedure Sumatoria(cmp: tedit);
    { Private declarations }
  public
    { Public declarations }
  end;

var
   frmConsumodeCombustible: TfrmConsumodeCombustible;
   UtGrid:TicDbGrid;
   botonpermiso: tbotonespermisos;
   sOpcion, lStatusOrig: string;
   FechaReporte: TDateTime;
   BuscaObjeto: Tzquery;
implementation

{$R *.dfm}

procedure TfrmConsumodeCombustible.FormShow(Sender: TObject);
begin
  zQuery1.ParamByName('Contrato').AsString := Global_Contrato_Barco;
  zQuery1.Open;
  zQuery2.Open;
  
  zq_Equipos.Active := False;
  zq_Equipos.ParamByName('Contrato').AsString := Global_Contrato_Barco;
  zq_Equipos.Open;

  BuscaObjeto := TZQuery.Create(nil);
  BuscaObjeto.Connection := Connection.zConnection;

  zq_recursos.Active := False;
  zq_recursos.Open;

  zqryPlatafomas.Active := False;
  zqryPlatafomas.Open;

  zq_proveedores.Active := False;
  zq_proveedores.Open;

  zqFolios.Active := False;
  zqFolios.ParamByName('contrato').AsString := global_contrato;
  zqFolios.Open;

  sMenuP:=stMenu;
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cPlataformas', PopupPrincipal);
  OpcButton := '' ;
  frmbarra1.btnCancel.Click ;
  UtGrid:=TicdbGrid.create(grid_PLATAFORMAS);
  BotonPermiso.permisosBotones(frmBarra1);
  tSeleccionarFecha.DateTime := FechaReporte;

  zq_ConsumosPorEquipo.Active := False;
  zq_ConsumosPorEquipo.ParamByName('Fecha').AsDateTime := FechaReporte;
  zq_ConsumosPorEquipo.ParamByName('Contrato').AsString := param_global_contrato;
  zq_ConsumosPorEquipo.Open;

  //por alguna razon la barra al hacer el primer insert o edit marca error
  frmBarra1btnAddClick(frmBarra1.btnadd);
  frmBarra1btnCancelClick(frmBarra1.btnCancel);
end;

procedure TfrmConsumodeCombustible.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  zq_ConsumosPorEquipo.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmConsumodeCombustible.tSeleccionarFechaExit(Sender: TObject);
begin
  zq_ConsumosPorEquipo.Active := False;
  zq_ConsumosPorEquipo.ParamByName('Fecha').AsDateTime := tSeleccionarFecha.DateTime;
  zq_ConsumosPorEquipo.ParamByName('Contrato').AsString := param_global_contrato;
  zq_ConsumosPorEquipo.Open;
end;

procedure TfrmConsumodeCombustible.tsIdEquipoKeyPress(Sender: TObject;
  var Key: Char);
var
  sDescripcion  : string;
  sTipoVigencia : string;
begin
    if key =#13 then
    begin
        Connection.qryBusca.Active := False;
        Connection.qryBusca.SQL.Clear;
        Connection.qryBusca.SQL.Add('select sIdEquipo, sDescripcion, sMedida, sNumeroSerie from equipos where sContrato = :Contrato and sIdEquipo = :Equipo and lAplicaDiesel = "Si"');
        Connection.qryBusca.Params.ParamByName('Contrato').DataType := ftString;
        Connection.qryBusca.Params.ParamByName('Contrato').Value    := global_contrato_barco;
        Connection.qryBusca.Params.ParamByName('Equipo').DataType   := ftString;
        Connection.qryBusca.Params.ParamByName('Equipo').Value      := Edtbusca.Text;
        Connection.qryBusca.Open;

        if Connection.qryBusca.RecordCount > 0 then
        begin
          Edtbusca.Text := Connection.qryBusca.FieldByName('sIdEquipo').AsString;
          zq_consumosporequipo.FieldByName('sIdEquipo').AsString := Connection.qryBusca.FieldByName('sIdEquipo').AsString;
          zq_consumosporequipo.FieldByName('sDescripcion').AsString := Connection.qryBusca.FieldByName('sDescripcion').AsString;
          zq_consumosporequipo.FieldByName('sMedida').AsString := Connection.qryBusca.FieldByName('sMedida').AsString;
          zq_consumosporequipo.FieldByName('sNumerosDeSerie').AsString := Connection.qryBusca.FieldByName('sNumeroSerie').AsString;
          tsdCantidad.SetFocus;
        end
        else
            if Trim(Edtbusca.Text) <> '' then
            begin
              VCancelar := True;
              sDescripcion := '%' + Trim(Edtbusca.Text) + '%';
              BuscaObjeto.Active := False;
              gridListaObjeto.Columns[0].FieldName := 'sIdEquipo';
              gridListaObjeto.Columns[1].FieldName := 'sDescripcion';
              DsBuscaobjeto.DataSet := BuscaObjeto;

              BuscaObjeto.SQL.Clear;
              BuscaObjeto.SQL.Add('Select * from equipos Where ' +
                'sContrato = :Contrato And (sIdEquipo Like :Descripcion Or sdescripcion Like :Descripcion) and laplicadiesel = "Si" Order by sDescripcion');
              BuscaObjeto.Params.ParamByName('Contrato').DataType := ftString;
              BuscaObjeto.Params.ParamByName('Contrato').Value := global_contrato_barco;
              BuscaObjeto.Params.ParamByName('Descripcion').DataType := ftString;
              BuscaObjeto.Params.ParamByName('Descripcion').Value := sDescripcion;
              BuscaObjeto.Open;

              Panel.Visible := True;
              gridListaObjeto.SetFocus
            end
    end;

end;

procedure TfrmConsumodeCombustible.tsMedidaEnter(Sender: TObject);
begin
    tsMedida.Color := global_color_entrada;
end;

procedure TfrmConsumodeCombustible.tsMedidaExit(Sender: TObject);
begin
    tsMedida.Color := global_color_salida;
end;

procedure TfrmConsumodeCombustible.tsNumeroSerieEnter(Sender: TObject);
begin
    tsNumeroSerie.Color := global_color_entrada;
end;

procedure TfrmConsumodeCombustible.tsNumeroSerieExit(Sender: TObject);
begin
    tsNumeroSerie.Color := global_color_salida;
end;

procedure TfrmConsumodeCombustible.zq_consumosporequipoAfterScroll(
  DataSet: TDataSet);
begin
  EdtBusca.Text := zq_consumosporequipo.FieldByName('sidequipo').AsString;
  Sumatoria(EdtSuma);
end;

procedure TfrmConsumodeCombustible.zq_consumosporequipoCalcFields(
  DataSet: TDataSet);
begin
    if zq_equipos.RecordCount > 0 then
    begin
        connection.QryBusca2.Active := False;
        connection.QryBusca2.SQL.Clear;
        connection.QryBusca2.SQL.Add('select sDescripcion from equipos where sContrato =:contrato and sIdEquipo =:Equipo');
        connection.QryBusca2.ParamByName('contrato').AsString := global_contrato_barco;
        connection.QryBusca2.ParamByName('Equipo').AsString   := zq_consumosporequipo.FieldValues['sIdEquipo'];
        connection.QryBusca2.Open;

        if connection.QryBusca2.RecordCount > 0 then
           zq_consumosporequiposDescripcion.Text := connection.QryBusca2.FieldValues['sDescripcion'];
    end;
end;

procedure TfrmConsumodeCombustible.dblkcbbsIdEquipoEnter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmConsumodeCombustible.dblkcbbsIdEquipoExit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmConsumodeCombustible.frmBarra1btnAddClick(Sender: TObject);
begin
   activapop(frmConsumodeCombustible, popupprincipal);
   zq_recursos.First;
   zqryPlatafomas.First;
   frmBarra1.btnAddClick(Sender);
   PnlPosterior.Enabled := True;
   zq_ConsumosPorEquipo.Append;
   zq_ConsumosPorEquipo.FieldValues['sIdEquipo']      := '';
   zq_ConsumosPorEquipo.FieldValues['sContrato']      := param_global_contrato;
   zq_ConsumosPorEquipo.FieldValues['dIdFecha']       := FechaReporte;
   zq_ConsumosPorEquipo.FieldValues['sIdPlataforma']  := zqryPlatafomas.FieldValues['sIdPlataforma'];
   zq_ConsumosPorEquipo.FieldValues['iIdTipoConsumo'] := zq_recursos.FieldValues['iIdRecursoExistencia'];
   //frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;

   EdtBusca.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;
end;

procedure TfrmConsumodeCombustible.frmBarra1btnEditClick(Sender: TObject);
begin
   activapop(frmConsumodeCombustible, popupprincipal);
   frmBarra1.btnEditClick(Sender);
   PnlPosterior.Enabled := True;
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   sOpcion := 'Edit';
   try
     zq_ConsumosPorEquipo.Edit ;
   except
     on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al agregar registro', 0);
     frmbarra1.btnCancel.Click ;
     end;
   end ;
   EdtBusca.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := False;
end;

procedure TfrmConsumodeCombustible.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
   panel.visible := False;
   PanelMateriales.Visible := False;
    {Continua insercion de datos..}
  
  try
      //desactivapop(popupprincipal);
      zq_ConsumosPorEquipo.Post ;
      PnlPosterior.Enabled := False;
      frmBarra1.btnPostClick(Sender);
  except
      on e : exception do begin
          UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al salvar registro', 0);
          frmbarra1.btnCancel.Click ;
      end;
  end;
  //BotonPermiso.permisosBotones(frmBarra1);

  if sOpcion = 'Edit' then
  begin
      grid_plataformas.Enabled := True;
      sOpcion := '';
  end
  else
     frmBarra1.btnAddClick(Sender);
end;

procedure TfrmConsumodeCombustible.frmBarra1btnCancelClick(Sender: TObject);
begin
   panel.visible := False;
   PanelMateriales.Visible := False;
   desactivapop(popupprincipal);
   frmBarra1.btnCancelClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   zq_ConsumosPorEquipo.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   grid_plataformas.Enabled := True;
   sOpcion := '';
   PnlPosterior.Enabled := False;
   if (zq_consumosporequipo.State = dsBrowse) and (zq_consumosporequipo.RecordCount > 0) then
     EdtBusca.text := zq_consumosporequipo.FieldByName('sidequipo').AsString
   else
     EdtBusca.text := '';
end;

procedure TfrmConsumodeCombustible.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If zq_ConsumosPorEquipo.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        zq_ConsumosPorEquipo.Delete ;
      except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Municipios/Localidades/Plataformas ..', 'Al eliminar registro', 0);
       end;
      end
    end
end;

procedure TfrmConsumodeCombustible.frmBarra1btnRefreshClick(Sender: TObject);
begin
  zq_ConsumosPorEquipo.refresh ;
end;

procedure TfrmConsumodeCombustible.frmBarra1btnExitClick(Sender: TObject);
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


procedure TfrmConsumodeCombustible.gridListaObjetoCellClick(Column: TColumn);
begin
    tsdCantidad.SetFocus;
end;

procedure TfrmConsumodeCombustible.gridListaObjetoDblClick(Sender: TObject);
begin
    tsdCantidad.SetFocus;
    VCancelar := False;
    BtnSeleccionarMat.Click;
end;

procedure TfrmConsumodeCombustible.gridListaObjetoKeyPress(Sender: TObject;
  var Key: Char);
begin
   if key =#13 then
   begin
      VCancelar := False;
      tsdCantidad.SetFocus;
      BtnSeleccionarMat.Click;
   end;
   if Key = #27 then
   begin
     VCancelar := True;
     EdtBusca.SetFocus;
     BtnCancelSel.Click;
   end;
end;

procedure TfrmConsumodeCombustible.grid_plataformasCellClick(Column: TColumn);
begin
  if frmbarra1.btnCancel.Enabled = True then
      frmbarra1.btnCancel.Click ;
end;

procedure TfrmConsumodeCombustible.grid_plataformasMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmConsumodeCombustible.grid_plataformasMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmConsumodeCombustible.grid_plataformasTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmConsumodeCombustible.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click;
end;

procedure TfrmConsumodeCombustible.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click 
end;

procedure TfrmConsumodeCombustible.NuevoMaterialClick(Sender: TObject);
begin
  DsBuscaobjeto.DataSet.Append;
  PanelMateriales.Visible := True;
end;

procedure TfrmConsumodeCombustible.Paste1Click(Sender: TObject);
begin
   UtGrid.AddRowsFromClip;
end;

procedure TfrmConsumodeCombustible.raerconsumosdeldaanterior1Click(
  Sender: TObject);
begin
  Connection.QryBusca.Active := False;
  Connection.QryBusca.SQL.Text := 'SELECT * FROM consumosdecombustibleporequipo WHERE sContrato = :Contrato AND dIdFecha = :Fecha ';
  Connection.QryBusca.Params.ParamByName('Contrato').AsString := param_global_contrato;
  Connection.QryBusca.Params.ParamByName('Fecha').AsDateTime := IncDay(tSeleccionarFecha.DateTime, -1);
  Connection.QryBusca.Open;
  if Connection.QryBusca.RecordCount > 0 then begin
    while Not Connection.QryBusca.Eof do begin
      connection.zCommand.Active := False;
      Connection.zCommand.SQL.Text := 'INSERT INTO consumosdecombustibleporequipo (sIdEquipo, sNumerosDeSerie, iIdTipoConsumo, sMedida, dCantidad, sContrato, dIdFecha, sIdPlataforma) VALUES ' +
                                      '(:sIdEquipo, :NumerosDeSerie, :IdConsumo, :Medida, :Cantidad, :Contrato, :Fecha, :Plataforma) ';
      Connection.zCommand.Params.ParamByName('sIdEquipo').AsString := Connection.QryBusca.FieldByName('sIdEquipo').AsString;
      Connection.zCommand.Params.ParamByName('NumerosDeSerie').AsString := Connection.QryBusca.FieldByName('sNumerosDeSerie').AsString;
      Connection.zCommand.Params.ParamByName('IdConsumo').AsInteger := Connection.QryBusca.FieldByName('iIdTipoConsumo').AsInteger;
      Connection.zCommand.Params.ParamByName('Medida').AsString := Connection.QryBusca.FieldByName('sMedida').AsString;
      Connection.zCommand.Params.ParamByName('Cantidad').AsFloat := Connection.QryBusca.FieldByName('dCantidad').AsFloat;
      Connection.zCommand.Params.ParamByName('Contrato').AsString := Connection.QryBusca.FieldByName('sContrato').AsString;
      Connection.zCommand.Params.ParamByName('Fecha').AsDateTime := tSeleccionarFecha.DateTime;
      Connection.zCommand.Params.ParamByName('Plataforma').AsString := Connection.QryBusca.FieldByName('sIdPlataforma').AsString;
      Connection.zCommand.ExecSQL;
      Connection.QryBusca.Next;
    end;
    zq_ConsumosPorEquipo.Refresh;
  end else begin
    ShowMessage('No se encontraron consumos de equipos reportados el día anterior.');
  end;
end;

procedure TfrmConsumodeCombustible.Copy1Click(Sender: TObject);
begin
    UtGrid.CopyRowsToClip;
end;

procedure TfrmConsumodeCombustible.tsdCantidadEnter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmConsumodeCombustible.tsdCantidadExit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmConsumodeCombustible.tsdCantidadKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    cmbConsumo.SetFocus;
end;

procedure TfrmConsumodeCombustible.tsDescripcionEnter(Sender: TObject);
begin
    tsDescripcion.Color := global_color_entrada;
end;

procedure TfrmConsumodeCombustible.tsDescripcionExit(Sender: TObject);
begin
    tsDescripcion.Color := global_color_salida;
end;

procedure TfrmConsumodeCombustible.cmbConsumoEnter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmConsumodeCombustible.cmbConsumoExit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmConsumodeCombustible.cmbConsumoKeyPress(Sender: TObject;
  var Key: Char);
begin
   if key =#13 then
      cmbPlataforma.SetFocus;
end;

procedure TfrmConsumodeCombustible.cmbPlataformaEnter(Sender: TObject);
begin
      cmbPlataforma.Color := global_color_entrada;
end;

procedure TfrmConsumodeCombustible.cmbPlataformaExit(Sender: TObject);
begin
   cmbPlataforma.Color := global_color_salida;
end;

procedure TfrmConsumodeCombustible.cmbPlataformaKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key =#13 then
       cmbFolios.SetFocus
end;

procedure TfrmConsumodeCombustible.cmbFoliosEnter(Sender: TObject);
begin
    cmbFolios.Color := global_color_entrada;
end;

procedure TfrmConsumodeCombustible.cmbFoliosExit(Sender: TObject);
begin
    cmbFolios.Color := global_color_salida;
end;

procedure TfrmConsumodeCombustible.cmbFoliosKeyPress(Sender: TObject;
  var Key: Char);
begin
    if Key =#13 then
      if zq_consumosporequipo.State = dsBrowse then
          tsNumeroSerie.SetFocus
       else
          frmBarra1btnPostClick(Sender);
end;

procedure TfrmConsumodeCombustible.DBMemo1Enter(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Entrada;
end;

procedure TfrmConsumodeCombustible.DBMemo1Exit(Sender: TObject);
begin
  TEdit(Sender).Color := Global_Color_Salida;
end;

procedure TfrmConsumodeCombustible.ds_CombustiblesStateChange(Sender: TObject);
begin
  if (ds_Combustibles.State = dsInsert) and (frmBarra1.btnPost.Enabled = False) then
    ds_Combustibles.dataset.cancel;
end;

procedure TfrmConsumodeCombustible.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click 
end;

procedure TfrmConsumodeCombustible.EdtBuscaEnter(Sender: TObject);
begin
     EdtBusca.Color := global_color_entrada;
end;

procedure TfrmConsumodeCombustible.EdtBuscaExit(Sender: TObject);
begin
    EdtBusca.Color := global_color_salida;
end;

procedure TfrmConsumodeCombustible.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click
end;

procedure TfrmConsumodeCombustible.BtnAgregarMAtClick(Sender: TObject);
begin
  NuevoMaterial.Click;
end;

procedure TfrmConsumodeCombustible.BtnCancelMatClick(Sender: TObject);
begin
  DsBuscaobjeto.DataSet.Cancel;
  PanelMateriales.Visible := False;
end;

procedure TfrmConsumodeCombustible.BtnCancelSelClick(Sender: TObject);
begin
  Panel.Visible := False;
end;

procedure TfrmConsumodeCombustible.BtnGuardarMatClick(Sender: TObject);
var iitemorden:Integer;
begin
  connection.QryBusca.Active := False;
  connection.QryBusca.SQL.text := 'select ifnull(max(iitemorden)+1,1) as mx from equipos where sContrato = :scontrato';
  connection.QryBusca.Open;
  iitemorden :=  connection.QryBusca.FieldByName('mx').AsInteger;
  DsBuscaobjeto.dataset.FieldByName('sdescripcion').AsString := DsBuscaobjeto.DataSet.FieldByName('sdescripciondiesel').AsString;
  DsBuscaobjeto.dataset.FieldByName('scontrato').AsString := global_contrato_barco;
  DsBuscaobjeto.dataset.FieldByName('sidtipoequipo').AsString  := 'PU';
  DsBuscaobjeto.dataset.FieldByName('lprorrateo').AsString :='Si';
  DsBuscaobjeto.dataset.FieldByName('lcobro').AsString :='Si';
  DsBuscaobjeto.dataset.FieldByName('limprime').AsString :='Si';
  DsBuscaobjeto.dataset.FieldByName('ijornada').AsInteger := 1;
  DsBuscaobjeto.dataset.FieldByName('ldistribuye').AsString  := 'Si';
  DsBuscaobjeto.dataset.FieldByName('lcuadraequipo').AsString := 'No';
  DsBuscaobjeto.dataset.FieldByName('laplicadiesel').AsString := 'Si';
  DsBuscaobjeto.dataset.FieldByName('lsumasolicitado').AsString := 'Si';
  DsBuscaobjeto.dataset.FieldByName('iitemorden').AsInteger := iitemorden;
  DsBuscaobjeto.dataset.FieldByName('dcantidad').AsInteger := 0;
  DsBuscaobjeto.dataset.FieldByName('dcostomn').AsInteger := 0;
  DsBuscaobjeto.dataset.FieldByName('dcostodll').AsInteger := 0;
  DsBuscaobjeto.dataset.FieldByName('dventamn').AsInteger := 0;
  DsBuscaobjeto.dataset.FieldByName('dventadll').AsInteger := 0;
  DsBuscaobjeto.dataset.FieldByName('dfechainicio').AsDateTime := Now;
  DsBuscaobjeto.dataset.FieldByName('dfechafinal').AsDateTime := Now;
  DsBuscaobjeto.DataSet.Post;
  DsBuscaobjeto.DataSet.Refresh;
  PanelMateriales.Visible := False;
end;

procedure TfrmConsumodeCombustible.BtnSeleccionarMatClick(Sender: TObject);
begin
    if (BuscaObjeto.RecordCount > 0) then
    begin
        Edtbusca.Text := BuscaObjeto.FieldByName('sIdEquipo').AsString;
        zq_consumosporequipo.FieldByName('sIdEquipo').AsString :=  BuscaObjeto.FieldByName('sIdEquipo').AsString;
        zq_consumosporequipo.FieldByName('sDescripcion').AsString :=  BuscaObjeto.FieldByName('sDescripcion').AsString;
        zq_consumosporequipo.FieldByName('sMedida').AsString :=  BuscaObjeto.FieldByName('sMedida').AsString;
        zq_consumosporequipo.FieldByName('sNumerosDeSerie').AsString :=  BuscaObjeto.FieldByName('sNumeroSerie').AsString;
    end;
    Panel.Visible := False;
end;

procedure TfrmConsumodeCombustible.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click 
end;

procedure TfrmConsumodeCombustible.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click 
end;

procedure TfrmConsumodeCombustible.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click 
end;

procedure TfrmConsumodeCombustible.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click 
end;

procedure TfrmConsumodeCombustible.Sumatoria(cmp:tedit);
begin
  connection.QryBusca2.Active := False;
  connection.QryBusca2.SQL.Clear;
  connection.QryBusca2.sql.Text := 'SELECT ifnull(sum(dcantidad),0) as suma FROM consumosdecombustibleporequipo WHERE dIdFecha = :Fecha AND sContrato = :Contrato ';
  connection.QryBusca2.ParamByName('Fecha').AsDateTime := tSeleccionarFecha.DateTime;
  connection.QryBusca2.ParamByName('Contrato').AsString := param_global_contrato;
  connection.QryBusca2.Open;

  if connection.QryBusca2.RecordCount = 1 then
    cmP.text := connection.QryBusca2.FieldByName('suma').AsString
  else
    cmp.Text := '';  


end;

end.

