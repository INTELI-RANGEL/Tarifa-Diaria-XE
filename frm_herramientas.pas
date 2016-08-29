unit frm_herramientas;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, global, ComCtrls, ToolWin,
  StdCtrls, ExtCtrls, DBCtrls, Mask, frm_barra, adodb, db, Menus, OleCtrls,
  frxClass, frxDBSet, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  rxToolEdit, RXDBCtrl,udbgrid,unitexcepciones, unittbotonespermisos,
  UnitValidaTexto, UnitExcel, ComObj, UnitTablasImpactadas,unitactivapop,
  DBDateTimePicker, UnitValidacion;

type
  TfrmHerramientas = class(TForm)
    grid_GruposIsometrico: TDBGrid;
    Label2: TLabel;
    Label9: TLabel;
    frmBarra1: TfrmBarra;
    ds_Herramientas: TDataSource;
    tsIdGrupo: TDBEdit;
    tsDescripcion: TDBEdit;
    PopupPrincipal: TPopupMenu;
    Insertar1: TMenuItem;
    Editar1: TMenuItem;
    N1: TMenuItem;
    Registrar1: TMenuItem;
    Can1: TMenuItem;
    N2: TMenuItem;
    Eliminar1: TMenuItem;
    Refresh1: TMenuItem;
    N4: TMenuItem;
    Cut1: TMenuItem;
    Copy1: TMenuItem;
    N3: TMenuItem;
    Salir1: TMenuItem;
    herramientas: TZQuery;
    mnGeneraGrupos: TMenuItem;
    Label1: TLabel;
    Label3: TLabel;
    tdVentaMN: TDBEdit;
    Label4: TLabel;
    Label5: TLabel;
    tdVentaDLL: TDBEdit;
    Label6: TLabel;
    dbCostomn: TDBEdit;
    dbCostoDLL: TDBEdit;
    dbMedida: TDBEdit;
    Label7: TLabel;
    tdSimbolo: TDBEdit;
    herramientassContrato: TStringField;
    herramientassIdHerramientas: TStringField;
    herramientassDescripcion: TStringField;
    herramientassMedida: TStringField;
    herramientasdCostoMN: TFloatField;
    herramientasdCostoDLL: TFloatField;
    herramientasdVentaMN: TFloatField;
    herramientasdVentaDLL: TFloatField;
    herramientassSimbolo: TStringField;
    Label8: TLabel;
    herramientasfFecha: TDateField;
    ExportaaPlantillaExcel1: TMenuItem;
    SaveDialog1: TSaveDialog;
    tdFecha: TDBDateTimePicker;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure grid_GruposIsometricoCellClick(Column: TColumn);
    procedure tsIdPersonalKeyPress(Sender: TObject; var Key: Char);
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
    procedure grid_GruposIsometricoEnter(Sender: TObject);
    procedure grid_GruposIsometricoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grid_GruposIsometricoKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure tsIdGrupoEnter(Sender: TObject);
    procedure tsIdGrupoExit(Sender: TObject);
    procedure tsIdGrupoKeyPress(Sender: TObject; var Key: Char);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure tsDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure tlFaseKeyPress(Sender: TObject; var Key: Char);
    procedure Imprimir1Click(Sender: TObject);
    procedure mnGeneraGruposClick(Sender: TObject);
    procedure tdCostomnEnter(Sender: TObject);
    procedure tdCostomnExit(Sender: TObject);
    procedure tdCostomnKeyPress(Sender: TObject; var Key: Char);
    procedure tdCostoDLLEnter(Sender: TObject);
    procedure tdCostoDLLExit(Sender: TObject);
    procedure dbCostomnEnter(Sender: TObject);
    procedure dbCostomnExit(Sender: TObject);
    procedure dbCostoDLLEnter(Sender: TObject);
    procedure dbCostoDLLExit(Sender: TObject);
    procedure dbMedidaEnter(Sender: TObject);
    procedure dbMedidaExit(Sender: TObject);
    procedure tdSimboloEnter(Sender: TObject);
    procedure tdSimboloExit(Sender: TObject);
    procedure dbCostomnKeyPress(Sender: TObject; var Key: Char);
    procedure dbCostoDLLKeyPress(Sender: TObject; var Key: Char);
    procedure dbMedidaKeyPress(Sender: TObject; var Key: Char);
    procedure tdSimboloKeyPress(Sender: TObject; var Key: Char);
    procedure tdVentaDLLKeyPress(Sender: TObject; var Key: Char);
    procedure tdVentaMNKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaKeyPress(Sender: TObject; var Key: Char);
    procedure tdFechaEnter(Sender: TObject);
    procedure tdFechaExit(Sender: TObject);
    procedure grid_GruposIsometricoMouseMove(Sender: TObject;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_GruposIsometricoMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure grid_GruposIsometricoTitleClick(Column: TColumn);
    procedure Cut1Click(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure formatoEncabezado();
    procedure ExportaaPlantillaExcel1Click(Sender: TObject);
    function tablasDependientes(idOrig: string): boolean;
    function posibleBorrar(idOrig: string): boolean;
    procedure herramientasdCostoMNSetText(Sender: TField; const Text: string);
    procedure herramientasdCostoDLLSetText(Sender: TField; const Text: string);
    procedure herramientasdVentaMNSetText(Sender: TField; const Text: string);
    procedure herramientasdVentaDLLSetText(Sender: TField; const Text: string);
    procedure dbCostomnChange(Sender: TObject);
    procedure dbCostoDLLChange(Sender: TObject);
    procedure tdVentaMNChange(Sender: TObject);
    procedure tdVentaDLLChange(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmHerramientas : TfrmHerramientas;
  utdbgrid:ticdbgrid;
  botonpermiso: tbotonespermisos;
  sOpcion : string;

  //Exporta elementos a Excel..
  Excel, Libro, Hoja: Variant;
  sIdOrig : string;

implementation

{$R *.dfm}

procedure TfrmHerramientas.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := cafree ;
  utdbgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmHerramientas.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'Herramientas1', PopupPrincipal);
  OpcButton := '' ;
  sIdOrig := '';
  frmbarra1.btnCancel.Click ;
  herramientas.Active := False ;
  herramientas.Params.ParamByName('Contrato').DataType := ftString ;
  herramientas.Params.ParamByName('Contrato').Value := Global_Contrato ;
  herramientas.Open ;
  Grid_GruposIsometrico.SetFocus;
  UtdbGrid:=TicdbGrid.create(grid_gruposisometrico);
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
end;
procedure TfrmHerramientas.grid_GruposIsometricoCellClick(Column: TColumn);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure TfrmHerramientas.tsIdPersonalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsDescripcion.SetFocus ;
end;

procedure TfrmHerramientas.frmBarra1btnAddClick(Sender: TObject);
begin
  frmBarra1.btnAddClick(Sender);
  Insertar1.Enabled := False ;
  Editar1.Enabled := False ;
  Registrar1.Enabled := True ;
  Can1.Enabled := True ;
  Eliminar1.Enabled := False ;
  Refresh1.Enabled := False ;
  Salir1.Enabled := False ;
  herramientas.Append ;
  activapop(frmHerramientas,popupprincipal);
  herramientas.FieldValues['sIdHerramientas'] := '' ;
  herramientas.FieldValues['sDescripcion'] := '' ;
  herramientas.FieldValues['dVentaMN']     := '0' ;
  herramientas.FieldValues['dVentaDLL']    := '0' ;
  herramientas.FieldValues['dCostoMN']     := '0' ;
  herramientas.FieldValues['dCostoDLL']    := '0' ;
  herramientas.FieldValues['sSimbolo']     := '*' ;
  herramientas.FieldValues['fFecha']       := date ;
  tsIdGrupo.SetFocus ;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
  grid_gruposisometrico.Enabled := False;
end;

procedure TfrmHerramientas.frmBarra1btnEditClick(Sender: TObject);
begin
    If herramientas.RecordCount > 0 Then
    Begin
       TRY
       sIdOrig := herramientas.FieldByName('sIdHerramientas').AsString;
       frmBarra1.btnEditClick(Sender);
       Insertar1.Enabled := False ;
       Editar1.Enabled := False ;
       Registrar1.Enabled := True ;
       Can1.Enabled := True ;
       Eliminar1.Enabled := False ;
       Refresh1.Enabled := False ;
       Salir1.Enabled := False ;
       sOpcion := 'Edit';
       herramientas.Edit ;
       activapop(frmHerramientas,popupprincipal);
       tsIdGrupo.SetFocus;
          EXCEPT
           on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Herramientas', 'Al agregar registro', 0);
           end;
       END;
    End;
BotonPermiso.permisosBotones(frmBarra1);
frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmHerramientas.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
  lEdita: boolean;
begin
  {Validaciones de campos}
  nombres:=TStringList.Create;cadenas:=TStringList.Create;
  nombres.Add('Herramienta');  nombres.Add('Descripcion');  nombres.Add('Costo MN');
  cadenas.Add(tsIdGrupo.Text); cadenas.Add(tsDescripcion.Text); cadenas.Add(dbCostoMN.Text);

  nombres.Add('Costo DLL');     nombres.Add('Medida'); 
  cadenas.Add(dbCostoDLL.Text); cadenas.Add(dbMedida.Text);

  nombres.Add('Venta MN');     nombres.Add('Venta DLL');    nombres.Add('Simbolo');
  cadenas.Add(tdVentaMN.Text); cadenas.Add(tdVentaDLL.Text); cadenas.Add(tdSimbolo.Text);

  if not validaTexto(nombres, cadenas, 'Herramienta',tsIdGrupo.Text) then
  begin
     MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
     exit;
  end;

  {Continua insercion de datos..}

  lEdita := false;
  if herramientas.State = dsEdit then
    lEdita := true;

  Try
     herramientas.FieldValues['sContrato'] := global_contrato ;
     herramientas.FieldValues['fFecha'] := tdfecha.date;
     herramientas.Post ;
     Insertar1.Enabled := True ;
     Editar1.Enabled := True ;
     Registrar1.Enabled := False ;
     Can1.Enabled := False ;
     Eliminar1.Enabled := True ;
     Refresh1.Enabled := True ;
     desactivapop(popupprincipal);
     Salir1.Enabled := True ;
     frmBarra1.btnPostClick(Sender);
  except
      on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Herramientas', 'Al salvar registro', 0);
           frmBarra1.btnCancel.Click ;
           lEdita := false;//cancelar la actualizacion de tablas dependientes
      end;
  end;
   if (lEdita) and (herramientas.FieldByName('sIdHerramientas').AsString <> sIdOrig) then
   begin
       tablasDependientes(sIdOrig);
   end;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
  if sOpcion =  'Edit' then
  begin
      grid_gruposisometrico.Enabled := True;
      sOpcion := '';
  end;
end;

procedure TfrmHerramientas.frmBarra1btnCancelClick(Sender: TObject);
begin
   frmBarra1.btnCancelClick(Sender);
   herramientas.Cancel ;
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   desactivapop(popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled  := False;
   grid_gruposisometrico.Enabled := True;
   sOpcion := '';
end;

procedure TfrmHerramientas.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If herramientas.RecordCount > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if not posibleBorrar(herramientas.FieldByName('sIdHerramientas').AsString) then
      begin
        MessageDlg('No es posible eliminar el registro, existen registros dependientes.', mtInformation, [mbOk], 0);
        exit;
      end;
      try
          herramientas.Delete ;
      except
         on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Herramientas', 'Al eliminar registro', 0);
         end;
      end
    end
end;

procedure TfrmHerramientas.frmBarra1btnRefreshClick(Sender: TObject);
begin
  herramientas.refresh ;
end;

procedure TfrmHerramientas.frmBarra1btnExitClick(Sender: TObject);
begin
   frmBarra1.btnExitClick(Sender);
   Insertar1.Enabled := True ;
   Editar1.Enabled := True ;
   Registrar1.Enabled := False ;
   Can1.Enabled := False ;
   Eliminar1.Enabled := True ;
   Refresh1.Enabled := True ;
   Salir1.Enabled := True ;
   Close
end;

procedure TfrmHerramientas.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure TfrmHerramientas.Copy1Click(Sender: TObject);
begin
    UtdbGrid.AddRowsFromClip;
end;

procedure TfrmHerramientas.Cut1Click(Sender: TObject);
begin
  UtdbGrid.CopyRowsToClip;
end;

procedure TfrmHerramientas.dbCostoDLLChange(Sender: TObject);
begin
  tdbEditChangef(dbCostoDLL, 'Costo DLL');
end;

procedure TfrmHerramientas.dbCostoDLLEnter(Sender: TObject);
begin
   dbCostoDLL.Color := global_color_entrada ;
end;

procedure TfrmHerramientas.dbCostoDLLExit(Sender: TObject);
begin
   dbCostoDLL.Color := global_color_salida ;
end;

procedure TfrmHerramientas.dbCostoDLLKeyPress(Sender: TObject; var Key: Char);
begin
if not keyFiltroTdbedit(dbCostoDLL, key) then
  key:=#0;
    If Key = #13 then
        dbMedida.SetFocus
end;

procedure TfrmHerramientas.dbCostomnChange(Sender: TObject);
begin
  tdbeditChangef(dbCostomn, 'Costo MN');
end;

procedure TfrmHerramientas.dbCostomnEnter(Sender: TObject);
begin
   dbCostomn.Color := global_color_entrada ;
end;

procedure TfrmHerramientas.dbCostomnExit(Sender: TObject);
begin
   dbCostomn.Color := global_color_salida ;
end;

procedure TfrmHerramientas.dbCostomnKeyPress(Sender: TObject; var Key: Char);
begin
    if not keyFiltroTdbedit(dbCostomn, key) then
    key:=#0;
    
    If Key = #13 then
        dbCostoDLL.SetFocus
end;

procedure TfrmHerramientas.tdSimboloEnter(Sender: TObject);
begin
    tdSimbolo.Color := global_color_entrada
end;

procedure TfrmHerramientas.tdSimboloExit(Sender: TObject);
begin
    tdSimbolo.Color := global_color_salida
end;

procedure TfrmHerramientas.tdSimboloKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 then
        tdFecha.SetFocus
end;

procedure TfrmHerramientas.tdVentaDLLChange(Sender: TObject);
begin
  tdbeditchangef(tdVentaDLL, 'Venta DLL');
end;

procedure TfrmHerramientas.tdVentaDLLKeyPress(Sender: TObject; var Key: Char);
begin
if not keyFiltroTdbedit(tdVentaDLL, key) then
  key:=#0;

    If Key = #13 Then
        tdSimbolo.SetFocus
end;

procedure TfrmHerramientas.tdVentaMNChange(Sender: TObject);
begin
  tdbeditchangef(tdVentaMN,'Venta MN');
end;

procedure TfrmHerramientas.tdVentaMNKeyPress(Sender: TObject; var Key: Char);
begin
  if not keyFiltroTdbedit(tdVentaMN, key) then
    key:=#0;
  if key = #13 then
    tdVentaDLL.SetFocus ;
end;

procedure TfrmHerramientas.dbMedidaEnter(Sender: TObject);
begin
   dbMedida.Color := global_color_entrada ;
end;

procedure TfrmHerramientas.dbMedidaExit(Sender: TObject);
begin
   dbMedida.Color := global_color_salida ;
end;

procedure TfrmHerramientas.dbMedidaKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 then
        tdVentaMN.SetFocus
end;

procedure TfrmHerramientas.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click
end;

procedure TfrmHerramientas.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click
end;

procedure TfrmHerramientas.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click
end;

procedure TfrmHerramientas.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click
end;

procedure TfrmHerramientas.ExportaaPlantillaExcel1Click(Sender: TObject);
Var
  CadError, OrdenVigencia: String;
//////////////////////////////////// PLANTILAS DE IMPORTACION //////////////////
Function GenerarPlantilla: Boolean;
Var
  Resultado: Boolean;

Procedure DatosPlantilla;
Var
  CadFecha, tmpNombre, cadena : String;
  fs: tStream;
  Alto : Extended;
  Ren, nivel : integer;
Begin
    Ren := 2;
  // Realizar los ajustes visuales y de formato de hoja
    Excel.ActiveWindow.Zoom := 100;
//  if rAnexoC.Checked then
//  begin
      Excel.Columns['A:A'].ColumnWidth := 20;
      Excel.Columns['B:B'].ColumnWidth := 10;
      Excel.Columns['C:C'].ColumnWidth := 40;
      Excel.Columns['D:H'].ColumnWidth := 12;

      // Colocar los encabezados de la plantilla...
      Hoja.Range['A1:A1'].Select;
      Excel.Selection.Value := 'Contrato';
      FormatoEncabezado;
      Hoja.Range['B1:B1'].Select;
      Excel.Selection.Value := 'Id_Herramienta';
      FormatoEncabezado;
      Hoja.Range['C1:C1'].Select;
      Excel.Selection.Value := 'Descripcion';
      FormatoEncabezado;
      Hoja.Range['D1:D1'].Select;
      Excel.Selection.Value := 'Medida';
      FormatoEncabezado;
      Hoja.Range['E1:E1'].Select;
      Excel.Selection.Value := 'Costo MN';
      FormatoEncabezado;
      Hoja.Range['F1:F1'].Select;
      Excel.Selection.Value := 'Costo DLL';
      FormatoEncabezado;
      Hoja.Range['G1:G1'].Select;
      Excel.Selection.Value := 'Venta DLL';
      FormatoEncabezado;
      Hoja.Range['H1:H1'].Select;
      Excel.Selection.Value := 'Venta MN';
      FormatoEncabezado;

      connection.QryBusca.Active := False ;
      Connection.QryBusca.Filtered := False;
      connection.QryBusca.SQL.Clear ;
      connection.QryBusca.SQL.Add('select * from herramientas where sContrato =:Contrato order by sIdHerramientas ');
      connection.QryBusca.Params.ParamByName('Contrato').DataType := ftString ;
      connection.QryBusca.Params.ParamByName('Contrato').Value    := global_contrato ;
      connection.QryBusca.Open ;

      if connection.QryBusca.RecordCount > 0 then
      begin
           while not connection.QryBusca.Eof do
           begin
                Hoja.Cells[Ren,1].Select;
                Excel.Selection.Value := global_contrato;
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;
                Excel.Selection.Font.Size := 11;
                Excel.Selection.Font.Bold := False;
                Excel.Selection.Font.Name := 'Calibri';

                Hoja.Cells[Ren,2].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['sIdHerramientas'];
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;

                Hoja.Cells[Ren,3].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['sDescripcion'];
                Alto := Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight;
                Hoja.Cells[Ren,3].Value := '';

                if Alto > 15 then
                   Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := Alto
                Else
                   Excel.Rows[IntToStr(Ren) + ':' + IntToStr(Ren)].RowHeight := 15;

                Excel.Selection.Value := connection.QryBusca.FieldValues['sDescripcion'];

                Hoja.Cells[Ren,4].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['sMedida'];
                Excel.Selection.HorizontalAlignment := xlCenter;
                Excel.Selection.VerticalAlignment   := xlCenter;

                Hoja.Cells[Ren,5].Select;
                Excel.Selection.NumberFormat := '@';
                Excel.Selection.Value := connection.QryBusca.FieldValues['dCostoMN'];
                Excel.Selection.HorizontalAlignment := xlRight;
                Excel.Selection.VerticalAlignment   := xlCenter;

                Hoja.Cells[Ren,6].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['dCostoDLL'];

                Hoja.Cells[Ren,7].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['dVentaMN'];

                Hoja.Cells[Ren,8].Select;
                Excel.Selection.Value := connection.QryBusca.FieldValues['dVentaDLL'];

                connection.QryBusca.Next;
                Inc(Ren);
           end;
      end;
      Hoja.Cells[2,2].Select;


  Hoja.Range['A1:H1'].Select;
  // Formato general de encabezado de datos..
  Excel.Selection.HorizontalAlignment                   := xlCenter;
  Excel.Selection.VerticalAlignment                     := xlCenter;
  Excel.Selection.Interior.ColorIndex := 5;
  Excel.Selection.Font.color          := clWhite;
  Excel.Selection.Interior.Pattern    := xlSolid;

  Hoja.Range['A1:A1'].Select;
End;

Begin
  Resultado := True;
  Try
    Hoja := Libro.Sheets[1];
    Hoja.Select;
    try
       Hoja.Name := 'HERRAMIENTA '+ global_contrato;
    Except
       Hoja.Name := 'HERRAMIENTA '+ global_contrato;
    end;
    DatosPlantilla;
  Except
    on e:exception do
    Begin
      Resultado := False;
      CadError := 'Se ha producido el siguiente error al generar la Plantilla de Herramientas' + #10 + #10 + e.Message;
    End;
  End;

  Result := Resultado;
End;

begin
  // Solicitarle al usuario la ruta del archivo en donde desea grabar el reporte
  If Not SaveDialog1.Execute Then
    Exit;

  // Generar el ambiente de excel
  try
    Excel := CreateOleObject('Excel.Application');
  except
    FreeAndNil(Excel);
    showmessage('No es posible generar el ambiente de EXCEL, informe de esto al administrador del sistema.');
    Exit;
  end;

  Excel.Visible := True;
  Excel.DisplayAlerts := False;
  Excel.ScreenUpdating := True;

  Libro := Excel.Workbooks.Add;    // Crear el libro sobre el que se ha de trabajar

  // Verificar si cuenta con las hojas necesarias
  while Libro.Sheets.Count < 2 do
    Libro.Sheets.Add;

  // Verificar si se pasa de hojas necesarias
  Libro.Sheets[1].Select;
  while Libro.Sheets.Count > 1 do
    Excel.ActiveWindow.SelectedSheets.Delete;

  // Proceder a generar la hoja REPORTE
  CadError := '';

  if GenerarPlantilla then
    // Grabar el archivo de excel con el nombre dado
    messageDlg('La Plantilla se Genero Correctamente!' , mtConfirmation, [mbOk], 0) ;

      Excel := '';

  if CadError <> '' then
    showmessage(CadError);

end;

procedure TfrmHerramientas.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmHerramientas.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

procedure TfrmHerramientas.grid_GruposIsometricoEnter(Sender: TObject);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure TfrmHerramientas.grid_GruposIsometricoKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure TfrmHerramientas.grid_GruposIsometricoKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure TfrmHerramientas.grid_GruposIsometricoMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtdbGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmHerramientas.grid_GruposIsometricoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
   UtdbGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmHerramientas.grid_GruposIsometricoTitleClick(Column: TColumn);
begin
  UtdbGrid.DbGridTitleClick(Column);
end;

procedure TfrmHerramientas.herramientasdCostoDLLSetText(Sender: TField;
  const Text: string);
begin
  sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmHerramientas.herramientasdCostoMNSetText(Sender: TField;
  const Text: string);
begin
  sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmHerramientas.herramientasdVentaDLLSetText(Sender: TField;
  const Text: string);
begin
  sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmHerramientas.herramientasdVentaMNSetText(Sender: TField;
  const Text: string);
begin
  sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmHerramientas.tsIdGrupoEnter(Sender: TObject);
begin
    tsIdGrupo.Color := global_color_entrada
end;

procedure TfrmHerramientas.tsIdGrupoExit(Sender: TObject);
begin
    tsIdGrupo.Color := global_color_salida
end;

procedure TfrmHerramientas.tsIdGrupoKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsDescripcion.SetFocus
end;

procedure TfrmHerramientas.tsDescripcionEnter(Sender: TObject);
begin
    tsDescripcion.Color := global_color_entrada
end;

procedure TfrmHerramientas.tsDescripcionExit(Sender: TObject);
begin
    tsDescripcion.Color := global_color_salida
end;

procedure TfrmHerramientas.tsDescripcionKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 then
        dbCostomn.SetFocus
end;

function TfrmHerramientas.tablasDependientes(idOrig: string): boolean;
var
  ParamNamesSET,ParamValuesSET,ParamNamesWHERE,ParamValuesWHERE: TStringList;
begin
  result := true;
  ParamNamesSET:=TStringList.Create;ParamValuesSET:=TStringList.Create;ParamNamesWHERE:=TStringList.Create;ParamValuesWHERE:=TStringList.Create;
  ParamNamesSET.Add('sIdHerramientas');ParamValuesSET.Add(Herramientas.FieldByName('sIdHerramientas').AsString);
  ParamNamesWHERE.Add('sContrato');ParamValuesWHERE.Add(global_contrato);
  ParamNamesWHERE.Add('sIdHerramientas');ParamValuesWHERE.Add(idOrig);
  if not UnitTablasImpactadas.impactar('herramientas',ParamNamesSET,ParamValuesSET,ParamNamesWHERE,ParamValuesWHERE) then
  begin
    result := false;
    showmessage('Ocurrio un error al actualizar las tablas dependientes: ' + UnitTablasImpactadas.xError);
  end;
end;

function TfrmHerramientas.posibleBorrar(idOrig: string): boolean;
var
  ParamNamesWHERE,ParamValuesWHERE: TStringList;
begin
  result := true;
  ParamNamesWHERE:=TStringList.Create;ParamValuesWHERE:=TStringList.Create;
  ParamNamesWHERE.Add('sContrato');ParamValuesWHERE.Add(global_contrato);
  ParamNamesWHERE.Add('sIdHerramientas');ParamValuesWHERE.Add(idOrig);
  result := not UnitTablasImpactadas.hayDependientes('herramientas',ParamNamesWHERE,ParamValuesWHERE);
end;

procedure TfrmHerramientas.tdCostoDLLEnter(Sender: TObject);
begin
   tdVentaDLL.Color := global_color_entrada ;
end;

procedure TfrmHerramientas.tdCostoDLLExit(Sender: TObject);
begin
   tdVentaDLL.Color := global_color_salida ;
end;

procedure TfrmHerramientas.tdCostomnEnter(Sender: TObject);
begin
   tdVentaMn.Color := global_color_entrada ;
end;

procedure TfrmHerramientas.tdCostomnExit(Sender: TObject);
begin
  tdVentaMn.Color := global_color_salida ;
end;

procedure TfrmHerramientas.tdCostomnKeyPress(Sender: TObject; var Key: Char);
begin
  if key = #13 then
    tdVentaDLL.SetFocus ;
end;

procedure TfrmHerramientas.tdFechaEnter(Sender: TObject);
begin
    tdFecha.Color := global_color_entrada
end;

procedure TfrmHerramientas.tdFechaExit(Sender: TObject);
begin
    tdFecha.Color := global_color_salida
end;

procedure TfrmHerramientas.tdFechaKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 then
        tsIdGrupo.SetFocus
end;

procedure TfrmHerramientas.tlFaseKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 then
        tsIdGrupo.SetFocus
end;

procedure TfrmHerramientas.Imprimir1Click(Sender: TObject);
begin
    frmBarra1.btnPrinter.Click
end;

procedure TfrmHerramientas.mnGeneraGruposClick(Sender: TObject);
begin
    connection.QryBusca.Active := False ;
    connection.QryBusca.SQL.Clear ;
    connection.QryBusca.SQL.Add('select sNumeroActividad, mDescripcion from actividadesxanexo ' +
                                'Where sContrato = :Contrato and sIdConvenio = :Convenio and sTipoActividad = "Paquete" and iNivel= 1 Order By iItemOrden') ;
    connection.QryBusca.Params.ParamByName('contrato').DataType := ftString ;
    connection.QryBusca.Params.ParamByName('contrato').Value := global_contrato ;
    connection.QryBusca.Params.ParamByName('convenio').DataType := ftString ;
    connection.QryBusca.Params.ParamByName('convenio').Value := '' ;
    connection.QryBusca.Open ;
    While NOT connection.QryBusca.Eof Do
    Begin
          Try
              connection.zCommand.Active := False ;
              connection.zCommand.SQL.Clear ;
              connection.zCommand.SQL.Add('insert into gruposisometrico values (:Contrato, :Grupo, :Descripcion)') ;
              connection.zCommand.Params.ParamByName('contrato').DataType := ftString ;
              connection.zCommand.Params.ParamByName('contrato').Value := global_contrato ;
              connection.zCommand.Params.ParamByName('grupo').DataType := ftString ;
              connection.zCommand.Params.ParamByName('grupo').Value := connection.QryBusca.FieldValues['sNumeroActividad'] ;
              connection.zCommand.Params.ParamByName('descripcion').DataType := ftString ;
              connection.zCommand.Params.ParamByName('descripcion').Value := connection.QryBusca.FieldValues['mDescripcion'] ;
              connection.zCommand.ExecSQL ;
          except
          End ;
          connection.QryBusca.Next ;
    End ;
    herramientas.Active := False ;
    herramientas.Open ;

end;

procedure TfrmHerramientas.formatoEncabezado;
begin
      Excel.Selection.MergeCells := False;
      Excel.Selection.HorizontalAlignment := xlCenter;
      Excel.Selection.VerticalAlignment := xlCenter;
      Excel.Selection.Font.Size := 12;
      Excel.Selection.Font.Bold := False;
      Excel.Selection.Font.Name := 'Calibri';
end;

end.
