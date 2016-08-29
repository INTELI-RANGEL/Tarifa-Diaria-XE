unit frm_reaperturadefolios;

interface
             
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, DBGrids, frm_connection, global, ComCtrls, ToolWin,
  StdCtrls, ExtCtrls, DBCtrls, Mask, frm_barra, adodb, db, Menus, OleCtrls,
  frxClass, frxDBSet, ZAbstractRODataset, ZAbstractDataset, ZDataset,
  udbgrid, unitexcepciones, unittbotonespermisos, UnitValidaTexto,
  unitactivapop, rxToolEdit, rxCurrEdit;

type
  Tfrmreaperturadefolios = class(TForm)
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
    qryestatus: TZQuery;
    ds_estatus: TDataSource;
    Panel1: TPanel;
    Panel2: TPanel;
    Label2: TLabel;
    tsDescripcion: TDBEdit;
    frmBarra1: TfrmBarra;
    grid_estatus: TDBGrid;
    tsFolioReabrir: TDBEdit;
    Label1: TLabel;
    Label3: TLabel;
    RxCedtAvance: TRxCalcEdit;
    Label4: TLabel;
    RxCalcEdit1: TRxCalcEdit;
    RxCalPorcentaje: TRxCalcEdit;
    qryestatusiIdFolio: TIntegerField;
    qryestatussFolioNuevo: TStringField;
    qryestatussFolioReabierto: TStringField;
    qryestatusdGlobalFolioReabrir: TFloatField;
    qryestatusdPorcentajeRestante: TFloatField;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure grid_estatusCellClick(Column: TColumn);
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
    procedure grid_estatusEnter(Sender: TObject);
    procedure grid_estatusKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure grid_estatusKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure tsIdGrupoKeyPress(Sender: TObject; var Key: Char);
    procedure tsDescripcionEnter(Sender: TObject);
    procedure tsDescripcionExit(Sender: TObject);
    procedure grid_estatusMouseMove(Sender: TObject;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_estatusMouseUp(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure grid_estatusTitleClick(Column: TColumn);
    procedure Cut1Click(Sender: TObject);
    procedure Copy1Click(Sender: TObject);
    procedure tsDescripcionKeyPress(Sender: TObject; var Key: Char);
    procedure tiColoresEnter(Sender: TObject);
    procedure tiColoresExit(Sender: TObject);
    procedure qryestatusAfterScroll(DataSet: TDataSet);
    procedure qryestatusAfterInsert(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmreaperturadefolios : Tfrmreaperturadefolios;
  utgrid           : ticdbgrid;
  sOldId           : string;
  botonpermiso     : tbotonespermisos;
  sOpcion          : string;

  implementation
uses
    frm_consumibles;
{$R *.dfm}

procedure Tfrmreaperturadefolios.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
  end;

procedure Tfrmreaperturadefolios.FormShow(Sender: TObject);
begin
   BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'MnuFamiliadePro', PopupPrincipal);
   OpcButton := '' ;
   frmbarra1.btnCancel.Click ;

   qryestatus.Active := False ;
   qryestatus.Open ;

   grid_estatus.SetFocus;
   UtGrid:=TicdbGrid.create(grid_estatus);

   BotonPermiso.permisosBotones(frmBarra1);
   frmbarra1.btnPrinter.Enabled := False;

   //Showmessage( grid_estatus.Columns.Items[3].)
end;
procedure Tfrmreaperturadefolios.grid_estatusCellClick(Column: TColumn);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;

   RxCalPorcentaje.Value := qryestatus.FieldByName('dPorcentajeRestante').AsFloat;
   RxCedtAvance.Value := qryestatus.FieldByName('dGlobalFolioReabrir').AsFloat;
end;

procedure Tfrmreaperturadefolios.tsIdPersonalKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tsDescripcion.SetFocus ;
end;

procedure Tfrmreaperturadefolios.frmBarra1btnAddClick(Sender: TObject);
begin
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   qryestatus.Append ;
   qryestatus.FieldValues['sFolioNuevo'] := '';
   tsDescripcion.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   frmbarra1.btnPrinter.Enabled := False;
   grid_estatus.Enabled := False;

end;

procedure Tfrmreaperturadefolios.frmBarra1btnEditClick(Sender: TObject);
begin
    If qryestatus.RecordCount > 0 Then
    Begin
        try
           frmBarra1.btnEditClick(Sender);
           Insertar1.Enabled := False ;
           Editar1.Enabled := False ;
           Registrar1.Enabled := True ;
           Can1.Enabled := True ;
           Eliminar1.Enabled := False ;
           Refresh1.Enabled := False ;
           Salir1.Enabled := False ;
           sOpcion := 'Edit';
           tsDescripcion.SetFocus;
           qryEstatus.Edit;
           grid_estatus.Enabled := False;
        except
           on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Clasificacion de Familia de Materiales', 'Al agregar registro', 0);
           end;
        end;
    End;
   BotonPermiso.permisosBotones(frmBarra1);
   frmbarra1.btnPrinter.Enabled := False;
end;

procedure Tfrmreaperturadefolios.frmBarra1btnPostClick(Sender: TObject);
var
    lEdicion : boolean;
begin
    frmBarra1.btnPostClick(Sender);
    if (trim(tsDescripcion.Text) = '') and (Trim(tsFolioReabrir.Text) = '')then
    begin
       MessageDlg('La descripcion debe tener un valor!', mtInformation, [mbOk], 0);
       exit;
    end;

    qryestatus.FieldByName('dglobalfolioreabrir').Asfloat := RxCedtAvance.Value;
    qryestatus.FieldByName('dporcentajerestante').Asfloat := RxCalPorcentaje.Value;
    qryestatus.Post;

    Insertar1.Enabled  := True ;
    Editar1.Enabled    := True ;
    Registrar1.Enabled := False ;
    Can1.Enabled       := False ;
    Eliminar1.Enabled  := True ;
    Refresh1.Enabled   := True ;
    Salir1.Enabled     := True ;
    grid_estatus.Enabled := True;


    desactivapop(popupprincipal);

end;

procedure Tfrmreaperturadefolios.frmBarra1btnCancelClick(Sender: TObject);
begin

   frmBarra1.btnCancelClick(Sender);
   qryestatus.Cancel;

//   Insertar1.Enabled := True ;
//   Editar1.Enabled := True ;
//   Registrar1.Enabled := False ;
//   Can1.Enabled := False ;
//   Eliminar1.Enabled := True ;
//   Refresh1.Enabled := True ;
//   Salir1.Enabled := True ;
   desactivapop(popupprincipal);
   BotonPermiso.permisosBotones(frmBarra1);
   frmbarra1.btnPrinter.Enabled := False;
   grid_estatus.Enabled := True;
   sOpcion := '';
end;

procedure Tfrmreaperturadefolios.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If qryestatus.RecordCount > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
     try
         qryestatus.Delete;
      except
         on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'Catalogo de Estatus de Empleados', 'Al eliminar registro', 0);
         end;
      end
    end
end;

procedure Tfrmreaperturadefolios.frmBarra1btnRefreshClick(Sender: TObject);
begin
    qryestatus.refresh ;
end;

procedure Tfrmreaperturadefolios.frmBarra1btnExitClick(Sender: TObject);
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

procedure Tfrmreaperturadefolios.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure Tfrmreaperturadefolios.qryestatusAfterInsert(DataSet: TDataSet);
begin

   // qryestatus.FieldValues['acceso'] := 'true' ;
end;

procedure Tfrmreaperturadefolios.qryestatusAfterScroll(DataSet: TDataSet);
begin
    if qryestatus.RecordCount > 0 then
    begin
        if (qryestatus.State <> dsInsert) then
         // tiColores.ItemIndex := qryestatus.FieldValues['iColor'];
    end;
    RxCalPorcentaje.Value := qryestatus.FieldByName('dPorcentajeRestante').AsFloat;
    RxCedtAvance.Value := qryestatus.FieldByName('dGlobalFolioReabrir').AsFloat;
end;

procedure Tfrmreaperturadefolios.Copy1Click(Sender: TObject);
begin
UtGrid.AddRowsFromClip;
end;

procedure Tfrmreaperturadefolios.Cut1Click(Sender: TObject);
begin
  UtGrid.CopyRowsToClip;
end;

procedure Tfrmreaperturadefolios.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click
end;

procedure Tfrmreaperturadefolios.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click
end;

procedure Tfrmreaperturadefolios.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click
end;

procedure Tfrmreaperturadefolios.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click
end;

procedure Tfrmreaperturadefolios.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure Tfrmreaperturadefolios.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

procedure Tfrmreaperturadefolios.grid_estatusEnter(Sender: TObject);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure Tfrmreaperturadefolios.grid_estatusKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure Tfrmreaperturadefolios.grid_estatusKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  If frmbarra1.btnCancel.Enabled = True then
        frmBarra1.btnCancel.Click ;
end;

procedure Tfrmreaperturadefolios.grid_estatusMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure Tfrmreaperturadefolios.grid_estatusMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure Tfrmreaperturadefolios.grid_estatusTitleClick(Column: TColumn);
begin
   UtGrid.DbGridTitleClick(Column);
end;

procedure Tfrmreaperturadefolios.tsIdGrupoKeyPress(Sender: TObject;
  var Key: Char);
begin
    If Key = #13 Then
        tsDescripcion.SetFocus
end;

procedure Tfrmreaperturadefolios.tiColoresEnter(Sender: TObject);
begin
   // tiColores.Color := $00E6FEFF;
end;

procedure Tfrmreaperturadefolios.tiColoresExit(Sender: TObject);
begin
   // tiColores.Color := global_color_salidaERP;
end;

procedure Tfrmreaperturadefolios.tsDescripcionEnter(Sender: TObject);
begin
   // tsDescripcion.Color := global_color_entradaERP
end;

procedure Tfrmreaperturadefolios.tsDescripcionExit(Sender: TObject);
begin
   // tsDescripcion.Color := global_color_salidaERP;
end;

procedure Tfrmreaperturadefolios.tsDescripcionKeyPress(Sender: TObject;
  var Key: Char);
begin
    if frmbarra1.btnPost.Enabled = True then
       if key =#13 then
          frmBarra1.btnPost.SetFocus;
end;

end.
