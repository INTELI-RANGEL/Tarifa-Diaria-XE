unit frm_factordecosto;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, frm_connection, global, frm_barra, Grids, DBGrids, StdCtrls,
  ExtCtrls, DBCtrls, Mask, DB, Menus, 
  ADODB, ZAbstractRODataset, ZAbstractDataset, ZDataset, ComCtrls,
  udbgrid, unitexcepciones, unittbotonespermisos, UnitValidaTexto,UnitValidacion
  ,unitactivapop;

type
  TfrmFactordeCosto = class(TForm)
    grid_factordecosto: TDBGrid;
    Label1: TLabel;
    tdFactor: TDBEdit;
    frmBarra1: TfrmBarra;
    Label3: TLabel;
    ds_factordecosto: TDataSource;
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
    Deshacer1: TMenuItem;
    N4: TMenuItem;
    Salir1: TMenuItem;
    FactordeCosto: TZQuery;
    FactordeCostosContrato: TStringField;
    FactordeCostoiAnno: TIntegerField;
    FactordeCostoiMes: TIntegerField;
    FactordeCostodFactor: TFloatField;
    FactordeCostosMes: TStringField;
    tiMes: TComboBox;
    upanno: TUpDown;
    tiAnno: TMaskEdit;
    Label2: TLabel;
    Label4: TLabel;
    procedure tsIdEmbarcacionKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure grid_factordecostoCellClick(Column: TColumn);
    procedure frmBarra1btnAddClick(Sender: TObject);
    procedure frmBarra1btnEditClick(Sender: TObject);
    procedure frmBarra1btnPostClick(Sender: TObject);
    procedure frmBarra1btnCancelClick(Sender: TObject);
    procedure frmBarra1btnDeleteClick(Sender: TObject);
    procedure frmBarra1btnRefreshClick(Sender: TObject);
    procedure frmBarra1btnExitClick(Sender: TObject);
    procedure tdFactorKeyPress(Sender: TObject; var Key: Char);
    procedure Insertar1Click(Sender: TObject);
    procedure Editar1Click(Sender: TObject);
    procedure Registrar1Click(Sender: TObject);
    procedure Can1Click(Sender: TObject);
    procedure Eliminar1Click(Sender: TObject);
    procedure Refresh1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
    procedure tdFactorEnter(Sender: TObject);
    procedure tdFactorExit(Sender: TObject);
    procedure tiAnnoEnter(Sender: TObject);
    procedure tiAnnoExit(Sender: TObject);
    procedure tiAnnoKeyPress(Sender: TObject; var Key: Char);
    procedure tiMesEnter(Sender: TObject);
    procedure tiMesExit(Sender: TObject);
    procedure tiMesKeyPress(Sender: TObject; var Key: Char);
    procedure FactordeCostoCalcFields(DataSet: TDataSet);
    procedure tiMesChange(Sender: TObject);
    procedure tiAnnoChange(Sender: TObject);
    procedure upannoClick(Sender: TObject; Button: TUDBtnType);
    procedure FactordeCostoAfterScroll(DataSet: TDataSet);
    procedure FactordeCostoBeforeInsert(DataSet: TDataSet);
    procedure grid_factordecostoMouseMove(Sender: TObject; Shift: TShiftState;
      X, Y: Integer);
    procedure grid_factordecostoMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure grid_factordecostoTitleClick(Column: TColumn);
    procedure Copy1Click(Sender: TObject);
    procedure Paste1Click(Sender: TObject);
    procedure tdFactorChange(Sender: TObject);
    procedure FactordeCostodFactorSetText(Sender: TField; const Text: string);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmFactordeCosto: TfrmFactordeCosto;
  x : Byte ;
  utgrid:ticdbgrid;
  botonpermiso:tbotonespermisos;
implementation

{$R *.dfm}

procedure TfrmFactordeCosto.tsIdEmbarcacionKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    tdFactor.SetFocus
end;

procedure TfrmFactordeCosto.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Factordecosto.Cancel ;
  action := cafree ;
  utgrid.Destroy;
  botonpermiso.Free;
end;

procedure TfrmFactordeCosto.FormShow(Sender: TObject);
begin
  BotonPermiso := TBotonesPermisos.Create(Self, connection.zConnection, global_grupo, 'cFactorCosto', PopupPrincipal);
  UtGrid:=TicdbGrid.create(grid_factordecosto);
  OpcButton := '' ;
  frmBarra1.btnCancel.Click ;
  FactordeCosto.Active := False ;
  FactordeCosto.Params.ParamByName('Contrato').DataType := ftString ;
  FactordeCosto.Params.ParamByName('Contrato').Value := global_contrato ;
  FactordeCosto.Open ;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;   
end;

procedure TfrmFactordeCosto.grid_factordecostoCellClick(Column: TColumn);
begin
  if frmBarra1.btnCancel.Enabled = True then
      frmBarra1.btnCancel.Click ;
end;

procedure TfrmFactordeCosto.grid_factordecostoMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
    UtGrid.dbGridMouseMoveCoord(x,y);
end;

procedure TfrmFactordeCosto.grid_factordecostoMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  UtGrid.DbGridMouseUp(Sender,Button,Shift,X, Y);
end;

procedure TfrmFactordeCosto.grid_factordecostoTitleClick(Column: TColumn);
begin
  UtGrid.DbGridTitleClick(Column);
end;

procedure TfrmFactordeCosto.frmBarra1btnAddClick(Sender: TObject);
begin
   activapop(frmFactordeCosto, popupprincipal);
   frmBarra1.btnAddClick(Sender);
   Insertar1.Enabled   := False ;
   Editar1.Enabled     := False ;
   Registrar1.Enabled  := True ;
   Can1.Enabled        := True ;
   Eliminar1.Enabled   := False ;
   Refresh1.Enabled    := False ;
   Salir1.Enabled      := False ;

   tiAnno.SetFocus;
   Factordecosto.Append ;
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmFactordeCosto.frmBarra1btnEditClick(Sender: TObject);
begin
   activapop(frmFactordeCosto, popupprincipal);
   frmBarra1.btnEditClick(Sender);
   Insertar1.Enabled := False ;
   Editar1.Enabled := False ;
   Registrar1.Enabled := True ;
   Can1.Enabled := True ;
   Eliminar1.Enabled := False ;
   Refresh1.Enabled := False ;
   Salir1.Enabled := False ;
   try
       Factordecosto.Edit ;
   except
       on e : exception do begin
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_factordecosto', 'Al editar registro', 0);
        frmBarra1.btnCancel.Click ;
       end;
   end ;
   tiAnno.SetFocus;
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmFactordeCosto.frmBarra1btnPostClick(Sender: TObject);
var
  nombres, cadenas: TStringList;
begin
   {Validacion de campos}
   desactivapop(popupprincipal);
   nombres:=TStringList.Create;cadenas:=TStringList.Create;
   nombres.Add('Año');nombres.Add('Mes');nombres.Add('Factor');
   cadenas.Add(tiAnno.Text);cadenas.Add(tiMes.Text);cadenas.Add(tdFactor.Text);
   if not validaTexto(nombres, cadenas, '', '') then
   begin
     MessageDlg(UnitValidaTexto.errorValidaTexto, mtInformation, [mbOk], 0);
     exit;
   end;
   {Continua insercion de datos}
   try
      FactordeCosto.FieldValues['sContrato'] := global_contrato ;
      FactordeCosto.FieldValues['iAnno']     := tiAnno.Text ;
      FactordeCosto.FieldValues['iMes']      := tiMes.ItemIndex + 1 ;
      Factordecosto.Post ;
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
        UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_factordecosto', 'Al salvar registro', 0);
        frmBarra1.btnCancel.Click ;
       end;
   end;
  BotonPermiso.permisosBotones(frmBarra1);
  frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmFactordeCosto.frmBarra1btnCancelClick(Sender: TObject);
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
   Factordecosto.Cancel ;
   BotonPermiso.permisosBotones(frmBarra1);
   frmBarra1.btnPrinter.Enabled := False;
end;

procedure TfrmFactordeCosto.frmBarra1btnDeleteClick(Sender: TObject);
begin
  If Factordecosto.RecordCount  > 0 then
    if MessageDlg('Desea eliminar el Registro Activo?',
        mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      try
        Factordecosto.Delete ;
      except
       on e : exception do begin
           UnitExcepciones.manejarExcep(E.Message, E.ClassName, 'frm_factordecosto', 'Al eliminar registro', 0);
       end;
      end
    end
end;

procedure TfrmFactordeCosto.frmBarra1btnRefreshClick(Sender: TObject);
begin
  Factordecosto.Active := False ;
  FactordeCosto.Open ;
end;

procedure TfrmFactordeCosto.frmBarra1btnExitClick(Sender: TObject);
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

procedure TfrmFactordeCosto.tdFactorKeyPress(Sender: TObject;
  var Key: Char);
begin
  if not keyFiltroTdbedit(tdFactor,key) then
   key:=#0;
  if key = #13 then
    tiAnno.SetFocus
end;

procedure TfrmFactordeCosto.Insertar1Click(Sender: TObject);
begin
    frmBarra1.btnAdd.Click
end;

procedure TfrmFactordeCosto.Paste1Click(Sender: TObject);
begin
  UtGrid.AddRowsFromClip;
end;

procedure TfrmFactordeCosto.Copy1Click(Sender: TObject);
begin
  UtGrid.CopyRowsToClip;
end;

procedure TfrmFactordeCosto.Editar1Click(Sender: TObject);
begin
    frmBarra1.btnEdit.Click
end;

procedure TfrmFactordeCosto.Registrar1Click(Sender: TObject);
begin
    frmBarra1.btnPost.Click
end;

procedure TfrmFactordeCosto.Can1Click(Sender: TObject);
begin
    frmBarra1.btnCancel.Click
end;

procedure TfrmFactordeCosto.Eliminar1Click(Sender: TObject);
begin
    frmBarra1.btnDelete.Click
end;

procedure TfrmFactordeCosto.Refresh1Click(Sender: TObject);
begin
    frmBarra1.btnRefresh.Click
end;

procedure TfrmFactordeCosto.Salir1Click(Sender: TObject);
begin
    frmBarra1.btnExit.Click
end;

procedure TfrmFactordeCosto.tdFactorChange(Sender: TObject);
begin
  tdbeditchangef(tdFactor,'Factor');  
end;

procedure TfrmFactordeCosto.tdFactorEnter(Sender: TObject);
begin
    tdFactor.Color := global_color_entrada
end;

procedure TfrmFactordeCosto.tdFactorExit(Sender: TObject);
begin
    tdFactor.Color := global_color_salida
end;

procedure TfrmFactordeCosto.tiAnnoEnter(Sender: TObject);
begin
    tiAnno.Color := global_color_Entrada
end;

procedure TfrmFactordeCosto.tiAnnoExit(Sender: TObject);
begin
    tiAnno.Color := global_color_salida
end;

procedure TfrmFactordeCosto.tiAnnoKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tiMes.SetFocus
end;

procedure TfrmFactordeCosto.tiMesEnter(Sender: TObject);
begin
    tiMes.Color := global_color_Entrada
end;

procedure TfrmFactordeCosto.tiMesExit(Sender: TObject);
begin
    tiMes.Color := global_color_salida
end;

procedure TfrmFactordeCosto.tiMesKeyPress(Sender: TObject; var Key: Char);
begin
    If Key = #13 Then
        tdFactor.SetFocus
end;

procedure TfrmFactordeCosto.FactordeCostoCalcFields(DataSet: TDataSet);
begin
    If FactordeCosto.FieldValues['iMes'] > 0 Then
        Case FactordeCosto.FieldValues['iMes'] Of
           1 : FactordeCostosMes.Text  := 'ENERO' ;
           2 : FactordeCostosMes.Text  := 'FEBRERO' ;
           3 : FactordeCostosMes.Text  := 'MARZO' ;
           4 : FactordeCostosMes.Text  := 'ABRIL' ;
           5 : FactordeCostosMes.Text  := 'MAYO' ;
           6 : FactordeCostosMes.Text  := 'JUNIO' ;
           7 : FactordeCostosMes.Text  := 'JULIO' ;
           8 : FactordeCostosMes.Text  := 'AGOSTO' ;
           9 : FactordeCostosMes.Text  := 'SEPTIEMBRE' ;
           10 : FactordeCostosMes.Text := 'OCTUBRE' ;
           11 : FactordeCostosMes.Text := 'NOVIEMBRE' ;
           12 : FactordeCostosMes.Text := 'DICIEMBRE' ;
        End
end;



procedure TfrmFactordeCosto.FactordeCostodFactorSetText(Sender: TField;
  const Text: string);
begin
   sender.Value:=abs(StrToFloatDef(text,0));
end;

procedure TfrmFactordeCosto.tiMesChange(Sender: TObject);
begin
  //If frmBarra1.btnAdd.Enabled = False Then
  //  If (FactordeCosto.State = dsInsert) Or (FactordeCosto.State = dsEdit) Then
  //      FactordeCosto.FieldValues['iMes'] := tiMes.ItemIndex + 1 ;
end;

procedure TfrmFactordeCosto.tiAnnoChange(Sender: TObject);
begin
//  If frmBarra1.btnAdd.Enabled = False Then
//   If (FactordeCosto.State = dsInsert) Or (FactordeCosto.State = dsEdit) Then
//       FactordeCosto.FieldValues['iAnno'] := tiAnno.Text
end;

procedure TfrmFactordeCosto.upannoClick(Sender: TObject;
  Button: TUDBtnType);
begin
  //  If frmBarra1.btnAdd.Enabled = False Then
  //    FactordeCosto.FieldValues['iAnno'] := tiAnno.Text
end;

procedure TfrmFactordeCosto.FactordeCostoAfterScroll(DataSet: TDataSet);
begin
   If FactordeCosto.RecordCount > 0 Then
       If (FactordeCosto.State <> dsInsert) and (FactordeCosto.State <> dsEdit) Then
           tiAnno.Text := FactordeCosto.FieldValues['iAnno']
end;

procedure TfrmFactordeCosto.FactordeCostoBeforeInsert(DataSet: TDataSet);
begin
   x := 90 ;
end;

end.
